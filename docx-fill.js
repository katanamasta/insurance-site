// Заполнение оригинальных бланков Word: распаковка .docx, вставка значений в ячейки
// рядом с подписями граф, сборка архива обратно. Всё выполняется в браузере.
(function () {
  const crcTable = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(buf) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  async function deflateRaw(bytes) {
    const s = new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(s).arrayBuffer());
  }

  async function inflateRaw(bytes) {
    const s = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(s).arrayBuffer());
  }

  // Читает архив, сохраняя сжатые данные всех записей — нетронутые файлы переносим как есть.
  async function readZip(arrayBuffer) {
    const buf = new Uint8Array(arrayBuffer);
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let eo = -1;
    for (let i = buf.length - 22; i >= 0; i--) { if (dv.getUint32(i, true) === 0x06054b50) { eo = i; break; } }
    if (eo < 0) throw new Error('файл повреждён или это не .docx');
    const count = dv.getUint16(eo + 10, true);
    let p = dv.getUint32(eo + 16, true);
    const entries = [];
    for (let k = 0; k < count; k++) {
      const flags = dv.getUint16(p + 8, true);
      const method = dv.getUint16(p + 10, true);
      const crc = dv.getUint32(p + 16, true);
      const csize = dv.getUint32(p + 20, true);
      const usize = dv.getUint32(p + 24, true);
      const nameLen = dv.getUint16(p + 28, true);
      const extraLen = dv.getUint16(p + 30, true);
      const commLen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(buf.slice(p + 46, p + 46 + nameLen));
      const lnl = dv.getUint16(lho + 26, true);
      const lel = dv.getUint16(lho + 28, true);
      const start = lho + 30 + lnl + lel;
      entries.push({ name, method, crc, csize, usize, flags, data: buf.slice(start, start + csize) });
      p += 46 + nameLen + extraLen + commLen;
    }
    return entries;
  }

  function writeZip(entries) {
    const enc = new TextEncoder();
    const parts = [];
    const central = [];
    let offset = 0;
    entries.forEach(e => {
      const nameBytes = enc.encode(e.name);
      const lh = new Uint8Array(30 + nameBytes.length);
      const dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true);
      dv.setUint16(6, 0, true);
      dv.setUint16(8, e.method, true);
      dv.setUint32(14, e.crc, true);
      dv.setUint32(18, e.csize, true);
      dv.setUint32(22, e.usize, true);
      dv.setUint16(26, nameBytes.length, true);
      lh.set(nameBytes, 30);
      parts.push(lh, e.data);
      const cd = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(cd.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(10, e.method, true);
      cv.setUint32(16, e.crc, true);
      cv.setUint32(20, e.csize, true);
      cv.setUint32(24, e.usize, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint32(42, offset, true);
      cd.set(nameBytes, 46);
      central.push(cd);
      offset += lh.length + e.data.length;
    });
    const cdSize = central.reduce((a, c) => a + c.length, 0);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, entries.length, true);
    ev.setUint16(10, entries.length, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, offset, true);
    return new Blob([...parts, ...central, eocd], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }

  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const plain = xml => xml
    .replace(/<w:instrText[\s\S]*?<\/w:instrText>/g, '')
    .replace(/<[^>]+>/g, '')
    .replace(/\u00a0/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();

  // Вставляет значение в ячейку: сначала в пустое текстовое поле, иначе новым прогоном.
  function writeIntoCell(cellXml, value) {
    const run = '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + esc(value) + '</w:t></w:r>';
    const empty = cellXml.match(/<w:t(?:\s[^>]*)?\/>|<w:t(?:\s[^>]*)?>\s*<\/w:t>/);
    if (empty) {
      return cellXml.replace(empty[0], '<w:t xml:space="preserve">' + esc(value) + '</w:t>');
    }
    const lastP = cellXml.lastIndexOf('</w:p>');
    if (lastP > -1) return cellXml.slice(0, lastP) + run + cellXml.slice(lastP);
    return cellXml;
  }

  // rules: [[regexp подписи графы, значение], ...]
  function fillDocumentXml(xml, rules) {
    let filled = 0;
    const out = xml.replace(/<w:tr[\s>][\s\S]*?<\/w:tr>/g, (tr) => {
      const cells = [];
      tr.replace(/<w:tc[\s>][\s\S]*?<\/w:tc>/g, (tc, idx) => { cells.push({ xml: tc, at: idx }); return tc; });
      if (!cells.length) return tr;
      let result = tr;
      for (const [re, value] of rules) {
        if (!value) continue;
        const i = cells.findIndex(c => re.test(plain(c.xml)));
        if (i < 0) continue;
        const target = cells[i + 1] && !plain(cells[i + 1].xml) ? cells[i + 1] : cells[i];
        const before = target.xml;
        const after = writeIntoCell(before, value);
        if (after !== before && result.indexOf(before) > -1) {
          result = result.replace(before, after);
          filled++;
        }
        break;
      }
      return result;
    });
    return { xml: out, filled };
  }

  async function fillDocx(url, rules, opts) {
    const res = await fetch(url);
    if (!res.ok) throw new Error('шаблон не найден: ' + url);
    const entries = await readZip(await res.arrayBuffer());
    const doc = entries.find(e => e.name === 'word/document.xml');
    if (!doc) throw new Error('в шаблоне нет word/document.xml');
    const raw = doc.method === 0 ? doc.data : await inflateRaw(doc.data);
    const xml = new TextDecoder().decode(raw);
    const { xml: filledXml, filled } = fillDocumentXml(xml, rules);
    const bytes = new TextEncoder().encode(filledXml);
    const packed = await deflateRaw(bytes);
    doc.method = 8;
    doc.data = packed;
    doc.csize = packed.length;
    doc.usize = bytes.length;
    doc.crc = crc32(bytes);
    return { blob: writeZip(entries), filled, name: (opts && opts.name) || 'zayavlenie.docx' };
  }

  window.PinsDocx = { fillDocx };
})();
