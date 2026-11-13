/**
 * Minimal DOCX writer for tests: builds a valid ZIP (stored entries) containing
 * just word/document.xml. Keeps fixtures readable and dependency-free.
 */
import { deflateRawSync } from 'node:zlib';

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (const byte of buf) c = CRC_TABLE[(c ^ byte) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function zipStored(files) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const nameBuf = Buffer.from(name, 'utf8');
    const data = Buffer.from(content, 'utf8');
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt32LE(0, 10);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    locals.push(local, nameBuf, data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt32LE(0, 12);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);
    offset += local.length + nameBuf.length + data.length;
  }
  const centralBuf = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(Object.keys(files).length, 8);
  end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(centralBuf.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);
  return Buffer.concat([...locals, centralBuf, end]);
}

const esc = (s) => String(s ?? '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;');

function para(text, { superscript = false } = {}) {
  const rPr = superscript ? '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>' : '';
  return `<w:p><w:r>${rPr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
}

/**
 * cell: string | { text, gridSpan, vMerge: 'restart'|'continue', nested: rows[][] }
 */
function cellXml(cell) {
  const spec = typeof cell === 'string' ? { text: cell } : cell || {};
  const pr = [];
  if (spec.gridSpan > 1) pr.push(`<w:gridSpan w:val="${spec.gridSpan}"/>`);
  if (spec.vMerge) pr.push(`<w:vMerge${spec.vMerge === 'restart' ? ' w:val="restart"' : ''}/>`);
  const tcPr = pr.length ? `<w:tcPr>${pr.join('')}</w:tcPr>` : '';
  const paragraphs = spec.paragraphs || [spec.text ?? ''];
  const inner = paragraphs.map((p) => para(p)).join('');
  const nested = (spec.nested || []).map((rows) => tableXml(rows)).join('');
  return `<w:tc>${tcPr}${inner}${nested}</w:tc>`;
}

/** rows: array of arrays of cells; row may be { cells, repeatHeader } */
function tableXml(rows, { cols } = {}) {
  const colCount = cols || Math.max(...rows.map((r) => {
    const cs = Array.isArray(r) ? r : r.cells;
    return cs.reduce((n, c) => n + (typeof c === 'object' && c.gridSpan ? c.gridSpan : 1), 0);
  }));
  const grid = `<w:tblGrid>${'<w:gridCol w:w="1000"/>'.repeat(colCount)}</w:tblGrid>`;
  const trs = rows.map((r) => {
    const spec = Array.isArray(r) ? { cells: r } : r;
    const trPr = spec.repeatHeader ? '<w:trPr><w:tblHeader/></w:trPr>' : '';
    return `<w:tr>${trPr}${spec.cells.map(cellXml).join('')}</w:tr>`;
  }).join('');
  return `<w:tbl>${grid}${trs}</w:tbl>`;
}

/** blocks: [{ p: 'text' } | { table: rows }] */
export function buildDocx(blocks) {
  const body = blocks.map((b) => (b.p !== undefined ? para(b.p) : tableXml(b.table, b))).join('');
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${body}<w:sectPr/></w:body>
</w:document>`;
  return zipStored({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
    'word/document.xml': documentXml,
  });
}

export { tableXml };
