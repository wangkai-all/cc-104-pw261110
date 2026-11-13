/**
 * DOCX -> structured document model.
 *
 * Only the parts needed for survey-table import are extracted:
 *   - body blocks in reading order (paragraphs and tables)
 *   - table grids with horizontal/vertical merges resolved
 *   - nested tables kept separate from their host cell's text
 *   - real Word footnotes/endnotes (marked, never merged into cell values)
 *   - "table notes" (rows/paragraphs starting with 注/说明/…) flagged separately
 */
import { createHash } from 'node:crypto';
import { readZip } from './zip.js';
import { parseXml, firstChild, childrenOf, findFirst, textContent, attr } from './xml.js';

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const FOOTNOTE_RE = /^(注|备注|说明|附注|资料来源|数据来源|来源|填写说明|填表说明)\s*[:：]/;
const SUBTOTAL_RE = /^(合计|小计|总计|总\s*计|合\s*计|小\s*计)\s*[:：]?/;
const MARKER_RE = /^(?:[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]+|[*＊※†‡]+|\[\d+\]|\(\d+\)|注\s*\d+)$/;

export function normalizeText(text) {
  return String(text ?? '')
    .replace(/ /g, ' ')
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/^\s+|\s+$/g, '');
}

function blockText(el, opts = {}) {
  // Text of a w:p or w:tc, excluding nested tables and deleted runs.
  let out = '';
  const vertAlign = (node) => {
    // true when the run is superscript/subscript (used for footnote markers)
    return !!findFirst(node, 'vertAlign', W_NS);
  };
  const walk = (node) => {
    if (!node || node.type === 'text') return;
    const ns = node.ns;
    if (ns !== W_NS) {
      for (const c of node.children) if (c.type !== 'text') walk(c);
      return;
    }
    switch (node.local) {
      case 'tbl':
        return;
      case 'del':
        return;
      case 'delText':
        return;
      case 'instrText':
        return;
      case 't':
        out += textContent(node);
        return;
      case 'tab':
        out += '\t';
        return;
      case 'br':
      case 'cr':
        out += '\n';
        return;
      case 'noBreakHyphen':
        out += '-';
        return;
      case 'p':
        for (const c of node.children) if (c.type !== 'text') walk(c);
        out += '\n';
        return;
      default:
        for (const c of node.children) if (c.type !== 'text') walk(c);
    }
  };
  for (const c of el.children) if (c.type !== 'text') walk(c);
  return normalizeText(out);
}

/** Superscript-only text in a paragraph, e.g. footnote markers like ① or * */
function superscriptMarkers(p) {
  const markers = [];
  const walk = (node, inSup) => {
    if (!node || node.type === 'text') return;
    if (node.ns === W_NS && node.local === 'r') {
      const rPr = firstChild(node, 'rPr', W_NS);
      const va = rPr ? firstChild(rPr, 'vertAlign', W_NS) : null;
      const isSup = va ? attr(va, 'val') === 'superscript' : inSup;
      const text = childrenOf(node, 't', W_NS).map((t) => textContent(t)).join('');
      if (isSup && text) markers.push(normalizeText(text));
      return;
    }
    if (node.ns === W_NS && (node.local === 'del' || node.local === 'tbl')) return;
    for (const c of node.children) if (c.type !== 'text') walk(c, inSup);
  };
  walk(p, false);
  return markers.filter((m) => MARKER_RE.test(m));
}

function paragraphInfo(p) {
  const text = blockText(p);
  const style = attr(firstChild(p, 'pPr', W_NS) ? firstChild(firstChild(p, 'pPr', W_NS), 'pStyle', W_NS) : null, 'val') || '';
  const footnoteRefs = [];
  const walk = (node) => {
    if (!node || node.type === 'text') return;
    if (node.ns === W_NS) {
      if (node.local === 'footnoteReference') {
        const id = attr(node, 'id');
        if (id) footnoteRefs.push(id);
      }
      if (node.local === 'endnoteReference') {
        const id = attr(node, 'id');
        if (id) footnoteRefs.push(`e${id}`);
      }
      if (node.local === 'del' || node.local === 'tbl') return;
    }
    for (const c of node.children) if (c.type !== 'text') walk(c);
  };
  walk(p);
  return { text, style, footnoteRefs, markers: superscriptMarkers(p) };
}

function cellProps(tc) {
  const tcPr = firstChild(tc, 'tcPr', W_NS);
  let gridSpan = 1;
  let vMerge = null;
  if (tcPr) {
    const span = firstChild(tcPr, 'gridSpan', W_NS);
    if (span) {
      const n = parseInt(attr(span, 'val') || '1', 10);
      if (Number.isFinite(n) && n > 0) gridSpan = n;
    }
    const vm = firstChild(tcPr, 'vMerge', W_NS);
    if (vm) vMerge = attr(vm, 'val') === 'restart' ? 'restart' : 'continue';
    if (firstChild(tcPr, 'hMerge', W_NS)) {
      // Legacy horizontal merge: treat like continuation of previous span.
      const hm = firstChild(tcPr, 'hMerge', W_NS);
      if (attr(hm, 'val') !== 'restart') vMerge = vMerge || null;
    }
  }
  return { gridSpan, vMerge };
}

function stripMarkers(text, markers) {
  let out = text;
  for (const m of markers) {
    if (!m) continue;
    // only strip when it sits at the very end of a line, so real data like "A①区" survives
    const re = new RegExp(`(^|\\n)\\s*${m.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*$`);
    out = out.replace(re, '$1');
  }
  return normalizeText(out);
}

function parseTable(tbl, ctx, parent = null) {
  const id = ctx.nextTableId(parent);
  const table = {
    id,
    kind: 'table',
    parentTableId: parent ? parent.id : null,
    parentCell: parent ? ctx.lastCellRef : null,
    caption: '',
    title: '',
    blockIndex: null,
    rows: [],
    nestedTableIds: [],
    footnotes: [],
    notes: [],
  };

  const gridCols = childrenOf(firstChild(tbl, 'tblGrid', W_NS), 'gridCol', W_NS);
  const declaredCols = gridCols.length;

  const rows = [];
  const vMergeTrack = new Map(); // col -> origin cell

  childrenOf(tbl, 'tr', W_NS).forEach((tr, rowIndex) => {
    const trPr = firstChild(tr, 'trPr', W_NS);
    const repeatHeader = !!firstChild(trPr, 'tblHeader', W_NS);
    let gridBefore = 0;
    const gb = firstChild(trPr, 'gridBefore', W_NS);
    if (gb) gridBefore = Math.max(0, parseInt(attr(gb, 'val') || '0', 10) || 0);

    const cells = [];
    let col = gridBefore;
    for (const tc of childrenOf(tr, 'tc', W_NS)) {
      const { gridSpan, vMerge } = cellProps(tc);
      const paragraphs = childrenOf(tc, 'p', W_NS).map(paragraphInfo);
      const markers = paragraphs.flatMap((p) => p.markers);
      const rawText = normalizeText(paragraphs.map((p) => p.text).join('\n'));
      const text = stripMarkers(rawText, markers);

      ctx.lastCellRef = { tableId: id, row: rowIndex, col, colSpan: gridSpan };
      const nested = childrenOf(tc, 'tbl', W_NS).map((t) => parseTable(t, ctx, table));
      const footnoteRefs = [...new Set(paragraphs.flatMap((p) => p.footnoteRefs))];

      const cell = {
        row: rowIndex,
        col,
        colSpan: gridSpan,
        rowSpan: 1,
        text,
        rawText,
        markers,
        paragraphs: paragraphs.map((p) => ({ text: p.text, style: p.style })),
        footnoteRefs,
        vMerge,
        continuation: false,
        origin: null,
        nestedTables: nested.map((t) => t.id),
        hasNestedTable: nested.length > 0,
      };

      if (vMerge === 'continue' && vMergeTrack.has(col)) {
        const originCell = vMergeTrack.get(col);
        cell.continuation = true;
        cell.origin = { row: originCell.row, col: originCell.col };
        cell.text = originCell.text;
        cell.paragraphs = originCell.paragraphs;
        originCell.rowSpan += 1;
      } else {
        if (vMerge === 'restart') vMergeTrack.set(col, cell);
        else vMergeTrack.delete(col);
        cells.push(cell);
      }

      for (let i = 1; i < gridSpan; i++) vMergeTrack.delete(col + i);
      col += gridSpan;
    }
    rows.push({ index: rowIndex, repeatHeader, gridBefore, cells, colCount: col });
  });

  const colCount = Math.max(declaredCols, ...rows.map((r) => r.colCount), 1);
  table.rows = rows;
  table.rowCount = rows.length;
  table.colCount = colCount;
  table.declaredCols = declaredCols || null;
  table.nestedTableIds = rows.flatMap((r) => r.cells.flatMap((c) => c.nestedTables));
  table.merged = rows.some((r) => r.cells.some((c) => c.colSpan > 1 || c.rowSpan > 1));
  ctx.tables.push(table);
  return table;
}

function collectFootnotes(zip) {
  const out = new Map();
  for (const [part, prefix] of [['word/footnotes.xml', 'f'], ['word/endnotes.xml', 'e']]) {
    if (!zip.has(part)) continue;
    let xml;
    try {
      xml = parseXml(zip.read(part).toString('utf8'));
    } catch {
      continue;
    }
    for (const note of childrenOf(findFirst(xml, 'footnotes') || findFirst(xml, 'endnotes'), 'footnote', W_NS)
      .concat(childrenOf(findFirst(xml, 'footnotes') || findFirst(xml, 'endnotes'), 'endnote', W_NS))) {
      const id = attr(note, 'id');
      if (id === undefined || id === '-1' || id === '0') continue;
      const type = attr(note, 'type');
      if (type === 'separator' || type === 'continuationSeparator') continue;
      const text = childrenOf(note, 'p', W_NS).map(paragraphInfo).map((p) => p.text).filter(Boolean).join('\n');
      if (text) out.set(`${prefix}${id}`, { id: `${prefix}${id}`, text });
    }
  }
  return out;
}

function tableNotes(table, followingParagraphs) {
  const notes = [];
  for (const row of table.rows) {
    const nonEmpty = row.cells.filter((c) => c.text);
    if (!nonEmpty.length) continue;
    const joined = nonEmpty.map((c) => c.text).join(' ');
    const first = normalizeText(nonEmpty[0].text);
    const allInFirst = nonEmpty.length === 1;
    if (FOOTNOTE_RE.test(first) || (allInFirst && /^[*＊※]/.test(first)) || /^[①-⑳]\s*/.test(first) && allInFirst) {
      notes.push({ row: row.index, text: joined, kind: 'table-note' });
    }
  }
  for (const p of followingParagraphs) {
    if (FOOTNOTE_RE.test(p.text) || /^[*＊※]/.test(p.text)) {
      notes.push({ text: p.text, kind: 'table-note', afterTable: true });
    }
  }
  return notes;
}

export function parseDocx(buffer) {
  const sha256 = createHash('sha256').update(buffer).digest('hex');
  let zip;
  try {
    zip = readZip(buffer);
  } catch (err) {
    throw new Error(`无法读取 DOCX：${err.message}`);
  }
  if (!zip.has('word/document.xml')) throw new Error('文件缺少 word/document.xml，不是有效的 DOCX');
  let documentXml;
  try {
    documentXml = parseXml(zip.read('word/document.xml').toString('utf8'));
  } catch (err) {
    throw new Error(`DOCX 内容解析失败：${err.message}`);
  }

  const footnotesById = collectFootnotes(zip);
  const body = findFirst(documentXml, 'body', W_NS);
  if (!body) throw new Error('DOCX 缺少文档正文');

  let tableCounter = 0;
  const nestedCounters = new Map();
  const ctx = {
    tables: [],
    lastCellRef: null,
    nextTableId(parent) {
      if (!parent) {
        tableCounter += 1;
        return `t${tableCounter}`;
      }
      const key = parent.id;
      const n = (nestedCounters.get(key) || 0) + 1;
      nestedCounters.set(key, n);
      return `${key}.n${n}`;
    },
  };

  const blocks = [];
  const pendingParagraphs = [];
  let blockIndex = 0;
  let title = '';

  const walkBlocks = (container) => {
    for (const child of container.children) {
      if (child.type === 'text') continue;
      if (child.ns !== W_NS) {
        walkBlocks(child);
        continue;
      }
      if (child.local === 'p') {
        const info = paragraphInfo(child);
        pendingParagraphs.push(info);
        blocks.push({
          kind: 'paragraph',
          index: blockIndex++,
          text: info.text,
          style: info.style,
          footnoteRefs: info.footnoteRefs,
        });
        if (!title && info.text) title = info.text;
      } else if (child.local === 'tbl') {
        pendingParagraphs.splice(0);
        // The immediately preceding paragraph is a caption only when it reads
        // like one; form meta lines (填报单位/填表人…) must not become titles.
        const prev = [...blocks].reverse().find((b) => b.kind === 'paragraph' && b.text);
        const prevText = prev ? prev.text : '';
        const prevStyle = prev?.style || '';
        const isCaption = !!prevText && prevText.length <= 60 &&
          !/^(Heading|Title|Subtitle|标题|标题 \d)/i.test(prevStyle) &&
          !/^(填报单位|填报人|填表人|制表人|审核人|负责人|联系电话|单位负责人|日期|时间)\s*[:：]/.test(prevText) &&
          (/^(表|附表|续表|附件|清单)/.test(prevText) || /表\s*[0-9一二三四五六七八九十]*$/.test(prevText.trim()) || /表\s*[0-9一二三四五六七八九十]*[：:（(]/.test(prevText));
        ctx.lastCellRef = null;
        const table = parseTable(child, ctx);
        table.blockIndex = blockIndex++;
        table.caption = isCaption ? prevText : '';
        // Many surveys put the table title in a merged first row. Use it when
        // there is no caption paragraph, so the UI can show a meaningful name.
        if (!table.caption) {
          const first = table.rows[0];
          const nonEmpty = (first?.cells || []).filter((c) => c.text.trim());
          if (nonEmpty.length === 1 && (nonEmpty[0].colSpan || 1) >= Math.max(2, table.colCount - 1)) {
            table.titleRow = 0;
            table.caption = nonEmpty[0].text.trim().slice(0, 80);
          }
        }
        table.title = table.caption;
        blocks.push({ kind: 'table', index: table.blockIndex, tableId: table.id });
        // attach real DOCX footnotes to cells that reference them, plus note markers
        for (const row of table.rows) {
          for (const cell of row.cells) {
            for (const ref of cell.footnoteRefs) {
              const note = footnotesById.get(ref);
              if (note) table.footnotes.push({ ...note, row: row.index, col: cell.col, cell: cell.text });
            }
          }
        }
      } else if (child.local === 'sdt') {
        const content = firstChild(child, 'sdtContent', W_NS);
        if (content) walkBlocks(content);
      } else if (child.local === 'customXml' || child.local === 'smartTag') {
        walkBlocks(child);
      }
    }
  };
  walkBlocks(body);

  // Notes: table-internal note rows plus paragraphs between this table and the next.
  const topTables = ctx.tables.filter((t) => !t.parentTableId);
  for (const table of topTables) {
    const nextTableIndex = blocks.find((b) => b.kind === 'table' && b.index > table.blockIndex)?.index ?? Infinity;
    const paras = blocks.filter((b) => b.kind === 'paragraph' && b.index > table.blockIndex && b.index < nextTableIndex);
    table.notes = tableNotes(table, paras.filter((p) => p.text));
    table.noteParagraphs = paras.filter((p) => p.text).map((p) => p.text);
  }
  for (const table of ctx.tables.filter((t) => t.parentTableId)) {
    table.notes = tableNotes(table, []);
    table.noteParagraphs = [];
  }

  // Assemble nested table references metadata.
  for (const table of ctx.tables) {
    for (const row of table.rows) {
      for (const cell of row.cells) {
        if (!cell.nestedTables.length) continue;
        cell.nestedTableMeta = cell.nestedTables.map((id) => ({
          id,
          row: cell.row,
          col: cell.col,
        }));
      }
    }
  }

  const plainText = blocks.filter((b) => b.kind === 'paragraph').map((b) => b.text).filter(Boolean);
  return {
    sha256,
    title: title || '未命名文档',
    blocks,
    paragraphs: plainText,
    tables: ctx.tables,
    footnotes: [...footnotesById.values()],
    stats: {
      tableCount: topTables.length,
      nestedTableCount: ctx.tables.length - topTables.length,
      paragraphCount: plainText.length,
    },
  };
}

export { FOOTNOTE_RE, SUBTOTAL_RE };
