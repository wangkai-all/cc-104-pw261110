/**
 * Structural analysis of parsed DOCX tables:
 *   - header detection
 *   - data-segment detection (subtotal / note / section / repeated-header boundaries)
 *   - downward group (地区) inheritance inside one segment only
 *   - pending-confirmation items for anything ambiguous
 */
import { FOOTNOTE_RE, SUBTOTAL_RE } from './docx.js';

const NOTE_ROW_RE = /^(注|备注|说明|附注|资料来源|数据来源|来源|填写说明|填表说明)\s*[:：]/;
const SUBTOTAL_TOKEN_RE = /(^|[\s:：])(合计|小计|总计)$/;

export function isNoteText(text) {
  const t = String(text || '').trim();
  return NOTE_ROW_RE.test(t) || FOOTNOTE_RE.test(t);
}

export function isSubtotalText(text) {
  const t = String(text || '').trim().replace(/\s+/g, '');
  if (!t) return false;
  return SUBTOTAL_RE.test(t) || SUBTOTAL_TOKEN_RE.test(String(text || '').trim()) ||
    /^(合计|小计|总计)/.test(t) && t.length <= 12;
}

export function buildGrid(table) {
  const rows = Math.max(table.rowCount || table.rows.length, 1);
  const cols = Math.max(table.colCount || 1, 1);
  const grid = Array.from({ length: rows }, () => Array(cols).fill(null));
  for (const row of table.rows) {
    for (const cell of row.cells) {
      const rs = Math.max(1, Math.min(cell.rowSpan || 1, rows - row.index));
      const cs = Math.max(1, Math.min(cell.colSpan || 1, cols - cell.col));
      for (let r = row.index; r < row.index + rs; r++) {
        for (let c = cell.col; c < cell.col + cs; c++) {
          if (!grid[r][c]) grid[r][c] = { cell, row: row.index, col: cell.col, rowSpan: rs, colSpan: cs };
        }
      }
    }
  }
  return grid;
}

export function cellTextAt(grid, row, col) {
  const hit = grid?.[row]?.[col];
  return hit ? hit.cell.text : '';
}

export function rowTexts(table, rowIndex) {
  const row = table.rows.find((r) => r.index === rowIndex);
  if (!row) return [];
  return row.cells.map((c) => ({ col: c.col, text: c.text, colSpan: c.colSpan, cell: c }));
}

export function headerSignature(table, rowIndex) {
  return rowTexts(table, rowIndex)
    .filter((c) => c.text)
    .map((c) => c.text.replace(/\s+/g, ' ').trim());
}

function sameSignature(a, b) {
  if (!a.length || !b.length) return false;
  if (a.length !== b.length) return false;
  let same = 0;
  for (let i = 0; i < a.length; i++) if (a[i] === b[i]) same++;
  return same === a.length || same / a.length >= 0.8;
}

function rowNonEmpty(table, rowIndex) {
  return rowTexts(table, rowIndex).filter((c) => c.text.trim());
}

function isMergedTitleRow(table, rowIndex) {
  const cells = rowTexts(table, rowIndex).filter((c) => c.text.trim());
  if (!cells.length) return false;
  if (cells.length === 1) {
    const c = cells[0];
    return c.colSpan >= Math.max(2, Math.ceil(table.colCount * 0.6));
  }
  return false;
}

function isSectionRow(table, rowIndex) {
  const cells = rowTexts(table, rowIndex).filter((c) => c.text.trim());
  if (cells.length !== 1) return false;
  const c = cells[0];
  return c.colSpan >= Math.max(2, Math.ceil(table.colCount * 0.6)) &&
    !isNoteText(c.text) && !isSubtotalText(c.text);
}

/** Candidate header rows: label-like rows near the top, not data rows. */
export function headerCandidates(table) {
  const out = [];
  const grid = buildGrid(table);
  // only the first rows of a table can be its header; a repeated header further
  // down is a boundary, not a competing candidate
  const max = Math.min(table.rowCount, 8);
  for (let r = 0; r < max; r++) {
    const cells = rowNonEmpty(table, r);
    if (cells.length < 2) continue;
    if (isMergedTitleRow(table, r)) continue;
    const continuations = cells.filter((c) => c.cell.continuation).length;
    const carriesMerge = rowCarriesMergeFromAbove(table, r);
    const subHeader = continuations >= Math.ceil(cells.length * 0.6) || carriesMerge;
    let shortLabels = 0;
    let numeric = 0;
    for (const c of cells) {
      const t = c.text.trim();
      if (/^(填报单位|单位|制表|审核|填表人|负责人|联系电话)/.test(t)) continue;
      if (/^[-+]?\d+(?:\.\d+)?%?$/.test(t)) numeric++;
      else if (t.length <= 30) shortLabels++;
    }
    const distinct = new Set(cells.map((c) => c.text.trim()));
    if (distinct.size < 2) continue;
    const ratio = cells.length / Math.max(1, table.colCount);
    if (ratio < 0.35) continue;
    if (numeric > 0 && numeric >= shortLabels) continue;
    if (shortLabels < 2) continue;
    const spansDown = cells.some((c) => {
      const hit = grid[r]?.[c.col];
      return hit && hit.rowSpan > 1;
    });
    let score = shortLabels - numeric * 3 - (table.colCount - cells.length) * 0.5;
    if (r <= 2) score += 3;
    if (spansDown) score += 2;
    if (subHeader) score -= 5;
    // a row that repeats the top header verbatim is a repeat-header boundary
    const firstCandidate = out.find((c) => !c.subHeader);
    const repeatedFromAbove = firstCandidate && sameSignature(firstCandidate.texts, headerSignature(table, r));
    if (repeatedFromAbove) score -= 8;
    out.push({ row: r, texts: headerSignature(table, r), spansDown, subHeader, repeatedFromAbove, score });
  }
  return out;
}

function rowCarriesMergeFromAbove(table, rowIndex) {
  const prev = table.rows.find((r) => r.index === rowIndex - 1);
  if (!prev) return false;
  return prev.cells.some((c) => c.row >= rowIndex - 1 && c.row + (c.rowSpan || 1) > rowIndex);
}

export function detectHeader(table) {
  const candidates = headerCandidates(table);
  const repeated = table.rows.filter((r) => r.repeatHeader).map((r) => r.index);
  let chosen = null;
  let ambiguous = false;
    if (!repeated.length && candidates.length) {
    // header rows sit above the data: prefer the earliest candidate that scores
    // within 75% of the best candidate
    const bestScore = Math.max(...candidates.map((c) => c.score));
    const viable = candidates.filter((c) => c.score >= bestScore * 0.75);
    chosen = Math.min(...viable.map((c) => c.row));
  }
  if (chosen === null) {
    for (let r = 0; r < table.rowCount; r++) {
      if (rowNonEmpty(table, r).length >= 2) { chosen = r; break; }
    }
  }
  if (chosen === null) chosen = 0;
  const chosenScore = candidates.find((c) => c.row === chosen)?.score ?? 0;
  const chosenTexts = candidates.find((c) => c.row === chosen)?.texts || [];
  const conflicting = candidates.find(
    (c) => c.row > chosen && c.row <= chosen + 2 && !c.subHeader &&
      c.score >= chosenScore * 0.9 && !sameSignature(chosenTexts, c.texts)
  );
  // An exact re-statement of the chosen header further down is a repeated header,
  // not an ambiguity.
  if (conflicting && !candidates.some((c) => c.row > chosen && sameSignature(chosenTexts, c.texts) && c.score >= chosenScore)) {
    ambiguous = true;
  }
  return { headerRow: chosen, ambiguous, candidates, repeatedHeaderRows: repeated };
}

/** How many rows the header occupies (1 unless merged header cells span downward). */
export function headerRowCountAt(table, headerRow, candidates = []) {
  const grid = buildGrid(table);
  const row = table.rows.find((r) => r.index === headerRow);
  if (!row) return 1;
  let count = 1;
  for (let r = headerRow; r < table.rowCount - 1; r++) {
    const nextRow = table.rows.find((x) => x.index === r + 1);
    if (!nextRow) break;
    // the header continues if any cell starting in this row spans into the next row
    // AND the next row contains header-like labels (not data values)
    const spansIntoNext = table.rows
      .find((x) => x.index === r)
      ?.cells.some((c) => c.col >= 0 && c.rowSpan >= 2 && c.row === r);
    if (!spansIntoNext) break;
    // require the next row to add sub-headings rather than data
    const subCells = rowNonEmpty(table, r + 1);
    const numeric = subCells.filter((c) => /^[-+]?\d+(?:\.\d+)?%?$/.test(c.text.trim()));
    if (numeric.length > subCells.length / 2) break;
    if (!subCells.length) break;
    count++;
    if (count >= 4) break;
  }
  return count;
}

function inferFieldType(table, col, startRow, endRow) {
  let numeric = 0;
  let text = 0;
  let total = 0;
  const grid = buildGrid(table);
  for (let r = startRow; r <= endRow; r++) {
    if (r >= table.rowCount) break;
    const value = cellTextAt(grid, r, col).trim();
    if (!value) continue;
    total++;
    if (/^[-+]?\d+(?:\.\d+)?%?$/.test(value)) numeric++;
    else text++;
  }
  if (!total) return 'text';
  return numeric / total >= 0.6 ? 'number' : 'text';
}

const GROUP_HEADER = /(地区|区域|乡镇|街道|乡镇街道|辖区|行政区|县|市|省|单位|部门|机构)/;

/**
 * Pick the column whose names are inherited downward. A column only qualifies if
 * it shows vertical merges / blank cells (the thing inheritance exists for) or
 * its header clearly names a region. Otherwise a plain label column such as
 * 检测项 in a key-value table would be mislabelled as the region column.
 */
export function detectGroupColumn(table, headerRow, startRow, endRow, headerRowCount = 1) {
  const grid = buildGrid(table);
  const colCount = table.colCount;
  const headerGridRow = Math.min(headerRow + Math.max(1, headerRowCount) - 1, table.rowCount - 1);
  const candidates = [];
  for (let c = 0; c < colCount; c++) {
    let textCount = 0;
    let numCount = 0;
    let blankOrMerged = 0;
    let merges = 0;
    for (let r = startRow; r <= Math.min(endRow, table.rowCount - 1); r++) {
      const hit = grid[r]?.[c];
      if (!hit) continue;
      const val = hit.cell.text.trim();
      if (hit.cell.rowSpan > 1) merges++;
      if (hit.cell.continuation) blankOrMerged++;
      else if (!val) blankOrMerged++;
      else if (/^[-+]?\d+(?:\.\d+)?%?$/.test(val)) numCount++;
      else textCount++;
    }
    const headerText = (grid[headerRow]?.[c]?.cell.text || '') + ' ' + (grid[headerGridRow]?.[c]?.cell.text || '');
    const headerHint = GROUP_HEADER.test(headerText);
    const score = textCount * 3 + blankOrMerged * 2 + merges * 4 - numCount * 2;
    candidates.push({ col: c, textCount, numCount, blankOrMerged, merges, headerHint, score });
  }
  // Prefer a column with real merge/blank evidence, then one whose header names
  // a region; leftmost wins on ties.
  const eligible = candidates.filter((c) => c.textCount > 0 && c.numCount <= c.textCount);
  const best = eligible
    .filter((c) => c.merges > 0 || c.blankOrMerged > 0 || c.headerHint)
    .sort((a, b) => (b.merges - a.merges) || (Number(b.headerHint) - Number(a.headerHint)) || (a.col - b.col))[0];
  return best ? best.col : null;
}

function makeField(key, label, col, role = 'value', type = 'text') {
  return { key, label, col, role, type };
}

export function suggestFields(table, headerRow, headerRowCount, startRow, endRow) {
  const grid = buildGrid(table);
  const headerGridRow = Math.min(headerRow + headerRowCount - 1, table.rowCount - 1);
  const fields = [];
  for (let c = 0; c < table.colCount; c++) {
    const hit = grid[headerGridRow]?.[c];
    if (!hit || hit.col !== c) continue; // only at a cell's own starting column
    const own = (hit.cell.text || '').replace(/\s+/g, ' ').trim();
    // For a two-row header, the top row often carries the group label; join them.
    let label = own;
    if (!own && headerRowCount > 1) {
      const parents = [];
      for (let r = headerRow; r < headerGridRow; r++) {
        const parent = grid[r]?.[c];
        if (parent && parent.cell.text.trim()) parents.push(parent.cell.text.trim());
      }
      label = parents.join(' / ');
    }
    if (!label) continue;
    fields.push(makeField(`c${c}`, label, c, 'value', inferFieldType(table, c, startRow, endRow)));
  }
  if (!fields.length) {
    for (let c = 0; c < Math.min(table.colCount, 8); c++) {
      const label = (cellTextAt(grid, headerGridRow, c) || '').trim() || `第${c + 1}列`;
      fields.push(makeField(`c${c}`, label, c, 'value', 'text'));
    }
  }
  return fields;
}

export function trimDataEnd(table, headerRow, headerRowCount) {
  let end = table.rowCount - 1;
  while (end > headerRow + headerRowCount - 1) {
    const cells = rowNonEmpty(table, end);
    const mergedTitle = isMergedTitleRow(table, end);
    const note = cells.length && isNoteText(cells[0].text);
    if (!cells.length || mergedTitle || note) end--;
    else break;
  }
  return Math.max(end, headerRow + headerRowCount - 1);
}

/** Suggest one data region for a table. */
export function suggestRegion(table) {
  const detected = detectHeader(table);
  const headerRow = detected.headerRow ?? 0;
  const headerRowCount = headerRowCountAt(table, headerRow, detected.candidates);
  const startRow = Math.min(headerRow + headerRowCount, Math.max(0, table.rowCount - 1));
  const endRow = trimDataEnd(table, headerRow, headerRowCount);
  const groupCol = detectGroupColumn(table, headerRow, startRow, endRow, headerRowCount);
  const fields = suggestFields(table, headerRow, headerRowCount, startRow, endRow);
  if (groupCol !== null && fields.some((f) => f.col === groupCol)) {
    const f = fields.find((f) => f.col === groupCol);
    f.role = 'group';
    f.type = 'text';
  } else if (groupCol !== null) {
    fields.unshift(makeField(`c${groupCol}`, `第${groupCol + 1}列`, groupCol, 'group', 'text'));
  }
  return {
    tableId: table.id,
    headerRow,
    headerRowCount,
    startRow,
    endRow,
    groupCol,
    inheritGroup: true,
    skipSubtotalRows: true,
    fields,
    confirmed: { header: !detected.ambiguous, range: false, groupCol: false },
    auto: { headerAmbiguous: !!detected.ambiguous, candidates: detected.candidates },
  };
}

function repeatHeaderOf(table, rowIndex) {
  const row = table.rows.find((r) => r.index === rowIndex);
  return !!row?.repeatHeader;
}

/** A single merged cell spanning the whole table, sitting above the header. */
export function isTitleRow(table, rowIndex) {
  const row = table.rows.find((r) => r.index === rowIndex);
  if (!row || !row.cells.length) return false;
  const nonEmpty = row.cells.filter((c) => c.text.trim());
  if (nonEmpty.length !== 1) return false;
  const cell = nonEmpty[0];
  return (cell.colSpan || 1) >= Math.max(2, table.colCount - 1);
}

/**
 * Classify each row in [startRow, endRow] and the boundaries between data segments.
 * Types: header | repeat-header | word-repeat-header | subtotal | note | blank | section | data | nested
 */
export function classifyRows(table, region) {
  const grid = buildGrid(table);
  const headerRow = Math.max(0, Math.min(region.headerRow ?? 0, table.rowCount - 1));
  const headerRowCount = Math.max(1, region.headerRowCount || 1);
  const start = Math.max(0, Math.min(region.startRow ?? headerRow + headerRowCount, table.rowCount));
  const end = Math.max(start - 1, Math.min(region.endRow ?? table.rowCount - 1, table.rowCount - 1));
  const headerTexts = headerSignature(table, headerRow);
  const types = [];
  for (let r = 0; r < table.rowCount; r++) {
    const rowCells = rowTexts(table, r);
    const cells = rowCells.filter((c) => c.text.trim());
    const nestedIds = rowCells.flatMap((c) => c.cell.nestedTables || []);
    const hasNested = nestedIds.length > 0;
    let type = 'data';
    if (r >= headerRow && r < headerRow + headerRowCount) type = 'header';
    else if (r < headerRow && isTitleRow(table, r)) type = 'title';
    else if (r < start || r > end) type = 'outside';
    else if (repeatHeaderOf(table, r)) type = 'word-repeat-header';
    else if (!cells.length) type = hasNested ? 'data' : 'blank';
    else if (cells.length === 1 && isNoteText(cells[0].text)) type = hasNested ? 'data' : 'note';
    else if (cells.some((c) => isSubtotalText(c.text))) type = 'subtotal';
    else if (isSectionRow(table, r)) type = 'section';
    else if (sameSignature(headerSignature(table, r), headerTexts)) type = 'repeat-header';
    // A row whose cell contains a nested table is still a data row; the nested
    // content is flagged as a pending item instead of being mixed into the value.
    types.push({ row: r, type, cells, nested: nestedIds });
  }
  return { types, start, end, headerRow, headerRowCount, headerTexts };
}

const BOUNDARY_TYPES = new Set(['header', 'repeat-header', 'subtotal', 'note', 'blank', 'section']);
const TRANSPARENT_TYPES = new Set(['word-repeat-header']);

export function buildSegments(classification) {
  const segments = [];
  let current = null;
  for (const item of classification.types) {
    if (item.type === 'outside' || TRANSPARENT_TYPES.has(item.type)) continue;
    if (item.type === 'data') {
      if (!current) {
        current = { rows: [], boundaryBefore: null };
        segments.push(current);
      }
      current.rows.push(item.row);
    } else {
      current = null;
      if (BOUNDARY_TYPES.has(item.type)) {
        segments.push({ rows: [], boundary: item });
      }
    }
  }
  // keep only non-empty segments for the caller, but report boundary rows separately
  return segments.filter((s) => s.rows.length);
}

export function parentCellOf(doc, table) {
  if (!table.parentTableId) return null;
  return table.parentCell || null;
}

export function tableLabel(doc, table) {
  const title = table.title || table.caption || '';
  const suffix = table.parentTableId ? `（嵌套于 ${table.parentTableId}）` : '';
  return `${table.id}${title ? ` · ${title}` : ''}${suffix}`;
}

/**
 * Full analysis of a mapping region: rows, group inheritance, exclusions, pending items.
 */
export function analyzeRegion(doc, table, region, regionLabel, context = {}) {
  const pending = [];
  const warnings = [];
  const grid = buildGrid(table);
  const classification = classifyRows(table, region);
  const boundaryRows = [];
  for (const item of classification.types) {
    if (item.type === 'subtotal') boundaryRows.push({ row: item.row, type: 'subtotal', text: item.cells.map((c) => c.text).join(' ') });
    if (item.type === 'note') boundaryRows.push({ row: item.row, type: 'note', text: item.cells.map((c) => c.text).join(' ') });
    if (item.type === 'repeat-header') boundaryRows.push({ row: item.row, type: 'repeat-header', text: item.cells.map((c) => c.text).join(' ') });
    if (item.type === 'section') boundaryRows.push({ row: item.row, type: 'section', text: item.cells.map((c) => c.text).join(' ') });
  }

  const pendingBase = { regionId: region.id, tableId: table.id, regionLabel };

  if (!classification.headerTexts.length) {
    pending.push({
      ...pendingBase,
      id: `${region.id}:no-header`,
      type: 'no-header',
      severity: 'blocking',
      title: '未识别到表头',
      message: `表格 ${table.id} 的第 ${(region.headerRow ?? 0) + 1} 行没有可用表头文字，请指定表头行。`,
      action: 'set-header',
    });
  }
  if ((region.startRow ?? 0) < classification.headerRow + classification.headerRowCount) {
    pending.push({
      ...pendingBase,
      id: `${region.id}:range-overlaps-header`,
      type: 'range-overlaps-header',
      severity: 'blocking',
      title: '数据区与表头重叠',
      message: `数据区从第 ${(region.startRow ?? 0) + 1} 行开始，但表头占到第 ${classification.headerRow + classification.headerRowCount} 行。`,
      action: 'set-range',
    });
  }
  if (classification.end < classification.start) {
    pending.push({
      ...pendingBase,
      id: `${region.id}:empty-range`,
      type: 'empty-range',
      severity: 'blocking',
      title: '数据区为空',
      message: '所选数据区没有可导入的行。',
      action: 'set-range',
    });
  }
  if (region.auto?.headerAmbiguous && !region.confirmed?.header) {
    const cand = (region.auto.candidates || []).map((c) => `第 ${c.row + 1} 行`).join('、');
    pending.push({
      ...pendingBase,
      id: `${region.id}:ambiguous-header`,
      type: 'ambiguous-header',
      severity: 'blocking',
      title: '表头不唯一',
      message: `检测到多个可能的表头行（${cand}），请确认使用哪一行。`,
      action: 'confirm-header',
      candidates: region.auto.candidates || [],
    });
  }

  // nested tables inside the selected data range
  const dataRowsAll = classification.types.filter((t) => t.type === 'data');
  const nestedInside = [];
  for (const item of dataRowsAll) {
    for (const nestedId of item.nested || []) {
      const nestedTable = doc.tables.find((t) => t.id === nestedId);
      if (!nestedTable) continue;
      const resolved = context.nestedResolutions?.[nestedId];
      nestedInside.push({ nestedId, row: item.row, col: null, table: nestedTable, resolved });
      if (!resolved) {
        pending.push({
          ...pendingBase,
          id: `${region.id}:nested:${nestedId}`,
          type: 'nested-table',
          severity: 'blocking',
          title: '数据区内有嵌套表格',
          message: `第 ${item.row + 1} 行单元格内有嵌套表格 ${nestedId}。当前区域只读取该单元格自身的文字，请选择跳过该嵌套表，或把它作为单独数据区导入。`,
          action: 'resolve-nested',
          nestedTableId: nestedId,
          nestedTableTitle: nestedTable.title || nestedTable.caption || '',
          row: item.row,
        });
      }
    }
  }
  if (table.nestedTableIds?.length && !dataRowsAll.some((r) => (r.nested || []).length)) {
    warnings.push({
      ...pendingBase,
      type: 'nested-outside-range',
      severity: 'info',
      title: '表格内还有其他嵌套表',
      message: `该表含嵌套表 ${table.nestedTableIds.join('、')}，不在当前数据区范围内，将不会被导入。`,
    });
  }

  // group inheritance
  const groupCol = region.groupCol;
  const segments = buildSegments(classification);
  const records = [];
  const unresolvedGroups = [];
  const rowSources = new Map();
  for (const segment of segments) {
    let currentGroup = null;
    let currentSource = null;
    for (const row of segment.rows) {
      const hit = groupCol === null || groupCol === undefined ? null : grid[row]?.[groupCol];
      const cell = hit?.cell;
      // Covered by a vertical merge that started above: the visible text belongs
      // to the merge origin, so this row inherits it rather than owning it.
      const coveredFromAbove = !!(hit && hit.row < row);
      const explicit = !!(cell && cell.text.trim() && (hit.row === row || coveredFromAbove));
      let groupText = null;
      let source = null;
      if (explicit) {
        groupText = cell.text.trim();
        source = coveredFromAbove
          ? { kind: 'inherited', row: hit.row, col: hit.col, via: 'merge' }
          : { kind: 'explicit', row: row, col: groupCol };
      } else if (region.groupOverrides && Object.prototype.hasOwnProperty.call(region.groupOverrides, String(row))) {
        groupText = String(region.groupOverrides[String(row)]).trim() || null;
        if (groupText) source = { kind: 'manual', row, col: groupCol, note: '用户指定' };
      } else if (region.inheritGroup !== false && currentGroup !== null) {
        groupText = currentGroup;
        source = { kind: 'inherited', row: currentSource.row, col: currentSource.col, via: 'carry-down' };
      }
      if (groupText !== null) {
        currentGroup = groupText;
        // Keep the original owner of the value for later carry-down reporting.
        currentSource = coveredFromAbove
          ? { kind: 'explicit', row: hit.row, col: groupCol }
          : source;
        rowSources.set(row, { groupText, source });
      } else if (groupCol !== null && groupCol !== undefined) {
        rowSources.set(row, { groupText: null, source: null });
        unresolvedGroups.push(row);
      } else {
        rowSources.set(row, { groupText: null, source: null });
      }
      records.push({ row, groupText: source ? { value: groupText, ...source } : null });
    }
  }
  if (groupCol === null || groupCol === undefined) {
    pending.push({
      ...pendingBase,
      id: `${region.id}:no-group-column`,
      type: 'no-group-column',
      severity: 'warning',
      title: '未设置地区列',
      message: '未指定用于向下继承的地区列。若每行自带地区信息，可忽略；否则请选择地区列。',
      action: 'set-group-column',
    });
  } else if (unresolvedGroups.length) {
    for (const row of unresolvedGroups) {
      pending.push({
        ...pendingBase,
        id: `${region.id}:group-gap:${row}`,
        type: 'group-gap',
        severity: 'blocking',
        title: '地区名称缺失，无法向下继承',
        message: `第 ${row + 1} 行的地区单元格为空，且上方没有同一数据区内可继承的值（小计行、重复表头、空行或表格边界会中断继承）。请填写该行地区、调整数据区起点，或改选地区列。`,
        action: 'fill-group',
        row,
      });
    }
  }

  for (const b of boundaryRows) {
    const label = { subtotal: '小计/合计行', note: '附注行', 'repeat-header': '重复表头行', section: '分节标题行' }[b.type];
    warnings.push({
      ...pendingBase,
      type: `excluded-${b.type}`,
      severity: 'info',
      title: `已排除${label}`,
      message: `第 ${b.row + 1} 行「${b.text.slice(0, 40)}」不会被当作数据导入，也不会参与地区继承。`,
      row: b.row,
    });
  }

  // irregular row structure
  for (const row of table.rows) {
    if (row.index < classification.start || row.index > classification.end) continue;
    if (row.gridBefore > 0) {
      warnings.push({
        ...pendingBase,
        type: 'irregular-row',
        severity: 'warning',
        title: '行结构不规整',
        message: `第 ${row.index + 1} 行前有 ${row.gridBefore} 个空网格，列位置按文档网格解析，请核对预览。`,
        row: row.index,
      });
    }
  }
  if (classification.types.some((t) => t.type === 'data') === false) {
    pending.push({
      ...pendingBase,
      id: `${region.id}:no-data`,
      type: 'no-data',
      severity: 'blocking',
      title: '没有数据行',
      message: '所选范围内没有可导入的数据行。',
      action: 'set-range',
    });
  }

  return {
    regionId: region.id,
    tableId: table.id,
    tableOrdinal: region.tableOrdinal ?? null,
    classification: {
      headerRow: classification.headerRow,
      headerRowCount: classification.headerRowCount,
      start: classification.start,
      end: classification.end,
      headerTexts: classification.headerTexts,
      types: classification.types.map((t) => ({ row: t.row, type: t.type, nested: t.nested })),
    },
    segments: segments.filter((s) => s.rows.length),
    segmentsMap: segments.filter((s) => s.rows.length),
    rowSources,
    unresolvedGroups,
    pending,
    warnings,
    boundaryRows,
    nestedInside: nestedInside.map((n) => ({ nestedId: n.nestedId, row: n.row, col: n.col, resolved: !!n.resolved })),
    records,
  };
}

export function tableStructure(table, headerRow, headerRowCount) {
  const grid = buildGrid(table);
  const lastHeader = Math.min((headerRow ?? 0) + (headerRowCount || 1) - 1, table.rowCount - 1);
  const headers = [];
  for (let c = 0; c < table.colCount; c++) {
    const hit = grid[lastHeader]?.[c];
    headers.push({
      col: c,
      text: hit ? hit.cell.text : '',
      colSpan: hit ? hit.colSpan : 1,
      isOrigin: hit ? hit.cell.col === c : false,
    });
  }
  return {
    colCount: table.colCount,
    headers,
    headerRow,
    headerRowCount,
    nestedTableIds: table.nestedTableIds || [],
    hasMerges: !!table.merged,
  };
}

export function normalizeHeader(text) {
  return String(text || '').replace(/\s+/g, ' ').trim();
}
