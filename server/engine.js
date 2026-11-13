/**
 * Mapping spec evaluation: structure diffs, region extraction, value provenance.
 */
import { buildGrid, classifyRows, buildSegments, analyzeRegion, tableStructure, normalizeHeader, isNoteText, isSubtotalText, suggestRegion } from './structure.js';

export const BLANK_KIND = 'blank';
export const NONE_KIND = 'none';
export const ZERO_KIND = 'zero';
export const VALUE_KIND = 'value';

export function cleanCellText(text) {
  return String(text ?? '').replace(/[　 ]/g, ' ').replace(/\s+/g, ' ').trim();
}

const NUMERIC_RE = /^[-+]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][-+]?\d+)?%?$/;

export function interpretValue(rawText, field) {
  const raw = String(rawText ?? '');
  const text = cleanCellText(raw);
  if (!text) return { kind: BLANK_KIND, value: null, raw, display: '' };
  if (text === '无') return { kind: NONE_KIND, value: '无', raw, display: '无' };
  if (field?.type === 'number') {
    if (NUMERIC_RE.test(text)) {
      const isPercent = text.endsWith('%');
      const n = Number(isPercent ? text.slice(0, -1) : text);
      if (Number.isFinite(n)) {
        if (n === 0) return { kind: ZERO_KIND, value: 0, raw, display: text, number: 0 };
        return { kind: VALUE_KIND, value: isPercent ? text : n, raw, display: text, number: n };
      }
    }
    return { kind: VALUE_KIND, value: text, raw, display: text, invalidNumber: true };
  }
  if (NUMERIC_RE.test(text)) {
    const n = Number(text.replace('%', ''));
    if (Number.isFinite(n) && n === 0) return { kind: ZERO_KIND, value: text, raw, display: text, number: 0 };
  }
  return { kind: VALUE_KIND, value: text, raw, display: text };
}

function tableSignatureOf(table, headerRow, headerRowCount) {
  const st = tableStructure(table, headerRow, headerRowCount);
  return {
    colCount: st.colCount,
    headerRowCount,
    cells: st.headers
      .filter((h) => h.isOrigin)
      .map((h) => ({ col: h.col, text: normalizeHeader(h.text), colSpan: h.colSpan })),
  };
}

export function specRegionSignature(region) {
  return {
    colCount: region.colCount,
    headerRowCount: region.headerRowCount,
    cells: (region.headerSignature || []).map((h) => ({
      col: h.col,
      text: normalizeHeader(h.text),
      colSpan: h.colSpan || 1,
    })),
  };
}

/**
 * Find the table a saved region belongs to.
 *
 * The ordinal is only a hint: if the new document reordered its tables, blindly
 * using the ordinal would line the mapping up with the wrong table and import
 * wrong columns. We therefore score every table by header-text overlap, column
 * count, header depth and id hint, and report whether the table moved.
 */
function findTableForRegion(doc, region) {
  const all = doc.tables;
  const ordinal = region.tableOrdinal ?? region.tableIndex;
  const expected = specRegionSignature(region);
  const expectedTexts = new Set(expected.cells.map((c) => normalizeHeader(c.text)).filter(Boolean));
  const candidates = [];
  all.forEach((table, idx) => {
    if (!table) return;
    const headerRow = Math.max(0, Math.min(region.headerRow ?? 0, table.rowCount - 1));
    const counts = [...new Set([region.headerRowCount || 1, 1])];
    let bestSig = null;
    let bestOverlap = -1;
    let bestScore = -Infinity;
    for (const count of counts) {
      const actual = tableSignatureOf(table, headerRow, Math.min(count, table.rowCount - headerRow));
      const actualTexts = new Set(actual.cells.map((c) => normalizeHeader(c.text)).filter(Boolean));
      let overlap = 0;
      for (const t of expectedTexts) if (actualTexts.has(t)) overlap += 1;
      let score = overlap * 3;
      if (actual.colCount === expected.colCount) score += 4;
      if ((actual.headerRowCount || 1) === (expected.headerRowCount || 1)) score += 2;
      if (idx === ordinal) score += 2;
      if (region.tableIdHint && table.id === region.tableIdHint) score += 3;
      if (score > bestScore) { bestScore = score; bestSig = actual; bestOverlap = overlap; }
    }
    candidates.push({ table, ordinal: idx, actual: bestSig, overlap: bestOverlap, score: bestScore });
  });
  candidates.sort((a, b) => b.score - a.score || a.ordinal - b.ordinal);
  const best = candidates[0];
  const exactOrdinal = Number.isInteger(ordinal) && ordinal >= 0 && ordinal < all.length;
  if (!best) return { table: null, ordinal: null, exactOrdinal: false, moved: false, sig: null, candidates };
  // Require either a real header overlap or an ordinal hit with the same column count.
  const plausible = best.overlap >= 1 || (best.ordinal === ordinal && exactOrdinal && best.actual.colCount === expected.colCount);
  if (!plausible) {
    return { table: null, ordinal: null, exactOrdinal, moved: false, sig: expected, candidates };
  }
  const moved = best.ordinal !== ordinal;
  return {
    table: best.table,
    ordinal: best.ordinal,
    exactOrdinal: !moved,
    moved,
    overlap: best.overlap,
    sig: best.actual,
    candidates,
  };
}

function headerTextsAt(table, row, colCount) {
  if (!table) return [];
  const grid = buildGrid(table);
  const out = [];
  for (let c = 0; c < colCount; c++) {
    const hit = grid[row]?.[c];
    out.push(hit ? normalizeHeader(hit.cell.text) : '');
  }
  return out;
}

/**
 * Diff two header signatures as a column *edit script* rather than by raw index.
 * A single inserted column should read as one "column added" instead of shifting
 * every later header and reporting a rename for each of them.
 */
function diffHeaderCells(expected, actual) {
  const diffs = [];
  const exp = expected.map((c, i) => ({ ...c, pos: i }));
  const act = actual.map((c, i) => ({ ...c, pos: i }));
  const key = (c) => normalizeHeader(c.text);
  const n = exp.length;
  const m = act.length;
  // longest common subsequence on header text
  const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = key(exp[i]) === key(act[j]) && key(exp[i])
        ? dp[i + 1][j + 1] + 1
        : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const pairs = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (key(exp[i]) === key(act[j]) && key(exp[i])) {
      pairs.push([i, j]);
      i += 1; j += 1;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      i += 1;
    } else {
      j += 1;
    }
  }
  pairs.push([n, m]); // sentinel
  let ei = 0;
  let ai = 0;
  for (const [pe, pa] of pairs) {
    const removed = exp.slice(ei, pe);
    const added = act.slice(ai, pa);
    // Pair by textual similarity first so a rename reads as one change, then
    // pair whatever is left by position.
    const usedRemoved = new Set();
    const usedAdded = new Set();
    for (let r = 0; r < removed.length; r++) {
      let best = -1;
      let bestSim = 0;
      for (let a = 0; a < added.length; a++) {
        if (usedAdded.has(a)) continue;
        const sim = headerSimilarity(removed[r].text, added[a].text);
        if (sim > bestSim) { bestSim = sim; best = a; }
      }
      if (best >= 0 && bestSim >= 0.5) {
        usedRemoved.add(r);
        usedAdded.add(best);
        diffs.push({
          code: 'header-label-changed',
          severity: 'blocking',
          col: removed[r].col,
          message: `表头第 ${removed[r].col + 1} 列由「${removed[r].text || '(空)'}」变为「${added[best].text || '(空)'}」。`,
          expected: removed[r],
          actual: added[best],
        });
      }
    }
    const restRemoved = removed.filter((_c, idx) => !usedRemoved.has(idx));
    const restAdded = added.filter((_c, idx) => !usedAdded.has(idx));
    const common = Math.min(restRemoved.length, restAdded.length);
    for (let k = 0; k < common; k++) {
      diffs.push({
        code: 'header-label-changed',
        severity: 'blocking',
        col: restRemoved[k].col,
        message: `表头第 ${restRemoved[k].col + 1} 列由「${restRemoved[k].text || '(空)'}」变为「${restAdded[k].text || '(空)'}」。`,
        expected: restRemoved[k],
        actual: restAdded[k],
      });
    }
    for (let k = common; k < restAdded.length; k++) {
      diffs.push({
        code: 'column-added',
        severity: 'blocking',
        col: restAdded[k].col,
        message: `新增了第 ${restAdded[k].col + 1} 列「${restAdded[k].text || '(空)'}」。`,
        expected: null,
        actual: restAdded[k],
      });
    }
    for (let k = common; k < restRemoved.length; k++) {
      diffs.push({
        code: 'column-removed',
        severity: 'blocking',
        col: restRemoved[k].col,
        message: `缺少了第 ${restRemoved[k].col + 1} 列「${restRemoved[k].text || '(空)'}」，方案中的对应字段无法落位。`,
        expected: restRemoved[k],
        actual: null,
      });
    }
    if (pe < n && pa < m) {
      const e = exp[pe];
      const a = act[pa];
      if ((e.colSpan || 1) !== (a.colSpan || 1)) {
        diffs.push({
          code: 'header-span-changed',
          severity: 'blocking',
          col: e.col,
          message: `表头第 ${e.col + 1} 列的合并跨度由 ${e.colSpan || 1} 变为 ${a.colSpan || 1}，列对应关系不再可靠。`,
          expected: e,
          actual: a,
        });
      }
    }
    ei = pe + 1;
    ai = pa + 1;
  }
  return diffs;
}

/**
 * Cheap similarity for Chinese header labels: shared characters (bigrams when
 * long enough) plus a containment bonus, so 水质 → 水质评价 counts as a rename.
 */
function headerSimilarity(a, b) {
  const x = normalizeHeader(a);
  const y = normalizeHeader(b);
  if (!x || !y) return x === y ? 1 : 0;
  if (x === y) return 1;
  if (x.length >= 2 && (y.includes(x) || x.includes(y))) return 0.9;
  const grams = (s) => {
    const set = new Set();
    if (s.length === 1) set.add(s);
    for (let i = 0; i < s.length - 1; i++) set.add(s.slice(i, i + 2));
    return set;
  };
  const ga = grams(x);
  const gb = grams(y);
  let shared = 0;
  for (const g of ga) if (gb.has(g)) shared += 1;
  return (2 * shared) / (ga.size + gb.size);
}

/**
 * Compare a saved mapping against a new document.
 * Blocking diffs mean the spec cannot be applied without a human remap.
 */
export function diffSpecAgainstDoc(doc, spec) {
  const diffs = [];
  const regions = Array.isArray(spec.regions) ? spec.regions : [];
  if (!regions.length) {
    diffs.push({ code: 'no-regions', severity: 'blocking', message: '映射方案中没有数据区。' });
  }
  const seenOrdinals = new Set();
  for (const region of regions) {
    const label = region.name || region.key || '数据区';
    const found = findTableForRegion(doc, region);
    if (!found.table) {
      diffs.push({
        code: 'table-missing', severity: 'blocking', regionKey: region.key,
        message: `数据区「${label}」对应第 ${(region.tableOrdinal ?? 0) + 1} 个表格，但当前文件只有 ${doc.tables.length} 个表格。`,
        expected: { tableOrdinal: region.tableOrdinal, headerSignature: specRegionSignature(region).cells },
        actual: { tableCount: doc.tables.length },
      });
      continue;
    }
    seenOrdinals.add(found.ordinal);
    const table = found.table;
    const expected = specRegionSignature(region);
    const actual = found.sig;
    if (found.moved) {
      diffs.push({
        code: 'table-moved',
        severity: 'warning',
        regionKey: region.key,
        message: `数据区「${label}」原来指向第 ${(region.tableOrdinal ?? 0) + 1} 个表格，当前文件中表头相同的是第 ${found.ordinal + 1} 个表格（${table.id}）。已按表头文字重新匹配，请核对。`,
        expected: { tableOrdinal: region.tableOrdinal }, actual: { tableOrdinal: found.ordinal, tableId: table.id },
      });
    } else if (found.overlap !== undefined && found.overlap < expected.cells.length) {
      diffs.push({
        code: 'header-partial-match',
        severity: 'warning',
        regionKey: region.key,
        message: `数据区「${label}」在第 ${found.ordinal + 1} 个表格（${table.id}）中只匹配到 ${found.overlap}/${expected.cells.length} 个表头，请核对差异。`,
        expected: { cells: expected.cells.length }, actual: { overlap: found.overlap, tableId: table.id },
      });
    }

    if (region.headerRow >= table.rowCount) {
      diffs.push({
        code: 'header-row-out-of-range', severity: 'blocking', regionKey: region.key,
        message: `数据区「${label}」的表头行（第 ${region.headerRow + 1} 行）超出表格 ${table.id} 的行数（${table.rowCount} 行）。`,
        expected: { headerRow: region.headerRow }, actual: { rowCount: table.rowCount },
      });
      continue;
    }
    if (actual.colCount !== expected.colCount) {
      diffs.push({
        code: 'column-count-changed', severity: 'blocking', regionKey: region.key,
        message: `数据区「${label}」的列数由 ${expected.colCount} 变为 ${actual.colCount}，按列位置套用会导致错列。`,
        expected: { colCount: expected.colCount }, actual: { colCount: actual.colCount },
        field: { tableId: table.id },
      });
    }
    const expCells = expected.cells;
    const actCells = actual.cells;
    for (const d of diffHeaderCells(expCells, actCells)) {
      diffs.push({ ...d, message: `数据区「${label}」${d.message}`, regionKey: region.key, tableId: table.id });
    }
    if ((expected.headerRowCount || 1) !== (actual.headerRowCount || 1)) {
      diffs.push({
        code: 'header-row-count-changed', severity: 'blocking', regionKey: region.key,
        message: `数据区「${label}」的表头行数由 ${expected.headerRowCount || 1} 变为 ${actual.headerRowCount || 1}。`,
        expected: { headerRowCount: expected.headerRowCount || 1 },
        actual: { headerRowCount: actual.headerRowCount || 1 },
        tableId: table.id,
      });
    }

    // row range sanity (warnings: extra or missing data rows are normal)
    if (region.endRow !== undefined && region.endRow !== null && region.endRow >= table.rowCount) {
      diffs.push({
        code: 'range-shrunk', severity: 'warning', regionKey: region.key,
        message: `数据区「${label}」原来截止到第 ${region.endRow + 1} 行，当前表格只有 ${table.rowCount} 行；将按当前表格内容重新确定结束位置。`,
        expected: { endRow: region.endRow }, actual: { rowCount: table.rowCount }, tableId: table.id,
      });
    }
    if (region.startRow >= table.rowCount) {
      diffs.push({
        code: 'range-missing', severity: 'blocking', regionKey: region.key,
        message: `数据区「${label}」的起始行（第 ${region.startRow + 1} 行）超出当前表格 ${table.id} 的行数（${table.rowCount} 行）。`,
        expected: { startRow: region.startRow }, actual: { rowCount: table.rowCount }, tableId: table.id,
      });
    }
    const nestedExpected = region.nestedTableOrdinals || [];
    const nestedActual = (table.nestedTableIds || []).length;
    if (nestedExpected.length !== nestedActual) {
      diffs.push({
        code: 'nested-tables-changed', severity: 'warning', regionKey: region.key,
        message: `数据区「${label}」的嵌套表数量由 ${nestedExpected.length} 变为 ${nestedActual}。嵌套表内容不会混合进单元格数值。`,
        expected: { count: nestedExpected.length }, actual: { count: nestedActual }, tableId: table.id,
      });
    }
    const fieldCols = (region.fields || []).map((f) => f.col);
    for (const f of region.fields || []) {
      if (f.col >= table.colCount) {
        diffs.push({
          code: 'field-column-out-of-range', severity: 'blocking', regionKey: region.key, col: f.col,
          message: `字段「${f.label}」指向第 ${f.col + 1} 列，但表格 ${table.id} 只有 ${table.colCount} 列。`,
          expected: { field: f.label, col: f.col }, actual: { colCount: table.colCount }, tableId: table.id,
        });
      }
      if (f.role === 'group' && !headerTextsAt(table, region.headerRow, table.colCount)[f.col]) {
        diffs.push({
          code: 'group-column-empty-header', severity: 'warning', regionKey: region.key, col: f.col,
          message: `地区列「${f.label}」对应的第 ${f.col + 1} 列表头为空，请确认。`,
          expected: { field: f.label }, actual: null, tableId: table.id,
        });
      }
    }
    if (region.groupCol !== null && region.groupCol !== undefined && region.groupCol >= table.colCount) {
      diffs.push({
        code: 'group-column-out-of-range', severity: 'blocking', regionKey: region.key,
        message: `地区列（第 ${region.groupCol + 1} 列）超出表格 ${table.id} 的列数。`,
        expected: { groupCol: region.groupCol }, actual: { colCount: table.colCount }, tableId: table.id,
      });
    }
    if (fieldCols.length !== new Set(fieldCols).size) {
      const dupes = fieldCols.filter((c, i) => fieldCols.indexOf(c) !== i);
      diffs.push({
        code: 'duplicate-field-columns', severity: 'warning', regionKey: region.key,
        message: `数据区「${label}」有多个字段映射到同一列（第 ${[...new Set(dupes)].map((c) => c + 1).join('、')} 列）。`,
        expected: null, actual: { cols: dupes }, tableId: table.id,
      });
    }
  }

  // extra tables not covered by the mapping (nested tables belong to their parent)
  const extra = doc.tables
    .map((t, i) => ({ t, i }))
    .filter(({ i }) => !seenOrdinals.has(i))
    .filter(({ t }) => !t.parentTableId)
    .filter(({ t }) => t.rowCount > 0);
  if (extra.length && regions.length) {
    diffs.push({
      code: 'extra-tables', severity: 'warning',
      message: `当前文件有 ${extra.length} 个表格未被映射方案覆盖：${extra.map(({ t, i }) => `${i + 1}.${t.id}`).join('、')}。这些表格不会导入。`,
      expected: null, actual: { tables: extra.map(({ t, i }) => ({ ordinal: i, id: t.id, title: t.title })) },
    });
  }
  const blocking = diffs.filter((d) => d.severity === 'blocking');
  return { ok: blocking.length === 0, diffs, blocking, warnings: diffs.filter((d) => d.severity === 'warning') };
}

/**
 * Extract records for one region. Never throws on data problems; reports them.
 */
export function extractRegion(doc, region, context = {}) {
  const table = context.table || doc.tables[region.tableOrdinal];
  if (!table) {
    return {
      regionKey: region.key,
      regionName: region.name,
      error: `找不到数据区「${region.name || region.key}」对应的表格。`,
      rows: [], values: [], issues: [], pending: [], warnings: [], blocking: [],
      stats: zeroStats(),
      nFields: 0,
      rowRange: { startRow: null, endRow: null },
    };
  }
  const normalized = {
    ...region,
    headerRow: region.headerRow,
    headerRowCount: region.headerRowCount || 1,
    startRow: region.startRow,
    endRow: region.endRow === null || region.endRow === undefined ? table.rowCount - 1 : region.endRow,
    groupCol: region.groupCol,
    inheritGroup: region.inheritGroup !== false,
    groupOverrides: region.groupOverrides || {},
    nestedResolutions: region.nestedResolutions || {},
  };
  if (normalized.startRow === null || normalized.startRow === undefined) {
    normalized.startRow = normalized.headerRow + normalized.headerRowCount;
  }
  if (normalized.endRow === null || normalized.endRow === undefined) normalized.endRow = table.rowCount - 1;

  const analysis = analyzeRegion(doc, table, normalized, region.name || region.key, {
    nestedResolutions: normalized.nestedResolutions,
  });
  const grid = buildGrid(table);
  const classification = analysis.classification;
  const fields = (region.fields || []).filter((f) => f.include !== false && f.role !== 'ignore');
  const issues = [];
  const pending = [...analysis.pending];
  const warnings = [...analysis.warnings];

  const seenKeys = new Set();
  for (const f of fields) {
    if (seenKeys.has(f.key)) {
      issues.push({ severity: 'blocking', code: 'duplicate-field-key', fieldKey: f.key, message: `字段 key「${f.key}」重复。` });
    }
    seenKeys.add(f.key);
  }
  const groupFields = fields.filter((f) => f.role === 'group');
  if (groupFields.length > 1) {
    issues.push({ severity: 'blocking', code: 'multiple-group-fields', message: '一个数据区只能有一个地区继承字段。' });
  }
  if (!fields.length) {
    issues.push({ severity: 'blocking', code: 'no-fields', message: `数据区「${region.name || region.key}」没有启用任何字段。` });
  }
  for (const f of fields) {
    if (f.col === null || f.col === undefined) {
      issues.push({ severity: 'blocking', code: 'field-without-column', fieldKey: f.key, message: `字段「${f.label}」没有指定来源列。` });
      continue;
    }
    if (f.col >= table.colCount) {
      issues.push({ severity: 'blocking', code: 'field-column-out-of-range', fieldKey: f.key, message: `字段「${f.label}」的来源列超出表格列数。` });
    }
  }
  if (normalized.groupCol !== null && normalized.groupCol !== undefined && !fields.some((f) => f.col === normalized.groupCol)) {
    warnings.push({
      severity: 'warning', code: 'group-column-not-exported',
      message: `地区列（第 ${normalized.groupCol + 1} 列）没有作为字段导出，继承结果只用于空白地区单元格的判定。`,
    });
  }

  // data rows only; boundaries already exclude subtotal/note/repeat-header/section rows
  const dataRowSet = new Set(analysis.segments.flatMap((s) => s.rows));
  const rows = [];
  const values = [];
  for (const row of table.rows) {
    if (!dataRowSet.has(row.index)) continue;
    const groupInfo = analysis.rowSources.get(row.index);
    const record = {};
    const provenance = {};
    const rowIssues = [];
    for (const f of fields) {
      if (f.role === 'group') {
        const g = groupInfo;
        if (!g || g.groupText === null) {
          record[f.key] = null;
          provenance[f.key] = {
            kind: BLANK_KIND,
            raw: '',
            cellText: '',
            fileId: context.fileId,
            fileName: context.fileName,
            sha256: context.sha256,
            tableId: table.id,
            tableOrdinal: region.tableOrdinal,
            row: row.index,
            col: normalized.groupCol,
            displayRow: row.index + 1,
            displayCol: (normalized.groupCol ?? 0) + 1,
            header: f.label,
            regionKey: region.key,
            groupField: true,
            inherited: false,
          };
          rowIssues.push({ fieldKey: f.key, code: 'group-missing', message: '地区为空且无法继承。', severity: 'blocking' });
        } else {
          const src = g.source;
          record[f.key] = g.groupText;
          provenance[f.key] = {
            kind: VALUE_KIND,
            raw: g.groupText,
            cellText: g.groupText,
            value: g.groupText,
            fileId: context.fileId,
            fileName: context.fileName,
            sha256: context.sha256,
            tableId: table.id,
            tableOrdinal: region.tableOrdinal,
            row: src.row,
            col: src.col,
            displayRow: src.row + 1,
            displayCol: src.col + 1,
            valueRow: row.index,
            valueDisplayRow: row.index + 1,
            header: f.label,
            regionKey: region.key,
            groupField: true,
            inherited: src.kind === 'inherited',
            inheritKind: src.kind,
            inheritedFromRow: src.kind === 'inherited' ? src.row : null,
            inheritedFromDisplayRow: src.kind === 'inherited' ? src.row + 1 : null,
          };
        }
        {
          const prov = provenance[f.key];
          // `row` identifies the imported data row; the physical source cell is
          // kept separately so a carried-down group value still points at it.
          values.push({ ...prov, fieldKey: f.key, sourceRow: prov.row, sourceCol: prov.col, row: row.index });
        }
        continue;
      }
      const hit = grid[row.index]?.[f.col];
      const cell = hit?.cell;
      const rawText = cell ? cell.text : '';
      const interpreted = interpretValue(rawText, f);
      let kind = interpreted.kind;
      let value = interpreted.value;
      let sourceRow = cell ? cell.row : row.index;
      let sourceCol = cell ? cell.col : f.col;
      let inherited = false;

      // A vertically merged value cell explicitly continues the origin cell's value.
      if (cell && cell.continuation && interpreted.kind === VALUE_KIND && !interpreted.invalidNumber) {
        inherited = true;
      }
      const p = {
        kind,
        raw: interpreted.raw,
        cellText: cell ? cell.rawText ?? cell.text : '',
        value: value,
        fileId: context.fileId,
        fileName: context.fileName,
        sha256: context.sha256,
        tableId: table.id,
        tableOrdinal: region.tableOrdinal,
        row: sourceRow,
        col: sourceCol,
        displayRow: sourceRow + 1,
        displayCol: sourceCol + 1,
        valueRow: row.index,
        valueDisplayRow: row.index + 1,
        header: f.label,
        regionKey: region.key,
        inherited,
        fromRowMerge: !!(cell && cell.rowSpan > 1),
        rowSpan: cell ? cell.rowSpan : 1,
        colSpan: cell ? cell.colSpan : 1,
      };
      if (kind === BLANK_KIND) {
        p.kind = BLANK_KIND;
      } else if (kind === NONE_KIND) {
        p.kind = NONE_KIND;
      } else if (kind === ZERO_KIND) {
        p.kind = ZERO_KIND;
      }
      if (interpreted.invalidNumber) {
        rowIssues.push({ fieldKey: f.key, code: 'invalid-number', message: `「${interpreted.raw.trim()}」不是数字，已按原文保留。`, severity: 'warning' });
      }
      record[f.key] = value;
      provenance[f.key] = p;
      // One entry per imported data row; physical cell position stays in the
      // provenance (`row`/`col`), and `sourceRow` mirrors it for quick lookups.
      values.push({ ...p, fieldKey: f.key, sourceRow: p.row, sourceCol: p.col, row: row.index });
    }

    // nested table content never becomes a cell value
    for (const cell of row.cells) {
      if (cell.hasNestedTable) {
        rowIssues.push({
          fieldKey: null, code: 'nested-table-in-row', severity: 'warning',
          message: `第 ${row.index + 1} 行第 ${cell.col + 1} 列含嵌套表格；只导入了单元格自身文字，嵌套表内容未混合。`,
        });
      }
    }
    rows.push({ row: row.index, record, provenance, issues: rowIssues });
  }

  const blockingIssues = [...issues.filter((i) => i.severity === 'blocking'), ...pending.filter((p) => p.severity === 'blocking')];
  const stats = computeStats(rows, fields, values);
  return {
    regionKey: region.key,
    regionName: region.name,
    tableId: table.id,
    tableOrdinal: region.tableOrdinal,
    analysis,
    rows,
    values,
    issues,
    pending,
    warnings,
    blocking: blockingIssues,
    stats,
    nFields: fields.length,
    rowRange: { startRow: normalized.startRow, endRow: normalized.endRow },
    headerRow: normalized.headerRow,
    headerRowCount: normalized.headerRowCount,
    groupCol: normalized.groupCol,
  };
}

function zeroStats() {
  return { rows: 0, values: 0, blank: 0, none: 0, zero: 0, filled: 0 };
}

function computeStats(rows, fields, values) {
  const stats = { rows: rows.length, values: values.length, blank: 0, none: 0, zero: 0, filled: 0, fields: fields.length };
  for (const v of values) {
    if (v.kind === BLANK_KIND) stats.blank++;
    else if (v.kind === NONE_KIND) stats.none++;
    else if (v.kind === ZERO_KIND) stats.zero++;
    else stats.filled++;
  }
  return stats;
}

export function evaluateSpec(doc, spec, context = {}) {
  const diff = diffSpecAgainstDoc(doc, spec);
  const regions = [];
  for (const region of spec.regions || []) {
    // Match by header content first so a reordered table is still read correctly.
    const found = findTableForRegion(doc, region);
    const table = found.table;
    const extracted = extractRegion(
      doc,
      found.moved && table ? { ...region, tableOrdinal: found.ordinal } : region,
      { ...context, table }
    );
    if (found.moved && table) {
      extracted.tableMoved = { from: region.tableOrdinal, to: found.ordinal, tableId: table.id };
    }
    regions.push(extracted);
  }
  const pending = regions.flatMap((r) => r.pending.map((p) => ({ ...p, regionKey: r.regionKey })));
  const warnings = regions.flatMap((r) => r.warnings.map((w) => ({ ...w, regionKey: r.regionKey })));
  const issues = regions.flatMap((r) => r.issues.map((i) => ({ ...i, regionKey: r.regionKey })));
  const blocking = [
    ...diff.blocking.map((d) => ({ ...d, severity: 'blocking', source: 'structure' })),
    ...regions.flatMap((r) => r.blocking.map((b) => ({ ...b, regionKey: r.regionKey, source: b.source || 'region' }))),
  ];
  const totals = regions.reduce((acc, r) => {
    acc.rows += r.stats.rows;
    acc.values += r.stats.values;
    acc.blank += r.stats.blank;
    acc.none += r.stats.none;
    acc.zero += r.stats.zero;
    acc.filled += r.stats.filled;
    return acc;
  }, { rows: 0, values: 0, blank: 0, none: 0, zero: 0, filled: 0 });
  return {
    ok: blocking.length === 0,
    diff,
    regions,
    pending,
    warnings,
    issues,
    blocking,
    totals,
    specHash: hashSpec(spec),
  };
}

export function hashSpec(spec) {
  const canonical = JSON.stringify({
    regions: (spec.regions || []).map((r) => ({
      key: r.key,
      tableOrdinal: r.tableOrdinal,
      headerRow: r.headerRow,
      headerRowCount: r.headerRowCount,
      startRow: r.startRow,
      endRow: r.endRow,
      groupCol: r.groupCol,
      inheritGroup: r.inheritGroup !== false,
      headerSignature: specRegionSignature(r).cells,
      fields: (r.fields || []).map((f) => ({ key: f.key, label: f.label, col: f.col, role: f.role, type: f.type, include: f.include !== false })),
      groupOverrides: r.groupOverrides || {},
      nestedResolutions: r.nestedResolutions || {},
    })),
    options: spec.options || {},
  });
  let h = 2166136261;
  for (let i = 0; i < canonical.length; i++) {
    h ^= canonical.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return (h >>> 0).toString(16).padStart(8, '0');
}

/** Suggest a spec from a document (one region per top-level table with data). */
export function suggestSpec(doc) {
  const regions = [];
  let key = 0;
  doc.tables.forEach((table, ordinal) => {
    if (table.parentTableId) return;
    if (table.rowCount < 2) return;
    key += 1;
    const suggestion = suggestRegion(table);
    regions.push(buildRegionFromSuggestion(doc, table, ordinal, suggestion, `r${key}`, `数据区 ${key}`));
  });
  return {
    version: 1,
    name: '',
    regions,
    options: { preserveBlank: true, preserveNone: true, preserveZero: true },
  };
}

export function buildRegionFromSuggestion(doc, table, tableOrdinal, suggestion, key, name) {
  const fields = suggestion.fields.map((f) => ({
    key: f.key, label: f.label, col: f.col, role: f.role, type: f.type, include: true,
  }));
  const sig = tableSignatureOf(table, suggestion.headerRow, suggestion.headerRowCount);
  return {
    key,
    name,
    tableOrdinal,
    tableIdHint: table.id,
    headerRow: suggestion.headerRow,
    headerRowCount: suggestion.headerRowCount,
    startRow: suggestion.startRow,
    endRow: suggestion.endRow,
    groupCol: suggestion.groupCol,
    inheritGroup: suggestion.inheritGroup,
    headerSignature: sig.cells,
    colCount: sig.colCount,
    nestedTableOrdinals: (table.nestedTableIds || []).map((id) => doc.tables.findIndex((t) => t.id === id)).filter((i) => i >= 0),
    fields,
    confirmed: suggestion.confirmed,
    groupOverrides: {},
    nestedResolutions: {},
  };
}
