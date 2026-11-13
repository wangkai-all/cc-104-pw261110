import { api } from './api.js';
import { el, clear, toast, fmtSize, fmtTime, debounce } from './ui.js';

const state = {
  step: 'upload',
  file: null,
  doc: null,
  spec: null,
  specId: null,
  activeRegion: 0,
  tableOrdinalForEditor: null,
  analyses: new Map(), // region key -> latest server analysis of that region
  preview: null,
  previousBatchId: null,
  batches: [],
  samples: [],
  dirty: false,
};

const $ = (sel) => document.querySelector(sel);
const STEPS = ['upload', 'map', 'preview', 'batches'];

// ---------------- navigation ----------------
function setStep(step) {
  state.step = step;
  for (const s of STEPS) {
    const panel = $(`#step-${s}`);
    if (panel) panel.hidden = s !== step;
    const btn = document.querySelector(`.step[data-step="${s}"]`);
    if (btn) {
      btn.classList.toggle('active', s === step);
      btn.classList.toggle('done', STEPS.indexOf(s) < STEPS.indexOf(step));
    }
  }
  if (step === 'batches') loadBatches();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

document.querySelectorAll('.step').forEach((btn) => {
  btn.addEventListener('click', () => {
    const target = btn.dataset.step;
    if (target === 'map' && !state.file) return toast('请先上传一个 DOCX 文件。', 'error');
    if (target === 'preview' && !state.spec) return toast('请先建立字段对应关系。', 'error');
    setStep(target);
    if (target === 'map') renderMapping();
    if (target === 'preview') renderPreview();
  });
});

// ---------------- upload ----------------
const dropzone = $('#dropzone');
const fileInput = $('#file-input');

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
});
for (const evt of ['dragenter', 'dragover']) {
  dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add('drag'); });
}
for (const evt of ['dragleave', 'drop']) {
  dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove('drag'); });
}
dropzone.addEventListener('drop', (e) => {
  const file = e.dataTransfer?.files?.[0];
  if (file) handleUpload(file);
});
fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file) handleUpload(file);
  fileInput.value = '';
});

async function handleUpload(file) {
  const status = $('#upload-status');
  if (!/\.docx$/i.test(file.name)) {
    status.className = 'status error';
    status.textContent = '只支持 .docx 文件。若为 .doc，请先在 Word 中另存为 .docx。';
    return;
  }
  status.className = 'status';
  status.innerHTML = '<span class="spinner"></span>正在解析文档…';
  try {
    const result = await api.upload(file);
    await loadFile(result.file.id, result);
    if (result.duplicate) {
      status.className = 'status ok';
      status.textContent = `该文件此前已上传（内容一致），直接载入已有记录，未生成新文件。已有批次 ${result.batches.length} 个。`;
    } else {
      status.className = 'status ok';
      status.textContent = '解析完成。已识别表格结构，请核对右侧概览后进入字段对应。';
    }
  } catch (err) {
    status.className = 'status error';
    status.textContent = err.message;
    toast(err.message, 'error');
  }
}

document.querySelectorAll('.sample-open').forEach((btn) => {
  btn.addEventListener('click', async () => {
    const name = btn.dataset.sample;
    try {
      const res = await fetch(`/samples/${name}`);
      if (!res.ok) throw new Error('示例文件不可用');
      const blob = await res.blob();
      const file = new File([blob], name, { type: blob.type });
      await handleUpload(file);
    } catch (err) {
      toast(`载入示例失败：${err.message}`, 'error');
    }
  });
});

$('#download-sample').addEventListener('click', () => { /* native download */ });
$('#load-sample-btn').addEventListener('click', async () => {
  const name = state.samples[0] || 'sample_survey.docx';
  const btn = document.querySelector(`.sample-open[data-sample="${name}"]`);
  if (btn) btn.click();
  else {
    try {
      const res = await fetch(`/samples/${name}`);
      const blob = await res.blob();
      await handleUpload(new File([blob], name));
    } catch (e) { toast('示例文件不可用。', 'error'); }
  }
});

$('#help-btn').addEventListener('click', () => $('#help-dialog').showModal());

async function loadFile(fileId, uploadResult) {
  const data = await api.file(fileId);
  state.file = data.file;
  state.doc = data.doc;
  state.analyses = new Map();
  state.previousBatchId = null;
  const suggested = await api.suggest(fileId);
  state.spec = normalizeSpec(suggested.spec, fileId);
  state.specId = null;
  $('#spec-name').value = state.spec.name || '';
  renderDocView($('#doc-view'), state.doc, { highlightTables: true });
  $('#doc-title').textContent = state.doc.title;
  renderUploadSummary(uploadResult);
  $('#upload-preview').hidden = false;
  $('#download-original').onclick = () => {
    window.location.href = api.originalUrl(fileId);
  };
  state.dirty = true;
  renderMapping();
  setStep('upload');
}

function normalizeSpec(spec, fileId) {
  const s = JSON.parse(JSON.stringify(spec || { version: 1, regions: [], options: {} }));
  s.version = 1;
  s.options = s.options || { preserveBlank: true, preserveNone: true, preserveZero: true };
  s.regions = (s.regions || []).map((r, i) => ({
    ...r,
    key: r.key || `r${i + 1}`,
    name: r.name || `数据区 ${i + 1}`,
    inheritGroup: r.inheritGroup !== false,
    headerRowCount: r.headerRowCount || 1,
    fields: (r.fields || []).map((f) => ({ ...f, include: f.include !== false })),
    groupOverrides: r.groupOverrides || {},
    nestedResolutions: r.nestedResolutions || {},
    confirmed: r.confirmed || {},
  }));
  s.fileId = fileId;
  return s;
}

function renderUploadSummary(uploadResult) {
  const doc = state.doc;
  const spec = state.spec;
  const container = clear($('#upload-summary'));
  const tables = doc.tables.filter((t) => !t.parentTableId);
  const nested = doc.tables.filter((t) => t.parentTableId);
  container.append(
    el('div', { class: 'kv' },
      el('div', {}, `文件：${state.file.name}（${fmtSize(state.file.size)}）`),
      el('div', {}, `内容指纹 SHA-256：`, el('code', { text: state.file.sha256.slice(0, 24) + '…' })),
      el('div', {}, `标题：${doc.title}`),
      el('div', {}, `表格：${tables.length} 个顶层表${nested.length ? `，${nested.length} 个嵌套表` : ''}`),
    ),
    el('div', { class: 'cards', style: 'margin-top:14px' },
      card('数据区', String(spec.regions.length)),
      card('可导入行', String(spec.regions.reduce((n, r) => n + Math.max(0, (r.endRow ?? 0) - (r.startRow ?? 0) + 1), 0)), '估算'),
      card('附注', String(doc.tables.reduce((n, t) => n + (t.notes?.length || 0), 0))),
      card('待确认', String(countPendingSuggestions()), countPendingSuggestions() ? '需在下一步处理' : '自动识别', countPendingSuggestions() ? 'warn' : ''),
    ),
  );
  if (uploadResult?.batches?.length) {
    container.append(el('p', { class: 'note-list', style: 'margin-top:14px' },
      `已有 ${uploadResult.batches.length} 个批次：${uploadResult.batches.map((b) => `${b.status === 'committed' ? '已导入' : b.status === 'replaced' ? '已替换' : '预览'}（${b.recordCount} 行）`).join('、')}`));
  }
  container.append(el('p', { class: 'muted', style: 'margin-top:14px' },
    '表格中带底色的是自动识别的数据行；小计、附注和重复表头不会被当成数据。下一步可逐项确认。'));
}

function countPendingSuggestions() {
  if (!state.spec) return 0;
  return state.spec.regions.reduce((n, r) => n + (r.confirmed?.header === false ? 1 : 0), 0);
}

function card(k, v, hint = '', tone = '') {
  return el('div', { class: `card ${tone}`.trim() },
    el('div', { class: 'k', text: k }),
    el('div', { class: 'v', text: v }),
    hint ? el('div', { class: 'k', text: hint }) : null);
}

// ---------------- document view ----------------
/**
 * Render the parsed document, optionally annotating rows with their region type.
 * options: { highlightTables, region, classification, onRowClick, selectedTableId }
 */
function renderDocView(container, doc, options = {}) {
  clear(container);
  const {
    region = null, classification = null, onRowClick = null,
    onlyTableId = null, markCol = null, classes = null,
  } = options;
  const blocks = onlyTableId
    ? [{ kind: 'table', tableId: onlyTableId, index: 0 }]
    : doc.blocks;
  if (!blocks.length) {
    container.append(el('p', { class: 'empty', text: '文档中没有可显示的表格或段落。' }));
    return;
  }
  let lastWasTable = false;
  for (const block of blocks) {
    if (block.kind === 'paragraph') {
      if (!block.text) continue;
      container.append(el('p', { class: 'doc-para', text: block.text }));
      lastWasTable = false;
      continue;
    }
    const table = doc.tables.find((t) => t.id === block.tableId);
    if (!table) continue;
    container.append(renderTable(doc, table, { region, classification, onRowClick, markCol, classes, container }));
    lastWasTable = true;
  }
  if (!lastWasTable && !container.childElementCount) {
    container.append(el('p', { class: 'empty', text: '没有可显示的表格。' }));
  }
}

function typeOf(classification, rowIndex) {
  if (!classification) return null;
  return classification.types.find((t) => t.row === rowIndex)?.type || null;
}

function renderTable(doc, table, ctx) {
  const section = el('div', { class: 'doc-section' });
  // When the caption was read from the table's own merged title row, printing it
  // again in the heading just repeats what is already visible below.
  const titleText = table.titleRow !== undefined && table.titleRow !== null ? '' : (table.title || table.caption || '');
  section.append(el('h3', {},
    el('span', { class: 'tid', text: tableDisplayName(doc.tables.indexOf(table)) }),
    titleText ? el('span', { text: titleText }) : null,
    table.parentTableId ? el('span', { class: 'cellnote', text: `嵌套于 ${table.parentTableId}` }) : null,
    table.repeatHeaderRows ? null : null,
  ));
  const wrapper = el('div', { class: 'table-scroll' });
  const t = el('table', { class: 'doc-table' });
  const grid = buildClientGrid(table);
  const showTypeCol = !!ctx.classification;
  for (const row of table.rows) {
    const type = typeOf(ctx.classification, row.index);
    const cls = ['selectable'];
    if (type) cls.push(`type-${type}`);
    const tr = el('tr', { class: cls.join(' ') });
    if (type === 'data' || type === 'subtotal' || type === 'note' || type === 'repeat-header') {
      tr.classList.add('in-range');
    }
    if (ctx.onRowClick) {
      tr.addEventListener('click', () => ctx.onRowClick(row.index, table));
      tr.title = '点击将该行设为数据区起点；按住 Shift 点击设为终点';
    }
    if (showTypeCol) {
      const label = {
        header: '表头', title: '标题', data: '数据', subtotal: '小计', note: '附注',
        'repeat-header': '重复表头', 'word-repeat-header': '重复表头', section: '分节', blank: '空行', outside: '范围外',
      }[type] || '';
      tr.append(el('td', { class: 'rowtype', text: label }));
    }
    for (let c = 0; c < table.colCount; c++) {
      const hit = grid[row.index]?.[c];
      if (!hit) {
        tr.append(el('td', { class: 'merged' }));
        continue;
      }
      if (hit.cell.col !== c) continue; // covered by a horizontal merge
      const td = el('td', {
        colspan: hit.cell.colSpan > 1 ? hit.cell.colSpan : null,
        rowspan: hit.cell.rowSpan > 1 ? hit.cell.rowSpan : null,
        class: [hit.cell.colSpan > 1 || hit.cell.rowSpan > 1 ? 'merged' : ''].filter(Boolean).join(' '),
      });
      if (c === 0) td.append(el('span', { class: 'rownum', text: `R${row.index + 1}` }));
      const text = hit.cell.text;
      if (text) td.append(document.createTextNode(text));
      else td.append(el('span', { class: 'cellblank', text: '（空）' }));
      if (hit.cell.hasNestedTable) td.append(el('span', { class: 'cellnote', text: '含嵌套表' }));
      if (ctx.markCol !== null && ctx.markCol !== undefined && c === ctx.markCol) td.classList.add('col-highlight');
      td.dataset.row = String(row.index);
      td.dataset.col = String(c);
      td.dataset.table = table.id;
      tr.append(td);
    }
    t.append(tr);
  }
  wrapper.append(t);
  section.append(wrapper);
  if (table.notes?.length) {
    const ul = el('ul', { class: 'note-list' });
    for (const n of table.notes) ul.append(el('li', { text: n.text }));
    section.append(el('div', { class: 'muted', style: 'font-size:11px;margin-top:4px', text: '附注（不作为数值导入）' }), ul);
  }
  return section;
}

function buildClientGrid(table) {
  const rows = Math.max(table.rowCount, 1);
  const cols = Math.max(table.colCount, 1);
  const grid = Array.from({ length: rows }, () => Array(cols).fill(null));
  for (const row of table.rows) {
    for (const cell of row.cells) {
      const rs = Math.max(1, Math.min(cell.rowSpan || 1, rows - row.index));
      const cs = Math.max(1, Math.min(cell.colSpan || 1, cols - cell.col));
      for (let r = row.index; r < row.index + rs; r++) {
        for (let c = cell.col; c < cell.col + cs; c++) {
          if (!grid[r][c]) grid[r][c] = { cell, row: row.index, col: cell.col };
        }
      }
    }
  }
  return grid;
}

// ---------------- mapping ----------------
let mapShiftAnchor = null;

function currentRegion() {
  return state.spec?.regions?.[state.activeRegion] || null;
}

async function refreshAnalysis(region, table) {
  try {
    const body = {
      tableOrdinal: region.tableOrdinal,
      key: region.key,
      name: region.name,
      headerRow: region.headerRow,
      headerRowCount: region.headerRowCount,
      startRow: region.startRow,
      endRow: region.endRow,
      groupCol: region.groupCol ?? null,
      inheritGroup: region.inheritGroup !== false,
      nestedResolutions: region.nestedResolutions || {},
    };
    const res = await api.analyze(state.file.id, body);
    state.analyses.set(region.key, res);
    return res;
  } catch (err) {
    toast(err.message, 'error');
    return null;
  }
}

async function renderMapping() {
  if (!state.file || !state.spec) return;
  $('#map-file').textContent = state.file.name;
  $('#map-region-count').textContent = `${state.spec.regions.length} 个数据区`;
  const tableSelect = $('#table-select');
  clear(tableSelect);
  const tops = state.doc.tables.filter((t) => !t.parentTableId);
  for (const [i, t] of tops.entries()) {
    const ordinal = state.doc.tables.indexOf(t);
    tableSelect.append(el('option', { value: String(ordinal) }, tableOptionLabel(t, i)));
  }
  // A region may target a nested table (imported as its own data area).
  const current = currentRegion();
  if (current) tableSelect.value = String(current.tableOrdinal);
  // A region may target a nested table (imported as its own data area).
  const regionTable = state.doc.tables[current?.tableOrdinal];
  if (regionTable?.parentTableId && !tops.includes(regionTable)) {
    tableSelect.append(el('option', { value: String(current.tableOrdinal) }, `${tableDisplayName(current.tableOrdinal)}（${regionTable.id}）· ${regionTable.rowCount} 行`));
    tableSelect.value = String(current.tableOrdinal);
  }
  tableSelect.onchange = async () => {
    const ord = Number(tableSelect.value);
    const existing = state.spec.regions.find((r) => r.tableOrdinal === ord);
    if (existing) {
      state.activeRegion = state.spec.regions.indexOf(existing);
    } else {
      const suggested = await api.suggest(state.file.id, { tableOrdinal: ord });
      const nr = { ...suggested.region, key: `r${state.spec.regions.length + 1}`, name: `数据区 ${state.spec.regions.length + 1}` };
      nr.nestedResolutions = {};
      nr.groupOverrides = {};
      state.spec.regions.push(nr);
      state.activeRegion = state.spec.regions.length - 1;
    }
    await renderMapping();
  };
  renderRegionTabs();
  await renderRegionEditor();
  await renderMapDoc();
}

function tableOptionLabel(t, i) {
  const title = t.title || t.caption || '';
  return `${tableDisplayName(state.doc.tables.indexOf(t))}（${t.id}）${title ? ` · ${title.slice(0, 28)}` : ''} · ${t.rowCount} 行`;
}

function renderRegionTabs() {
  const tabs = clear($('#region-tabs'));
  state.spec.regions.forEach((r, i) => {
    const analysis = state.analyses.get(r.key);
    const blocking = analysis?.result?.blocking?.length || 0;
    const warnings = analysis?.result?.warnings?.length || 0;
    const dotClass = blocking ? 'bad' : warnings ? 'warn' : '';
    tabs.append(el('button', {
      class: `rtab ${i === state.activeRegion ? 'active' : ''}`,
      type: 'button',
      onclick: async () => {
        state.activeRegion = i;
        await renderMapping();
      },
    }, el('span', { class: `dot ${dotClass}` }), r.name, el('span', { class: 'muted', text: tableDisplayName(r.tableOrdinal) })));
  });
  if (state.spec.regions.length < state.doc.tables.filter((t) => !t.parentTableId).length) {
    tabs.append(el('span', { class: 'muted', style: 'font-size:11.5px;align-self:center', text: '可在左侧“表格”下拉中添加数据区' }));
  }
}

async function renderRegionEditor() {
  const container = clear($('#region-editor'));
  const region = currentRegion();
  if (!region) {
    container.append(el('p', { class: 'empty', text: '尚未建立数据区。请在左上角选择表格。' }));
    return;
  }
  const table = state.doc.tables[region.tableOrdinal];
  const analysisRes = await refreshAnalysis(region, table);
  const result = analysisRes?.result;
  const nf = await renderFieldsBlock(region, table, result);
  container.append(nf);

  // group + options block
  const opts = el('div', { class: 'editor-block' },
    el('h3', { text: '数据区设置' }),
    el('div', { class: 'grid3' },
      field('表头行（第几行）', numberInput(region.headerRow + 1, async (v) => {
        region.headerRow = Math.max(0, v - 1);
        region.confirmed.header = true;
        await syncRegionWithTable(region, table);
        await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
      })),
      field('表头行数', numberInput(region.headerRowCount, async (v) => {
        region.headerRowCount = Math.max(1, v);
        await syncRegionWithTable(region, table);
        await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
      })),
      field('数据区起始行', numberInput(region.startRow + 1, async (v) => {
        region.startRow = Math.max(0, v - 1);
        if (region.startRow > region.endRow) region.endRow = region.startRow;
        region.confirmed.range = true;
        await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
      })),
      field('数据区结束行', numberInput(region.endRow + 1, async (v) => {
        region.endRow = Math.min(table.rowCount - 1, Math.max(region.startRow, v - 1));
        region.confirmed.range = true;
        await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
      })),
      field('地区（向下继承）列', selectCols(table, region.groupCol, async (v) => {
        const col = v === '' ? null : Number(v);
        setGroupCol(region, table, col);
        await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
      })),
      field('地区列作用', el('div', {},
        el('label', { class: 'checkline' },
          el('input', {
            type: 'checkbox', checked: region.inheritGroup !== false,
            onchange: async (e) => {
              region.inheritGroup = e.target.checked;
              await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
            },
          }),
          el('span', { text: '空白地区沿用同一数据区内上方的值' })))),
    ),
    el('p', { class: 'muted', style: 'font-size:11.5px;margin-top:8px' },
      '继承只发生在同一数据区内：小计行、重复表头、附注行、空行或另一个表格都会中断继承。'),
  );
  container.append(opts);

  // pending items
  renderPendingPanel(result, region);
}

function field(label, control) {
  return el('label', { class: 'field' }, el('span', { text: label }), control);
}

function numberInput(value, onChange) {
  const input = el('input', { type: 'text', inputmode: 'numeric', value: String(value) });
  input.addEventListener('change', () => {
    const n = parseInt(input.value, 10);
    if (Number.isFinite(n) && n >= 0) onChange(n);
    else input.value = String(value);
  });
  return input;
}

function selectCols(table, current, onChange) {
  const sel = el('select', {});
  sel.append(el('option', { value: '' }, '（不设置）'));
  for (let c = 0; c < table.colCount; c++) {
    const label = headerLabel(table, c);
    sel.append(el('option', { value: String(c), selected: current === c }, `第 ${c + 1} 列 · ${label || '(空表头)'}`));
  }
  sel.addEventListener('change', () => onChange(sel.value));
  return sel;
}

function headerLabel(table, col) {
  const region = currentRegion();
  const row = Math.min((region?.headerRow ?? 0) + (region?.headerRowCount || 1) - 1, table.rowCount - 1);
  const grid = buildClientGrid(table);
  const hit = grid[row]?.[col];
  return hit ? hit.cell.text.replace(/\s+/g, ' ').trim() : '';
}

/**
 * Human-facing table number. Top-level tables are numbered in document order;
 * nested tables are shown under their parent's number so that the tab label and
 * the table picker never disagree (doc.tables includes nested tables in between).
 */
function tableDisplayName(ordinal) {
  const doc = state.doc;
  const table = doc?.tables?.[ordinal];
  if (!table) return `表格 ${ordinal + 1}`;
  const tops = doc.tables.filter((t) => !t.parentTableId);
  if (!table.parentTableId) return `表 ${tops.indexOf(table) + 1}`;
  const parent = doc.tables.find((t) => t.id === table.parentTableId);
  const parentNo = tops.indexOf(parent) + 1;
  const siblings = doc.tables.filter((t) => t.parentTableId === table.parentTableId);
  return `表 ${parentNo} · 嵌套表${siblings.length > 1 ? siblings.indexOf(table) + 1 : ''}`;
}

function setGroupCol(region, table, col) {
  region.groupCol = col;
  for (const f of region.fields) f.role = f.role === 'group' ? 'value' : f.role;
  if (col === null) return;
  const f = region.fields.find((x) => x.col === col);
  if (f) { f.role = 'group'; f.type = 'text'; return; }
  const label = headerLabel(table, col) || `第${col + 1}列`;
  region.fields.unshift({ key: `c${col}`, label, col, role: 'group', type: 'text', include: true });
}

async function syncRegionWithTable(region, table) {
  // header change invalidates the signature, so recompute it from the table
  const res = await api.analyze(state.file.id, {
    tableOrdinal: region.tableOrdinal, key: region.key, name: region.name,
    headerRow: region.headerRow, headerRowCount: region.headerRowCount,
    startRow: region.startRow, endRow: region.endRow,
    groupCol: region.groupCol ?? null, inheritGroup: region.inheritGroup !== false,
    nestedResolutions: region.nestedResolutions || {},
  });
  region.headerSignature = res.region.headerSignature;
  region.colCount = res.region.colCount;
  state.analyses.set(region.key, res);
}

async function renderFieldsBlock(region, table, result) {
  const block = el('div', { class: 'editor-block' },
    el('h3', {}, '字段对应 ', el('span', { class: 'muted', style: 'font-weight:400;font-size:11.5px' }, `来自表头第 ${region.headerRow + 1}${region.headerRowCount > 1 ? `–${region.headerRow + region.headerRowCount}` : ''} 行`)),
  );
  const tableEl = el('table', { class: 'fields-table' });
  tableEl.append(el('thead', {}, el('tr', {},
    el('th', { text: '导入字段' }), el('th', { text: '来源列' }), el('th', { text: '类型' }), el('th', { text: '角色' }), el('th', { text: '' }))));
  const tbody = el('tbody');
  for (const [i, f] of region.fields.entries()) {
    const colOptions = el('select', {});
    for (let c = 0; c < table.colCount; c++) {
      colOptions.append(el('option', { value: String(c), selected: f.col === c }, `第 ${c + 1} 列 · ${headerLabel(table, c) || '(空)'}`));
    }
    colOptions.addEventListener('change', async () => {
      const newCol = Number(colOptions.value);
      const dup = region.fields.find((x, j) => j !== i && x.col === newCol);
      f.col = newCol;
      if (!f.label || /^第\d+列$/.test(f.label)) f.label = headerLabel(table, newCol) || f.label;
      if (dup) toast(`注意：第 ${newCol + 1} 列同时被「${dup.label}」使用。`, 'error');
      await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
    });
    const labelInput = el('input', { type: 'text', value: f.label });
    labelInput.addEventListener('change', () => { f.label = labelInput.value.trim() || f.label; state.dirty = true; });
    const typeSel = el('select', {},
      el('option', { value: 'text', selected: f.type === 'text' }, '文本'),
      el('option', { value: 'number', selected: f.type === 'number' }, '数值'),
    );
    typeSel.addEventListener('change', async () => { f.type = typeSel.value; await renderRegionEditor(); });
    const roleSel = el('select', {},
      el('option', { value: 'value', selected: f.role !== 'group' }, '数据字段'),
      el('option', { value: 'group', selected: f.role === 'group' }, '地区（继承）'),
    );
    roleSel.addEventListener('change', async () => {
      if (roleSel.value === 'group') { setGroupCol(region, table, f.col); }
      else {
        f.role = 'value';
        if (region.groupCol === f.col) region.groupCol = null;
      }
      await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
    });
    const del = el('button', {
      class: 'icon-btn', type: 'button', title: '不导入该字段',
      onclick: async () => {
        region.fields.splice(i, 1);
        if (region.groupCol === f.col) region.groupCol = null;
        await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
      },
    }, '✕');
    tbody.append(el('tr', {},
      el('td', {}, labelInput),
      el('td', {}, colOptions),
      el('td', {}, typeSel),
      el('td', {}, roleSel),
      el('td', {}, del)));
  }
  tableEl.append(tbody);
  block.append(tableEl);
  block.append(el('p', { class: 'muted', style: 'font-size:11.5px;margin-top:8px' },
    result ? `当前数据区可导入 ${result.stats.rows} 行 · ${result.stats.filled} 个非空值 · 空白 ${result.stats.blank} · “无” ${result.stats.none} · 零 ${result.stats.zero}` : '正在分析…'));
  return block;
}

function renderPendingPanel(result, region) {
  const panel = clear($('#pending-panel'));
  if (!result) return;
  const items = [
    ...(result.pending || []).map((p) => ({ ...p, kind: 'pending' })),
    ...(result.issues || []).filter((i) => i.severity === 'blocking').map((i) => ({ ...i, kind: 'issue', title: '字段问题', message: i.message })),
    ...(result.warnings || []).map((w) => ({ ...w, kind: 'warning' })),
  ];
  if (!items.length) {
    panel.append(el('div', { class: 'pending-item auto' },
      el('div', { class: 'pt' }, '✓ 没有待确认项'),
      el('div', { class: 'why', text: '表头、数据区、地区继承和字段对应均已识别。可继续预览。' })));
    return;
  }
  panel.append(el('h3', { style: 'font-size:12.5px;color:var(--ink-2)', text: `待确认项（${items.filter((i) => i.severity === 'blocking').length} 项阻止导入）` }));
  for (const item of items) {
    const cls = item.kind === 'warning' || item.severity === 'warning' || item.severity === 'info' ? 'warning' : '';
    const node = el('div', { class: `pending-item ${cls}` },
      el('div', { class: 'pt' },
        el('span', { text: item.title || item.message }),
        el('span', { class: 'kind', text: item.type || item.code || '' })),
      item.title ? el('div', { class: 'why', text: item.message }) : null);
    const actions = el('div', { class: 'actions' });
    if (item.action === 'resolve-nested' && item.nestedTableId) {
      actions.append(
        el('button', {
          class: 'btn small', type: 'button',
          onclick: async () => {
            region.nestedResolutions = { ...(region.nestedResolutions || {}), [item.nestedTableId]: 'skip' };
            await renderMapping();
            toast('已选择：跳过该嵌套表，只导入单元格自身文字。', 'ok');
          },
        }, '跳过该嵌套表'),
        el('button', {
          class: 'btn ghost small', type: 'button',
          onclick: async () => {
            await addNestedAsRegion(item.nestedTableId);
          },
        }, '作为单独数据区导入'),
      );
    }
    if (item.action === 'confirm-header' && item.candidates) {
      for (const cand of item.candidates.slice(0, 4)) {
        actions.append(el('button', {
          class: 'btn ghost small', type: 'button',
          onclick: async () => {
            region.headerRow = cand.row;
            region.confirmed.header = true;
            await syncRegionWithTable(region, state.doc.tables[region.tableOrdinal]);
            await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
          },
        }, `用第 ${cand.row + 1} 行作表头`));
      }
    }
    if (item.action === 'fill-group' && Number.isInteger(item.row)) {
      const input = el('input', { type: 'text', placeholder: '填写该行地区名称' });
      actions.append(input, el('button', {
        class: 'btn small', type: 'button',
        onclick: async () => {
          const v = input.value.trim();
          if (!v) return toast('请先填写地区名称。', 'error');
          region.groupOverrides = { ...(region.groupOverrides || {}), [String(item.row)]: v };
          await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
        },
      }, '应用'));
      actions.append(el('button', {
        class: 'btn ghost small', type: 'button',
        onclick: async () => {
          region.startRow = item.row;
          region.confirmed.range = true;
          await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
          toast(`数据区已从第 ${item.row + 1} 行开始，该行之前的地区不再继承。`, 'ok');
        },
      }, `从第 ${item.row + 1} 行重新开始`));
    }
    if (item.action === 'set-range' || item.action === 'range-overlaps-header') {
      actions.append(el('button', {
        class: 'btn ghost small', type: 'button',
        onclick: async () => {
          const res = await api.suggest(state.file.id, { tableOrdinal: region.tableOrdinal });
          Object.assign(region, res.region, { key: region.key, name: region.name, groupOverrides: {}, nestedResolutions: region.nestedResolutions || {} });
          await renderRegionEditor(); await renderMapDoc(); renderRegionTabs();
        },
      }, '恢复自动识别的范围'));
    }
    if (actions.childElementCount) node.append(actions);
    panel.append(node);
  }
}

async function addNestedAsRegion(nestedId) {
  const ordinal = state.doc.tables.findIndex((t) => t.id === nestedId);
  if (ordinal < 0) return toast('找不到该嵌套表。', 'error');
  const existing = state.spec.regions.find((r) => r.tableOrdinal === ordinal);
  if (existing) {
    state.activeRegion = state.spec.regions.indexOf(existing);
    await renderMapping();
    return toast('该嵌套表已经是一个数据区。', 'ok');
  }
  const suggested = await api.suggest(state.file.id, { tableOrdinal: ordinal });
  const nested = state.doc.tables[ordinal];
  const region = {
    ...suggested.region,
    key: `r${state.spec.regions.length + 1}`,
    name: `嵌套表 ${nestedId}`,
    nestedResolutions: {},
    groupOverrides: {},
  };
  state.spec.regions.push(region);
  // mark the parent cell's nested resolution as "separate" so the parent stops asking
  for (const r of state.spec.regions) {
    if (r === region) continue;
    if (state.doc.tables[r.tableOrdinal]?.nestedTableIds?.includes(nestedId)) {
      r.nestedResolutions = { ...(r.nestedResolutions || {}), [nestedId]: 'separate' };
    }
  }
  state.activeRegion = state.spec.regions.length - 1;
  await renderMapping();
  toast('已把嵌套表添加为单独数据区。请核对它的表头与地区列。', 'ok');
}

async function renderMapDoc() {
  const region = currentRegion();
  const container = $('#map-doc');
  if (!region) return;
  const analysis = state.analyses.get(region.key);
  const classification = analysis?.analysis ? {
    types: analysis.analysis.classification.types,
  } : null;
  renderDocView(container, state.doc, {
    region,
    classification,
    onlyTableId: state.doc.tables[region.tableOrdinal]?.id,
    markCol: region.groupCol,
    onRowClick: async (rowIndex, table) => {
      if (mapShiftAnchor !== null) {
        region.startRow = Math.min(mapShiftAnchor, rowIndex);
        region.endRow = Math.max(mapShiftAnchor, rowIndex);
        region.confirmed.range = true;
        mapShiftAnchor = null;
      } else {
        if (rowIndex <= region.headerRow + region.headerRowCount - 1) {
          region.headerRow = rowIndex;
          region.headerRowCount = 1;
          region.startRow = rowIndex + 1;
          region.endRow = Math.max(region.endRow, region.startRow);
          region.confirmed.header = true;
        } else {
          region.startRow = rowIndex;
          if (region.endRow < region.startRow) region.endRow = table.rowCount - 1;
        }
      }
      mapShiftAnchor = rowIndex;
      await renderRegionEditor();
      await renderMapDoc();
      renderRegionTabs();
    },
  });
}

// shift-click support: remember anchor on plain click
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') mapShiftAnchor = null; });

// ---------------- preview ----------------
$('#goto-preview').addEventListener('click', async () => {
  await buildPreview();
  setStep('preview');
});

async function buildPreview() {
  if (!state.spec?.regions?.length) return toast('至少需要一个数据区。', 'error');
  $('#preview-file').textContent = state.file.name;
  try {
    const res = await api.preview({
      fileId: state.file.id,
      spec: state.spec,
      previousBatchId: state.previousBatchId,
    });
    state.preview = res;
    renderPreview();
  } catch (err) {
    toast(err.message, 'error');
  }
}

function renderPreview() {
  if (!state.preview) { buildPreview(); return; }
  const { evaluation, diff } = state.preview;
  $('#preview-stats').textContent = `${evaluation.totals.rows} 行 · ${evaluation.totals.values} 个取值`;
  renderBanner(evaluation);
  renderDiff(diff);
  renderPreviewDoc();
  renderPreviewTable();
}

function renderBanner(evaluation) {
  const banner = $('#preview-banner');
  banner.hidden = false;
  banner.className = 'banner';
  clear(banner);
  const blocking = evaluation.blocking || [];
  const warnings = (evaluation.pending || []).filter((p) => p.severity !== 'blocking');
  if (blocking.length) {
    banner.classList.add('bad');
    banner.append(el('div', { class: 'grow' },
      el('strong', { text: `还有 ${blocking.length} 项待确认，暂不能导入。` }),
      el('ul', { class: 'note-list' },
        blocking.slice(0, 6).map((b) => el('li', { text: b.message || b.title || b.code })),
        blocking.length > 6 ? el('li', { text: `…另有 ${blocking.length - 6} 项` }) : null)));
    banner.append(el('button', {
      class: 'btn small', type: 'button', onclick: () => { setStep('map'); },
    }, '回到字段对应处理'));
  } else {
    banner.classList.add('ok');
    banner.append(el('div', { class: 'grow' },
      el('strong', { text: '可以导入。' }),
      el('span', { text: `共 ${evaluation.totals.rows} 行数据，${evaluation.totals.values} 个取值：空白 ${evaluation.totals.blank}、填“无” ${evaluation.totals.none}、零 ${evaluation.totals.zero}，三者分别保留。` })));
  }
  if (evaluation.diff && !evaluation.diff.ok) {
    banner.append(el('div', { class: 'grow', style: 'font-size:12px' },
      `结构差异 ${evaluation.diff.blocking.length} 项：${evaluation.diff.blocking.slice(0, 3).map((d) => d.message).join('；')}`));
  }
  const warnItems = [...(evaluation.warnings || []).filter((w) => w.severity === 'info' || w.severity === 'warning')];
  if (warnItems.length) {
    banner.append(el('details', { class: 'grow', style: 'font-size:12px' },
      el('summary', { text: `${warnItems.length} 条说明（已排除行、嵌套表等）` }),
      el('ul', { class: 'note-list' }, warnItems.slice(0, 12).map((w) => el('li', { text: w.message })))));
  }
}

function renderDiff(diff) {
  const panel = $('#preview-diff');
  if (!diff) { panel.hidden = true; return; }
  panel.hidden = false;
  clear(panel);
  panel.append(el('h3', {},
    '与上一批次的差异 ',
    el('span', { class: 'muted', style: 'font-weight:400;font-size:12px' },
      `（对比 ${diff.previousBatch.id.slice(0, 8)} · ${diff.previousBatch.status === 'committed' ? '已导入' : diff.previousBatch.status === 'replaced' ? '已替换' : '预览'}）`)));
  panel.append(el('p', { class: 'muted', style: 'font-size:12px', text: diff.summary }));
  for (const region of diff.regions) {
    const box = el('div', { class: 'diff-region' });
    box.append(el('strong', { text: `${region.regionName}（${region.regionKey}）` }),
      el('span', { class: 'muted', text: `：${region.oldStats?.rows ?? 0} 行 → ${region.newStats?.rows ?? 0} 行` }));
    for (const s of region.structural) {
      box.append(el('div', { class: 'diff-structural', text: `${labelOfKey(s.key)}：${JSON.stringify(s.before)} → ${JSON.stringify(s.after)}` }));
    }
    for (const ch of region.changes.slice(0, 40)) {
      if (ch.type === 'value-changed') {
        box.append(el('div', { class: 'diff-change' },
          el('span', { class: 'ab', text: `R${ch.row + 1}` }), ' ',
          el('span', { text: `${ch.fieldLabel}：` }),
          el('span', { class: 'ab before', text: fmtVal(ch.before) }), ' → ',
          el('span', { class: 'ab after', text: fmtVal(ch.after) }),
          ch.afterSource ? el('span', { class: 'muted', text: `（新来源 第${ch.afterSource.row + 1}行第${ch.afterSource.col + 1}列）` }) : null));
      } else {
        box.append(el('div', { class: 'diff-change', text: ch.message }));
      }
    }
    if (region.changes.length > 40) box.append(el('div', { class: 'muted', style: 'font-size:12px', text: `…另有 ${region.changes.length - 40} 处` }));
    panel.append(box);
  }
  if (!diff.totalChanges) panel.append(el('p', { class: 'muted', text: '没有数值变化。如果这是更正映射后重新导入，确认后旧批次会被标记为“已替换”，但不会被删除。' }));
  panel.append(el('p', { class: 'muted', style: 'font-size:12px;margin-top:8px' },
    '确认导入会生成新批次；旧批次及其原文件保持不变，只标记为“已替换”。'));
}

function labelOfKey(k) {
  return { headerRow: '表头行', headerRowCount: '表头行数', startRow: '起始行', endRow: '结束行', groupCol: '地区列' }[k] || k;
}

function fmtVal(v) {
  if (v === null || v === undefined || v === '') return '（空）';
  return String(v);
}

function renderPreviewDoc() {
  const container = $('#preview-doc');
  const first = state.preview?.evaluation.regions?.[0];
  const region = state.spec.regions.find((r) => r.regionKey === first?.regionKey) ||
    state.spec.regions[state.activeRegion] || state.spec.regions[0];
  if (!region) return;
  const analysis = state.analyses.get(region.key);
  const classification = analysis?.analysis ? { types: analysis.analysis.classification.types } : null;
  renderDocView(container, state.doc, {
    region,
    classification,
    onlyTableId: state.doc.tables[region.tableOrdinal]?.id,
    markCol: region.groupCol,
  });
}

function flattenPreviewRows() {
  const rows = [];
  for (const region of state.preview?.evaluation.regions || []) {
    const specRegion = state.spec.regions.find((r) => r.key === region.regionKey);
    const fields = specRegion?.fields || [];
    for (const row of region.rows || []) {
      rows.push({ region, row, fields, specRegion });
    }
  }
  return rows;
}

function renderPreviewTable() {
  const container = clear($('#preview-table-wrap'));
  const allRows = flattenPreviewRows();
  const filter = $('#preview-filter').value;
  const matches = (row) => {
    if (filter === 'all') return true;
    return Object.entries(row.provenance).some(([k, p]) => {
      if (filter === 'inherited') return p.inherited && p.groupField;
      return p.kind === filter;
    });
  };
  if (!allRows.length) {
    container.append(el('p', { class: 'empty', text: '没有可预览的数据行。' }));
    return;
  }
  const byRegion = dedupeRegions(allRows);
  let shown = 0;
  for (const { region, specRegion } of byRegion) {
    const rows = (region.rows || []).filter(matches);
    const fields = (specRegion?.fields || []).filter((f) => f.include !== false);
    const box = el('div', { class: 'preview-region' });
    const range = region.rowRange?.startRow === null
      ? ''
      : `第 ${region.rowRange.startRow + 1}–${region.rowRange.endRow + 1} 行`;
    box.append(el('h3', { class: 'preview-region-head' },
      el('span', { class: 'tid', text: tableDisplayName(region.tableOrdinal) }),
      el('span', { text: specRegion?.name || region.regionKey }),
      el('span', { class: 'muted', text: `${region.tableId} · ${range} · ${rows.length} 行` })));
    if (!rows.length) {
      box.append(el('p', { class: 'empty', text: '当前筛选条件下该数据区没有取值。' }));
      container.append(box);
      continue;
    }
    const t = el('table', { class: 'data-table' });
    t.append(el('thead', {}, el('tr', {},
      el('th', { text: '#' }),
      el('th', { text: '原表位置' }),
      fields.map((f) => el('th', { text: f.label, title: f.key })))));
    const body = el('tbody');
    for (const row of rows) {
      const tr = el('tr', {},
        el('td', { class: 'rowno', text: String(row.row + 1) }),
        el('td', { class: 'rowno' },
          el('button', {
            class: 'link', type: 'button', style: 'font-size:11px',
            onclick: () => locateCell(region.tableId, row.row, specRegion?.groupCol ?? 0),
          }, `R${row.row + 1}`)));
      for (const f of fields) {
        const prov = row.provenance[f.key];
        if (!prov) { tr.append(el('td', { class: 'muted', text: '—' })); continue; }
        tr.append(renderValueCell(region, row, f, prov));
      }
      body.append(tr);
    }
    t.append(body);
    box.append(t);
    shown += rows.length;
    container.append(box);
  }
  if (!shown) {
    container.append(el('p', { class: 'empty', text: '当前筛选条件下没有任何取值。' }));
  }
}

function dedupeRegions(rows) {
  const seen = new Set();
  const out = [];
  for (const r of rows) {
    if (seen.has(r.region.regionKey)) continue;
    seen.add(r.region.regionKey);
    out.push(r);
  }
  return out;
}

function renderValueCell(region, row, field, prov) {
  const td = el('td', { class: 'val' });
  if (prov.kind === 'blank') {
    td.append(el('span', { class: 'v-blank', title: '原单元格为空' }));
  } else if (prov.kind === 'none') {
    td.append(el('span', { class: 'v-none', text: '无' }));
  } else if (prov.kind === 'zero') {
    td.append(el('span', { class: 'v-zero', text: '0' }));
  } else {
    td.append(el('span', { class: `v-value ${prov.inherited ? 'v-inherit' : ''}`.trim(), text: prov.value === null ? '' : String(prov.value) }));
  }
  if (prov.inherited && prov.groupField) {
    td.append(el('span', { class: 'prov-hint inherited', text: `← 继承 R${(prov.row ?? 0) + 1}` }));
  } else if (prov.kind !== 'blank') {
    td.append(el('span', { class: 'prov-hint', text: `R${(prov.row ?? 0) + 1}·C${(prov.col ?? 0) + 1}` }));
  }
  td.title = '点击定位到原表格';
  td.addEventListener('click', () => {
    locateCell(prov.tableId, prov.row, prov.col);
    showProvenance(prov, field, region, row);
  });
  return td;
}

function showProvenance(prov, field, region, row) {
  const panel = $('#prov-panel');
  panel.hidden = false;
  clear(panel);
  const kindText = { blank: '空白（原单元格为空）', none: '填写了“无”', zero: '填写了数字 0', value: '有值' }[prov.kind] || prov.kind;
  panel.append(
    el('div', { class: 'prov-head' },
      el('strong', { text: `${field.label} · 第 ${row.row + 1} 行` }),
      el('button', { class: 'icon-btn', type: 'button', onclick: () => { panel.hidden = true; } }, '✕')),
    el('div', { class: 'prov-grid' },
      provRow('取值状态', kindText),
      provRow('文件', `${prov.fileName || state.file.name}`),
      provRow('内容指纹', `${(prov.sha256 || state.file.sha256).slice(0, 20)}…`),
      provRow('表格', `第 ${(prov.tableOrdinal ?? 0) + 1} 个表格（${prov.tableId}）`),
      provRow('单元格', `第 ${(prov.row ?? 0) + 1} 行 · 第 ${(prov.col ?? 0) + 1} 列`),
      prov.inherited && prov.groupField ? provRow('继承自', `第 ${(prov.inheritedFromRow ?? 0) + 1} 行（同一数据区内）`) : null,
      prov.inherited && prov.groupField && prov.inheritKind === 'inherited' ? provRow('说明', '该行地区单元格空白或跨行合并，沿用上方地区') : null,
      prov.rowSpan > 1 ? provRow('合并', `纵向合并 ${prov.rowSpan} 行`) : null,
      prov.colSpan > 1 ? provRow('合并', `横向合并 ${prov.colSpan} 列`) : null,
      provRow('原始文字', prov.cellText || '（空）'),
      provRow('数据区', `${region.regionName || region.regionKey} · ${region.tableId}`),
    ),
  );
}

function provRow(k, v) {
  return el('div', { class: 'prov-kv' }, el('span', { class: 'k', text: k }), el('span', { class: 'v', text: v }));
}

function locateCell(tableId, row, col) {
  const view = $('#preview-doc');
  if (!view) return;
  const cells = view.querySelectorAll(`td[data-table="${cssEscape(tableId)}"][data-row="${row}"][data-col="${col}"]`);
  const target = cells[0] || view.querySelector(`td[data-table="${cssEscape(tableId)}"]`);
  if (!target) return;
  target.scrollIntoView({ block: 'center', behavior: 'smooth' });
  view.querySelectorAll('.cell-flash').forEach((n) => n.classList.remove('cell-flash'));
  target.classList.add('cell-flash');
  setTimeout(() => target.classList.remove('cell-flash'), 1500);
}

function cssEscape(s) {
  return String(s).replace(/["\\]/g, '\\$&');
}

$('#preview-filter').addEventListener('change', () => { if (state.preview) renderPreviewTable(); });

$('#preview-back').addEventListener('click', () => { setStep('map'); renderMapping(); });

$('#export-csv').addEventListener('click', async () => {
  if (!state.preview?.committedBatchId && !state.lastBatchId) {
    // export current preview via a temporary client-side CSV
    downloadPreviewCsv();
    return;
  }
  window.location.href = `/api/batches/${state.lastBatchId}/csv`;
});

function downloadPreviewCsv() {
  const rows = flattenPreviewRows();
  const fieldCols = [];
  for (const { region, specRegion } of rows) {
    for (const f of specRegion?.fields || []) {
      if (!fieldCols.some((c) => c.regionKey === region.regionKey && c.key === f.key)) {
        fieldCols.push({ regionKey: region.regionKey, key: f.key, label: f.label });
      }
    }
  }
  const esc = (v) => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [['__region', '__table', '__row', ...fieldCols.map((f) => f.label)].map(esc).join(',')];
  for (const { region, row, specRegion } of rows) {
    const line = [specRegion?.name || region.regionKey, region.tableId, row.row + 1];
    for (const f of fieldCols) {
      if (f.regionKey !== region.regionKey) { line.push(''); continue; }
      const p = row.provenance[f.key];
      line.push(p ? (p.kind === 'blank' ? '' : p.kind === 'zero' ? '0' : p.value) : '');
    }
    lines.push(line.map(esc).join(','));
  }
  const blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${state.file.name.replace(/\.docx$/i, '')}-预览.csv`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

$('#commit-btn').addEventListener('click', async () => {
  if (!state.preview) return toast('请先生成预览。', 'error');
  if (!state.preview.evaluation.ok) return toast('仍有待确认项，处理后才能导入。', 'error');
  const btn = $('#commit-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>正在导入…';
  try {
    const res = await api.createBatch({
      fileId: state.file.id,
      spec: state.spec,
      specId: state.specId,
      commit: true,
      previousBatchId: state.previousBatchId,
    });
    if (res.duplicate) {
      toast(res.message, 'ok');
    } else {
      state.lastBatchId = res.batch.id;
      const status = $('#preview-status');
      status.className = 'status ok';
      status.textContent = `已导入批次 ${res.batch.id.slice(0, 8)}：${res.batch.recordCount} 行。${
        res.batch.supersedesBatchId ? '上一批次已标记为“已替换”，数据仍保留。' : ''}`;
      toast('导入完成。可在“批次记录”查看新旧批次。', 'ok');
      setStep('batches');
    }
  } catch (err) {
    toast(err.message, 'error');
    if (err.payload?.evaluation) {
      state.preview.evaluation = err.payload.evaluation;
      renderPreview();
    }
  } finally {
    btn.disabled = false;
    btn.textContent = '确认导入';
  }
});

// ---------------- save / load specs ----------------
$('#spec-name').addEventListener('input', () => { state.dirty = true; });

$('#save-spec').addEventListener('click', async () => {
  const name = $('#spec-name').value.trim();
  if (!name) return toast('请填写方案名称。', 'error');
  if (!state.spec.regions.length) return toast('至少需要一个数据区。', 'error');
  try {
    const saved = await api.saveSpec({ id: state.specId, name, spec: { ...state.spec, name } });
    state.specId = saved.id;
    state.spec.name = name;
    toast(`方案「${name}」已保存，可应用到结构相同的文件。`, 'ok');
  } catch (err) {
    toast(err.message, 'error');
  }
});

$('#spec-library-btn').addEventListener('click', async () => {
  if (!state.file) return toast('请先上传一个文件，再把方案应用到它。', 'error');
  try {
    const { specs } = await api.specs();
    const list = clear($('#spec-list'));
    if (!specs.length) {
      list.append(el('p', { class: 'empty', text: '还没有保存的方案。在字段对应页填写名称并“保存映射方案”。' }));
    }
    for (const spec of specs) {
      list.append(el('div', { class: 'spec-item' },
        el('div', {},
          el('div', { class: 'name', text: spec.name }),
          el('div', { class: 'meta' },
            `${spec.spec.regions?.length || 0} 个数据区 · 更新 ${fmtTime(spec.updatedAt)} · `,
            el('code', { text: spec.id.slice(0, 8) }))),
        el('div', { class: 'spec-actions' },
          el('button', {
            class: 'btn small', type: 'button',
            onclick: async () => { $('#spec-dialog').close(); await applySavedSpec(spec); },
          }, '应用到当前文件'),
          el('button', {
            class: 'btn ghost small', type: 'button',
            onclick: async () => {
              if (!confirm(`删除方案「${spec.name}」？已导入的批次不会被删除。`)) return;
              try {
                await api.deleteSpec(spec.id);
                toast('方案已删除。', 'ok');
                $('#spec-library-btn').click();
              } catch (err) { toast(err.message, 'error'); }
            },
          }, '删除'))));
    }
    $('#spec-dialog').showModal();
  } catch (err) {
    toast(err.message, 'error');
  }
});

async function applySavedSpec(spec) {
  const panel = $('#spec-diff-panel');
  try {
    const res = await api.applySpec(spec.id, state.file.id);
    state.spec = normalizeSpec(res.spec, state.file.id);
    state.specId = spec.id;
    state.spec.name = spec.name;
    state.activeRegion = 0;
    $('#spec-name').value = spec.name;
    state.analyses = new Map();
    panel.hidden = false;
    clear(panel);
    panel.append(el('h3', {}, `已套用方案「${spec.name}」`),
      el('p', { class: 'muted', style: 'font-size:12px', text: '结构核对通过，字段按保存的对应关系落位。可继续预览。' }));
    if (res.warnings?.length) {
      panel.append(el('div', { class: 'diff-list' },
        res.warnings.map((d) => el('div', { class: 'diff-line warning', text: d.message }))));
    }
    await renderMapping();
    toast(`方案「${spec.name}」已套用。`, 'ok');
  } catch (err) {
    panel.hidden = false;
    clear(panel);
    if (err.status === 409 && err.payload?.diffs) {
      const blocking = err.payload.blocking || [];
      const warnings = err.payload.warnings || [];
      panel.append(
        el('h3', { style: 'color:var(--danger)' }, `方案「${spec.name}」与当前文件结构不一致，未套用`),
        el('p', { class: 'muted', style: 'font-size:12px', text: '为避免按列错位导入，差异未确认前不会应用该方案。请逐项核对后重新建立对应关系。' }),
      );
      const list = el('div', { class: 'diff-list' });
      for (const d of blocking) list.append(el('div', { class: 'diff-line blocking', text: d.message }));
      for (const d of warnings) list.append(el('div', { class: 'diff-line warning', text: d.message }));
      panel.append(list);
      toast(`方案未套用：发现 ${blocking.length} 处结构差异。`, 'error');
    } else {
      panel.append(el('h3', { style: 'color:var(--danger)', text: '套用方案失败' }), el('p', { class: 'muted', text: err.message }));
      toast(err.message, 'error');
    }
  }
}

// ---------------- batches ----------------
async function loadBatches() {
  const list = $('#batches-list');
  try {
    const { batches } = await api.batches(state.file?.id);
    state.batches = batches;
    clear(list);
    if (!batches.length) {
      list.append(el('p', { class: 'empty', text: '还没有批次。完成预览并“确认导入”后会出现在这里。' }));
    }
    for (const b of batches) list.append(renderBatchCard(b));
  } catch (err) {
    toast(err.message, 'error');
  }
}

$('#refresh-batches').addEventListener('click', loadBatches);

function renderBatchCard(b) {
  const specName = b.specName || b.spec?.name || '未命名方案';
  const card = el('div', { class: `batch-card ${b.status}` },
    el('div', {},
      el('div', { class: 'bt' },
        el('span', { class: `badge ${b.status}`, text: b.status === 'committed' ? '已导入' : b.status === 'replaced' ? '已替换' : '预览' }),
        el('strong', { text: specName }),
        el('span', { class: 'bid', text: b.id.slice(0, 8) }),
        b.supersedesBatchId ? el('span', { class: 'muted', text: `替换 ${b.supersedesBatchId.slice(0, 8)}` }) : null,
        b.replacedByBatchId ? el('span', { class: 'muted', text: `被 ${b.replacedByBatchId.slice(0, 8)} 替换` }) : null),
      el('div', { class: 'meta' },
        el('span', { text: `${b.recordCount} 行` }),
        b.totals ? el('span', { text: `空白 ${b.totals.blank} · 无 ${b.totals.none} · 零 ${b.totals.zero}` }) : null,
        el('span', { text: `创建 ${fmtTime(b.createdAt)}` }),
        b.committedAt ? el('span', { text: `导入 ${fmtTime(b.committedAt)}` }) : null)),
    el('div', { style: 'display:flex;gap:6px' },
      el('button', { class: 'btn ghost small', type: 'button', onclick: () => showBatch(b.id) }, '查看数据'),
      el('a', { class: 'btn ghost small', href: `/api/batches/${b.id}/csv`, download: '' }, 'CSV'),
      el('a', { class: 'btn ghost small', href: api.originalUrl(b.fileId), download: '' }, '原文件')));
  return card;
}

async function showBatch(id) {
  const detail = $('#batch-detail');
  detail.hidden = false;
  clear(detail);
  detail.append(el('p', { class: 'muted', text: '正在载入…' }));
  try {
    const { batch, records } = await api.batch(id);
    clear(detail);
    const summary = batch.totals || {};
    detail.append(el('h3', {},
      `批次 ${batch.id.slice(0, 8)} `,
      el('span', { class: `badge ${batch.status}`, text: batch.status === 'committed' ? '已导入' : batch.status === 'replaced' ? '已替换' : '预览' }),
      el('span', { class: 'muted', style: 'font-weight:400;font-size:12px;margin-left:8px', text: batch.specName || '' })));
    detail.append(el('div', { class: 'cards', style: 'margin-bottom:12px' },
      card('行数', String(batch.recordCount)),
      card('取值', String(summary.values ?? '')),
      card('空白', String(summary.blank ?? '')),
      card('“无”', String(summary.none ?? '')),
      card('零', String(summary.zero ?? ''))));
    if (batch.supersedesBatchId) detail.append(el('p', { class: 'muted', text: `本批次替换了 ${batch.supersedesBatchId.slice(0, 8)}；旧批次和原文件仍保留在列表中。` }));
    // Group rows by data region: columns differ between regions, so a single
    // wide table would show misleading blanks for columns that never existed.
    const regions = [];
    for (const r of records) {
      let group = regions.find((g) => g.regionKey === r.regionKey);
      if (!group) {
        group = { regionKey: r.regionKey, fields: [], records: [] };
        regions.push(group);
      }
      group.records.push(r);
      for (const v of r.values) {
        if (!group.fields.some((f) => f.fieldKey === v.fieldKey)) group.fields.push({ fieldKey: v.fieldKey, fieldLabel: v.fieldLabel });
      }
    }
    for (const group of regions) {
      const specRegion = (batch.spec?.regions || []).find((s) => s.key === group.regionKey);
      const wrap = el('div', { class: 'table-scroll', style: 'max-height:420px;overflow:auto;margin-bottom:12px' });
      const label = el('p', { class: 'muted', style: 'font-size:12px;margin:0 0 6px' },
        `${specRegion?.name || group.regionKey} · ${group.records[0].tableId} · ${group.records.length} 行`);
      const t = el('table', { class: 'data-table' });
      t.append(el('thead', {}, el('tr', {},
        el('th', { text: '#' }), el('th', { text: '位置' }),
        group.fields.map((f) => el('th', { text: f.fieldLabel })))));
      const tbody = el('tbody');
      group.records.forEach((r, i) => {
        const tr = el('tr', {}, el('td', { class: 'rowno', text: String(i + 1) }), el('td', { class: 'rowno', text: `${r.tableId} R${r.rowIndex + 1}` }));
        for (const f of group.fields) {
          const v = r.values.find((x) => x.fieldKey === f.fieldKey);
          const td = el('td', { class: 'val', title: v ? `第${v.provenance.row + 1}行第${v.provenance.col + 1}列` : '' });
          if (!v || v.kind === 'blank') td.append(el('span', { class: 'v-blank' }));
          else if (v.kind === 'none') td.append(el('span', { class: 'v-none', text: '无' }));
          else if (v.kind === 'zero') td.append(el('span', { class: 'v-zero', text: '0' }));
          else td.append(el('span', { class: 'v-value', text: String(v.value ?? '') }));
          tr.append(td);
        }
        tbody.append(tr);
      });
      t.append(tbody);
      wrap.append(t);
      detail.append(label, wrap);
    }
  } catch (err) {
    clear(detail);
    detail.append(el('p', { class: 'status error', text: err.message }));
  }
}

// ---------------- boot ----------------
(async function boot() {
  try {
    const health = await api.health();
    if (!health.ok) throw new Error('服务未就绪');
  } catch {
    toast('无法连接后端服务。', 'error');
  }
  try {
    const res = await fetch('/samples/index.json').catch(() => null);
    if (res && res.ok) {
      state.samples = (await res.json()).samples || [];
    }
  } catch { /* optional */ }
  setStep('upload');
})();

export { state };
