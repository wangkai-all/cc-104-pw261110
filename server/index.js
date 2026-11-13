import express from 'express';
import { fileURLToPath } from 'node:url';
import { dirname, join, resolve } from 'node:path';
import { existsSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { Store } from './store.js';
import { parseDocx } from './docx.js';
import {
  suggestSpec, evaluateSpec, extractRegion, diffSpecAgainstDoc, buildRegionFromSuggestion, hashSpec,
} from './engine.js';
import { suggestRegion, tableStructure } from './structure.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, '..');

export function createApp({ dataDir = join(ROOT, 'data'), staticDir = join(ROOT, 'public') } = {}) {
  const store = new Store(dataDir);
  const app = express();
  app.disable('x-powered-by');
  app.use(express.json({ limit: '4mb' }));

  const json = (res, status, body) => res.status(status).json(body);
  const fail = (res, status, message, extra = {}) => json(res, status, { error: { message, ...extra } });

  // ---------- meta ----------
  app.get('/api/health', (req, res) => json(res, 200, { ok: true, stats: store.stats() }));

  // ---------- files ----------
  app.post('/api/files', express.raw({ type: '*/*', limit: '60mb' }), (req, res) => {
    const body = req.body;
    if (!Buffer.isBuffer(body) || body.length === 0) return fail(res, 400, '没有收到文件内容。');
    const name = String(req.query.name || req.get('x-file-name') || 'uploaded.docx');
    if (!/\.docx$/i.test(name)) return fail(res, 400, '只支持 .docx 文件（.doc 旧格式请先在 Word 中另存为 .docx）。');
    const magic = body.subarray(0, 4).toString('hex');
    if (magic !== '504b0304' && magic !== '504b0506' && magic !== '504b0708') {
      return fail(res, 400, '文件不是有效的 DOCX（不是 ZIP 容器）。');
    }
    let doc;
    try {
      doc = parseDocx(body);
    } catch (err) {
      return fail(res, 422, err.message);
    }
    const existing = store.findFileBySha(doc.sha256);
    if (existing) {
      // Duplicate upload: never create a second file or a second batch of data.
      return json(res, 200, {
        duplicate: true,
        file: publicFile(existing),
        batches: store.listBatches({ fileId: existing.id }),
        message: '相同文件此前已上传，未生成新文件或新批次。可继续使用已有批次，或以新方案预览差异。',
      });
    }
    const file = store.saveFile({ sha256: doc.sha256, name, size: body.length, buffer: body, doc });
    return json(res, 201, {
      duplicate: false,
      file: publicFile(file),
      batches: store.listBatches({ fileId: file.id }),
    });
  });

  function publicFile(file) {
    return {
      id: file.id,
      name: file.name,
      sha256: file.sha256,
      size: file.size,
      uploadedAt: file.uploadedAt,
      title: file.doc.title,
      stats: file.doc.stats,
    };
  }

  app.get('/api/files', (req, res) => json(res, 200, { files: store.listFiles() }));

  app.get('/api/files/:id', (req, res) => {
    const file = store.getFile(req.params.id);
    if (!file) return fail(res, 404, '文件不存在。');
    return json(res, 200, { file: publicFile(file), doc: file.doc });
  });

  app.get('/api/files/:id/original', (req, res) => {
    const got = store.readOriginal(req.params.id);
    if (!got) return fail(res, 404, '原始文件不存在。');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(got.file.name)}`);
    return res.send(got.buffer);
  });

  // Suggest a mapping for a file; optional tableOrdinal narrows to one table.
  app.post('/api/files/:id/suggest', (req, res) => {
    const file = store.getFile(req.params.id);
    if (!file) return fail(res, 404, '文件不存在。');
    const doc = file.doc;
    const ordinal = req.body?.tableOrdinal;
    if (Number.isInteger(ordinal)) {
      const table = doc.tables[ordinal];
      if (!table) return fail(res, 404, '表格不存在。');
      const suggestion = suggestRegion(table);
      return json(res, 200, {
        region: buildRegionFromSuggestion(doc, table, ordinal, suggestion, 'r1', '数据区 1'),
        analysis: {
          headerAmbiguous: suggestion.auto.headerAmbiguous,
          candidates: suggestion.auto.candidates,
          table: tableStructure(table, suggestion.headerRow, suggestion.headerRowCount),
          notes: table.notes,
          footnotes: table.footnotes,
          nestedTableIds: table.nestedTableIds,
        },
      });
    }
    return json(res, 200, { spec: suggestSpec(doc) });
  });

  // Re-analyze a single table with explicit header/range choices (interactive editing).
  app.post('/api/files/:id/analyze', (req, res) => {
    const file = store.getFile(req.params.id);
    if (!file) return fail(res, 404, '文件不存在。');
    const { tableOrdinal, headerRow, headerRowCount, startRow, endRow, groupCol, inheritGroup, nestedResolutions } = req.body || {};
    const table = file.doc.tables[tableOrdinal];
    if (!table) return fail(res, 404, '表格不存在。');
    const suggestion = suggestRegion(table);
    const region = buildRegionFromSuggestion(file.doc, table, tableOrdinal, suggestion, req.body?.key || 'r1', req.body?.name || '数据区 1');
    if (Number.isInteger(headerRow)) {
      region.headerRow = headerRow;
      const sig = tableStructure(table, headerRow, headerRowCount || 1);
      region.headerSignature = sig.headers.filter((h) => h.isOrigin).map((h) => ({ col: h.col, text: h.text.trim(), colSpan: h.colSpan }));
      region.colCount = sig.colCount;
      region.confirmed.header = true;
    }
    if (Number.isInteger(headerRowCount)) region.headerRowCount = headerRowCount;
    if (Number.isInteger(startRow)) { region.startRow = startRow; region.confirmed.range = true; }
    if (Number.isInteger(endRow)) { region.endRow = endRow; region.confirmed.range = true; }
    if (Number.isInteger(groupCol) || groupCol === null) {
      region.groupCol = groupCol;
      for (const f of region.fields) f.role = f.role === 'group' ? 'value' : f.role;
      if (Number.isInteger(groupCol)) {
        const f = region.fields.find((x) => x.col === groupCol);
        if (f) { f.role = 'group'; f.type = 'text'; }
        else region.fields.unshift({ key: `c${groupCol}`, label: `第${groupCol + 1}列`, col: groupCol, role: 'group', type: 'text', include: true });
      }
    }
    if (typeof inheritGroup === 'boolean') region.inheritGroup = inheritGroup;
    region.nestedResolutions = nestedResolutions || {};
    const evaluated = extractRegion(file.doc, region, {
      fileId: file.id, fileName: file.name, sha256: file.sha256, table,
    });
    return json(res, 200, { region, analysis: evaluated.analysis, result: evaluated });
  });

  // ---------- specs ----------
  app.get('/api/specs', (req, res) => json(res, 200, { specs: store.listSpecs() }));

  app.get('/api/specs/:id', (req, res) => {
    const spec = store.getSpec(req.params.id);
    if (!spec) return fail(res, 404, '映射方案不存在。');
    return json(res, 200, spec);
  });

  app.post('/api/specs', (req, res) => {
    const { id, name, spec } = req.body || {};
    if (!spec || !Array.isArray(spec.regions)) return fail(res, 400, '方案缺少 regions。');
    const cleanName = String(name || spec.name || '').trim();
    if (!cleanName) return fail(res, 400, '请填写方案名称。');
    for (const region of spec.regions) {
      if (!Number.isInteger(region.tableOrdinal)) return fail(res, 400, '方案的每个数据区都需要 tableOrdinal。');
      if (!Array.isArray(region.fields) || !region.fields.length) return fail(res, 400, '方案的每个数据区至少需要一个字段。');
    }
    const saved = store.saveSpec({ id, name: cleanName, spec: { ...spec, name: cleanName } });
    return json(res, 201, saved);
  });

  app.delete('/api/specs/:id', (req, res) => {
    const existing = store.getSpec(req.params.id);
    if (!existing) return fail(res, 404, '映射方案不存在。');
    // Deleting a spec never deletes batches or files already produced from it.
    store.db.prepare('DELETE FROM specs WHERE id = ?').run(req.params.id);
    return json(res, 200, { ok: true });
  });

  // Apply a spec to a file as a preview (no commit).
  app.post('/api/specs/:id/apply', (req, res) => {
    const saved = store.getSpec(req.params.id);
    if (!saved) return fail(res, 404, '映射方案不存在。');
    const file = store.getFile(req.body?.fileId);
    if (!file) return fail(res, 404, '文件不存在。');
    const diff = diffSpecAgainstDoc(file.doc, saved.spec);
    if (!diff.ok) {
      return json(res, 409, {
        applied: false,
        diffs: diff.diffs,
        blocking: diff.blocking,
        warnings: diff.warnings,
        message: '文件结构与映射方案不一致，未套用，以免导入错列。请先处理下列差异。',
      });
    }
    const evaluation = evaluateSpec(file.doc, saved.spec, { fileId: file.id, fileName: file.name, sha256: file.sha256 });
    return json(res, 200, { applied: true, spec: saved.spec, evaluation, diffs: diff.diffs, warnings: diff.warnings });
  });

  // Apply an unsaved (in-editor) spec to a file as a preview.
  app.post('/api/preview', (req, res) => {
    const { fileId, spec, specId, previousBatchId } = req.body || {};
    const file = store.getFile(fileId);
    if (!file) return fail(res, 404, '文件不存在。');
    if (!spec || !Array.isArray(spec.regions)) return fail(res, 400, '缺少映射方案。');
    const evaluation = evaluateSpec(file.doc, spec, { fileId: file.id, fileName: file.name, sha256: file.sha256 });
    // Compare against an explicit batch if given, otherwise against the newest
    // committed batch for this file, so a corrected mapping always shows a diff.
    const previous = previousBatchId
      ? store.getBatch(previousBatchId)
      : store.latestCommittedBatch(file.id);
    let diff = null;
    if (previous) {
      diff = compareBatches(store, previous.id, evaluation);
    }
    return json(res, 200, { evaluation, previousBatch: previous ? summarizeBatch(previous) : null, diff });
  });

  // ---------- batches ----------
  app.get('/api/batches', (req, res) => {
    const fileId = req.query.fileId ? String(req.query.fileId) : undefined;
    return json(res, 200, { batches: store.listBatches({ fileId }) });
  });

  app.get('/api/batches/:id', (req, res) => {
    const batch = store.getBatch(req.params.id);
    if (!batch) return fail(res, 404, '批次不存在。');
    return json(res, 200, {
      batch: summarizeBatch(batch),
      evaluation: batch.evaluation,
      records: store.getRecordsWithValues(batch.id),
    });
  });

  app.get('/api/batches/:id/records', (req, res) => {
    const batch = store.getBatch(req.params.id);
    if (!batch) return fail(res, 404, '批次不存在。');
    return json(res, 200, { records: store.getRecordsWithValues(batch.id) });
  });

  app.get('/api/batches/:id/csv', (req, res) => {
    const batch = store.getBatch(req.params.id);
    if (!batch) return fail(res, 404, '批次不存在。');
    const file = store.getFile(batch.fileId);
    const records = store.getRecordsWithValues(batch.id);
    const csv = toCsv(batch, records, file);
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(`${stripExt(file?.name || 'import')}-${batch.id.slice(0, 8)}.csv`)}`);
    return res.send('﻿' + csv);
  });

  app.post('/api/batches', (req, res) => {
    const { fileId, spec, specId, commit, previousBatchId } = req.body || {};
    const file = store.getFile(fileId);
    if (!file) return fail(res, 404, '文件不存在。');
    if (!spec || !Array.isArray(spec.regions) || !spec.regions.length) return fail(res, 400, '缺少映射方案。');
    const evaluation = evaluateSpec(file.doc, spec, { fileId: file.id, fileName: file.name, sha256: file.sha256 });

    // Duplicate-import guard: same file + same mapping commit == no second batch.
    const specHash = hashSpec(spec);
    if (commit) {
      const same = store.listBatches({ fileId: file.id }).find(
        (b) => b.status === 'committed' && b.specHash === specHash
      );
      if (same) {
        return json(res, 200, {
          duplicate: true,
          batch: summarizeBatch(store.getBatch(same.id)),
          message: '相同文件与相同映射已导入过，未生成第二批数据。',
        });
      }
    }
    if (commit && !evaluation.ok) {
      return json(res, 409, {
        committed: false,
        evaluation,
        message: '仍有待确认项或结构差异，处理后才能导入。',
        blocking: evaluation.blocking,
      });
    }
    const batch = store.createBatch({ fileId: file.id, specId: specId || null, spec, evaluation, commit: !!commit });
    let diff = null;
    // A corrected import is diffed against the batch it replaces by default.
    const compareWith = previousBatchId
      ? store.getBatch(previousBatchId)
      : (commit ? null : store.latestCommittedBatch(file.id));
    if (compareWith) diff = compareBatches(store, compareWith.id, evaluation);
    return json(res, 201, {
      duplicate: false,
      batch: summarizeBatch(batch),
      evaluation,
      regions: batch.regions,
      diff,
    });
  });

  app.get('/api/batches/:id/diff', (req, res) => {
    const batch = store.getBatch(req.params.id);
    if (!batch) return fail(res, 404, '批次不存在。');
    const other = store.getBatch(String(req.query.against || ''));
    if (!other) return fail(res, 400, '缺少对比批次 against。');
    return json(res, 200, { diff: compareBatches(store, other.id, batch.evaluation), from: summarizeBatch(other), to: summarizeBatch(batch) });
  });

  app.delete('/api/batches/:id', (req, res) => {
    const batch = store.getBatch(req.params.id);
    if (!batch) return fail(res, 404, '批次不存在。');
    if (batch.status !== 'preview') return fail(res, 409, '只有未导入的预览批次可以丢弃；已导入批次会保留。');
    store.deletePreview(batch.id);
    return json(res, 200, { ok: true });
  });

  // ---------- provenance lookup ----------
  app.get('/api/values/:id', (req, res) => {
    const row = store.db.prepare('SELECT * FROM batch_values WHERE id = ?').get(req.params.id);
    if (!row) return fail(res, 404, '取值不存在。');
    const batch = store.getBatch(row.batch_id);
    const file = store.getFile(row.file_id);
    return json(res, 200, {
      value: { id: row.id, fieldKey: row.field_key, fieldLabel: row.field_label, kind: row.kind, value: row.value_json === null ? null : JSON.parse(row.value_json), raw: row.raw, provenance: JSON.parse(row.provenance_json) },
      batch: batch ? summarizeBatch(batch) : null,
      file: file ? publicFile(file) : null,
    });
  });

  app.use('/samples', express.static(join(ROOT, 'samples'), { extensions: ['docx'] }));
  app.use(express.static(staticDir, { extensions: ['html'] }));
  app.get(/^(?!\/api\/).*/, (req, res) => res.sendFile(join(staticDir, 'index.html')));
  app.use((err, req, res, next) => {
    if (res.headersSent) return next(err);
    if (err?.type === 'entity.too.large') return fail(res, 413, '文件或请求体过大。');
    console.error(err);
    return fail(res, 500, `服务器错误：${err?.message || '未知错误'}`);
  });

  app.locals.store = store;
  return app;
}

function stripExt(name) {
  return String(name).replace(/\.docx$/i, '');
}

function publicFile(file) {
  return { id: file.id, name: file.name, sha256: file.sha256, size: file.size, uploadedAt: file.uploadedAt, title: file.doc?.title };
}

function summarizeBatch(batch) {
  return {
    id: batch.id,
    fileId: batch.fileId,
    specId: batch.specId,
    specHash: batch.specHash,
    status: batch.status,
    supersedesBatchId: batch.supersedesBatchId,
    replacedByBatchId: batch.replacedByBatchId,
    createdAt: batch.createdAt,
    committedAt: batch.committedAt,
    ok: batch.evaluation?.ok,
    totals: batch.evaluation?.totals || null,
    specName: batch.evaluation?.specName || '',
    spec: batch.spec,
    recordCount: (batch.evaluation?.regions || []).reduce((n, r) => n + (r.rows || []).length, 0),
  };
}

export function toCsv(batch, records, file) {
  // Field keys are only unique inside a region (c0..cN are derived from column
  // positions), so columns are qualified by region when there is more than one.
  const regions = (batch.evaluation?.regions || []).filter((r) => (r.rows || []).length || (r.values || []).length);
  const multi = regions.length > 1;
  const columns = [];
  for (const region of regions) {
    const seen = new Set();
    for (const v of region.values || []) {
      if (seen.has(v.fieldKey)) continue;
      seen.add(v.fieldKey);
      columns.push({
        regionKey: region.regionKey,
        key: v.fieldKey,
        label: multi ? `${v.header || v.fieldKey}（${region.regionName || region.regionKey}）` : (v.header || v.fieldKey),
      });
    }
  }
  const esc = (v) => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [];
  lines.push(['__region', '__row', '__table', ...columns.map((c) => c.label)].map(esc).join(','));
  for (const rec of records) {
    const row = [rec.regionKey, rec.rowIndex + 1, rec.tableId];
    for (const c of columns) {
      // A row only carries its own region's fields; other regions stay empty.
      row.push(c.regionKey === rec.regionKey ? (rec.record[c.key] ?? '') : '');
    }
    lines.push(row.map(esc).join(','));
  }
  return lines.join('\r\n');
}

/**
 * Human-readable diff between an old stored batch and a freshly evaluated spec.
 * Values are paired by (regionKey, tableId, rowIndex, fieldKey); row shifts are reported.
 */
export function compareBatches(store, previousBatchId, evaluation) {
  const prevBatch = store.getBatch(previousBatchId);
  if (!prevBatch) return null;
  const prevRecords = store.getRecordsWithValues(prevBatch.id);
  const prevEval = prevBatch.evaluation;
  const regions = [];
  const regionKeys = new Set([
    ...(prevEval.regions || []).map((r) => r.regionKey),
    ...(evaluation.regions || []).map((r) => r.regionKey),
  ]);
  for (const key of regionKeys) {
    const oldRegion = (prevEval.regions || []).find((r) => r.regionKey === key);
    const newRegion = (evaluation.regions || []).find((r) => r.regionKey === key);
    const oldRows = prevRecords.filter((r) => r.regionKey === key);
    const newRecords = newRegion?.rows || [];
    const oldByRow = new Map(oldRows.map((r) => [r.rowIndex, r]));
    const newByRow = new Map(newRecords.map((r) => [r.row, r]));
    const changes = [];
    const rowKeys = [...new Set([...oldByRow.keys(), ...newByRow.keys()])].sort((a, b) => a - b);
    for (const row of rowKeys) {
      const o = oldByRow.get(row);
      const n = newByRow.get(row);
      if (o && !n) {
        changes.push({ type: 'row-removed', row, message: `第 ${row + 1} 行不再导入。`, before: o.record });
        continue;
      }
      if (!o && n) {
        changes.push({ type: 'row-added', row, message: `新增第 ${row + 1} 行。`, after: n.record });
        continue;
      }
      const fieldKeys = new Set([...Object.keys(o.record), ...Object.keys(n.record)]);
      for (const fk of fieldKeys) {
        const before = o.record[fk];
        const after = n.record[fk];
        if (JSON.stringify(before ?? null) === JSON.stringify(after ?? null)) continue;
        changes.push({
          type: 'value-changed',
          row,
          fieldKey: fk,
          fieldLabel: o.values.find((v) => v.fieldKey === fk)?.fieldLabel || n.provenance[fk]?.header || fk,
          before, after,
          beforeSource: o.values.find((v) => v.fieldKey === fk)?.provenance || null,
          afterSource: n.provenance?.[fk] || null,
          message: `第 ${row + 1} 行「${o.values.find((v) => v.fieldKey === fk)?.fieldLabel || fk}」：${fmt(before)} → ${fmt(after)}`,
        });
      }
    }
    const structural = [];
    if (oldRegion && newRegion) {
      for (const k of ['headerRow', 'headerRowCount', 'startRow', 'endRow', 'groupCol']) {
        if (JSON.stringify(oldRegion[k]) !== JSON.stringify(newRegion[k])) {
          structural.push({ key: k, before: oldRegion[k], after: newRegion[k] });
        }
      }
    }
    regions.push({
      regionKey: key,
      regionName: newRegion?.regionName || oldRegion?.regionName || key,
      status: !oldRegion ? 'added' : !newRegion ? 'removed' : 'compared',
      oldStats: oldRegion?.stats || null,
      newStats: newRegion?.stats || null,
      structural,
      changes,
      changeCount: changes.length,
    });
  }
  const totalChanges = regions.reduce((n, r) => n + r.changeCount, 0);
  return {
    previousBatch: summarizeBatch(prevBatch),
    totalChanges,
    regions,
    summary: totalChanges === 0 ? '与上一批次相比没有数值变化。' : `与上一批次相比有 ${totalChanges} 处变化。`,
  };
}

function fmt(v) {
  if (v === null || v === undefined) return '（空）';
  if (v === '') return '（空）';
  return String(v);
}

export function startServer({ port = Number(process.env.PORT) || 5173, dataDir } = {}) {
  const app = createApp({ dataDir });
  const server = app.listen(port, () => {
    const addr = server.address();
    console.log(`调查表整理工具已启动: http://localhost:${addr.port}`);
  });
  return { app, server, store: app.locals.store };
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  startServer({ dataDir: process.env.DATA_DIR });
}

export { existsSync, createHash };
