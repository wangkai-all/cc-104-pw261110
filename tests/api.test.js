import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createApp } from '../server/index.js';

const SAMPLES = join(dirname(fileURLToPath(import.meta.url)), '..', 'samples');
const sample = (name) => fs.readFileSync(join(SAMPLES, name));

let server;
let base;

test.before(async () => {
  const dataDir = fs.mkdtempSync(join(os.tmpdir(), 'survey-api-'));
  const app = createApp({ dataDir });
  await new Promise((resolve) => {
    server = app.listen(0, '127.0.0.1', resolve);
  });
  base = `http://127.0.0.1:${server.address().port}`;
});

test.after(() => new Promise((resolve) => server.close(resolve)));

async function api(path, { method = 'GET', body, raw, name, headers = {} } = {}) {
  const opts = { method, headers: { ...headers } };
  if (raw) {
    opts.body = raw;
    opts.headers['content-type'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    if (name) opts.headers['x-file-name'] = name;
  } else if (body !== undefined) {
    opts.body = JSON.stringify(body);
    opts.headers['content-type'] = 'application/json';
  }
  const res = await fetch(base + path, opts);
  const type = res.headers.get('content-type') || '';
  const payload = type.includes('application/json') ? await res.json() : await res.text();
  return { status: res.status, body: payload, headers: res.headers };
}

async function upload(name) {
  const res = await api('/api/files', { method: 'POST', raw: sample(name), name });
  assert.equal(res.status === 201 || res.status === 200, true, JSON.stringify(res.body));
  return res.body;
}

/** Build a suggestion and resolve nested tables so it is immediately importable. */
async function readySpec(fileId) {
  const sug = await api(`/api/files/${fileId}/suggest`, { method: 'POST', body: {} });
  const spec = sug.body.spec;
  for (const region of spec.regions) {
    const st = await api(`/api/files/${fileId}/suggest`, { method: 'POST', body: { tableOrdinal: region.tableOrdinal } });
    for (const id of st.body.analysis.nestedTableIds || []) region.nestedResolutions[id] = 'skip';
  }
  return spec;
}

let firstFile;
let firstSpec;
let firstBatch;

test('uploading the same bytes twice does not create a second file or batch', async () => {
  const a = await upload('sample_survey.docx');
  assert.equal(a.duplicate, false);
  assert.ok(a.file.id);
  firstFile = a.file;
  firstSpec = await readySpec(firstFile.id);

  const b = await upload('sample_survey.docx');
  assert.equal(b.duplicate, true);
  assert.equal(b.file.id, firstFile.id);
  assert.equal(b.batches.length, 0);

  const files = await api('/api/files');
  assert.equal(files.body.files.length, 1);
});

test('a preview batch is stored but not committed, and can be discarded', async () => {
  const preview = await api('/api/batches', {
    method: 'POST', body: { fileId: firstFile.id, spec: firstSpec, commit: false },
  });
  assert.equal(preview.status, 201);
  assert.equal(preview.body.batch.status, 'preview');
  assert.ok(preview.body.evaluation.totals.values > 0);

  const id = preview.body.batch.id;
  const del = await api(`/api/batches/${id}`, { method: 'DELETE' });
  assert.equal(del.status, 200);
  assert.equal((await api(`/api/batches/${id}`)).status, 404);
});

test('committing twice with the same spec returns a duplicate instead of a new batch', async () => {
  const commit = await api('/api/batches', {
    method: 'POST', body: { fileId: firstFile.id, spec: firstSpec, commit: true },
  });
  assert.equal(commit.status, 201);
  assert.equal(commit.body.batch.status, 'committed');
  firstBatch = commit.body.batch;

  const again = await api('/api/batches', {
    method: 'POST', body: { fileId: firstFile.id, spec: firstSpec, commit: true },
  });
  assert.equal(again.status, 200);
  assert.equal(again.body.duplicate, true);
  assert.equal(again.body.batch.id, firstBatch.id);

  const all = await api(`/api/batches?fileId=${firstFile.id}`);
  assert.equal(all.body.batches.filter((b) => b.status === 'committed').length, 1);
});

test('every stored value keeps a file/table/row/column trail back to the original', async () => {
  const detail = await api(`/api/batches/${firstBatch.id}`);
  assert.equal(detail.status, 200);
  const records = detail.body.records;
  assert.ok(records.length >= 7);
  const res = await api(`/api/files/${firstFile.id}/original`);
  assert.equal(res.status, 200);
  assert.equal(res.headers.get('content-type').includes('wordprocessingml'), true);

  let checked = 0;
  for (const record of records) {
    for (const value of record.values) {
      const p = value.provenance;
      assert.equal(p.fileId, firstFile.id);
      assert.equal(p.sha256, firstFile.sha256);
      assert.ok(p.tableId);
      assert.equal(Number.isInteger(p.row), true);
      assert.equal(Number.isInteger(p.col), true);
      assert.ok(p.header);
      const lookup = await api(`/api/values/${value.id}`);
      assert.equal(lookup.status, 200);
      assert.equal(lookup.body.value.provenance.tableId, p.tableId);
      assert.equal(lookup.body.file.id, firstFile.id);
      checked++;
    }
  }
  assert.ok(checked > 20);
});

test('a corrected mapping commits as a new batch while the old batch and file stay', async () => {
  const corrected = JSON.parse(JSON.stringify(firstSpec));
  // include an extra column and rename a field: a genuine mapping change
  const region = corrected.regions[0];
  const extra = region.fields.find((f) => f.include === false);
  if (extra) extra.include = true;
  else region.fields.push({ key: 'x1', label: '补充列', col: 1, role: 'value', type: 'text', include: true });
  region.name = `${region.name}（更正）`;

  const preview = await api('/api/preview', {
    method: 'POST', body: { fileId: firstFile.id, spec: corrected, previousBatchId: firstBatch.id },
  });
  assert.equal(preview.status, 200);
  assert.ok(preview.body.diff);
  assert.equal(preview.body.previousBatch.id, firstBatch.id);

  const commit = await api('/api/batches', {
    method: 'POST', body: { fileId: firstFile.id, spec: corrected, commit: true, previousBatchId: firstBatch.id },
  });
  assert.equal(commit.status, 201);
  assert.equal(commit.body.batch.status, 'committed');
  assert.notEqual(commit.body.batch.id, firstBatch.id);

  const old = await api(`/api/batches/${firstBatch.id}`);
  assert.equal(old.status, 200);
  assert.equal(old.body.batch.status, 'replaced');
  assert.equal(old.body.batch.replacedByBatchId, commit.body.batch.id);
  assert.ok(old.body.records.length > 0);

  const file = await api(`/api/files/${firstFile.id}`);
  assert.equal(file.status, 200);
  assert.equal(file.body.file.sha256, firstFile.sha256);
});

test('a saved spec applies cleanly to a structurally identical file', async () => {
  const saved = await api('/api/specs', {
    method: 'POST', body: { name: '村级供水模板', spec: firstSpec },
  });
  assert.equal(saved.status, 201);
  const specId = saved.body.id;

  const second = await upload('sample_survey_second.docx');
  assert.equal(second.duplicate, false);
  const apply = await api(`/api/specs/${specId}/apply`, {
    method: 'POST', body: { fileId: second.file.id },
  });
  assert.equal(apply.status, 200);
  assert.equal(apply.body.applied, true);
  assert.equal(apply.body.evaluation.ok, true);
  assert.equal(apply.body.evaluation.totals.rows, 7);
  const committed = await api(`/api/batches/${firstBatch.id}`);
  const oldNames = committed.body.records.map((r) => r.record.c1);
  const newNames = apply.body.evaluation.regions[0].rows.map((r) => r.record.c1);
  assert.notDeepEqual(newNames, oldNames);

  const listed = await api('/api/specs');
  assert.ok(listed.body.specs.some((s) => s.id === specId));
});

test('a changed structure is rejected with concrete diffs instead of importing wrong columns', async () => {
  const specs = await api('/api/specs');
  const specId = specs.body.specs.find((s) => s.name === '村级供水模板').id;
  const revised = await upload('sample_survey_revised.docx');
  assert.equal(revised.duplicate, false);

  const apply = await api(`/api/specs/${specId}/apply`, {
    method: 'POST', body: { fileId: revised.file.id },
  });
  assert.equal(apply.status, 409);
  assert.equal(apply.body.applied, false);
  assert.ok(apply.body.blocking.length > 0);
  const messages = apply.body.blocking.map((d) => d.message).join(' | ');
  assert.match(messages, /列|表头|水质评价/);
  // nothing was stored
  const batches = await api(`/api/batches?fileId=${revised.file.id}`);
  assert.equal(batches.body.batches.length, 0);
});

test('committing an evaluation with pending items is refused', async () => {
  const nestedFile = await upload('sample_nested.docx');
  assert.equal(nestedFile.duplicate, false);
  const sug = await api(`/api/files/${nestedFile.file.id}/suggest`, { method: 'POST', body: {} });
  // leave the nested table unresolved on purpose: it must block the import
  for (const region of sug.body.spec.regions) region.nestedResolutions = {};
  const res = await api('/api/batches', {
    method: 'POST', body: { fileId: nestedFile.file.id, spec: sug.body.spec, commit: true },
  });
  assert.equal(res.status, 409);
  assert.equal(res.body.committed, false);
  assert.ok(res.body.blocking.some((b) => b.type === 'nested-table'));

  // nothing committed, so the file still has no batch
  const batches = await api(`/api/batches?fileId=${nestedFile.file.id}`);
  assert.equal(batches.body.batches.length, 0);
});

test('CSV export includes the values and stays distinct for blank/无/zero', async () => {
  const csv = await api(`/api/batches/${firstBatch.id}/csv`);
  assert.equal(csv.status, 200);
  assert.match(csv.headers.get('content-type'), /text\/csv/);
  const text = csv.body.replace(/^﻿/, '');
  const lines = text.split('\r\n');
  assert.ok(lines.length >= 8);
  assert.match(lines[0], /__region/);
  assert.match(text, /甲镇/);
  // the "无" cell and the zero cells survive verbatim, not as empty strings
  assert.match(text, /(^|,)无(,|$)/m);
  assert.match(text, /(^|,)0(,|$)/m);
  // each region gets its own qualified columns; a region never fills another's
  assert.match(lines[0], /地区（数据区 1）/);
  assert.match(lines[0], /水源地名称（数据区 2）/);
  const r2 = lines.filter((l) => l.startsWith('r2,'));
  assert.ok(r2.length > 0);
  for (const line of r2) {
    assert.match(line, /一号井|三号泉|见单元格内嵌套表/);
  }
});

test('a preview diff highlights changes against the committed batch', async () => {
  const current = (await api(`/api/batches?fileId=${firstFile.id}`)).body.batches.find((b) => b.status === 'committed');
  const changed = JSON.parse(JSON.stringify(firstSpec));
  changed.regions[0].fields = changed.regions[0].fields.filter((f, i) => i !== 1);
  const preview = await api('/api/batches', {
    method: 'POST', body: { fileId: firstFile.id, spec: changed, commit: false, previousBatchId: current.id },
  });
  assert.equal(preview.status, 201);
  assert.ok(preview.body.diff);
  assert.ok(preview.body.diff.totalChanges > 0);
  assert.equal(preview.body.batch.status, 'preview');
});

test('a committed batch cannot be deleted', async () => {
  const committed = (await api(`/api/batches?fileId=${firstFile.id}`)).body.batches.find((b) => b.status === 'committed');
  const res = await api(`/api/batches/${committed.id}`, { method: 'DELETE' });
  assert.equal(res.status, 409);
  assert.equal((await api(`/api/batches/${committed.id}`)).status, 200);
});
