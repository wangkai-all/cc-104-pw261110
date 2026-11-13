/**
 * SQLite-backed persistence: uploaded files, mapping specs, batches, records, values.
 * Original DOCX bytes are kept on disk so every imported value can be traced back.
 */
import { DatabaseSync } from 'node:sqlite';
import { mkdirSync, writeFileSync, existsSync, readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { randomUUID, createHash } from 'node:crypto';

export class Store {
  constructor(dataDir) {
    this.dataDir = dataDir;
    this.uploadDir = join(dataDir, 'uploads');
    mkdirSync(this.uploadDir, { recursive: true });
    mkdirSync(join(dataDir, 'db'), { recursive: true });
    this.db = new DatabaseSync(join(dataDir, 'db', 'app.sqlite'));
    this.db.exec('PRAGMA journal_mode = WAL; PRAGMA foreign_keys = ON;');
    this.migrate();
  }

  migrate() {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS files (
        id TEXT PRIMARY KEY,
        sha256 TEXT NOT NULL UNIQUE,
        name TEXT NOT NULL,
        size INTEGER NOT NULL,
        uploaded_at TEXT NOT NULL,
        doc_json TEXT NOT NULL,
        path TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS specs (
        id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        spec_json TEXT NOT NULL,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS batches (
        id TEXT PRIMARY KEY,
        file_id TEXT NOT NULL REFERENCES files(id),
        spec_id TEXT,
        spec_json TEXT NOT NULL,
        spec_hash TEXT NOT NULL,
        status TEXT NOT NULL,
        supersedes_batch_id TEXT,
        replaced_by_batch_id TEXT,
        evaluation_json TEXT NOT NULL,
        created_at TEXT NOT NULL,
        committed_at TEXT
      );
      CREATE TABLE IF NOT EXISTS batch_records (
        id TEXT PRIMARY KEY,
        batch_id TEXT NOT NULL REFERENCES batches(id),
        file_id TEXT NOT NULL,
        region_key TEXT NOT NULL,
        table_id TEXT NOT NULL,
        row_index INTEGER NOT NULL,
        record_json TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS batch_values (
        id TEXT PRIMARY KEY,
        batch_id TEXT NOT NULL REFERENCES batches(id),
        record_id TEXT NOT NULL REFERENCES batch_records(id),
        field_key TEXT NOT NULL,
        field_label TEXT NOT NULL,
        kind TEXT NOT NULL,
        value_json TEXT,
        raw TEXT,
        file_id TEXT NOT NULL,
        table_id TEXT NOT NULL,
        row_index INTEGER NOT NULL,
        col_index INTEGER NOT NULL,
        source_row INTEGER NOT NULL,
        source_col INTEGER NOT NULL,
        inherited INTEGER NOT NULL DEFAULT 0,
        provenance_json TEXT NOT NULL
      );
      CREATE INDEX IF NOT EXISTS idx_values_batch ON batch_values(batch_id);
      CREATE INDEX IF NOT EXISTS idx_values_record ON batch_values(record_id);
      CREATE INDEX IF NOT EXISTS idx_records_batch ON batch_records(batch_id);
      CREATE INDEX IF NOT EXISTS idx_batches_file ON batches(file_id);
    `);
  }

  now() {
    return new Date().toISOString();
  }

  // ---- files ----
  findFileBySha(sha256) {
    const row = this.db.prepare('SELECT * FROM files WHERE sha256 = ?').get(sha256);
    return row ? this.hydrateFile(row) : null;
  }

  hydrateFile(row) {
    return {
      id: row.id,
      sha256: row.sha256,
      name: row.name,
      size: row.size,
      uploadedAt: row.uploaded_at,
      doc: JSON.parse(row.doc_json),
      path: row.path,
    };
  }

  listFiles() {
    return this.db.prepare('SELECT id, sha256, name, size, uploaded_at FROM files ORDER BY uploaded_at DESC').all()
      .map((r) => ({ id: r.id, sha256: r.sha256, name: r.name, size: r.size, uploadedAt: r.uploaded_at }));
  }

  getFile(id) {
    const row = this.db.prepare('SELECT * FROM files WHERE id = ?').get(id);
    return row ? this.hydrateFile(row) : null;
  }

  saveFile({ sha256, name, size, buffer, doc }) {
    const id = randomUUID();
    const path = join(this.uploadDir, `${sha256}.docx`);
    if (!existsSync(path)) writeFileSync(path, buffer);
    this.db.prepare(
      'INSERT INTO files (id, sha256, name, size, uploaded_at, doc_json, path) VALUES (?, ?, ?, ?, ?, ?, ?)'
    ).run(id, sha256, name, size, this.now(), JSON.stringify(doc), path);
    return this.getFile(id);
  }

  readOriginal(fileId) {
    const file = this.getFile(fileId);
    if (!file) return null;
    if (!existsSync(file.path)) return null;
    return { file, buffer: readFileSync(file.path) };
  }

  // ---- specs ----
  listSpecs() {
    return this.db.prepare('SELECT * FROM specs ORDER BY updated_at DESC').all().map((r) => ({
      id: r.id, name: r.name, spec: JSON.parse(r.spec_json), createdAt: r.created_at, updatedAt: r.updated_at,
    }));
  }

  getSpec(id) {
    const r = this.db.prepare('SELECT * FROM specs WHERE id = ?').get(id);
    if (!r) return null;
    return { id: r.id, name: r.name, spec: JSON.parse(r.spec_json), createdAt: r.created_at, updatedAt: r.updated_at };
  }

  saveSpec({ id, name, spec }) {
    const ts = this.now();
    const sid = id || randomUUID();
    const existing = id ? this.db.prepare('SELECT id FROM specs WHERE id = ?').get(id) : null;
    if (existing) {
      this.db.prepare('UPDATE specs SET name = ?, spec_json = ?, updated_at = ? WHERE id = ?')
        .run(name, JSON.stringify(spec), ts, sid);
    } else {
      this.db.prepare('INSERT INTO specs (id, name, spec_json, created_at, updated_at) VALUES (?, ?, ?, ?, ?)')
        .run(sid, name, JSON.stringify(spec), ts, ts);
    }
    return this.getSpec(sid);
  }

  // ---- batches ----
  listBatches({ fileId } = {}) {
    const sql = fileId
      ? 'SELECT * FROM batches WHERE file_id = ? ORDER BY created_at DESC'
      : 'SELECT * FROM batches ORDER BY created_at DESC';
    const rows = fileId ? this.db.prepare(sql).all(fileId) : this.db.prepare(sql).all();
    return rows.map((r) => this.hydrateBatchSummary(r));
  }

  hydrateBatchSummary(r) {
    const counts = this.db.prepare(
      'SELECT COUNT(*) AS records FROM batch_records WHERE batch_id = ?'
    ).get(r.id);
    const evalJson = JSON.parse(r.evaluation_json);
    return {
      id: r.id,
      fileId: r.file_id,
      specId: r.spec_id,
      specHash: r.spec_hash,
      status: r.status,
      supersedesBatchId: r.supersedes_batch_id,
      replacedByBatchId: r.replaced_by_batch_id,
      createdAt: r.created_at,
      committedAt: r.committed_at,
      totals: evalJson.totals || null,
      ok: evalJson.ok,
      blockingCount: (evalJson.blocking || []).length,
      recordCount: counts?.records || 0,
      specName: evalJson.specName || '',
    };
  }

  getBatch(id) {
    const r = this.db.prepare('SELECT * FROM batches WHERE id = ?').get(id);
    if (!r) return null;
    return {
      id: r.id,
      fileId: r.file_id,
      specId: r.spec_id,
      spec: JSON.parse(r.spec_json),
      specHash: r.spec_hash,
      status: r.status,
      supersedesBatchId: r.supersedes_batch_id,
      replacedByBatchId: r.replaced_by_batch_id,
      evaluation: JSON.parse(r.evaluation_json),
      createdAt: r.created_at,
      committedAt: r.committed_at,
    };
  }

  getBatchRecords(batchId) {
    const rows = this.db.prepare('SELECT * FROM batch_records WHERE batch_id = ? ORDER BY region_key, row_index').all(batchId);
    return rows.map((r) => ({
      id: r.id, batchId: r.batch_id, fileId: r.file_id, regionKey: r.region_key,
      tableId: r.table_id, rowIndex: r.row_index, record: JSON.parse(r.record_json),
    }));
  }

  getRecordsWithValues(batchId) {
    const records = this.getBatchRecords(batchId);
    const values = this.db.prepare('SELECT * FROM batch_values WHERE batch_id = ?').all(batchId);
    const byRecord = new Map();
    for (const v of values) {
      if (!byRecord.has(v.record_id)) byRecord.set(v.record_id, []);
      byRecord.get(v.record_id).push({
        id: v.id,
        fieldKey: v.field_key,
        fieldLabel: v.field_label,
        kind: v.kind,
        value: v.value_json === null ? null : JSON.parse(v.value_json),
        raw: v.raw,
        provenance: JSON.parse(v.provenance_json),
      });
    }
    return records.map((r) => ({ ...r, values: byRecord.get(r.id) || [] }));
  }

  /** The newest committed (or preview) batch for a file+region set, excluding one id. */
  previousBatchFor(fileId, excludeBatchId) {
    const row = this.db.prepare(
      `SELECT * FROM batches WHERE file_id = ? AND id != ? AND status IN ('committed','replaced','preview')
       ORDER BY created_at DESC LIMIT 1`
    ).get(fileId, excludeBatchId || '');
    return row ? this.getBatch(row.id) : null;
  }

  latestCommittedBatch(fileId) {
    const row = this.db.prepare(
      `SELECT * FROM batches WHERE file_id = ? AND status = 'committed' ORDER BY created_at DESC LIMIT 1`
    ).get(fileId);
    return row ? this.getBatch(row.id) : null;
  }

  /**
   * Persist an evaluated batch. A committed batch supersedes the previous committed
   * batch for the same file, but that old batch and its records stay intact.
   */
  createBatch({ fileId, specId, spec, evaluation, commit }) {
    const id = randomUUID();
    const ts = this.now();
    let supersedes = null;
    let replaceTarget = null;
    if (commit) {
      const prev = this.latestCommittedBatch(fileId);
      if (prev) {
        supersedes = prev.id;
        replaceTarget = prev.id;
      }
    }
    const specName = spec.name || `方案 ${specId || ''}`.trim();
    const evaluationWithName = { ...evaluation, specName };
    this.db.prepare(
      `INSERT INTO batches (id, file_id, spec_id, spec_json, spec_hash, status, supersedes_batch_id, replaced_by_batch_id, evaluation_json, created_at, committed_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, NULL, ?, ?, ?)`
    ).run(
      id, fileId, specId || null, JSON.stringify(spec), evaluation.specHash,
      commit ? 'committed' : 'preview', supersedes, JSON.stringify(evaluationWithName), ts, commit ? ts : null
    );

    const insertRecord = this.db.prepare(
      `INSERT INTO batch_records (id, batch_id, file_id, region_key, table_id, row_index, record_json)
       VALUES (?, ?, ?, ?, ?, ?, ?)`
    );
    const insertValue = this.db.prepare(
      `INSERT INTO batch_values (id, batch_id, record_id, field_key, field_label, kind, value_json, raw,
        file_id, table_id, row_index, col_index, source_row, source_col, inherited, provenance_json)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    );
    const regionStats = [];
    this.db.exec('BEGIN');
    try {
      for (const region of evaluation.regions || []) {
        const regionHasBlocking = (region.blocking || []).length > 0;
        for (const row of region.rows || []) {
          const recordId = randomUUID();
          insertRecord.run(recordId, id, fileId, region.regionKey, region.tableId, row.row, JSON.stringify(row.record));
          for (const [fieldKey, prov] of Object.entries(row.provenance)) {
            const label = prov.header || fieldKey;
            insertValue.run(
              randomUUID(), id, recordId, fieldKey, label, prov.kind,
              prov.value === undefined ? null : JSON.stringify(prov.value),
              prov.raw ?? null, fileId, prov.tableId, prov.valueRow ?? prov.row,
              prov.col ?? 0, prov.row ?? 0, prov.col ?? 0, prov.inherited ? 1 : 0,
              JSON.stringify(prov)
            );
          }
        }
        regionStats.push({
          regionKey: region.regionKey,
          regionName: region.regionName,
          tableId: region.tableId,
          rows: (region.rows || []).length,
          blocked: regionHasBlocking,
        });
      }
      if (replaceTarget) {
        this.db.prepare('UPDATE batches SET status = ?, replaced_by_batch_id = ? WHERE id = ?')
          .run('replaced', id, replaceTarget);
      }
      this.db.exec('COMMIT');
    } catch (err) {
      this.db.exec('ROLLBACK');
      throw err;
    }
    const batch = this.getBatch(id);
    batch.regions = regionStats;
    return batch;
  }

  deletePreview(batchId) {
    const batch = this.getBatch(batchId);
    if (!batch || batch.status !== 'preview') return false;
    this.db.exec('BEGIN');
    try {
      this.db.prepare('DELETE FROM batch_values WHERE batch_id = ?').run(batchId);
      this.db.prepare('DELETE FROM batch_records WHERE batch_id = ?').run(batchId);
      this.db.prepare('DELETE FROM batches WHERE id = ?').run(batchId);
      this.db.exec('COMMIT');
    } catch (e) {
      this.db.exec('ROLLBACK');
      throw e;
    }
    return true;
  }

  stats() {
    return {
      files: this.db.prepare('SELECT COUNT(*) c FROM files').get().c,
      specs: this.db.prepare('SELECT COUNT(*) c FROM specs').get().c,
      batches: this.db.prepare('SELECT COUNT(*) c FROM batches').get().c,
      committedBatches: this.db.prepare("SELECT COUNT(*) c FROM batches WHERE status='committed'").get().c,
      records: this.db.prepare('SELECT COUNT(*) c FROM batch_records').get().c,
      values: this.db.prepare('SELECT COUNT(*) c FROM batch_values').get().c,
    };
  }
}

export function shortHash(text) {
  return createHash('sha256').update(text).digest('hex').slice(0, 12);
}
