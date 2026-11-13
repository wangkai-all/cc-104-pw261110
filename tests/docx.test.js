import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseDocx } from '../server/docx.js';
import { buildDocx } from './helpers/docx-builder.js';

const SAMPLES = join(dirname(fileURLToPath(import.meta.url)), '..', 'samples');

test('parses merged headers, vertical merges and notes in the sample survey', () => {
  const doc = parseDocx(fs.readFileSync(join(SAMPLES, 'sample_survey.docx')));
  const t1 = doc.tables.find((t) => t.id === 't1');
  assert.equal(t1.colCount, 9);
  assert.equal(t1.rowCount, 10);

  // two-row header with 供水情况 spanning two columns
  assert.equal(t1.rows[1].cells.find((c) => c.text === '供水情况').colSpan, 2);
  assert.equal(t1.rows[1].cells.find((c) => c.text === '地区').rowSpan, 2);

  // 甲镇 covers rows 3..6 including the subtotal row; 乙乡 covers 7..8
  const jia = t1.rows[3].cells.find((c) => c.text === '甲镇');
  assert.equal(jia.rowSpan, 4);
  // Rows covered by the merge carry no separate cell of their own.
  assert.equal(t1.rows[4].cells.some((c) => c.col === 0), false);
  const yi = t1.rows[7].cells.find((c) => c.text === '乙乡');
  assert.equal(yi.rowSpan, 2);

  // a nested table lives in t2 and is not mixed into the host cell text
  const t2 = doc.tables.find((t) => t.id === 't2');
  assert.equal(t2.rowCount, 4);
  const nested = doc.tables.find((t) => t.parentTableId === 't2');
  assert.ok(nested);
  const hostCell = t2.rows[1].cells.find((c) => c.hasNestedTable);
  assert.equal(hostCell.text, '见单元格内嵌套表');
  assert.doesNotMatch(hostCell.text, /浊度/);

  // the in-table note is captured separately, never as a data value
  assert.equal(t1.notes.length, 1);
  assert.match(t1.notes[0].text, /^注：/);
  assert.equal(t1.notes[0].text.includes('浊度'), false);
});

test('table title comes from a caption paragraph or a merged title row', () => {
  const doc = parseDocx(fs.readFileSync(join(SAMPLES, 'sample_survey.docx')));
  const t1 = doc.tables.find((t) => t.id === 't1');
  assert.match(t1.title, /各村供水基本情况/);
  const t2 = doc.tables.find((t) => t.id === 't2');
  assert.match(t2.title, /水源地明细/);
});

test('repeated header rows are flagged from w:tblHeader', () => {
  const doc = parseDocx(buildDocx([
    { table: [
      { cells: ['地区', '数值'], repeatHeader: true },
      ['甲镇', '1'],
      { cells: ['地区', '数值'], repeatHeader: true },
      ['乙乡', '2'],
    ] },
  ]));
  const rows = doc.tables[0].rows;
  assert.equal(rows[0].repeatHeader, true);
  assert.equal(rows[1].repeatHeader, false);
  assert.equal(rows[2].repeatHeader, true);
});

test('superscript footnote markers are stripped from values', () => {
  const doc2 = parseDocx(buildDocx([
    { table: [
      { cells: ['项目', '数值'] },
      [{ paragraphs: ['A①区'] }, '3'],
    ] },
  ]));
  assert.equal(doc2.tables[0].rows[1].cells[0].text, 'A①区');
});

test('parsing is stable for the same bytes (content hash)', () => {
  const buf = buildDocx([{ table: [['a', 'b'], ['1', '2']] }]);
  const a = parseDocx(buf);
  const b = parseDocx(buf);
  assert.equal(a.sha256, b.sha256);
  assert.notEqual(a.sha256, parseDocx(buildDocx([{ table: [['a', 'b'], ['1', '3']] }])).sha256);
});
