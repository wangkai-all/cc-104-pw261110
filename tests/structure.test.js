import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseDocx } from '../server/docx.js';
import { suggestSpec, evaluateSpec, interpretValue } from '../server/engine.js';
import { buildDocx } from './helpers/docx-builder.js';

const SAMPLES = join(dirname(fileURLToPath(import.meta.url)), '..', 'samples');
const survey = parseDocx(fs.readFileSync(join(SAMPLES, 'sample_survey.docx')));
const spec = suggestSpec(survey);
for (const r of spec.regions) r.nestedResolutions = { 't2.n1': 'skip' };

function valuesOf(regionKey = 'r1') {
  const region = spec.regions.find((r) => r.key === regionKey);
  const ev = evaluateSpec(survey, { ...spec, regions: [region] });
  return ev.regions[0];
}

test('two-row header fields are flattened to leaf labels', () => {
  const r1 = spec.regions.find((r) => r.key === 'r1');
  assert.equal(r1.headerRow, 1);
  assert.equal(r1.headerRowCount, 2);
  assert.deepEqual(r1.fields.map((f) => f.label), ['地区', '村名', '供水人口', '检测频次', '水源类型', '水质', '数值', '计量单位', '备注']);
});

test('blank, 无 and zero are preserved as distinct kinds', () => {
  const region = valuesOf();
  const kinds = (field) => region.values.filter((v) => v.fieldKey === field).map((v) => [v.row, v.kind, v.raw]);
  // R5 供水人口 is blank (西河村)
  assert.deepEqual(kinds('c2').find(([row]) => row === 4), [4, 'blank', '']);
  // R4 备注 = 无
  assert.deepEqual(kinds('c8').find(([row]) => row === 3), [3, 'none', '无']);
  // R7 供水人口 = 0 and 年供水量 = 0
  assert.deepEqual(kinds('c2').find(([row]) => row === 6), [6, 'zero', '0']);
  assert.deepEqual(kinds('c6').find(([row]) => row === 6), [6, 'zero', '0']);
  assert.equal(region.stats.blank, 4);
  assert.equal(region.stats.none, 2);
  assert.equal(region.stats.zero, 2);
});

test('interpretValue keeps text "无" out of numeric zero handling', () => {
  assert.equal(interpretValue('无', { type: 'number' }).kind, 'none');
  assert.equal(interpretValue('0', { type: 'number' }).kind, 'zero');
  assert.equal(interpretValue('0.0', { type: 'number' }).kind, 'zero');
  assert.equal(interpretValue('', { type: 'number' }).kind, 'blank');
  assert.equal(interpretValue('  ', { type: 'number' }).kind, 'blank');
  assert.equal(interpretValue('12.5', { type: 'number' }).value, 12.5);
  assert.equal(interpretValue('暂无', { type: 'number' }).invalidNumber, true);
  const textZero = interpretValue('0', { type: 'text' });
  assert.equal(textZero.kind, 'zero');
  assert.equal(textZero.value, '0');
});

test('group names inherit downward only inside one data segment', () => {
  const region = valuesOf();
  const group = (row) => region.values.find((v) => v.fieldKey === 'c0' && v.row === row);
  // explicit owners
  assert.equal(group(3).value, '甲镇');
  assert.equal(group(3).inherited, false);
  // merged rows below the owner are marked inherited, pointing at the origin row
  assert.equal(group(4).inherited, true);
  assert.equal(group(4).row, 4);
  assert.equal(group(4).inheritedFromRow, 3);
  assert.equal(group(4).inheritKind, 'inherited');
  assert.equal(group(6).inherited, true);
  assert.equal(group(6).inheritedFromRow, 3);
  // a later block has its own explicit owner
  assert.equal(group(7).value, '乙乡');
  assert.equal(group(7).inherited, false);
  assert.equal(group(8).inherited, true);
  assert.equal(group(8).inheritedFromRow, 7);
});

test('subtotal, note and repeated-header rows are excluded from data', () => {
  const region = valuesOf();
  const rows = region.rows.map((r) => r.row);
  assert.deepEqual(rows, [3, 4, 6, 7, 8]);
  assert.equal(region.values.some((v) => v.raw.includes('小计')), false);
  assert.equal(region.values.some((v) => v.raw.startsWith('注：')), false);
});

test('footnote text never appears in imported values', () => {
  const region = valuesOf();
  for (const v of region.values) {
    assert.doesNotMatch(String(v.raw), /①②③|空白表示未填报/);
  }
});

test('nested table blocks import until resolved, then only cell text is taken', () => {
  const fresh = suggestSpec(survey);
  const r2 = fresh.regions.find((r) => r.key === 'r2');
  const unresolved = evaluateSpec(survey, { ...fresh, regions: [r2] });
  assert.equal(unresolved.ok, false);
  assert.ok(unresolved.blocking.some((b) => b.type === 'nested-table'));
  const resolved = evaluateSpec(survey, {
    ...fresh,
    regions: [{ ...r2, nestedResolutions: { 't2.n1': 'skip' } }],
  });
  assert.equal(resolved.ok, true);
  const cell = resolved.regions[0].values.find((v) => v.fieldKey === 'c3' && v.row === 1);
  assert.equal(cell.value, '见单元格内嵌套表');
  assert.doesNotMatch(cell.value, /浊度/);
});

test('structure change is reported and never silently applied', () => {
  const revised = parseDocx(fs.readFileSync(join(SAMPLES, 'sample_survey_revised.docx')));
  const ev = evaluateSpec(revised, spec);
  assert.equal(ev.ok, false);
  const codes = ev.diff.blocking.map((d) => d.code);
  assert.ok(codes.includes('column-count-changed'));
  assert.ok(codes.includes('header-label-changed'));
  // the renamed header is reported as a rename, not a cascade of shifted columns
  assert.ok(ev.diff.blocking.some((d) => d.message.includes('水质评价')));
  assert.ok(ev.diff.diffs.some((d) => d.code === 'table-moved'));
});

test('a structurally identical file applies with no blocking diffs', () => {
  const second = parseDocx(fs.readFileSync(join(SAMPLES, 'sample_survey_second.docx')));
  const ev = evaluateSpec(second, spec);
  assert.equal(ev.ok, true);
  assert.equal(ev.diff.blocking.length, 0);
  assert.equal(ev.totals.rows, 7);
  // different data, same shape
  const first = evaluateSpec(survey, spec);
  assert.notEqual(JSON.stringify(first.regions[0].values.map((v) => v.value)), JSON.stringify(ev.regions[0].values.map((v) => v.value)));
});

test('reordered tables are matched by header text, not ordinal', () => {
  const reordered = parseDocx(buildDocx([
    { table: [['名称', '状态'], ['一号井', '在用']] },
    { table: [['地区', '村名', '人口'], ['甲镇', '东河村', '1200']] },
  ]));
  const earlier = parseDocx(buildDocx([
    { table: [['地区', '村名', '人口'], ['甲镇', '东河村', '1200']] },
    { table: [['名称', '状态'], ['一号井', '在用']] },
  ]));
  const s = suggestSpec(earlier);
  const ev = evaluateSpec(reordered, s);
  assert.ok(ev.diff.diffs.some((d) => d.code === 'table-moved'));
  const moved = ev.regions.find((r) => r.regionKey === 'r1');
  assert.equal(moved.values.find((v) => v.fieldKey === 'c1').value, '东河村');
});

test('a blank group cell with nothing above raises a blocking item that a manual value can fix', () => {
  const doc = parseDocx(buildDocx([
    { table: [
      { cells: ['地区', '村名'] },
      ['', '东河村'],
      ['甲镇', '西河村'],
      ['', '南湾村'],
    ] },
  ]));
  const s = suggestSpec(doc);
  const ev = evaluateSpec(doc, s);
  assert.equal(ev.ok, false);
  const gap = ev.blocking.find((b) => b.type === 'group-gap');
  assert.ok(gap);
  assert.equal(gap.row, 1);
  // the second blank row sits inside a segment that already has an owner, so it
  // legitimately inherits 甲镇 instead of raising another gap
  assert.equal(ev.regions[0].values.find((v) => v.fieldKey === 'c0' && v.row === 3).value, '甲镇');

  // filling the gap by hand resolves it without inventing inheritance across it
  const fixed = JSON.parse(JSON.stringify(s));
  fixed.regions[0].groupOverrides = { [String(gap.row)]: '丙村' };
  const ev2 = evaluateSpec(doc, fixed);
  assert.equal(ev2.ok, true);
  assert.equal(ev2.regions[0].values.find((v) => v.fieldKey === 'c0' && v.row === gap.row).value, '丙村');
});

test('inheritance does not cross a subtotal row', () => {
  const doc = parseDocx(buildDocx([
    { table: [
      { cells: ['地区', '村名', '数值'] },
      [{ text: '甲镇', vMerge: 'restart' }, '东河村', '1'],
      ['', '小计', '1'],
      ['', '西河村', '2'],
    ] },
  ]));
  // Build a region where the group cell is genuinely blank on the last row:
  // the subtotal row splits the segment, so the blank group cannot carry across.
  const ev = evaluateSpec(doc, {
    regions: [{
      key: 'r1', name: '数据区 1', tableOrdinal: 0,
      headerRow: 0, headerRowCount: 1, startRow: 1, endRow: 3,
      groupCol: 0, inheritGroup: true,
      headerSignature: [{ col: 0, text: '地区', colSpan: 1 }, { col: 1, text: '村名', colSpan: 1 }, { col: 2, text: '数值', colSpan: 1 }],
      colCount: 3,
      fields: [
        { key: 'g', label: '地区', col: 0, role: 'group', type: 'text', include: true },
        { key: 'v', label: '村名', col: 1, role: 'value', type: 'text', include: true },
      ],
      groupOverrides: {}, nestedResolutions: {},
    }],
  });
  assert.equal(ev.ok, false);
  assert.ok(ev.blocking.some((b) => b.type === 'group-gap' && b.row === 3));
});
