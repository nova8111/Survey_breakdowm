import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const coreContext = { module: { exports: {} }, exports: {}, console };
vm.runInNewContext(fs.readFileSync(new URL('../survey-core.js', import.meta.url), 'utf8'), coreContext, { filename: 'survey-core.js' });
const SurveyCore = coreContext.module.exports;
const workerContext = { module: { exports: {} }, exports: {}, console, globalThis: {} };
workerContext.globalThis.SurveyCore = SurveyCore;
vm.runInNewContext(fs.readFileSync(new URL('../report-worker.js', import.meta.url), 'utf8'), workerContext, { filename: 'report-worker.js' });
const { handle } = workerContext.module.exports;

test('normalizes curly quotes, whitespace, dates, and labels without losing readable spelling', () => {
  assert.equal(SurveyCore.key('  North\u2019s   Campus '), "north's campus");
  assert.equal(SurveyCore.text(new Date('2025-01-02T03:04:05.000Z')), '2025-01-02T03:04:05.000Z');
  assert.deepEqual([...SurveyCore.labels('North; north ;South')], ['North', 'South']);
  assert.deepEqual([...SurveyCore.labels('No Response')], ['No Response']);
});

test('headers are collision safe, including pre-existing numbered names', () => {
  assert.deepEqual(SurveyCore.uniqueHeaders(['A', 'A', 'A (2)', 'A']), ['A', 'A (2)', 'A (2) (2)', 'A (3)']);
  const reserved = SurveyCore.uniqueHeaders(['__sourceRowNumber', '__linkedSiteKey']);
  assert.deepEqual(reserved, ['__sourceRowNumber (2)', '__linkedSiteKey (2)']);
});

test('summary counts each respondent once per answer and tracks blanks', () => {
  const result = SurveyCore.summarize([{q:'North'}, {q:'north'}, {q:''}, {q:' ; '}], 'q');
  assert.equal(result.respondentTotal, 2);
  assert.equal(result.answeredTotal, 2);
  assert.equal(result.missingCount, 2);
  assert.equal(result.items[0].count, 2);
});

test('supports app merge groups, blank-free percentages, and case-insensitive counts', () => {
  const merged = SurveyCore.summarize([{q:'A'}, {q:'B'}], 'q', {merges:[{name:'Combined',sources:new Set(['A','B'])}]});
  assert.equal(merged.items[0].response, 'Combined');
  const pct = SurveyCore.summarize([{q:'Yes'},{q:''},{q:''},{q:''}], 'q');
  assert.equal(pct.items[0].rowPercent, 100);
  assert.equal(pct.items[0].nonBlankPercent, 100);
  const ratings = SurveyCore.summarize([{q:'North'}, {q:'north'}], 'q');
  assert.equal(ratings.items[0].count, 2);
});

test('missing response labels do not collide with literal response text', () => {
  const summary = SurveyCore.summarize([{q:'(No response)'},{q:''}], 'q', {includeBlanks:true});
  assert.equal(summary.items.find(item => item.response === '(No response)').count, 1);
  assert.equal(summary.items.find(item => item.response === '(No response) (2)').count, 1);
  const comparison = SurveyCore.comparison([{p:'(No response)',c:'X'},{p:'',c:'X'}], 'p', 'c', {includeBlanks:true});
  assert.ok(comparison.labels.includes('(No response)'));
  assert.ok(comparison.labels.includes('(No response) (2)'));
});

test('labels honor explicit delimiters and arrays remain atomic', () => {
  assert.deepEqual([...SurveyCore.labels('A|B', {delimiter:'pipe'})], ['A','B']);
  assert.deepEqual([...SurveyCore.labels(['A|B'], {delimiter:'pipe'})], ['A|B']);
  assert.deepEqual([...SurveyCore.labels('A|B', {delimiter:'none'})], ['A|B']);
});

test('capped comparison counts overlapping Other memberships once per respondent', () => {
  const result = SurveyCore.comparison([{p:'A;B',c:'X;Y'},{p:'C',c:'X'}], 'p', 'c', {topMode:1});
  const other = result.labels.find(v => v.startsWith('Other'));
  assert.equal(result.primaryRespondentTotals.get(other), 2);
  assert.equal(result.matrix.get(other).get('X'), 2);
});

test('capped comparison uses per-respondent unions for Other cells and totals', () => {
  const result = SurveyCore.comparison([
    {P:'Top',C:'X'}, {P:'Top',C:'X'}, {P:'Top',C:'X'}, {P:'A;B',C:'X;Y'}
  ], 'P', 'C', {topMode:'1'});
  const pOther = result.labels.find(v => v.startsWith('Other'));
  const cOther = result.compareLabels.find(v => v.startsWith('Other'));
  assert.equal(result.matrix.get(pOther).get('X'), 1);
  assert.equal(result.primaryRespondentTotals.get(pOther), 1);
  assert.equal(result.compareRespondentTotals.get(cOther), 1);
});

test('report breakdowns include every multiselect group membership', () => {
  const result = SurveyCore.reportSection([{q:'Yes',g:'A;B'},{q:'Yes',g:'B'}], 'Q', 'q', ['g'], {answerOrder:['Yes']});
  const labels = result.groups.map(g => g.label);
  assert.deepEqual([...labels], ['A','B','(No response)']);
  assert.equal(result.aoa.find(r => r[0] === 'Yes')[1], 1);
  assert.equal(result.aoa.find(r => r[0] === 'Yes')[3], 2);
});

test('report includes blank breakdown groups and null percentages when no answers exist', () => {
  const result = SurveyCore.reportSection([{q:'',g:''}], 'Q', 'q', ['g'], {answerOrder:['Yes']});
  assert.ok(result.groups.some(g => g.label === '(No response)'));
  const yes = result.aoa.find(row => row[0] === 'Yes');
  assert.equal(yes[2], null);
});

test('report keeps literal missing text separate from blank breakdown membership', () => {
  const result = SurveyCore.reportSection([
    {q:'Yes',g:'(No response)'},
    {q:'Yes',g:''}
  ], 'Q', 'q', ['g'], {answerOrder:['Yes']});
  const groupLabels = [...result.groups].map(group => group.label);
  assert.ok(groupLabels.includes('(No response)'));
  assert.ok(groupLabels.includes('(No response) (2)'));
  const row = result.aoa.find(item => item[0] === 'Yes');
  assert.equal(row[1], 1);
  assert.equal(row[3], 1);
});

test('summary merges after deduplication and preserves literal No Response', () => {
  const result = SurveyCore.summarize([{q:'A;B'}, {q:'A'}, {q:'No Response'}], 'q', {merges: [['A', 'A=2']]});
  assert.deepEqual([...result.labels], ['A=2', 'B', 'No Response']);
  assert.deepEqual([...result.values], [2, 1, 1]);
});

test('comparison uses respondent unions and a collision-free Other bucket', () => {
  const rows = [{p:'A', c:'X;Y'}, {p:'B', c:'X'}, {p:'C', c:'Y'}, {p:'Other', c:'X'}];
  const result = SurveyCore.comparison(rows, 'p', 'c', {topMode: 2});
  assert.ok(result.labels.some(v => v.startsWith('Other (')));
  assert.equal(result.respondentTotal, 4);
  assert.equal(result.primaryRespondentTotals.get('A'), 1);
  assert.equal(result.compareRespondentTotals.get('X'), 3);
});

test('report section keeps unobserved answer scales and emits typed group headers and metadata', () => {
  const result = SurveyCore.reportSection(
    [{q:'Yes', group:'North'}, {q:'No', group:'South'}, {q:'Yes', group:'North'}],
    'Question', 'q', ['group'], {answerOrder:['Yes','No','Maybe']}
  );
  assert.deepEqual(result.answers, ['Yes','No','Maybe']);
  assert.ok(result.aoa.some(row => row[0] === 'Answered'));
  assert.ok(result.rows.flat().some(cell => cell.type === 'group-header' && cell.colspan === 2));
  assert.equal(result.aoa.find(row => row[0] === 'Maybe')[1], 0);
});

test('report worker returns sections and progress does not alter result', () => {
  const output = handle({id:'r1', rows:[{q:'Yes'}], questions:[{column:'q',display:'Q',answerOrder:['Yes','No']}], breakdownColumns:[]});
  assert.equal(output.id, 'r1');
  assert.equal(output.sections.length, 1);
  assert.equal(output.sections[0].aoa.find(row => row[0] === 'No')[1], 0);
});
