import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import vm from 'node:vm';

const source = await readFile(new URL('../chart-rules.js', import.meta.url), 'utf8');
const context = { globalThis: {} };
vm.runInNewContext(source, context);
const ChartRules = context.globalThis.ChartRules;

test('includes answered questions with no more than 15 unique values', () => {
  const columns = ['Question A', 'Question B', 'Question C'];
  const rows = Array.from({ length: 16 }, (_, index) => ({
    'Question A': index % 2 ? 'Yes' : 'No',
    'Question B': `Value ${index + 1}`,
    'Question C': index === 0 ? 'Only answer' : ''
  }));

  assert.deepEqual(
    Array.from(ChartRules.getEligibleChartColumns(rows, columns)),
    ['Question A']
  );
});

test('counts normalized response labels when applying the unique-value limit', () => {
  const rows = [
    { Question: '  YES ' },
    { Question: 'yes' },
    { Question: 'No' }
  ];

  assert.deepEqual(Array.from(ChartRules.getEligibleChartColumns(rows, ['Question'])), ['Question']);
  assert.equal(ChartRules.countUniqueAnswers(rows, 'Question'), 2);
});

test('splits semicolon-delimited multi-select responses into unique choices', () => {
  assert.deepEqual(
    Array.from(ChartRules.splitResponseLabels([' A; B ', 'B;C', 'Single'])),
    ['A', 'B', 'C', 'Single']
  );
  assert.equal(
    ChartRules.countUniqueAnswers([{ Question: 'A; B' }, { Question: 'B; C' }], 'Question'),
    3
  );
});

test('keeps single responses and ignores blank multi-select choices', () => {
  assert.deepEqual(Array.from(ChartRules.splitResponseLabels('One')), ['One']);
  assert.deepEqual(Array.from(ChartRules.splitResponseLabels(' A; ; B ')), ['A', 'B']);
  assert.deepEqual(Array.from(ChartRules.splitResponseLabels(' ; ')), []);
});

test('excludes questions with only one unique answer', () => {
  const rows = [
    { Question: 'Yes' },
    { Question: 'Yes' },
    { Question: '' }
  ];

  assert.deepEqual(Array.from(ChartRules.getEligibleChartColumns(rows, ['Question'])), []);
});

test('labels true binary zero-one responses as no and yes', () => {
  assert.equal(ChartRules.shouldUseYesNoLabels(['0', '1', '1']), true);
  assert.equal(ChartRules.getDisplayAnswerLabel('0', ['0', '1']), 'No');
  assert.equal(ChartRules.getDisplayAnswerLabel('1', ['0', '1']), 'Yes');
  assert.equal(ChartRules.getDisplayAnswerLabel('0', ['0', '1', '2']), '0');
  assert.equal(ChartRules.getDisplayAnswerLabel('0', ['0', '1'], false), '0');
  assert.equal(ChartRules.getDisplayAnswerLabel('0', [], true), 'No');
});

test('makes every non-metadata column selectable regardless of response count', () => {
  const rows = [
    { Single: 'Only answer', Open: 'One', Empty: '', 'Student ID': '1' },
    { Single: '', Open: 'Two', Empty: '', 'Student ID': '2' },
    { Single: '', Open: 'Three', Empty: '', 'Student ID': '3' }
  ];

  assert.deepEqual(
    Array.from(ChartRules.getSelectableChartColumns(Object.keys(rows[0]))),
    ['Single', 'Open', 'Empty']
  );
});

test('excludes empty and likely metadata columns from automatic charts', () => {
  const rows = [
    { 'Student ID': '1', Email: 'a@example.com', Timestamp: '2026-01-01', Name: 'A', Empty: '', Rating: 'Good' },
    { 'Student ID': '2', Email: 'b@example.com', Timestamp: '2026-01-02', Name: 'B', Empty: '', Rating: 'Great' }
  ];

  assert.deepEqual(
    Array.from(ChartRules.getEligibleChartColumns(rows, Object.keys(rows[0]))),
    ['Rating']
  );
});

test('returns only eligible columns that do not already have a chart', () => {
  const eligible = ['Question A', 'Question B', 'Question C'];
  const charts = [{ primaryColumn: 'Question A' }, { primaryColumn: 'Question C' }];

  assert.deepEqual(
    Array.from(ChartRules.getMissingChartColumns(eligible, charts)),
    ['Question B']
  );
});

test('returns selected chart columns in eligible-column order', () => {
  assert.deepEqual(
    ChartRules.getSelectedChartColumns(['Revenue', 'Units Sold', 'Margin'], new Set(['Margin', 'Revenue'])),
    ['Revenue', 'Margin']
  );
  assert.deepEqual(
    ChartRules.getSelectedChartColumns(['Revenue', 'Units Sold'], ['Units Sold']),
    ['Units Sold']
  );
});

test('uses supplied unique counts without rescanning rows', () => {
  const rows = [{ Question: 'Yes' }, { Question: 'No' }];
  const uniqueCounts = new Map([['Question', 99]]);

  assert.deepEqual(
    Array.from(ChartRules.getEligibleChartColumns(rows, ['Question'], { uniqueCounts })),
    []
  );
});

test('supports linked multi-select values when determining eligibility', () => {
  const rows = [
    { Strategies: ['Tutoring', 'Mentoring'] },
    { Strategies: ['Mentoring'] }
  ];

  assert.equal(ChartRules.countUniqueAnswers(rows, 'Strategies'), 2);
  assert.deepEqual(Array.from(ChartRules.getEligibleChartColumns(rows, ['Strategies'])), ['Strategies']);
});
