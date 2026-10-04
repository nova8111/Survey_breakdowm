import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import vm from 'node:vm';

const source = await readFile(new URL('../linked-survey.js', import.meta.url), 'utf8');
const context = { globalThis: {} };
vm.runInNewContext(source, context);
const LinkedSurvey = context.globalThis.LinkedSurvey;

test('matches shared values despite capitalization, spaces, and punctuation', () => {
  const result = LinkedSurvey.analyzeLink(
    [{ Site: ' Site-A ' }, { Site: 'SITE B' }, { Site: 'Café C' }],
    [{ Location: 'site a' }, { Location: 'site-b' }, { Location: 'Cafe C' }],
    'Site',
    'Location'
  );
  assert.equal(result.stats.matchedRows, 3);
  assert.equal(result.stats.matchRate, 100);
});

test('reports partial and missing matches', () => {
  const result = LinkedSurvey.analyzeLink(
    [{ Site: 'A' }, { Site: 'B' }, { Site: '' }],
    [{ Site: 'A' }],
    'Site',
    'Site'
  );
  assert.equal(result.stats.matchedRows, 1);
  assert.equal(result.stats.unmatchedRows, 2);
  assert.deepEqual(Array.from(result.unmatched, item => item.reason), ['No matching secondary record', 'Missing primary matching value']);
});

test('handles no matches without failing silently', () => {
  const result = LinkedSurvey.analyzeLink([{ Site: 'A' }], [{ Site: 'B' }], 'Site', 'Site');
  assert.equal(result.stats.matchedRows, 0);
  assert.equal(result.stats.unmatchedRate, 100);
});

test('detects duplicates and excludes ambiguous secondary keys', () => {
  const result = LinkedSurvey.analyzeLink(
    [{ Site: 'A' }, { Site: 'A' }, { Site: 'B' }],
    [{ Site: 'A' }, { Site: 'A' }, { Site: 'B' }],
    'Site',
    'Site'
  );
  assert.equal(result.duplicates.primary[0].count, 2);
  assert.equal(result.duplicates.secondary[0].count, 2);
  assert.equal(result.stats.matchedRows, 1);
  assert.equal(result.unmatched[0].reason, 'Duplicate value in secondary survey');
});

test('enriches single-select secondary questions', () => {
  const result = LinkedSurvey.analyzeLink([{ Site: 'A' }], [{ Site: 'A', Type: 'Urban' }], 'Site', 'Site');
  const enriched = LinkedSurvey.enrichMatchedRows(result, 'Type');
  assert.deepEqual(Array.from(enriched.rows[0][enriched.column]), ['Urban']);
});

test('uses an optional display question for the linked analysis column', () => {
  const result = LinkedSurvey.analyzeLink([{ Site: 'A' }], [{ Site: 'A', Type: 'Urban' }], 'Site', 'Site');
  const enriched = LinkedSurvey.enrichMatchedRows(result, 'Type', { displayQuestion: 'Which setting applies?' });
  assert.equal(enriched.column, 'Linked survey: Which setting applies?');
  assert.deepEqual(Array.from(enriched.rows[0][enriched.column]), ['Urban']);
});

test('expands multi-select answers into every applicable category', () => {
  const primary = [{ Site: 'A', Answer: 'Yes' }, { Site: 'A', Answer: 'No' }];
  const secondary = [{ Site: 'A', Strategies: 'Strategy A; Strategy C' }];
  const link = LinkedSurvey.analyzeLink(primary, secondary, 'Site', 'Site');
  const enriched = LinkedSurvey.enrichMatchedRows(link, 'Strategies');
  const summary = LinkedSurvey.buildCategorySummary(enriched.rows, enriched.column);
  assert.deepEqual(Array.from(enriched.rows[0][enriched.column]), ['Strategy A', 'Strategy C']);
  assert.deepEqual(JSON.parse(JSON.stringify(summary)), [
    { category: 'Strategy A', matchedSites: 1, surveyResponses: 2 },
    { category: 'Strategy C', matchedSites: 1, surveyResponses: 2 }
  ]);
});

test('single-survey data remains unchanged when linking is not invoked', () => {
  const rows = [{ Site: 'A', Answer: 'Yes' }];
  assert.deepEqual(rows, [{ Site: 'A', Answer: 'Yes' }]);
});

test('validates missing matching fields and invalid secondary questions', () => {
  assert.throws(
    () => LinkedSurvey.analyzeLink([{ Site: 'A' }], [{ Site: 'A' }], '', 'Site'),
    /matching field/
  );
  const link = LinkedSurvey.analyzeLink([{ Site: 'A' }], [{ Site: 'A', Type: 'Urban' }], 'Site', 'Site');
  assert.throws(() => LinkedSurvey.enrichMatchedRows(link, 'Missing question'), /invalid/);
});

test('rejects a secondary question when any matched record lacks the selected field', () => {
  const link = {
    matched: [
      { primary: { Site: 'A' }, secondary: { Site: 'A', Segment: 'Urban' } },
      { primary: { Site: 'B' }, secondary: { Site: 'B' } }
    ]
  };

  assert.throws(() => LinkedSurvey.enrichMatchedRows(link, 'Segment'), /invalid/);
});
