import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import vm from 'node:vm';

const source = await readFile(new URL('../data-dictionary.js', import.meta.url), 'utf8');
const context = { globalThis: {} };
vm.runInNewContext(source, context);
const DataDictionary = context.globalThis.DataDictionary;

test('maps trimmed sheet and header names to display questions per sheet', () => {
  const dictionary = DataDictionary.create([
    ['Sheet Name', 'Original/Header Name', 'Display Question'],
    ['  Staff Survey  ', '  Q1  ', '  How supported did you feel?  '],
    ['Parent Survey', 'Q1', 'How supported did your child feel?']
  ]);

  assert.equal(DataDictionary.getDisplayQuestion(dictionary, 'Staff Survey', 'Q1'), 'How supported did you feel?');
  assert.equal(DataDictionary.getDisplayQuestion(dictionary, 'Parent Survey', 'Q1'), 'How supported did your child feel?');
  assert.equal(DataDictionary.getDisplayQuestion(dictionary, 'Staff Survey', 'Missing'), '');
});

test('ignores blank questions and tolerates a dictionary without a header row', () => {
  const dictionary = DataDictionary.create([
    ['Survey', 'Q1', ''],
    ['Survey', 'Q2', '  Full question  '],
    ['', 'Q3', 'Should be ignored']
  ]);

  assert.equal(DataDictionary.getDisplayQuestion(dictionary, 'Survey', 'Q1'), '');
  assert.equal(DataDictionary.getDisplayQuestion(dictionary, ' Survey ', ' Q2 '), 'Full question');
  assert.equal(dictionary.hasEntries, true);
});

test('recognizes the optional data dictionary sheet without matching similarly named survey sheets', () => {
  assert.equal(DataDictionary.isDictionarySheetName('Data Dictionary'), true);
  assert.equal(DataDictionary.isDictionarySheetName(' data_dictionary '), true);
  assert.equal(DataDictionary.isDictionarySheetName('Dictionary'), false);
  assert.equal(DataDictionary.isDictionarySheetName('Survey Dictionary'), false);
});
