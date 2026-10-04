import test from 'node:test';
import assert from 'node:assert/strict';
import '../data-io.js';
const { parseCsv, csvEscape, toCsv, reportRows, buildExportRows, configureReportSheet } = globalThis.DataIO;

test('parses quoted multiline CSV and auto detects delimiters', () => {
  assert.deepEqual(parseCsv('Name\tNote\nAda\t"hello\nworld"', { delimiter: 'auto' }), [
    ['Name', 'Note'], ['Ada', 'hello\nworld']
  ]);
});

test('keeps empty CSV records and decodes explicit UTF-16 without a BOM', () => {
  const bytes = new TextEncoder().encode('x,y\n\n1,2');
  const utf16 = new Uint8Array('x,y\n1,2'.length * 2);
  [...'x,y\n1,2'].forEach((char, index) => { utf16[index * 2] = char.charCodeAt(0); });
  assert.deepEqual(parseCsv(bytes), [['x', 'y'], [''], ['1', '2']]);
  assert.deepEqual(parseCsv(utf16, { encoding: 'utf-16le' }), [['x', 'y'], ['1', '2']]);
});

test('auto decodes BOMless UTF-16 and falls back to Windows-1252 for invalid UTF-8', () => {
  const utf16 = new Uint8Array('name\ncafé'.length * 2);
  [...'name\ncafé'].forEach((char, index) => { utf16[index * 2] = char.charCodeAt(0); });
  assert.deepEqual(parseCsv(utf16), [['name'], ['café']]);
  assert.deepEqual(parseCsv(Uint8Array.from([0x6e, 0x61, 0x6d, 0x65, 0x0a, 0x63, 0x61, 0x66, 0xe9])), [['name'], ['café']]);
});

test('rejects malformed quoted fields with useful errors', () => {
  assert.throws(() => parseCsv('a,b\n"broken,b'), /CSV.*(quote|closed)/i);
});

test('escapes formula-like text but preserves numeric values', () => {
  assert.equal(csvEscape(' =SUM(A1:A2)'), "' =SUM(A1:A2)");
  assert.equal(csvEscape(-12), '-12');
  assert.match(toCsv([['x', -12], ['ümlaut', 'line\nvalue']]), /^\uFEFF/);
});

test('exports report metadata and count/percentage modes', () => {
  const rows = buildExportRows({ title: 'Report', sourceName: 'Book', sheetName: 'Data', filterLabel: 'All', breakdownLabel: 'Site', generatedAt: 'now', aoa: [['', 'Count', 'Percentage'], ['A', 2, 0.5]] }, 'both');
  assert.deepEqual(rows.slice(-2), [['', 'Count', 'Percentage'], ['A', 2, 0.5]]);
  assert.deepEqual(reportRows({ ...rows, aoa: [['', 'Count', 'Percentage'], ['A', 2, 0.5]] }, { mode: 'percentage' }).at(-1), ['A', 0.5]);
});

test('finds report headers after long metadata and preserves base rows', () => {
  const aoa = [['Report', 'R'], ['Source', 'S'], ['Answered', 10], ['Eligible', 12], [''], [''], [''], [''], [''], [''], ['', 'Count', 'Percentage', 'Count', 'Percentage'], ['A', 2, .2, 4, .4], ['Answered', 10, '', 12, '']];
  const percentage = reportRows({ aoa }, { mode: 'percentage' });
  assert.deepEqual(percentage.slice(0, 2), [['Report', 'R'], ['Source', 'S']]);
  assert.deepEqual(percentage.find(row => row[0] === 'A'), ['A', .2, .4]);
  assert.deepEqual(percentage.at(-1), ['Answered', 10, 12]);
});

test('formats only active report percentage columns, excluding metadata and linked counts', () => {
  const rows = [
    ['Answered', 10, 12], ['Eligible', 12, 14],
    ['', 'Count', 'Percentage', 'Count', 'Percentage'], ['A', 2, .2, 4, .4],
    ['Linked category summary'], ['Category', 'Matched sites', 'Survey responses'], ['North', 3, 9]
  ];
  const sheet = {};
  rows.forEach((row, r) => row.forEach((value, c) => {
    let n = c + 1; let col = '';
    while (n) { col = String.fromCharCode(65 + ((n - 1) % 26)) + col; n = Math.floor((n - 1) / 26); }
    sheet[`${col}${r + 1}`] = { v: value, t: typeof value === 'number' ? 'n' : 's' };
  }));
  configureReportSheet(sheet, rows);
  assert.equal(sheet.B1.z, undefined);
  assert.equal(sheet.C4.z, '0.0%');
  assert.equal(sheet.E4.z, '0.0%');
  assert.equal(sheet.C7.z, undefined);
});

test('can render report percentages as text for CSV without changing counts', () => {
  const rows = reportRows({ aoa: [['', 'Count', 'Percentage'], ['A', 2, .5], ['Answered', 4, '']] }, { mode: 'percentage', percentText: true });
  assert.deepEqual(rows, [['', 'Percentage'], ['A', '50%'], ['Answered', 4]]);
});
