import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';

const source = await readFile(new URL('../script.js', import.meta.url), 'utf8');
const context = vm.createContext({ console });
for (const name of ['survey-core.js', 'data-io.js', 'chart-rules.js', 'vendor/xlsx.full.min.js']) {
  vm.runInContext(await readFile(new URL(`../${name}`, import.meta.url), 'utf8'), context);
}
vm.runInContext(`const state = {responseDelimiter:'semicolon'}; const NO_RESPONSE='(No response)'; const COLORS=['green']; const normalizeForMatch=SurveyCore.key; const roundOne=n=>Math.round(n*10)/10;`, context);
for (const name of ['buildSingleColumnResult', 'buildComparisonResult', 'getResponseLabels', 'applyReportFilter']) {
  const start = source.indexOf(`  function ${name}(`);
  const end = source.indexOf('\n  function ', start + 1);
  assert.ok(start >= 0 && end > start);
  vm.runInContext(source.slice(start, end), context);
}

test('chart integration uses respondent denominators after case deduplication and merging', () => {
  const result = vm.runInContext(`buildSingleColumnResult([{q:'A;a;B'}, {q:''}], {primaryColumn:'q', delimiter:'semicolon', includeBlanks:true, merges:[{name:'Combined',sources:new Set(['A','B'])}]})`, context);
  assert.equal(result.respondentTotal, 2);
  assert.equal(result.items.find(item => item.response === 'Combined').count, 1);
  assert.equal(result.items.find(item => item.response === 'Combined').rowPercent, 50);
});

test('overlapping comparison choices cannot produce a misleading 100% stacked chart', () => {
  const result = vm.runInContext(`buildComparisonResult([{q:'A;B',g:'X;Y'}], {primaryColumn:'q',compareColumn:'g',delimiter:'semicolon',compareType:'stacked100',compareValueMode:'comparePercent'})`, context);
  assert.equal(result.overlapping, true);
  assert.equal(result.valueMode, 'comparePercent');
  assert.equal(result.datasets[0].data[0], 100);
});

test('report filter integration intersects columns and treats case variants consistently', () => {
  const rows = vm.runInContext(`applyReportFilter([{q:'A;B',g:'North'},{q:'B',g:'north'},{q:'A',g:'South'}], {filters:[{column:'q',values:new Set(['a'])},{column:'g',values:new Set(['north'])}]})`, context);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].g, 'North');
});

test('percentage Excel roundtrip preserves numeric base counts and literal text safely', () => {
  const result = vm.runInContext(`(() => {
    const section=SurveyCore.reportSection([{q:'=1+1',g:'North'},{q:'Yes',g:'South'}],'Question','q',['g']);
    const rows=DataIO.reportRows({aoa:section.aoa,includedRows:2,denominatorDescription:'Answered respondents'}, {mode:'percentage'});
    const sheet=XLSX.utils.aoa_to_sheet(rows); DataIO.configureReportSheet(sheet,rows);
    const book=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book,sheet,'Report');
    const parsed=XLSX.read(XLSX.write(book,{type:'array',bookType:'xlsx'}),{type:'array',cellNF:true});
    return {rows, sheet:parsed.Sheets.Report};
  })()`, context);
  const formulaRow = result.rows.findIndex(row => row[0] === '=1+1');
  assert.equal(result.sheet[`A${formulaRow+1}`].t, 's');
  assert.equal(result.sheet[`A${formulaRow+1}`].f, undefined);
  assert.equal(result.sheet[`B${formulaRow+1}`].z, '0.0%');
  const baseRow = result.rows.findIndex(row => row[0] === 'Answered');
  assert.equal(result.sheet[`B${baseRow+1}`].v, 1);
  assert.notEqual(result.sheet[`B${baseRow+1}`].z, '0.0%');
  assert.equal(result.sheet[`C${baseRow+1}`].v, 1);
});

