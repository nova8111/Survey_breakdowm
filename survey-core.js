(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.SurveyCore = api;
})(typeof self !== 'undefined' ? self : typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const NO_RESPONSE = 'No Response';
  const clean = value => value == null ? '' : String(value).replace(/[\u2018\u2019]/g, "'").replace(/[\u201c\u201d]/g, '"').replace(/\s+/g, ' ').trim();
  function text(value) {
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? '' : value.toISOString();
    if (value && typeof value.toISOString === 'function' && value.constructor && value.constructor.name === 'Date') return value.toISOString();
    return clean(value);
  }
  function key(value) { return text(value).toLowerCase(); }
  function labels(value, options) {
    const delimiter = options && options.delimiter;
    const split = delimiter === 'comma' ? /,/ : delimiter === 'pipe' ? /\|/ : delimiter === 'newline' ? /\r?\n/ : delimiter === 'semicolon' || delimiter == null ? /;/ : null;
    const values = Array.isArray(value) ? value : (value == null ? [] : [value]);
    const out = [], seen = new Set();
    values.forEach(item => {
      if (item == null) return;
      const parts = Array.isArray(value) || !split ? [item] : String(item).split(split);
      parts.forEach(part => { const readable = text(part); const k = key(readable); if (k && !seen.has(k)) { seen.add(k); out.push(readable); } });
    });
    return out;
  }
  function uniqueHeaders(headers) {
    const used = new Set();
    ['__sourceRowNumber', '__linkedSiteKey'].forEach(name => used.add(key(name)));
    return (headers || []).map(raw => {
      const base = text(raw) || 'Column'; let candidate = base; let n = 2;
      while (used.has(key(candidate))) candidate = `${base} (${n++})`;
      used.add(key(candidate)); return candidate;
    });
  }
  function mergeMap(merges) {
    const map = new Map();
    const add = (from, to) => { const f = key(from), t = text(to); if (f && t) map.set(f, t); };
    if (merges instanceof Map) merges.forEach((v, k) => add(k, v));
    else if (Array.isArray(merges)) merges.forEach(m => {
      if (Array.isArray(m)) return add(m[0], m[1]);
      if (!m) return;
      const target = m.to ?? m.target ?? m.name ?? m.label;
      if (m.sources && typeof m.sources.forEach === 'function') m.sources.forEach(source => add(source, target));
      else add(m.from ?? m.source, target);
    });
    else if (merges && typeof merges === 'object') Object.keys(merges).forEach(k => add(k, merges[k]));
    return map;
  }
  function merged(value, map) { let out = text(value), seen = new Set(); while (map.has(key(out)) && !seen.has(key(out))) { seen.add(key(out)); out = map.get(key(out)); } return out; }
  function setKeys(set) { const out = new Set(); if (set) for (const v of set) out.add(key(v)); return out; }
  function responseLabels(value, includeBlanks, blankLabel, options) { const result = labels(value, options); return result.length ? result : (includeBlanks ? [blankLabel] : []); }
  function stableDedupe(values) { const seen = new Set(); return values.filter(v => { const k = key(v); if (seen.has(k)) return false; seen.add(k); return true; }); }
  function chooseOther(existing) { let x = 'Other', n = 2; const keys = new Set(existing.map(key)); while (keys.has(key(x))) x = `Other (${n++})`; return x; }
  function unusedLabel(base, values) { let out = base, n = 2, used = new Set(values.map(key)); while (used.has(key(out))) out = `${base} (${n++})`; return out; }
  function cap(values, totals, limit) {
    if (!limit || values.length <= limit) return { shown: values.slice(), hidden: [] };
    const ranked = values.map((v, i) => ({ v, i, n: totals.get(v) || 0 })).sort((a,b) => b.n - a.n || a.i - b.i);
    const shown = ranked.slice(0, limit).map(x => x.v), hidden = values.filter(v => !shown.includes(v));
    return { shown, hidden };
  }
  function summarize(rows, column, options) {
    options = options || {}; const requestedBlank = options.blankLabel || '(No response)'; const observed = (rows || []).flatMap(row => labels(row && row[column], options)); const blankLabel = unusedLabel(requestedBlank, observed); const merge = mergeMap(options.merges); const hidden = setKeys(options.hiddenResponses); const counts = new Map(); const memberships = []; let missingCount = 0;
    (rows || []).forEach(row => {
      let vals = stableDedupe(responseLabels(row && row[column], !!options.includeBlanks, blankLabel, options));
      vals = stableDedupe(vals.map(v => merged(v, merge))).filter(v => !hidden.has(key(v)));
      const answered = vals.filter(v => key(v) !== key(blankLabel)); if (!answered.length) missingCount++;
      const canonical = []; vals.forEach(v => { const existing = rawKey(counts, v) || v; if (!canonical.some(x => key(x) === key(existing))) canonical.push(existing); });
      memberships.push(canonical);
      canonical.forEach(v => counts.set(v, (counts.get(v) || 0) + 1));
    });
    const answeredRows = memberships.filter(values => values.some(v => key(v) !== key(blankLabel))).length;
    const answeredTotal = answeredRows; const respondentTotal = options.includeBlanks ? (rows || []).length : answeredRows;
    const raw = Array.from(counts.keys()), total = new Map(raw.map(v => [v, counts.get(v)]));
    const limit = options.topMode === 'all' || options.topMode == null ? 0 : Number(options.topMode || 0); const capped = cap(raw, total, limit); let shown = capped.shown.slice();
    if (capped.hidden.length) { const other = chooseOther(raw); shown.push(other); counts.set(other, memberships.filter(values => values.some(v => capped.hidden.some(h => key(h) === key(v)))).length); }
    const nonBlank = answeredTotal;
    const round = value => Math.round(value * 10) / 10;
    const items = shown.map(response => ({ response, count: counts.get(response) || 0, rowPercent: respondentTotal ? round((counts.get(response)||0) / respondentTotal * 100) : 0, nonBlankPercent: nonBlank ? round((counts.get(response)||0) / nonBlank * 100) : 0 }));
    if (options.sortMode === 'alpha') items.sort((a,b) => a.response.localeCompare(b.response));
    else if (options.sortMode === 'asc') items.sort((a,b) => a.count-b.count || a.response.localeCompare(b.response));
    else if (options.sortMode === 'count' || options.sortMode === 'desc') items.sort((a,b) => b.count-a.count || a.response.localeCompare(b.response));
    return { items, labels: items.map(x=>x.response), values: items.map(x=>x.count), respondentTotal, answeredTotal, missingCount };
  }
  function rawKey(map, value) { const k = key(value); for (const candidate of map.keys()) if (key(candidate) === k) return candidate; return ''; }
  function comparison(rows, primary, compare, options) {
    options = options || {}; const requestedBlank = options.blankLabel || '(No response)'; const observed = (rows || []).flatMap(row => labels(row && row[primary], options).concat(labels(row && row[compare], options))); const blank = unusedLabel(requestedBlank, observed), hidden = setKeys(options.hiddenResponses), merge = mergeMap(options.merges); const records=[]; const pCanon=new Map(), cCanon=new Map();
    (rows || []).forEach(row=>{ let ps=stableDedupe(responseLabels(row&&row[primary],!!options.includeBlanks,blank,options)).map(v=>merged(v,merge)).filter(v=>!hidden.has(key(v))).map(v=>{const k=key(v);if(!pCanon.has(k))pCanon.set(k,v);return pCanon.get(k);}); let cs=stableDedupe(responseLabels(row&&row[compare],!!options.includeBlanks,blank,options)).map(v=>{const k=key(v);if(!cCanon.has(k))cCanon.set(k,v);return cCanon.get(k);}); ps=stableDedupe(ps);cs=stableDedupe(cs);if(ps.length&&cs.length)records.push({ps,cs}); });
    const ps=Array.from(new Set(records.flatMap(r=>r.ps))), cs=Array.from(new Set(records.flatMap(r=>r.cs))); const pResp=new Map(ps.map(p=>[p,records.filter(r=>r.ps.includes(p)).length])), cResp=new Map(cs.map(c=>[c,records.filter(r=>r.cs.includes(c)).length]));
    const limit=options.topMode==='all'||options.topMode==null?0:Number(options.topMode||0), pc=cap(ps,pResp,limit),cc=cap(cs,cResp,limit); let pl=pc.shown.slice(),cl=cc.shown.slice(); const pOther=pc.hidden.length?chooseOther(ps):null,cOther=cc.hidden.length?chooseOther(cs):null;if(pOther)pl.push(pOther);if(cOther)cl.push(cOther);
    const mapP=p=>pl.includes(p)?p:pOther, mapC=c=>cl.includes(c)?c:cOther;
    const matrix=new Map(pl.map(p=>[p,new Map(cl.map(c=>[c,0]))]));
    const cappedPRespondents=new Map(pl.map(p=>[p,0])), cappedCRespondents=new Map(cl.map(c=>[c,0]));
    records.forEach(r=>{const mappedP=new Set(r.ps.map(mapP)), mappedC=new Set(r.cs.map(mapC)); mappedP.forEach(p=>cappedPRespondents.set(p,cappedPRespondents.get(p)+1)); mappedC.forEach(c=>cappedCRespondents.set(c,cappedCRespondents.get(c)+1)); mappedP.forEach(p=>mappedC.forEach(c=>matrix.get(p).set(c,matrix.get(p).get(c)+1)));});
    const primaryTotals=new Map(pl.map(p=>[p,cl.reduce((s,c)=>s+matrix.get(p).get(c),0)])),compareTotals=new Map(cl.map(c=>[c,pl.reduce((s,p)=>s+matrix.get(p).get(c),0)]));
    return {type:'comparison',rows,labels:pl,compareLabels:cl,matrix,primaryTotals,compareTotals,primaryRespondentTotals:cappedPRespondents,compareRespondentTotals:cappedCRespondents,respondentTotal:records.length,total:Array.from(matrix.values()).flatMap(m=>Array.from(m.values())).reduce((a,b)=>a+b,0),datasets:cl.map(c=>({label:c,data:pl.map(p=>matrix.get(p).get(c)||0)}))};
  }
  function reportSection(rows, displayName, column, breakdownColumns, options) {
    options=options||{}; breakdownColumns=breakdownColumns||[]; const delimiter=options.delimiter ? {delimiter:options.delimiter} : {}; const blankByColumn=new Map(); const groupValues=breakdownColumns.map(c=>{const observed=(rows||[]).flatMap(r=>labels(r&&r[c],delimiter)); const blank=unusedLabel('(No response)',observed); blankByColumn.set(c,blank); const seen=new Map();observed.forEach(v=>{if(!seen.has(key(v)))seen.set(key(v),v);}); if(!seen.has(key(blank)))seen.set(key(blank),blank); return Array.from(seen.values());});
    let groups=[{values:[],label:'All rows'}]; if(breakdownColumns.length){groups=[];const build=(i,vals)=>{if(i===groupValues.length){groups.push({values:vals,label:vals.join(' / ')});return;}groupValues[i].forEach(v=>build(i+1,vals.concat(v)));};build(0,[]);}
    const answers=stableDedupe((options.answerOrder||[]).map(text)); const canon=new Map(answers.map(v=>[key(v),v])); const observed=[];(rows||[]).forEach(r=>labels(r&&r[column],delimiter).forEach(v=>{if(!canon.has(key(v))){canon.set(key(v),v);observed.push(v);}})); answers.push(...observed);
    const aggregates=groups.map(g=>({counts:new Map(),answered:0,eligible:0})); const groupIndex=new Map(groups.map((g,i)=>[g.values.map(key).join('\u001f'),i]));
    (rows||[]).forEach(r=>{const av=labels(r&&r[column],delimiter), unique=stableDedupe(av); const memberships=breakdownColumns.map(c=>{const vals=labels(r&&r[c],delimiter);return vals.length?vals:[blankByColumn.get(c)];}); const visit=(depth,chosen)=>{if(depth===memberships.length){const i=groupIndex.get(chosen.map(key).join('\u001f'));if(i==null)return;const a=aggregates[i];a.eligible++;if(unique.length)a.answered++;unique.forEach(v=>{const canonical=canon.get(key(v))||v;a.counts.set(canonical,(a.counts.get(canonical)||0)+1);});return;}memberships[depth].forEach(v=>visit(depth+1,chosen.concat(v)));}; visit(0,[]);});
    const width=1+groups.length*2, aoa=[], rendered=[]; const title=Array(width).fill('');title[0]=displayName;aoa.push(title);rendered.push([{value:displayName,type:'question',colspan:width}]); const h1=['',...groups.flatMap(g=>[g.label,''])],h2=['',...groups.flatMap(()=>['Count','Percentage'])];aoa.push(h1,h2);rendered.push([{value:'',type:'header'},...groups.map(g=>({value:g.label,type:'group-header',colspan:2}))],[{value:'',type:'header'},...groups.flatMap(()=>[{value:'Count',type:'header'},{value:'Percentage',type:'header'}])]);
    answers.forEach(answer=>{const row=Array(width).fill('');row[0]=answer;const cells=[{value:answer,type:'answer'}];aggregates.forEach((a,i)=>{const n=a.counts.get(answer)||0;row[1+i*2]=n;row[2+i*2]=a.answered?n/a.answered:null;cells.push({value:n,type:'count'},{value:row[2+i*2],type:'percent'});});aoa.push(row);rendered.push(cells);});
    ['Answered','Missing','Eligible'].forEach((label,mi)=>{const row=Array(width).fill('');row[0]=label;const cells=[{value:label,type:'metadata'}];aggregates.forEach((a,i)=>{const n=mi===0?a.answered:mi===1?a.eligible-a.answered:a.eligible;row[1+i*2]=n;cells.push({value:n,type:'count'},{value:'',type:'blank'});});aoa.push(row);rendered.push(cells);}); return {aoa,rows:rendered,answers,groups};
  }
  return {text,key,labels,uniqueHeaders,summarize,comparison,reportSection,NO_RESPONSE};
});
