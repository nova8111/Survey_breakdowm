/* Small, dependency-free data import/export helpers shared by the browser and tests. */
(function (root, factory) {
  const api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof globalThis !== 'undefined') globalThis.DataIO = api;
  if (root && root !== globalThis) root.DataIO = api;
})(typeof self !== 'undefined' ? self : globalThis, function () {
  function bytesOf(input) {
    if (typeof input === 'string') return null;
    if (input instanceof ArrayBuffer) return new Uint8Array(input);
    if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    throw new TypeError('CSV input must be a string, Buffer, ArrayBuffer, or typed array.');
  }

  function decode(input, encoding = 'auto') {
    if (typeof input === 'string') return input.replace(/^\uFEFF/, '');
    const bytes = bytesOf(input);
    let chosen = String(encoding || 'auto').toLowerCase();
    let automatic = chosen === 'auto';
    if (automatic) {
      if (bytes[0] === 0xff && bytes[1] === 0xfe) chosen = 'utf-16le';
      else if (bytes[0] === 0xfe && bytes[1] === 0xff) chosen = 'utf-16be';
      else {
        // BOMless UTF-16 commonly has NUL bytes in every other position.
        const sample = bytes.slice(0, Math.min(bytes.length, 256));
        let oddNuls = 0, evenNuls = 0;
        for (let i = 0; i < sample.length; i += 1) {
          if (sample[i] === 0) {
            if (i % 2) oddNuls += 1;
            else evenNuls += 1;
          }
        }
        if (sample.length > 3 && oddNuls > sample.length / 5 && oddNuls > evenNuls * 2) chosen = 'utf-16le';
        else if (sample.length > 3 && evenNuls > sample.length / 5 && evenNuls > oddNuls * 2) chosen = 'utf-16be';
        else chosen = 'utf-8';
      }
    }
    const hasUtf16Bom = (bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff);
    const skip = hasUtf16Bom ? 2 : (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0);
    if (typeof TextDecoder === 'undefined') throw new Error('TextDecoder is required to decode CSV input.');
    let label = chosen === 'cp1252' || chosen === 'windows-1252' ? 'windows-1252' : chosen;
    try { return new TextDecoder(label, { fatal: automatic && label === 'utf-8' }).decode(bytes.slice(skip)).replace(/^\uFEFF/, ''); }
    catch (error) {
      if (automatic && label === 'utf-8') {
        try { return new TextDecoder('windows-1252').decode(bytes).replace(/^\uFEFF/, ''); }
        catch { /* report the original UTF-8 error below */ }
      }
      throw new Error(`Unable to decode CSV as ${chosen}: ${error.message}`);
    }
  }

  function detectDelimiter(text) {
    const candidates = [',', '\t', ';'];
    const counts = Object.fromEntries(candidates.map(d => [d, 0]));
    let quoted = false;
    let records = 0;
    for (let i = 0; i < text.length && records < 8; i += 1) {
      const c = text[i];
      if (c === '"') {
        if (quoted && text[i + 1] === '"') i += 1;
        else quoted = !quoted;
      } else if (!quoted && candidates.includes(c)) counts[c] += 1;
      else if (!quoted && (c === '\n' || c === '\r')) records += 1;
    }
    return candidates.sort((a, b) => counts[b] - counts[a])[0] || ',';
  }

  function parseCsv(input, options = {}) {
    const text = decode(input, options.encoding || 'auto');
    const delimiter = options.delimiter && options.delimiter !== 'auto' ? String(options.delimiter) : detectDelimiter(text);
    if (delimiter.length !== 1) throw new Error('CSV delimiter must be one character.');
    const rows = [], row = [];
    let field = '', quoted = false, justClosed = false;
    const pushField = () => { row.push(field); field = ''; };
    const pushRow = () => { pushField(); rows.push(row.splice(0)); };
    for (let i = 0; i < text.length; i += 1) {
      const c = text[i];
      if (quoted) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 1; }
          else { quoted = false; justClosed = true; }
        } else field += c;
      } else if (c === '"' && field === '') {
        quoted = true; justClosed = false;
      } else if (c === '"') {
        throw new Error(`Malformed CSV: quote must start a field or be escaped at position ${i}.`);
      } else if (c === delimiter) {
        if (justClosed) justClosed = false;
        pushField();
      } else if (c === '\n' || c === '\r') {
        if (justClosed) justClosed = false;
        if (c === '\r' && text[i + 1] === '\n') i += 1;
        pushRow();
      } else if (justClosed && /\s/.test(c)) {
        // RFC 4180 permits insignificant whitespace before the delimiter/record end.
      } else if (justClosed) {
        throw new Error(`Malformed CSV: unexpected character ${JSON.stringify(c)} after a closing quote at position ${i}.`);
      } else { field += c; justClosed = false; }
    }
    if (quoted) throw new Error('Malformed CSV: quoted field was not closed.');
    if (field.length || row.length) pushRow();
    return rows;
  }

  function csvEscape(value) {
    const isNumber = typeof value === 'number' && Number.isFinite(value);
    let text = value == null ? '' : String(value);
    if (!isNumber && (/^[\s]*[=+\-@]/.test(text) || /^[\t\r]/.test(text))) text = `'${text}`;
    if (/[",\r\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
    return text;
  }

  function toCsv(rows) {
    return '\uFEFF' + rows.map(row => row.map(csvEscape).join(',')).join('\r\n');
  }

  function metadataRows(result) {
    const fields = [['Report', result.title], ['Source', result.sourceName], ['Sheet', result.sheetName], ['Filter', result.filterLabel], ['Breakdown', result.breakdownLabel], ['Included rows', result.includedCount ?? result.includedRows], ['Excluded rows', result.excludedCount ?? result.excludedRows], ['Denominator', result.denominatorDescription ?? result.denominator], ['Generated', result.generatedAt]];
    return fields.filter(([, value]) => value != null && value !== '').map(([label, value]) => [label, value]);
  }

  function selectReportColumns(rows, mode, options = {}) {
    if (mode === 'both') return rows.map(row => row.slice());
    const firstHeader = rows.findIndex(row => String(row[0] ?? '').trim() === '' && row.some(value => /^(count|percentage)$/i.test(String(value).trim())));
    const start = firstHeader < 0 ? 0 : firstHeader;
    const headerRows = rows.slice(start);
    const percentIndexes = new Set(), countIndexes = new Set();
    headerRows.forEach(row => row.forEach((value, index) => {
      if (String(value).trim().toLowerCase() === 'percentage') percentIndexes.add(index);
      if (String(value).trim().toLowerCase() === 'count') countIndexes.add(index);
    }));
    const keep = index => index === 0 || (mode === 'percentage' ? percentIndexes.has(index) : countIndexes.has(index));
    const selectedIndexes = [...(mode === 'percentage' ? percentIndexes : countIndexes)].sort((a, b) => a - b);
    const groupCount = Math.max(0, Math.ceil((Math.max(...selectedIndexes, ...percentIndexes, ...countIndexes, 0)) / 2));
    const compact = row => {
      if (row.length <= 2 || !groupCount) return row.slice();
      const output = [row[0]];
      for (let group = 0; group < groupCount; group += 1) {
        const countIndex = 1 + group * 2;
        const percentIndex = countIndex + 1;
        // Group headers carry their label in the count position in either mode.
        const source = mode === 'percentage' && row[percentIndex] !== '' && row[percentIndex] != null ? percentIndex : countIndex;
        output.push(row[source] ?? '');
      }
      return output;
    };
    return rows.map((row, rowIndex) => {
      const label = String(row[0] ?? '').trim().toLowerCase();
      // These base totals are useful in every display mode; their values live in
      // the count positions (1 + i*2), while percentage positions are blank.
      if (['answered', 'missing', 'eligible'].includes(label)) return compact(row);
      if (rowIndex < start) return row.length > 2 ? compact(row) : row.slice();
      if (options.percentText && mode === 'percentage') {
        return row.flatMap((value, sourceIndex) => {
          if (!keep(sourceIndex)) return [];
          if (sourceIndex > 0 && percentIndexes.has(sourceIndex) && typeof value === 'number') {
            const formatted = (value * 100).toFixed(1).replace(/\.0$/, '');
            return [`${formatted}%`];
          }
          return [value];
        });
      }
      return row.filter((_, index) => keep(index));
    });
  }

  function reportRows(result, options = {}) {
    const mode = options.mode || 'both';
    const aoa = result.aoa || result.rows || [];
    const linked = Array.isArray(result.linkedSummary) && result.linkedSummary.length
      ? [[], ['Linked category summary'], ['Category', 'Matched sites', 'Survey responses'], ...result.linkedSummary.map(item => [item.category ?? '', item.matchedSites ?? '', item.surveyResponses ?? ''])]
      : [];
    return [...metadataRows(result), ...(metadataRows(result).length ? [[]] : []), ...selectReportColumns(aoa, mode, options), ...linked];
  }

  function buildExportRows(result, mode = 'both') { return reportRows(result, { mode }); }

  function configureReportSheet(sheet, rows, options = {}) {
    if (!sheet || !rows) return sheet;
    let percentColumns = new Set();
    let inReport = false;
    rows.forEach((row, r) => {
      const normalized = row.map(value => String(value ?? '').trim().toLowerCase());
      const linkedReset = normalized.includes('linked category summary') || (normalized[0] === 'category' && normalized.includes('matched sites') && normalized.includes('survey responses'));
      if (linkedReset) { inReport = false; percentColumns = new Set(); }
      const headerPercent = new Set();
      normalized.forEach((value, index) => { if (value === 'percentage' && normalized[0] === '') headerPercent.add(index); });
      if (headerPercent.size) { inReport = true; percentColumns = headerPercent; }
      row.forEach((value, c) => {
      if (value == null) return;
      let column = '', n = c + 1;
      while (n) { const remainder = (n - 1) % 26; column = String.fromCharCode(65 + remainder) + column; n = Math.floor((n - 1) / 26); }
      const address = `${column}${r + 1}`;
      const cell = sheet[address];
      if (!cell) return;
      if (typeof value === 'string') { cell.t = 's'; cell.v = value; delete cell.f; }
      const metadataRow = ['answered', 'missing', 'eligible'].includes(normalized[0]);
      if (inReport && !metadataRow && percentColumns.has(c) && typeof value === 'number' && !headerPercent.size) cell.z = '0.0%';
      });
    });
    const widths = [];
    rows.forEach(row => row.forEach((value, index) => { widths[index] = Math.max(widths[index] || 0, String(value ?? '').length); }));
    sheet['!cols'] = widths.map((width, index) => ({ wch: Math.min(42, Math.max(index ? 12 : 18, width + 2)) }));
    return sheet;
  }

  return { parseCsv, csvEscape, toCsv, configureReportSheet, reportRows, buildExportRows };
});
