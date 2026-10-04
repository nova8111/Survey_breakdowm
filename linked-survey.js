(function (root, factory) {
  root.LinkedSurvey = factory();
})(typeof globalThis !== 'undefined' ? globalThis : window, function () {
  'use strict';

  const DEFAULT_PREFIX = 'Linked survey';

  function text(value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    return String(value).trim();
  }

  function normalizeMatchKey(value) {
    return text(value)
      .normalize('NFKD')
      .replace(/\p{M}+/gu, '')
      .toLocaleLowerCase()
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u2010-\u2015]/g, '-')
      .replace(/&/g, ' and ')
      .replace(/[^\p{L}\p{N}]+/gu, '')
      .trim();
  }

  function splitMultiSelect(value, delimiter = 'all') {
    if (Array.isArray(value)) {
      return unique(value.flat(Infinity).map(item => text(item)).filter(Boolean));
    }
    const normalized = text(value);
    if (!normalized) return [];
    const separators = {
      all: /\r?\n|\s*;\s*|\s*\|\s*|\s*,\s*/,
      semicolon: /\s*;\s*/,
      comma: /\s*,\s*/,
      pipe: /\s*\|\s*/,
      newline: /\r?\n/,
      none: null
    };
    const separator = Object.prototype.hasOwnProperty.call(separators, delimiter)
      ? separators[delimiter]
      : separators.all;
    return unique((separator ? normalized.split(separator) : [normalized])
      .map(item => item.trim()).filter(Boolean));
  }

  function unique(values) {
    const seen = new Set();
    return values.filter(value => {
      const key = text(value).toLocaleLowerCase();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function rowNumber(row, index) {
    return typeof row.__sourceRowNumber === 'number' && Number.isFinite(row.__sourceRowNumber)
      ? row.__sourceRowNumber
      : index + 2;
  }

  function collectKeys(rows, field, normalize) {
    const groups = new Map();
    rows.forEach((row, index) => {
      const rawValue = text(row[field]);
      const key = normalize(rawValue);
      if (!key) return;
      if (!groups.has(key)) groups.set(key, { key, value: rawValue, rowNumbers: [], rawValues: [] });
      const group = groups.get(key);
      group.rowNumbers.push(rowNumber(row, index));
      if (!group.rawValues.includes(rawValue)) group.rawValues.push(rawValue);
    });
    return groups;
  }

  function duplicateGroups(groups) {
    return Array.from(groups.values())
      .filter(group => group.rowNumbers.length > 1)
      .map(group => ({ key: group.key, value: group.value, rowNumbers: group.rowNumbers, count: group.rowNumbers.length }));
  }

  function analyzeLink(primaryRows, secondaryRows, primaryField, secondaryField, options = {}) {
    if (!Array.isArray(primaryRows) || !Array.isArray(secondaryRows)) {
      throw new Error('Both surveys must contain row arrays.');
    }
    if (!primaryField || !secondaryField) throw new Error('Select a matching field for both surveys.');
    if (primaryRows.length && !(primaryField in primaryRows[0])) throw new Error('The selected primary matching field is invalid.');
    if (secondaryRows.length && !(secondaryField in secondaryRows[0])) throw new Error('The selected secondary matching field is invalid.');

    const matchingMode = options.mode === 'exact' ? 'exact' : 'normalized';
    const normalize = matchingMode === 'exact' ? value => text(value) : normalizeMatchKey;
    const primaryGroups = collectKeys(primaryRows, primaryField, normalize);
    const secondaryGroups = collectKeys(secondaryRows, secondaryField, normalize);
    const primaryDuplicates = duplicateGroups(primaryGroups);
    const secondaryDuplicates = duplicateGroups(secondaryGroups);
    const ambiguousSecondaryKeys = new Set(secondaryDuplicates.map(group => group.key));
    const secondaryIndex = new Map();

    secondaryRows.forEach((row, index) => {
      const key = normalize(row[secondaryField]);
      if (key && !ambiguousSecondaryKeys.has(key)) secondaryIndex.set(key, { row, rowNumber: rowNumber(row, index) });
    });

    const matched = [];
    const unmatched = [];
    const matchedKeys = new Set();
    primaryRows.forEach((row, index) => {
      const rawValue = text(row[primaryField]);
      const key = normalize(rawValue);
      let reason = '';
      if (!key) reason = 'Missing primary matching value';
      else if (ambiguousSecondaryKeys.has(key)) reason = 'Duplicate value in secondary survey';
      else if (!secondaryIndex.has(key)) reason = 'No matching secondary record';

      if (reason) {
        unmatched.push({ row, rowNumber: rowNumber(row, index), value: rawValue, key, reason });
        return;
      }

      const secondary = secondaryIndex.get(key);
      matched.push({ primary: row, secondary: secondary.row, key, primaryRowNumber: rowNumber(row, index), secondaryRowNumber: secondary.rowNumber });
      matchedKeys.add(key);
    });

    const normalizationCollisions = matchingMode === 'normalized'
      ? Array.from(new Set([...primaryGroups.keys(), ...secondaryGroups.keys()])).map(key => {
        const primary = primaryGroups.get(key);
        const secondary = secondaryGroups.get(key);
        const primaryValues = primary ? primary.rawValues : [];
        const secondaryValues = secondary ? secondary.rawValues : [];
        const values = [...new Set([...primaryValues, ...secondaryValues])];
        return values.length > 1 ? { key, primaryValues, secondaryValues } : null;
      }).filter(Boolean)
      : [];
    const matchedRows = matched.length;
    const unmatchedRows = unmatched.length;
    const totalPrimaryRows = primaryRows.length;
    return {
      primaryField,
      secondaryField,
      matchingMode,
      mode: matchingMode,
      matched,
      unmatched,
      duplicates: { primary: primaryDuplicates, secondary: secondaryDuplicates },
      normalizationCollisions,
      stats: {
        totalPrimaryRows,
        totalSecondaryRows: secondaryRows.length,
        matchedRows,
        unmatchedRows,
        matchedSites: matchedKeys.size,
        matchRate: totalPrimaryRows ? (matchedRows / totalPrimaryRows) * 100 : 0,
        unmatchedRate: totalPrimaryRows ? (unmatchedRows / totalPrimaryRows) * 100 : 0
      }
    };
  }

  function linkedColumnName(question, prefix = DEFAULT_PREFIX) {
    return `${prefix}: ${question}`;
  }

  function enrichMatchedRows(linkResult, question, options = {}) {
    if (!linkResult || !Array.isArray(linkResult.matched)) throw new Error('Link the surveys before selecting a secondary question.');
    if (!question) throw new Error('Select a secondary survey question for disaggregation.');
    if (linkResult.matched.some(match => !match.secondary || !(question in match.secondary))) {
      throw new Error('The selected secondary survey question is invalid.');
    }
    const requestedColumn = linkedColumnName(options.displayQuestion || question, options.prefix);
    const existingColumns = new Set(linkResult.matched.flatMap(match => Object.keys(match.primary || {})));
    let column = requestedColumn;
    let suffix = 2;
    while (existingColumns.has(column)) column = `${requestedColumn} (${suffix++})`;
    const delimiter = options.delimiter === undefined ? 'semicolon' : options.delimiter;
    const rows = linkResult.matched.map(match => ({
      ...match.primary,
      [column]: splitMultiSelect(match.secondary[question], delimiter),
      __linkedSiteKey: match.key
    }));
    return { rows, column };
  }

  function buildCategorySummary(rows, column) {
    const categories = new Map();
    rows.forEach(row => {
      const values = Array.isArray(row[column])
        ? splitMultiSelect(row[column], 'none')
        : splitMultiSelect(row[column], 'semicolon');
      values.forEach(value => {
        const key = value.toLocaleLowerCase();
        if (!categories.has(key)) categories.set(key, { category: value, matchedSites: new Set(), surveyResponses: 0 });
        const item = categories.get(key);
        if (row.__linkedSiteKey) item.matchedSites.add(row.__linkedSiteKey);
        item.surveyResponses += 1;
      });
    });
    return Array.from(categories.values())
      .map(item => ({ category: item.category, matchedSites: item.matchedSites.size, surveyResponses: item.surveyResponses }))
      .sort((a, b) => a.category.localeCompare(b.category));
  }

  return {
    analyzeLink,
    buildCategorySummary,
    enrichMatchedRows,
    linkedColumnName,
    normalizeMatchKey,
    splitMultiSelect
  };
});
