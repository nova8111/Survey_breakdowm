(function (root, factory) {
  root.ChartRules = factory();
})(typeof globalThis !== 'undefined' ? globalThis : window, function () {
  'use strict';

  const DEFAULT_MAX_UNIQUE_VALUES = 15;
  const DEFAULT_MIN_UNIQUE_VALUES = 2;

  function text(value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    return String(value).trim();
  }

  function splitResponseLabels(value) {
    const values = Array.isArray(value)
      ? value.flatMap(splitResponseLabels)
      : text(value).split(';').map(item => item.trim());
    return Array.from(new Set(values.filter(Boolean)));
  }

  function answerLabels(value) {
    return splitResponseLabels(value);
  }

  function normalizeAnswer(value) {
    return text(value)
      .toLocaleLowerCase()
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u201c\u201d]/g, '"')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function countUniqueAnswers(rows, column) {
    const values = new Set();
    rows.forEach(row => {
      answerLabels(row[column]).forEach(value => values.add(normalizeAnswer(value)));
    });
    values.delete('');
    return values.size;
  }

  function shouldUseYesNoLabels(labels) {
    const normalized = new Set(labels.map(normalizeAnswer).filter(Boolean));
    return normalized.size === 2 && normalized.has('0') && normalized.has('1');
  }

  function getDisplayAnswerLabel(label, allLabels, useYesNoLabels = null) {
    if (!(useYesNoLabels ?? shouldUseYesNoLabels(allLabels))) return label;
    const normalized = normalizeAnswer(label);
    if (normalized === '0') return 'No';
    if (normalized === '1') return 'Yes';
    return label;
  }

  function isLikelyMetadataColumn(column) {
    return /^(id|student id|response id|tempid)$/i.test(text(column))
      || /date|timestamp|email|name|phone|address/i.test(text(column));
  }

  function getSelectableChartColumns(columns, options = {}) {
    const hiddenColumns = options.hiddenColumns instanceof Set ? options.hiddenColumns : new Set(options.hiddenColumns || []);
    const includeMetadata = options.includeMetadata === true;
    return columns.filter(column => !hiddenColumns.has(column)
      && (includeMetadata || !isLikelyMetadataColumn(column)));
  }

  function getEligibleChartColumns(rows, columns, options = {}) {
    const maxUniqueValues = options.maxUniqueValues ?? DEFAULT_MAX_UNIQUE_VALUES;
    const minUniqueValues = options.minUniqueValues ?? DEFAULT_MIN_UNIQUE_VALUES;
    const hiddenColumns = options.hiddenColumns instanceof Set ? options.hiddenColumns : new Set(options.hiddenColumns || []);
    const includeMetadata = options.includeMetadata === true;
    const uniqueCounts = options.uniqueCounts && typeof options.uniqueCounts.has === 'function'
      ? options.uniqueCounts
      : null;

    return getSelectableChartColumns(columns, { hiddenColumns, includeMetadata }).filter(column => {
      const uniqueCount = uniqueCounts?.has(column)
        ? uniqueCounts.get(column)
        : countUniqueAnswers(rows, column);
      return uniqueCount >= minUniqueValues && uniqueCount <= maxUniqueValues;
    });
  }

  function getMissingChartColumns(eligibleColumns, charts) {
    const existingColumns = new Set(charts.map(chart => chart.primaryColumn).filter(Boolean));
    return eligibleColumns.filter(column => !existingColumns.has(column));
  }

  function getSelectedChartColumns(eligibleColumns, selectedColumns) {
    const selected = selectedColumns instanceof Set ? selectedColumns : new Set(selectedColumns || []);
    return eligibleColumns.filter(column => selected.has(column));
  }

  return {
    DEFAULT_MAX_UNIQUE_VALUES,
    DEFAULT_MIN_UNIQUE_VALUES,
    countUniqueAnswers,
    getDisplayAnswerLabel,
    getEligibleChartColumns,
    getMissingChartColumns,
    getSelectableChartColumns,
    getSelectedChartColumns,
    isLikelyMetadataColumn,
    shouldUseYesNoLabels,
    splitResponseLabels
  };
});
