(function (root, factory) {
  root.DataDictionary = factory();
})(typeof globalThis !== 'undefined' ? globalThis : window, function () {
  'use strict';

  const KEY_SEPARATOR = '\u0000';

  function text(value) {
    if (value === null || value === undefined) return '';
    return String(value).trim();
  }

  function keyText(value) {
    return text(value)
      .toLocaleLowerCase()
      .replace(/[\s_-]+/g, ' ')
      .trim();
  }

  function mappingKey(sheetName, headerName) {
    return `${keyText(sheetName)}${KEY_SEPARATOR}${keyText(headerName)}`;
  }

  function isDictionarySheetName(sheetName) {
    return keyText(sheetName) === 'data dictionary';
  }

  function isHeaderRow(row) {
    const values = (row || []).slice(0, 3).map(keyText);
    return (values[0] === 'sheet name' || values[0] === 'sheet')
      && (values[1] === 'original/header name' || values[1] === 'original header name' || values[1] === 'header name')
      && (values[2] === 'display question' || values[2] === 'question');
  }

  function create(rows) {
    const mappings = new Map();
    (Array.isArray(rows) ? rows : []).forEach((row, index) => {
      if (index === 0 && isHeaderRow(row)) return;
      const sheetName = text(row?.[0]);
      const headerName = text(row?.[1]);
      const displayQuestion = text(row?.[2]);
      if (!sheetName || !headerName || !displayQuestion) return;

      const key = mappingKey(sheetName, headerName);
      if (!mappings.has(key)) mappings.set(key, displayQuestion);
    });

    return { mappings, hasEntries: mappings.size > 0 };
  }

  function getDisplayQuestion(dictionary, sheetName, headerName) {
    if (!dictionary?.mappings || !keyText(sheetName) || !keyText(headerName)) return '';
    return dictionary.mappings.get(mappingKey(sheetName, headerName)) || '';
  }

  return {
    create,
    getDisplayQuestion,
    isDictionarySheetName
  };
});
