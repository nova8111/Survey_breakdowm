'use strict';

importScripts('https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js');

self.onmessage = event => {
  try {
    const workbook = XLSX.read(event.data.buffer, {
      type: 'array',
      cellDates: true,
      raw: false
    });
    self.postMessage({ type: 'success', workbook });
  } catch (error) {
    self.postMessage({
      type: 'error',
      message: error?.message || 'The workbook could not be parsed.'
    });
  }
};
