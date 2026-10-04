'use strict';

// Keep parsing offline and deterministic. The build copies vendor/ recursively.
importScripts('vendor/xlsx.full.min.js');

self.onmessage = event => {
  try {
    const workbook = XLSX.read(event.data.buffer, {
      type: 'array',
      cellDates: true,
      raw: false
    });
    self.postMessage({ type: 'success', workbook });
  } catch (error) {
    self.postMessage({ type: 'error', message: `Workbook parse failed: ${error?.message || 'The workbook could not be parsed.'}` });
  }
};
