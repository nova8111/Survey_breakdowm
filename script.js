(function () {
  'use strict';

  const NO_RESPONSE = 'No Response';
  const TABLE_ROW_LIMIT = 10;
  const CHART_RESPONSE_DISPLAY_LIMIT = ChartRules.DEFAULT_MAX_UNIQUE_VALUES;
  const REPORT_UNIQUE_VALUE_LIMIT = 15;
  const REPORT_FILTER_UNIQUE_VALUE_LIMIT = 500;
  const UPLOADED_SOURCE_ID = 'uploaded-workbook';
  const rootStyles = getComputedStyle(document.documentElement);
  const COLORS = Array.from({ length: 8 }, (_, index) =>
    rootStyles.getPropertyValue(`--chart-${index + 1}`).trim()
  ).filter(Boolean);
  const CHART_TEXT = rootStyles.getPropertyValue('--chart-text').trim();
  const CHART_MUTED = rootStyles.getPropertyValue('--chart-muted').trim();
  const CHART_GRID = rootStyles.getPropertyValue('--chart-grid').trim();
  const CHART_ON_COLOR = rootStyles.getPropertyValue('--chart-on-color').trim();
  const sheetMatrixCache = new WeakMap();
  const sheetHeaderCache = new WeakMap();
  const sheetRecordsCache = new WeakMap();
  const workbookDictionaryCache = new WeakMap();
  const WORKBOOK_PARSE_OPTIONS = { type: 'array', cellDates: true, raw: false };
  const ANSWER_SETS = [
    ['Strongly disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly agree'],
    ['Not at all', 'Not really', 'Kind of', 'Definitely', 'Absolutely'],
    ['Never', 'Almost never', 'Sometimes', 'Lots of times', 'All the time'],
    ['1: Bad', '2: Okay', '3: Good', '4: Great', '5: Amazing'],
    ['1: Not confident at all', '2: Slightly confident', '3: Somewhat confident', '4: Very confident', '5: Extremely confident'],
    ['Poor', 'Fair', 'Good', 'Very Good', 'Excellent'],
    ['Not well at all', 'Slightly well', 'Somewhat well', 'Quite well', 'Extremely well'],
    ['Not at all', 'A little bit', 'Somewhat well', 'Quite well', 'Extremely well'],
    Array.from({ length: 10 }, (_, index) => String(index + 1)),
    ['Yes', 'No', "I don't know"],
    ['Yes', 'Maybe', 'No'],
    ['Very low extent', 'Low extent', 'Moderate extent', 'Great extent', 'Very great extent'],
    ['Not informed at all', 'Slightly informed', 'Somewhat informed', 'Quite informed', 'Extremely informed'],
    ['Not effective at all', 'Slightly effective', 'Somewhat effective', 'Quite effective', 'Extremely effective'],
    ['No growth at all', 'Slight growth', 'Some growth', 'A lot of growth', 'Substantial growth', "I don't know"],
    ['Never', 'Rarely', 'Occasionally', 'A moderate amount', 'A great deal'],
    ['Strongly disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly agree', 'Not applicable'],
    ['Not at all', 'Not really', 'Kind of', 'Definitely', 'Absolutely', "I didn't have any of these classes or activities"],
    ['0-10%', '11-25%', '26-50%', '51-75%', '76-100%'],
    ['I did not provide this support', '1 time during the program', '2-3 times during the program', 'Weekly', 'Daily'],
    ['Not helpful', 'Somewhat helpful', 'Helpful', 'Very Helpful'],
    ['Strongly disagree', 'Disagree', 'Neither agree nor disagree', 'Agree', 'Strongly agree'],
    ['Not Challenging', 'Somewhat Challenging', 'Very Challenging'],
    ['Selected', 'Not Selected'],
    ['Community or local organizations', 'Education newsletter or email blast', 'Grant database or website', 'Professional network', 'Social media', 'Other [please specify]'],
    ['Not at all', 'A little', 'Somewhat', 'Quite a bit', 'A great deal']
  ];

  const state = {
    workbook: null,
    fileName: '',
    sheetName: '',
    rows: [],
    columns: [],
    allRows: [],
    allColumns: [],
    columnOriginalHeaders: new Map(),
    hiddenAnalysisColumns: new Set(),
    columnStats: new Map(),
    rawColumnCount: 0,
    excludedChartColumns: [],
    eligibleChartColumns: [],
    selectedChartColumns: new Set(),
    chartGenerationInProgress: false,
    charts: [],
    nextChartNumber: 1,
    sources: [],
    reportResult: null,
    activeTab: 'charts',
    previewSearch: '',
    reportZoom: 1,
    linkedSurvey: {
      active: false,
      secondaryWorkbook: null,
      secondaryFileName: '',
      secondarySheet: '',
      result: null,
      question: '',
      column: ''
    }
  };

  const els = {
    fileInput: document.getElementById('fileInput'),
    fileDrop: document.getElementById('fileDrop'),
    uploadPanel: document.getElementById('uploadPanel'),
    statusMessage: document.getElementById('statusMessage'),
    fileDetails: document.getElementById('fileDetails'),
    sheetPickerWrap: document.getElementById('sheetPickerWrap'),
    sheetSelect: document.getElementById('sheetSelect'),
    fileStats: document.getElementById('fileStats'),
    datasetFileName: document.getElementById('datasetFileName'),
    mainTabs: document.getElementById('mainTabs'),
    tabButtons: Array.from(document.querySelectorAll('.tab-button')),
    dashboardSection: document.getElementById('dashboardSection'),
    chartGrid: document.getElementById('chartGrid'),
    addChartBtn: document.getElementById('addChartBtn'),
    generateSelectedChartsBtn: document.getElementById('generateSelectedChartsBtn'),
    selectAllCharts: document.getElementById('selectAllCharts'),
    chartColumnChecklist: document.getElementById('chartColumnChecklist'),
    selectedChartCount: document.getElementById('selectedChartCount'),
    chartGenerationStatus: document.getElementById('chartGenerationStatus'),
    chartEligibilityHint: document.getElementById('chartEligibilityHint'),
    emptyState: document.getElementById('emptyState'),
    sampleDataBtn: document.getElementById('sampleDataBtn'),
    changeSheetBtn: document.getElementById('changeSheetBtn'),
    replaceFileBtn: document.getElementById('replaceFileBtn'),
    clearDataBtn: document.getElementById('clearDataBtn'),
    chartTemplate: document.getElementById('chartCardTemplate'),
    filterTemplate: document.getElementById('filterTemplate'),
    reportSourceSelect: document.getElementById('reportSourceSelect'),
    publicSheetUrl: document.getElementById('publicSheetUrl'),
    loadPublicSheetBtn: document.getElementById('loadPublicSheetBtn'),
    reportStatus: document.getElementById('reportStatus'),
    reportQuestionSearch: document.getElementById('reportQuestionSearch'),
    selectAllQuestionsBtn: document.getElementById('selectAllQuestionsBtn'),
    selectMultipleChoiceBtn: document.getElementById('selectMultipleChoiceBtn'),
    clearQuestionsBtn: document.getElementById('clearQuestionsBtn'),
    selectedQuestionCount: document.getElementById('selectedQuestionCount'),
    reportNameInput: document.getElementById('reportNameInput'),
    reportDataSheetSelect: document.getElementById('reportDataSheetSelect'),
    questionChecklist: document.getElementById('questionChecklist'),
    reportColumnNote: document.getElementById('reportColumnNote'),
    primaryBreakdownSelect: document.getElementById('primaryBreakdownSelect'),
    reportFilterColumnSelect: document.getElementById('reportFilterColumnSelect'),
    reportFilterValues: document.getElementById('reportFilterValues'),
    reportFilterNote: document.getElementById('reportFilterNote'),
    generateReportBtn: document.getElementById('generateReportBtn'),
    downloadReportCsvBtn: document.getElementById('downloadReportCsvBtn'),
    downloadReportXlsxBtn: document.getElementById('downloadReportXlsxBtn'),
    reportOutputTitle: document.getElementById('reportOutputTitle'),
    reportOutputMeta: document.getElementById('reportOutputMeta'),
    reportContextBar: document.getElementById('reportContextBar'),
    distributionOutput: document.getElementById('distributionOutput'),
    dataPreviewSection: document.getElementById('dataPreviewSection'),
    previewSearch: document.getElementById('previewSearch'),
    previewStats: document.getElementById('previewStats'),
    columnProfiles: document.getElementById('columnProfiles'),
    previewResultCount: document.getElementById('previewResultCount'),
    dataPreviewTable: document.getElementById('dataPreviewTable'),
    zoomOutBtn: document.getElementById('zoomOutBtn'),
    zoomInBtn: document.getElementById('zoomInBtn'),
    zoomValue: document.getElementById('zoomValue'),
    fullscreenReportBtn: document.getElementById('fullscreenReportBtn'),
    toastRegion: document.getElementById('toastRegion'),
    confirmDialog: document.getElementById('confirmDialog'),
    confirmTitle: document.getElementById('confirmTitle'),
    confirmMessage: document.getElementById('confirmMessage'),
    confirmActionBtn: document.getElementById('confirmActionBtn'),
    linkedSurveyPanel: document.getElementById('linkedSurveyPanel'),
    toggleLinkedSurveyBtn: document.getElementById('toggleLinkedSurveyBtn'),
    linkedSurveySetup: document.getElementById('linkedSurveySetup'),
    linkHeadlineStatus: document.getElementById('linkHeadlineStatus'),
    linkPrimaryName: document.getElementById('linkPrimaryName'),
    linkPrimarySheet: document.getElementById('linkPrimarySheet'),
    linkPrimaryField: document.getElementById('linkPrimaryField'),
    linkSecondarySource: document.getElementById('linkSecondarySource'),
    linkSecondaryFileWrap: document.getElementById('linkSecondaryFileWrap'),
    linkSecondaryFile: document.getElementById('linkSecondaryFile'),
    linkSecondaryFileName: document.getElementById('linkSecondaryFileName'),
    linkSecondarySheet: document.getElementById('linkSecondarySheet'),
    linkSecondaryField: document.getElementById('linkSecondaryField'),
    linkQuestion: document.getElementById('linkQuestion'),
    createSurveyLinkBtn: document.getElementById('createSurveyLinkBtn'),
    clearSurveyLinkBtn: document.getElementById('clearSurveyLinkBtn'),
    linkValidation: document.getElementById('linkValidation'),
    linkStatusSummary: document.getElementById('linkStatusSummary'),
    linkWarnings: document.getElementById('linkWarnings'),
    linkDiagnosticActions: document.getElementById('linkDiagnosticActions'),
    viewUnmatchedBtn: document.getElementById('viewUnmatchedBtn'),
    downloadUnmatchedBtn: document.getElementById('downloadUnmatchedBtn'),
    viewDuplicatesBtn: document.getElementById('viewDuplicatesBtn'),
    linkDetailsDialog: document.getElementById('linkDetailsDialog'),
    linkDetailsTitle: document.getElementById('linkDetailsTitle'),
    linkDetailsBody: document.getElementById('linkDetailsBody'),
    closeLinkDetailsBtn: document.getElementById('closeLinkDetailsBtn')
  };

  Chart.register(ChartDataLabels);
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
  Chart.defaults.color = CHART_MUTED;

  els.fileInput.addEventListener('change', handleFileUpload);
  els.sheetSelect.addEventListener('change', () => loadSheet(els.sheetSelect.value));
  els.addChartBtn.addEventListener('click', () => addChart());
  els.generateSelectedChartsBtn.addEventListener('click', generateSelectedCharts);
  els.selectAllCharts.addEventListener('change', event => {
    state.selectedChartColumns = event.target.checked ? new Set(state.eligibleChartColumns) : new Set();
    renderChartSelection();
  });
  els.loadPublicSheetBtn.addEventListener('click', loadPublicGoogleSheet);
  els.reportSourceSelect.addEventListener('change', renderReportControls);
  els.reportDataSheetSelect.addEventListener('change', renderReportColumns);
  els.primaryBreakdownSelect.addEventListener('change', syncBreakdownQuestionSelection);
  els.reportFilterColumnSelect.addEventListener('change', renderReportFilterValues);
  els.generateReportBtn.addEventListener('click', generateDistributionReport);
  els.downloadReportCsvBtn.addEventListener('click', downloadDistributionCsv);
  els.downloadReportXlsxBtn.addEventListener('click', downloadDistributionXlsx);
  els.sampleDataBtn.addEventListener('click', loadSampleData);
  els.changeSheetBtn.addEventListener('click', showSheetPicker);
  els.replaceFileBtn.addEventListener('click', () => els.fileInput.click());
  els.clearDataBtn.addEventListener('click', async () => {
    if (await requestConfirmation('Clear this dataset?', 'Charts, filters, and the generated report will be removed. Your original file will not be changed.', 'Clear data')) {
      resetDataset();
      els.fileInput.value = '';
      showStatus('Dataset cleared. Choose another source when you are ready.', '');
      showToast('Dataset cleared.');
    }
  });
  els.tabButtons.forEach(button => button.addEventListener('click', () => setActiveTab(button.dataset.tab)));
  els.previewSearch.addEventListener('input', event => {
    state.previewSearch = event.target.value;
    renderDataPreview();
  });
  els.reportQuestionSearch.addEventListener('input', filterReportQuestions);
  els.selectAllQuestionsBtn.addEventListener('click', () => setVisibleReportQuestions(true));
  els.selectMultipleChoiceBtn.addEventListener('click', () => setVisibleReportQuestions(true));
  els.clearQuestionsBtn.addEventListener('click', () => setVisibleReportQuestions(false));
  document.querySelectorAll('[data-report-mode]').forEach(button => button.addEventListener('click', () => setReportMode(button.dataset.reportMode)));
  document.querySelectorAll('[data-density]').forEach(button => button.addEventListener('click', () => setReportDensity(button.dataset.density)));
  els.zoomOutBtn.addEventListener('click', () => setReportZoom(state.reportZoom - 0.1));
  els.zoomInBtn.addEventListener('click', () => setReportZoom(state.reportZoom + 0.1));
  els.fullscreenReportBtn.addEventListener('click', toggleReportFullscreen);
  els.toggleLinkedSurveyBtn.addEventListener('click', toggleLinkedSurveySetup);
  els.linkSecondarySource.addEventListener('change', () => {
    clearSurveyLink(false);
    renderLinkedSurveySource();
  });
  els.linkSecondaryFile.addEventListener('change', loadSecondarySurveyFile);
  els.linkPrimaryField.addEventListener('change', () => clearSurveyLink(false));
  els.linkSecondarySheet.addEventListener('change', () => {
    clearSurveyLink(false);
    renderLinkedSurveyFields();
  });
  els.linkSecondaryField.addEventListener('change', () => {
    clearSurveyLink(false);
    renderLinkedSurveyFields();
  });
  els.createSurveyLinkBtn.addEventListener('click', createSurveyLink);
  els.clearSurveyLinkBtn.addEventListener('click', () => clearSurveyLink(true));
  els.linkQuestion.addEventListener('change', applyLinkedQuestion);
  els.viewUnmatchedBtn.addEventListener('click', showUnmatchedRecords);
  els.downloadUnmatchedBtn.addEventListener('click', downloadUnmatchedRecords);
  els.viewDuplicatesBtn.addEventListener('click', showDuplicateValues);
  els.closeLinkDetailsBtn.addEventListener('click', () => els.linkDetailsDialog.close());

  ['dragenter', 'dragover'].forEach(type => els.fileDrop.addEventListener(type, event => {
    event.preventDefault();
    els.fileDrop.classList.add('is-dragging');
  }));
  ['dragleave', 'drop'].forEach(type => els.fileDrop.addEventListener(type, event => {
    event.preventDefault();
    els.fileDrop.classList.remove('is-dragging');
  }));
  els.fileDrop.addEventListener('drop', event => {
    const file = event.dataTransfer.files[0];
    if (file) loadFile(file);
  });
  document.addEventListener('keydown', handleGlobalKeydown);

  async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    await loadFile(file);
  }

  function yieldToBrowser() {
    return new Promise(resolve => window.setTimeout(resolve, 0));
  }

  function parseCsvWorkbook(buffer) {
    const bytes = new Uint8Array(buffer);
    const encoding = bytes[0] === 0xff && bytes[1] === 0xfe ? 'utf-16le'
      : (bytes[0] === 0xfe && bytes[1] === 0xff ? 'utf-16be' : 'utf-8');
    const text = new TextDecoder(encoding).decode(buffer).replace(/^\uFEFF/, '');
    const rows = [];
    let row = [];
    let field = '';
    let quoted = false;

    for (let index = 0; index < text.length; index += 1) {
      const character = text[index];
      if (quoted) {
        if (character === '"' && text[index + 1] === '"') {
          field += '"';
          index += 1;
        } else if (character === '"') {
          quoted = false;
        } else {
          field += character;
        }
      } else if (character === '"' && field === '') {
        quoted = true;
      } else if (character === ',') {
        row.push(field);
        field = '';
      } else if (character === '\n' || character === '\r') {
        if (character === '\r' && text[index + 1] === '\n') index += 1;
        row.push(field);
        if (row.some(value => value !== '')) rows.push(row);
        row = [];
        field = '';
      } else {
        field += character;
      }
    }

    if (field || row.length) {
      row.push(field);
      if (row.some(value => value !== '')) rows.push(row);
    }

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), 'Sheet1');
    return workbook;
  }

  function parseWorkbook(buffer, extension = '') {
    if (extension === 'csv') return Promise.resolve(parseCsvWorkbook(buffer));
    if (typeof Worker === 'undefined') return Promise.resolve(XLSX.read(buffer, WORKBOOK_PARSE_OPTIONS));

    return new Promise((resolve, reject) => {
      let settled = false;
      const worker = new Worker('data-worker.js?v=20260903-2');
      const finish = callback => value => {
        if (settled) return;
        settled = true;
        worker.terminate();
        callback(value);
      };
      const fallbackToMainThread = finish(value => {
        try {
          resolve(XLSX.read(buffer, WORKBOOK_PARSE_OPTIONS));
        } catch (error) {
          reject(error);
        }
      });

      worker.onmessage = event => {
        if (event.data?.type === 'success') {
          finish(resolve)(event.data.workbook);
        } else {
          fallbackToMainThread();
        }
      };
      worker.onerror = fallbackToMainThread;

      try {
        const transferableBuffer = buffer.slice(0);
        worker.postMessage({ buffer: transferableBuffer }, [transferableBuffer]);
      } catch (error) {
        fallbackToMainThread(error);
      }
    });
  }

  async function loadFile(file) {

    const extension = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(extension)) {
      showStatus('Please choose an .xlsx, .xls, or .csv file.', 'error');
      return;
    }

    showStatus('Reading your file...', 'loading');
    setButtonLoading(els.replaceFileBtn, true, 'Reading…');
    resetDataset();

    try {
      const buffer = await file.arrayBuffer();
      const workbook = await parseWorkbook(buffer, extension);
      showStatus('Analyzing rows and columns...', 'loading');
      await yieldToBrowser();

      if (!workbook.SheetNames.length) {
        throw new Error('No sheets were found in this file.');
      }

      upsertReportSource(UPLOADED_SOURCE_ID, `Uploaded: ${file.name}`, workbook);
      activateWorkbook(workbook, file.name, workbook.SheetNames[0]);

      if (state.rows.length > 50000) {
        showStatus('Large file warning: this app is designed for normal files up to about 50,000 rows. It may take longer to update charts.', 'warning');
        showToast('Large dataset loaded. Some updates may take longer.', 'warning');
      } else {
        showStatus('File loaded. Your data stays in this browser.', '');
        showToast(`${file.name} loaded successfully.`);
      }
    } catch (error) {
      console.error(error);
      resetDataset();
      showStatus('This file could not be opened. It may be damaged or in an unsupported format.', 'error');
      showToast('The file could not be opened.', 'error');
    } finally {
      setButtonLoading(els.replaceFileBtn, false);
    }
  }

  function activateWorkbook(workbook, fileName, sheetName) {
    state.workbook = workbook;
    state.fileName = fileName;
    const surveySheets = getSurveySheetNames(workbook.SheetNames);
    populateSheetSelector(surveySheets);
    const initialSheet = surveySheets.includes(sheetName) ? sheetName : (surveySheets[0] || sheetName || workbook.SheetNames[0]);
    loadSheet(initialSheet);
  }

  function populateSheetSelector(sheetNames) {
    els.sheetSelect.innerHTML = sheetNames.map(name => `<option value="${escapeAttr(name)}">${escapeHtml(name)}</option>`).join('');
    els.sheetPickerWrap.classList.toggle('hidden', sheetNames.length <= 1);
  }

  function loadSheet(sheetName) {
    const rawRows = getSheetMatrix(state.workbook, sheetName);

    state.sheetName = sheetName;
    state.rows = [];
    state.columns = [];
    state.allRows = [];
    state.allColumns = [];
    state.columnOriginalHeaders = new Map();
    state.hiddenAnalysisColumns = new Set();
    state.columnStats = new Map();
    state.rawColumnCount = 0;
    state.excludedChartColumns = [];
    state.eligibleChartColumns = [];
    state.selectedChartColumns = new Set();
    const hadLinkedSurvey = Boolean(state.linkedSurvey.result);
    resetLinkedSurveyMatch();
    if (hadLinkedSurvey) invalidateGeneratedReport();

    if (!rawRows.length) {
      renderDataset();
      showStatus('The selected sheet does not contain usable rows.', 'warning');
      return;
    }

    const headerDetails = getSheetHeaderDetails(state.workbook, sheetName, 0);
    const allColumns = headerDetails.columns;
    const allRows = rawRows.slice(1)
      .map(row => {
        const record = {};
        allColumns.forEach((column, index) => {
          record[column] = row[index] === undefined ? '' : row[index];
        });
        return record;
      });
    const columnStats = buildColumnStats(allRows, allColumns);
    const chartColumns = ChartRules.getSelectableChartColumns(allColumns);

    state.rawColumnCount = allColumns.length;
    state.excludedChartColumns = allColumns.filter(column => !chartColumns.includes(column));
    state.eligibleChartColumns = chartColumns;
    state.selectedChartColumns = new Set();
    state.allColumns = allColumns;
    state.columnOriginalHeaders = headerDetails.originalByColumn;
    state.allRows = allRows;
    state.columnStats = columnStats;
    updateAnalysisColumns();

    state.charts = [];
    state.nextChartNumber = 1;
    renderDataset();
    if (state.columns.length && state.rows.length) setActiveTab('charts');
  }

  function makeUniqueHeaders(headerRow) {
    const seen = new Map();
    return headerRow.map((header, index) => {
      const base = normalizeValue(header) || `Column ${index + 1}`;
      const count = seen.get(base) || 0;
      seen.set(base, count + 1);
      return count ? `${base} (${count + 1})` : base;
    });
  }

  function renderDataset() {
    const hasData = state.rows.length > 0 && state.columns.length > 0;
    const hasDataset = Boolean(state.workbook && state.allRows.length);
    els.emptyState.classList.toggle('hidden', hasDataset);
    els.fileDetails.classList.toggle('hidden', !state.workbook);
    els.mainTabs.classList.toggle('hidden', !hasDataset);
    els.uploadPanel.classList.toggle('is-compact', hasDataset);
    els.addChartBtn.disabled = !hasData;
    els.generateSelectedChartsBtn.disabled = !hasData || state.chartGenerationInProgress;

    if (hasDataset) setActiveTab(state.activeTab || 'charts');
    else {
      els.dashboardSection.classList.add('hidden');
      document.getElementById('distributionSection').classList.add('hidden');
      els.dataPreviewSection.classList.add('hidden');
    }

    renderFileStats();
    renderAllCharts();
    renderReportControls();
    renderDataPreview();
    renderLinkedSurveyPanel();
  }

  function renderFileStats() {
    if (!state.workbook) return;
    els.datasetFileName.textContent = state.fileName;
    const stats = [
      ['Sheet', state.sheetName || 'None'],
      ['Rows', formatNumber(state.allRows.length)],
      ['Columns', formatNumber(state.rawColumnCount)],
      ['Usable questions', formatNumber(state.columns.length)]
    ];
    els.fileStats.innerHTML = stats.map(([label, value]) => `
      <div class="stat">
        <span>${escapeHtml(label)}</span>
        <strong>${escapeHtml(value)}</strong>
      </div>
    `).join('');
  }

  function addChart(sourceConfig) {
    const chart = createChartConfig(sourceConfig);
    state.charts.push(chart);
    renderAllCharts();
  }

  async function generateSelectedCharts() {
    if (state.chartGenerationInProgress) return;
    const selectedColumns = ChartRules.getSelectedChartColumns(state.eligibleChartColumns, state.selectedChartColumns);
    if (!selectedColumns.length) {
      showChartGenerationStatus('Select at least one response column before generating charts.', 'warning');
      showToast('Select at least one chart column.', 'warning');
      return;
    }

    const missingColumns = ChartRules.getMissingChartColumns(selectedColumns, state.charts);
    if (!missingColumns.length) {
      showChartGenerationStatus('All selected columns already have charts.', '');
      showToast('Every selected column already has a chart.');
      return;
    }

    state.chartGenerationInProgress = true;
    setButtonLoading(els.generateSelectedChartsBtn, true, 'Generating…');
    showChartGenerationStatus(`Generating 0 of ${missingColumns.length} charts…`, 'loading');

    try {
      state.charts.forEach(chart => {
        if (selectedColumns.includes(chart.primaryColumn) && /^Chart \d+$/.test(chart.title)) {
          chart.title = getActiveColumnDisplayName(chart.primaryColumn);
        }
      });

      for (let index = 0; index < missingColumns.length; index += 1) {
        const column = missingColumns[index];
        const chart = createChartConfig();
        chart.title = getActiveColumnDisplayName(column);
        chart.primaryColumn = column;
        chart.topMode = String(CHART_RESPONSE_DISPLAY_LIMIT);
        chart.includeBlanks = false;
        chart.settingsCollapsed = true;
        state.charts.push(chart);
        showChartGenerationStatus(`Generating ${index + 1} of ${missingColumns.length} charts…`, 'loading');
        await new Promise(resolve => window.setTimeout(resolve, 0));
      }

      renderAllCharts();
      showChartGenerationStatus(`${missingColumns.length} chart${missingColumns.length === 1 ? '' : 's'} generated.`, '');
      showToast(`${missingColumns.length} chart${missingColumns.length === 1 ? '' : 's'} generated.`);
    } catch (error) {
      console.error(error);
      showChartGenerationStatus(error.message || 'Could not generate the selected charts.', 'error');
      showToast(error.message || 'Could not generate the selected charts.', 'error');
    } finally {
      state.chartGenerationInProgress = false;
      setButtonLoading(els.generateSelectedChartsBtn, false);
      updateChartGenerationControls();
    }
  }

  function renderChartEligibilitySummary() {
    if (!els.chartEligibilityHint) return;
    const eligibleCount = state.eligibleChartColumns.length;
    els.chartEligibilityHint.textContent = state.allRows.length
      ? `${formatNumber(eligibleCount)} selectable question${eligibleCount === 1 ? '' : 's'} · choose any response column · metadata excluded`
      : '';
    renderChartSelection();
  }

  function renderChartSelection() {
    if (!els.chartColumnChecklist) return;
    const eligibleColumns = state.eligibleChartColumns;
    const selected = new Set(eligibleColumns.filter(column => state.selectedChartColumns.has(column)));
    state.selectedChartColumns = selected;

    if (!eligibleColumns.length) {
      els.chartColumnChecklist.innerHTML = '<div class="checklist-empty">No response columns were found in this dataset.</div>';
    } else {
      els.chartColumnChecklist.innerHTML = eligibleColumns.map((column, index) => {
        const id = `chart-column-${hashString(column)}-${index}`;
        const displayName = getActiveColumnDisplayName(column);
        return `<label for="${id}" title="${escapeAttr(getColumnOptionLabel(column))}">
          <input id="${id}" type="checkbox" value="${escapeAttr(column)}" ${selected.has(column) ? 'checked' : ''}>
          <span>${escapeHtml(displayName)}</span>
          <small>${escapeHtml(getColumnStatsLabel(column))}</small>
        </label>`;
      }).join('');
      els.chartColumnChecklist.querySelectorAll('input[type="checkbox"]').forEach(input => {
        input.addEventListener('change', event => {
          if (event.target.checked) state.selectedChartColumns.add(event.target.value);
          else state.selectedChartColumns.delete(event.target.value);
          updateChartGenerationControls();
        });
      });
    }

    const selectedCount = selected.size;
    els.selectAllCharts.checked = eligibleColumns.length > 0 && selectedCount === eligibleColumns.length;
    els.selectAllCharts.indeterminate = selectedCount > 0 && selectedCount < eligibleColumns.length;
    updateChartGenerationControls();
  }

  function updateChartGenerationControls() {
    if (!els.selectedChartCount || !els.generateSelectedChartsBtn) return;
    const selectedCount = state.selectedChartColumns.size;
    els.selectedChartCount.textContent = `${selectedCount} chart${selectedCount === 1 ? '' : 's'} selected`;
    if (els.selectAllCharts) {
      els.selectAllCharts.checked = state.eligibleChartColumns.length > 0
        && selectedCount === state.eligibleChartColumns.length;
      els.selectAllCharts.indeterminate = selectedCount > 0
        && selectedCount < state.eligibleChartColumns.length;
    }
    els.generateSelectedChartsBtn.disabled = state.chartGenerationInProgress
      || !state.allRows.length
      || selectedCount === 0;
  }

  function showChartGenerationStatus(message, type = '') {
    if (!els.chartGenerationStatus) return;
    els.chartGenerationStatus.textContent = message;
    els.chartGenerationStatus.className = `status-message ${type}`.trim();
  }

  function createChartConfig(sourceConfig) {
    const firstColumn = state.columns[0] || '';
    const config = sourceConfig ? cloneChartConfig(sourceConfig) : {
      id: makeId(),
      title: `Chart ${state.nextChartNumber}`,
      collapsed: false,
      settingsCollapsed: false,
      primaryColumn: firstColumn,
      compareColumn: '',
      chartType: 'auto',
      compareType: 'grouped',
      compareValueMode: 'counts',
      sortMode: 'desc',
      topMode: String(CHART_RESPONSE_DISPLAY_LIMIT),
      showCounts: true,
      showPercentages: true,
      includeBlanks: false,
      filters: [],
      summarySearch: '',
      selectedResponses: new Set(),
      hiddenResponses: new Set(),
      merges: []
    };

    config.id = makeId();
    config.title = sourceConfig ? `${sourceConfig.title} copy` : config.title;
    state.nextChartNumber += 1;
    return config;
  }

  function cloneChartConfig(config) {
    return {
      ...config,
      filters: config.filters.map(filter => ({
        id: makeId(),
        column: filter.column,
        selected: new Set(filter.selected),
        search: filter.search || ''
      })),
      selectedResponses: new Set(),
      hiddenResponses: new Set(config.hiddenResponses),
      merges: config.merges.map(merge => ({ name: merge.name, sources: new Set(merge.sources) })),
      chartInstance: null
    };
  }

  function renderAllCharts() {
    renderChartEligibilitySummary();
    els.chartGrid.innerHTML = '';
    if (!state.charts.length && state.allRows.length) {
      const message = state.eligibleChartColumns.length
        ? '<strong>No charts generated yet.</strong><span>Select one or more response columns above, then choose Generate Selected Charts.</span>'
        : '<strong>No selectable chart columns found.</strong><span>Metadata columns are excluded; choose any remaining column to generate a chart.</span>';
      els.chartGrid.innerHTML = `<div class="chart-results-empty">${message}</div>`;
      return;
    }
    state.charts.forEach(chart => {
      const card = renderChartCard(chart);
      els.chartGrid.appendChild(card);
      updateChartCard(chart, card);
    });
  }

  function renderChartCard(chart) {
    const fragment = els.chartTemplate.content.cloneNode(true);
    const card = fragment.querySelector('.chart-card');
    card.dataset.chartId = chart.id;
    card.classList.toggle('collapsed', chart.collapsed);
    card.classList.toggle('settings-collapsed', chart.settingsCollapsed);

    const titleInput = card.querySelector('.chart-title-input');
    titleInput.value = chart.title;
    card.querySelector('.chart-title').textContent = chart.title;
    titleInput.addEventListener('input', event => {
      chart.title = event.target.value || 'Untitled chart';
      card.querySelector('.chart-title').textContent = chart.title;
    });
    titleInput.addEventListener('blur', () => finishTitleEdit(card));
    titleInput.addEventListener('keydown', event => {
      if (event.key === 'Enter') finishTitleEdit(card);
      if (event.key === 'Escape') {
        titleInput.value = chart.title;
        finishTitleEdit(card);
      }
    });
    card.querySelector('.edit-title').addEventListener('click', () => startTitleEdit(card));

    card.querySelector('.duplicate-chart').addEventListener('click', () => addChart(chart));
    card.querySelector('.collapse-chart').addEventListener('click', () => {
      chart.collapsed = !chart.collapsed;
      renderAllCharts();
    });
    card.querySelector('.delete-chart').addEventListener('click', async () => {
      if (await requestConfirmation(`Delete “${chart.title}”?`, 'This chart and its settings will be removed. Your dataset will not be changed.', 'Delete chart')) {
        if (chart.chartInstance) chart.chartInstance.destroy();
        state.charts = state.charts.filter(item => item.id !== chart.id);
        renderAllCharts();
        showToast('Chart deleted.');
      }
    });
    card.querySelector('.toggle-settings').addEventListener('click', () => {
      chart.settingsCollapsed = !chart.settingsCollapsed;
      card.classList.toggle('settings-collapsed', chart.settingsCollapsed);
    });
    card.querySelector('.expand-chart').addEventListener('click', () => toggleChartExpanded(card));

    populateColumnSelect(card.querySelector('.primary-column'), chart.primaryColumn, false);
    populateColumnSelect(card.querySelector('.compare-column'), chart.compareColumn, true);

    bindSelect(card, '.primary-column', chart, 'primaryColumn');
    bindSelect(card, '.compare-column', chart, 'compareColumn');
    bindSelect(card, '.chart-type', chart, 'chartType');
    bindSelect(card, '.compare-type', chart, 'compareType');
    bindSelect(card, '.compare-value-mode', chart, 'compareValueMode');
    bindSelect(card, '.sort-mode', chart, 'sortMode');
    bindSelect(card, '.top-mode', chart, 'topMode');
    bindCheckbox(card, '.show-counts', chart, 'showCounts');
    bindCheckbox(card, '.show-percentages', chart, 'showPercentages');
    bindCheckbox(card, '.include-blanks', chart, 'includeBlanks');

    const questionSearch = card.querySelector('.question-search');
    questionSearch.addEventListener('input', event => {
      filterColumnOptions(card.querySelector('.primary-column'), event.target.value, chart.primaryColumn, false);
    });
    card.querySelector('.compare-action').addEventListener('click', () => {
      if (chart.settingsCollapsed) {
        chart.settingsCollapsed = false;
        card.classList.remove('settings-collapsed');
      }
      card.querySelector('.compare-column').focus();
    });
    card.querySelectorAll('.show-response-editor').forEach(button => button.addEventListener('click', () => {
      card.querySelector('.response-tools').classList.remove('hidden');
      button.closest('details').removeAttribute('open');
      card.querySelector('.summary-table').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }));
    card.querySelector('.close-response-editor').addEventListener('click', () => card.querySelector('.response-tools').classList.add('hidden'));

    card.querySelector('.summary-search').value = chart.summarySearch;
    card.querySelector('.summary-search').addEventListener('input', event => {
      chart.summarySearch = event.target.value;
      updateChartCard(chart, card);
    });

    card.querySelector('.add-filter').addEventListener('click', () => {
      chart.filters.push({
        id: makeId(),
        column: state.columns[0] || '',
        selected: new Set(),
        search: ''
      });
      renderAllCharts();
    });

    card.querySelector('.clear-filters').addEventListener('click', () => {
      chart.filters = [];
      renderAllCharts();
    });

    card.querySelector('.merge-selected').addEventListener('click', () => {
      const name = normalizeValue(card.querySelector('.merge-name').value);
      if (!name || !chart.selectedResponses.size) return;
      const sources = new Set();
      chart.selectedResponses.forEach(label => {
        sources.add(label);
        chart.merges.forEach(merge => {
          if (merge.name === label) merge.sources.forEach(source => sources.add(source));
        });
      });
      chart.merges = chart.merges.filter(merge => !chart.selectedResponses.has(merge.name));
      chart.merges.push({ name, sources });
      chart.selectedResponses.clear();
      card.querySelector('.merge-name').value = '';
      updateChartCard(chart, card);
    });

    card.querySelector('.hide-selected').addEventListener('click', () => {
      chart.selectedResponses.forEach(value => chart.hiddenResponses.add(value));
      chart.selectedResponses.clear();
      updateChartCard(chart, card);
    });

    card.querySelector('.reset-responses').addEventListener('click', () => {
      chart.hiddenResponses.clear();
      chart.merges = [];
      chart.selectedResponses.clear();
      updateChartCard(chart, card);
      showToast('Original responses restored.');
    });

    card.querySelector('.export-png').addEventListener('click', () => exportChartPng(chart));
    card.querySelector('.export-summary').addEventListener('click', () => exportSummaryCsv(chart));
    card.querySelector('.export-filtered').addEventListener('click', () => exportFilteredDataCsv(chart));

    return card;
  }

  function bindSelect(card, selector, chart, key) {
    const input = card.querySelector(selector);
    input.value = chart[key];
    input.addEventListener('change', event => {
      chart[key] = event.target.value;
      chart.selectedResponses.clear();
      renderAllCharts();
    });
  }

  function bindCheckbox(card, selector, chart, key) {
    const input = card.querySelector(selector);
    input.checked = chart[key];
    input.addEventListener('change', event => {
      chart[key] = event.target.checked;
      updateChartCard(chart, card);
    });
  }

  function populateColumnSelect(select, selectedValue, includeNone) {
    const options = includeNone ? ['<option value="">No comparison</option>'] : [];
    options.push(...state.columns.map(column => `<option value="${escapeAttr(column)}">${escapeHtml(getColumnOptionLabel(column))}</option>`));
    select.innerHTML = options.join('');
    if (selectedValue && state.columns.includes(selectedValue)) select.value = selectedValue;
    else select.value = includeNone ? '' : (state.columns[0] || '');
  }

  function updateChartCard(chart, card) {
    if (!state.rows.length || !state.columns.length) return;
    if (!state.columns.includes(chart.primaryColumn)) chart.primaryColumn = state.columns[0] || '';
    if (chart.compareColumn && !state.columns.includes(chart.compareColumn)) chart.compareColumn = '';

    const isComparison = Boolean(chart.compareColumn);
    card.querySelectorAll('.single-setting').forEach(el => el.classList.toggle('hidden', isComparison));
    card.querySelectorAll('.compare-setting').forEach(el => el.classList.toggle('hidden', !isComparison));
    card.querySelector('.collapse-chart').textContent = chart.collapsed ? '+' : '−';
    card.querySelector('.collapse-chart').setAttribute('aria-label', chart.collapsed ? 'Expand chart' : 'Collapse chart');

    renderFilters(chart, card);
    renderActiveFilterChips(chart, card);

    const filteredRows = applyFilters(state.rows, chart.filters);
    card.querySelector('.row-count').textContent = `Showing ${formatNumber(filteredRows.length)} of ${formatNumber(state.rows.length)} rows`;

    const result = isComparison
      ? buildComparisonResult(filteredRows, chart)
      : buildSingleColumnResult(filteredRows, chart);

    const validResponses = filteredRows.filter(row => getResponseLabels(row[chart.primaryColumn]).some(label => label !== NO_RESPONSE)).length;
    const categoryTotal = result.type === 'single' ? result.items.length : result.labels.length;
    card.querySelector('.valid-response-count').textContent = `${formatNumber(validResponses)} valid response${validResponses === 1 ? '' : 's'}`;
    card.querySelector('.category-count').textContent = `${formatNumber(categoryTotal)} categor${categoryTotal === 1 ? 'y' : 'ies'}`;
    const activeFilterCount = chart.filters.filter(filter => filter.selected.size).length;
    card.querySelector('.analysis-summary').textContent = chart.compareColumn
      ? `Compared by ${getActiveColumnDisplayName(chart.compareColumn)}${activeFilterCount ? ` · ${activeFilterCount} filter${activeFilterCount === 1 ? '' : 's'}` : ''}`
      : (activeFilterCount ? `${activeFilterCount} active filter${activeFilterCount === 1 ? '' : 's'}` : 'No filters or comparison');
    card.querySelector('.question-detail').textContent = getColumnOptionLabel(chart.primaryColumn);

    renderChart(chart, card, result);
    renderSummaryTable(chart, card, result);
  }

  function renderFilters(chart, card) {
    const list = card.querySelector('.filters-list');
    list.innerHTML = '';

    chart.filters.forEach(filter => {
      if (!filter.column || !state.columns.includes(filter.column)) filter.column = state.columns[0] || '';
      const fragment = els.filterTemplate.content.cloneNode(true);
      const filterCard = fragment.querySelector('.filter-card');
      const columnSelect = filterCard.querySelector('.filter-column');
      populateColumnSelect(columnSelect, filter.column, false);

      columnSelect.addEventListener('change', event => {
        filter.column = event.target.value;
        filter.selected.clear();
        filter.search = '';
        renderAllCharts();
      });

      filterCard.querySelector('.remove-filter').addEventListener('click', () => {
        chart.filters = chart.filters.filter(item => item.id !== filter.id);
        renderAllCharts();
      });

      const search = filterCard.querySelector('.filter-search');
      search.value = filter.search || '';
      search.addEventListener('input', event => {
        filter.search = event.target.value;
        renderAllCharts();
      });

      const values = getUniqueValues(state.rows, filter.column);
      const searchText = (filter.search || '').toLowerCase();
      const visibleValues = values.filter(value => value.toLowerCase().includes(searchText));
      const valuesWrap = filterCard.querySelector('.filter-values');
      valuesWrap.innerHTML = visibleValues.map(value => {
        const id = `${filter.id}-${hashString(value)}`;
        return `
          <label for="${id}">
            <input id="${id}" type="checkbox" value="${escapeAttr(value)}" ${filter.selected.has(value) ? 'checked' : ''}>
            <span>${escapeHtml(value)}</span>
          </label>
        `;
      }).join('') || '<p class="muted">No matching values.</p>';

      valuesWrap.querySelectorAll('input[type="checkbox"]').forEach(input => {
        input.addEventListener('change', event => {
          if (event.target.checked) filter.selected.add(event.target.value);
          else filter.selected.delete(event.target.value);
          updateChartCard(chart, card);
        });
      });

      list.appendChild(filterCard);
    });
  }

  function renderActiveFilterChips(chart, card) {
    const container = card.querySelector('.active-filters');
    const activeFilters = chart.filters.filter(filter => filter.column && filter.selected.size);
    container.innerHTML = activeFilters.map(filter => {
      const selected = Array.from(filter.selected);
      const preview = selected.slice(0, 3).join(', ');
      const extra = selected.length > 3 ? ` +${selected.length - 3}` : '';
      const displayName = getActiveColumnDisplayName(filter.column);
      return `<div class="filter-chip"><span title="${escapeAttr(`${displayName}: ${selected.join(', ')}`)}"><strong>${escapeHtml(displayName)}:</strong> ${escapeHtml(preview)}${escapeHtml(extra)}</span><button type="button" data-filter-id="${escapeAttr(filter.id)}" aria-label="Remove ${escapeAttr(displayName)} filter">×</button></div>`;
    }).join('');
    container.querySelectorAll('button[data-filter-id]').forEach(button => button.addEventListener('click', () => {
      chart.filters = chart.filters.filter(filter => filter.id !== button.dataset.filterId);
      renderAllCharts();
    }));
  }

  function applyFilters(rows, filters) {
    const activeFilters = filters.filter(filter => filter.column && filter.selected.size);
    if (!activeFilters.length) return rows;

    return rows.filter(row => activeFilters.every(filter => {
      return getResponseLabels(row[filter.column]).some(value => filter.selected.has(value));
    }));
  }

  function buildSingleColumnResult(rows, chart) {
    const counts = new Map();
    let nonBlank = 0;
    const columnLabels = getColumnNonBlankLabels(rows, chart.primaryColumn);
    const useYesNoLabels = ChartRules.shouldUseYesNoLabels(columnLabels);

    rows.forEach(row => {
      const originalLabels = getDisplayResponseLabelsForColumn(row[chart.primaryColumn], columnLabels, useYesNoLabels);
      if (originalLabels.some(label => label !== NO_RESPONSE)) nonBlank += 1;
      originalLabels.forEach(originalLabel => {
        if (!chart.includeBlanks && originalLabel === NO_RESPONSE) return;
        const label = getMergedLabel(originalLabel, chart.merges);
        if (chart.hiddenResponses.has(label)) return;
        counts.set(label, (counts.get(label) || 0) + 1);
      });
    });

    let items = Array.from(counts, ([response, count]) => ({
      response,
      count,
      rowPercent: nonBlank ? roundOne((count / nonBlank) * 100) : 0,
      nonBlankPercent: nonBlank ? roundOne((count / nonBlank) * 100) : 0
    }));

    items = sortItems(items, chart.sortMode);
    items = applyTopGrouping(items, chart.topMode, nonBlank, nonBlank);

    return {
      type: 'single',
      rows,
      labels: items.map(item => item.response),
      values: items.map(item => item.count),
      items
    };
  }

  function buildComparisonResult(rows, chart) {
    const matrix = new Map();
    const comparisonLabels = new Set();
    const rawPrimaryRespondentTotals = new Map();
    const rawCompareRespondentTotals = new Map();
    let total = 0;
    let respondentTotal = 0;
    const primaryColumnLabels = getColumnNonBlankLabels(rows, chart.primaryColumn);
    const compareColumnLabels = getColumnNonBlankLabels(rows, chart.compareColumn);
    const usePrimaryYesNoLabels = ChartRules.shouldUseYesNoLabels(primaryColumnLabels);
    const useCompareYesNoLabels = ChartRules.shouldUseYesNoLabels(compareColumnLabels);

    rows.forEach(row => {
      const primaries = uniqueList(getDisplayResponseLabelsForColumn(row[chart.primaryColumn], primaryColumnLabels, usePrimaryYesNoLabels).map(label => getMergedLabel(label, chart.merges)))
        .filter(primary => !chart.hiddenResponses.has(primary));
      const comparisons = uniqueList(getDisplayResponseLabelsForColumn(row[chart.compareColumn], compareColumnLabels, useCompareYesNoLabels));
      if (!chart.includeBlanks && (primaries.includes(NO_RESPONSE) || comparisons.includes(NO_RESPONSE))) return;
      if (!primaries.length || !comparisons.length) return;

      primaries.forEach(primary => {
        rawPrimaryRespondentTotals.set(primary, (rawPrimaryRespondentTotals.get(primary) || 0) + 1);
      });
      comparisons.forEach(comparison => {
        rawCompareRespondentTotals.set(comparison, (rawCompareRespondentTotals.get(comparison) || 0) + 1);
      });
      respondentTotal += 1;
      primaries.forEach(primary => comparisons.forEach(comparison => {
        if (!matrix.has(primary)) matrix.set(primary, new Map());
        matrix.get(primary).set(comparison, (matrix.get(primary).get(comparison) || 0) + 1);
        comparisonLabels.add(comparison);
        total += 1;
      }));
    });

    const rawPrimaryLabels = Array.from(matrix.keys());
    const rawCompareLabels = Array.from(comparisonLabels);
    const rawPrimaryTotals = new Map(rawPrimaryLabels.map(label => [
      label,
      rawCompareLabels.reduce((sum, compare) => sum + (matrix.get(label).get(compare) || 0), 0)
    ]));
    const rawCompareTotals = new Map(rawCompareLabels.map(label => [
      label,
      rawPrimaryLabels.reduce((sum, primary) => sum + (matrix.get(primary).get(label) || 0), 0)
    ]));
    const primaryLabels = capLabels(rawPrimaryLabels, rawPrimaryTotals, TABLE_ROW_LIMIT);
    const compareLabels = capLabels(rawCompareLabels, rawCompareTotals, TABLE_ROW_LIMIT);
    const cappedMatrix = new Map(primaryLabels.map(label => [label, new Map(compareLabels.map(compare => [compare, 0]))]));

    rawPrimaryLabels.forEach(primary => {
      const primaryLabel = primaryLabels.includes(primary) ? primary : 'Other';
      rawCompareLabels.forEach(compare => {
        const compareLabel = compareLabels.includes(compare) ? compare : 'Other';
        const count = matrix.get(primary).get(compare) || 0;
        cappedMatrix.get(primaryLabel).set(compareLabel, cappedMatrix.get(primaryLabel).get(compareLabel) + count);
      });
    });

    const primaryTotals = new Map(primaryLabels.map(label => [
      label,
      compareLabels.reduce((sum, compare) => sum + (cappedMatrix.get(label).get(compare) || 0), 0)
    ]));
    const compareTotals = new Map(compareLabels.map(label => [
      label,
      primaryLabels.reduce((sum, primary) => sum + (cappedMatrix.get(primary).get(label) || 0), 0)
    ]));
    const primaryRespondentTotals = aggregateCappedTotals(primaryLabels, rawPrimaryLabels, rawPrimaryRespondentTotals);
    const compareRespondentTotals = aggregateCappedTotals(compareLabels, rawCompareLabels, rawCompareRespondentTotals);

    const valueMode = chart.compareType === 'stacked100' ? 'primaryPercent' : chart.compareValueMode;
    const datasets = compareLabels.map((compare, index) => ({
      label: compare,
      data: primaryLabels.map(primary => {
        const count = cappedMatrix.get(primary).get(compare) || 0;
        if (valueMode === 'primaryPercent') return primaryRespondentTotals.get(primary) ? roundOne((count / primaryRespondentTotals.get(primary)) * 100) : 0;
        if (valueMode === 'comparePercent') return compareRespondentTotals.get(compare) ? roundOne((count / compareRespondentTotals.get(compare)) * 100) : 0;
        if (valueMode === 'totalPercent') return respondentTotal ? roundOne((count / respondentTotal) * 100) : 0;
        return count;
      }),
      backgroundColor: COLORS[index % COLORS.length],
      borderColor: COLORS[index % COLORS.length],
      borderWidth: 1
    }));

    return {
      type: 'comparison',
      rows,
      labels: primaryLabels,
      compareLabels,
      matrix: cappedMatrix,
      primaryTotals,
      compareTotals,
      primaryRespondentTotals,
      compareRespondentTotals,
      total,
      respondentTotal,
      datasets,
      valueMode
    };
  }

  function aggregateCappedTotals(cappedLabels, rawLabels, rawTotals) {
    return new Map(cappedLabels.map(label => [
      label,
      rawLabels.reduce((sum, rawLabel) => {
        const cappedLabel = cappedLabels.includes(rawLabel) ? rawLabel : 'Other';
        return cappedLabel === label ? sum + (rawTotals.get(rawLabel) || 0) : sum;
      }, 0)
    ]));
  }

  function renderChart(chart, card, result) {
    const canvas = card.querySelector('canvas');
    const empty = card.querySelector('.chart-empty');
    const chartArea = card.querySelector('.chart-area');
    const hasData = result.type === 'single'
      ? result.items.length > 0
      : result.labels.length > 0 && result.compareLabels.length > 0;

    if (chart.chartInstance) {
      chart.chartInstance.destroy();
      chart.chartInstance = null;
    }

    empty.classList.toggle('hidden', hasData);
    const tableOnly = result.type === 'comparison' ? chart.compareType === 'table' : chart.chartType === 'table';
    chartArea.classList.toggle('hidden', tableOnly);
    if (!hasData || tableOnly) return;

    const horizontal = result.type === 'single'
      ? getResolvedChartType(chart, result) === 'horizontalBar'
      : shouldUseHorizontalBars(result.labels);
    const dynamicHeight = horizontal ? Math.min(760, Math.max(360, 150 + (result.labels.length * 46))) : 420;
    chartArea.style.setProperty('--chart-height', `${dynamicHeight}px`);

    const context = canvas.getContext('2d');
    const config = result.type === 'comparison'
      ? getComparisonChartConfig(chart, result)
      : getSingleChartConfig(chart, result);
    chart.chartInstance = new Chart(context, config);
  }

  function getSingleChartConfig(chart, result) {
    const resolvedType = getResolvedChartType(chart, result);
    const type = resolvedType === 'horizontalBar' ? 'bar' : resolvedType;
    const values = result.values;

    return {
      type,
      data: {
        labels: result.labels.map(label => truncateLabel(label)),
        datasets: [{
          label: 'Responses',
          data: values,
          backgroundColor: result.labels.map((_, index) => COLORS[index % COLORS.length]),
          borderColor: result.labels.map((_, index) => COLORS[index % COLORS.length]),
          borderWidth: 1,
          tension: 0.25,
          fill: resolvedType === 'line' ? false : true
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: resolvedType === 'horizontalBar' ? 'y' : 'x',
        layout: {
          padding: { top: 24, right: resolvedType === 'horizontalBar' ? 64 : 18 }
        },
        plugins: {
          legend: { display: ['pie', 'doughnut'].includes(resolvedType) },
          tooltip: {
            callbacks: {
              title: items => items.length ? result.items[items[0].dataIndex].response : '',
              label: context => {
                const item = result.items[context.dataIndex];
                return `${item.response}: ${formatNumber(item.count)} (${item.rowPercent}%)`;
              }
            }
          },
          datalabels: {
            color: ['pie', 'doughnut'].includes(resolvedType) ? CHART_ON_COLOR : CHART_TEXT,
            anchor: ['pie', 'doughnut'].includes(resolvedType) ? 'center' : 'end',
            align: ['pie', 'doughnut'].includes(resolvedType) ? 'center' : (type === 'line' ? 'top' : 'end'),
            clamp: true,
            clip: false,
            textAlign: 'center',
            font: context => ({
              weight: '700',
              size: ['pie', 'doughnut'].includes(resolvedType) && context.dataset.data.length > 8 ? 10 : 11
            }),
            formatter: (value, context) => {
              const item = result.items[context.dataIndex];
              const parts = [];
              if (chart.showCounts) parts.push(formatNumber(value));
              if (chart.showPercentages) parts.push(`${item.rowPercent}%`);
              return ['pie', 'doughnut'].includes(resolvedType) ? parts.join('\n') : parts.join(' | ');
            }
          }
        },
        scales: ['pie', 'doughnut'].includes(resolvedType) ? {} : {
          x: { beginAtZero: true, grace: resolvedType === 'horizontalBar' ? '12%' : 0, grid: { color: CHART_GRID } },
          y: { beginAtZero: true, grace: resolvedType === 'horizontalBar' ? 0 : '12%', grid: { display: resolvedType !== 'horizontalBar', color: CHART_GRID } }
        }
      }
    };
  }

  function getComparisonChartConfig(chart, result) {
    const stacked = chart.compareType === 'stacked' || chart.compareType === 'stacked100';
    const horizontal = shouldUseHorizontalBars(result.labels);
    return {
      type: 'bar',
      data: {
        labels: result.labels.map(label => truncateLabel(label)),
        datasets: result.datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: horizontal ? 'y' : 'x',
        plugins: {
          legend: { position: 'bottom' },
          tooltip: {
            callbacks: {
              title: items => items.length ? result.labels[items[0].dataIndex] : '',
              label: context => `${context.dataset.label}: ${context.formattedValue}${result.valueMode === 'counts' ? '' : '%'}`
            }
          },
          datalabels: {
            color: CHART_TEXT,
            anchor: 'end',
            align: 'end',
            formatter: value => value ? `${value}${result.valueMode === 'counts' ? '' : '%'}` : ''
          }
        },
        scales: {
          x: {
            stacked,
            beginAtZero: true,
            max: horizontal && (chart.compareType === 'stacked100' || result.valueMode !== 'counts') ? 100 : undefined,
            grid: { color: CHART_GRID }
          },
          y: {
            stacked,
            beginAtZero: true,
            max: !horizontal && (chart.compareType === 'stacked100' || result.valueMode !== 'counts') ? 100 : undefined,
            grid: { color: CHART_GRID }
          }
        }
      }
    };
  }

  function getResolvedChartType(chart, result) {
    if (chart.chartType !== 'auto') return chart.chartType;
    return shouldUseHorizontalBars(result.labels) ? 'horizontalBar' : 'bar';
  }

  function shouldUseHorizontalBars(labels) {
    return labels.length > 5 || labels.some(label => String(label).length > 18);
  }

  function truncateLabel(label, maxLength = 34) {
    const text = String(label);
    return text.length > maxLength ? `${text.slice(0, maxLength - 1)}…` : text;
  }

  function renderSummaryTable(chart, card, result) {
    if (result.type === 'comparison') {
      renderComparisonTable(chart, card, result);
      return;
    }

    const searchText = chart.summarySearch.trim().toLowerCase();
    const matchingItems = result.items.filter(item => item.response.toLowerCase().includes(searchText));
    const visibleItems = matchingItems.slice(0, TABLE_ROW_LIMIT);
    renderTableNote(card, matchingItems.length, 'responses');
    const html = `
      <table>
        <thead>
          <tr>
            <th>Select</th>
            <th>Response</th>
            <th class="number">Count</th>
            <th class="number">Percentage</th>
          </tr>
        </thead>
        <tbody>
          ${visibleItems.map(item => `
            <tr>
              <td>
                <label class="selected-cell">
                  <input type="checkbox" class="response-select" value="${escapeAttr(item.response)}" ${chart.selectedResponses.has(item.response) ? 'checked' : ''}>
                  <span>Select</span>
                </label>
              </td>
              <td>${escapeHtml(item.response)}</td>
              <td class="number">${formatNumber(item.count)}</td>
              <td class="number">${item.rowPercent}%</td>
            </tr>
          `).join('') || '<tr><td colspan="4">No matching responses.</td></tr>'}
        </tbody>
      </table>
    `;
    card.querySelector('.summary-table').innerHTML = html;
    card.querySelectorAll('.response-select').forEach(input => {
      input.addEventListener('change', event => {
        if (event.target.checked) chart.selectedResponses.add(event.target.value);
        else chart.selectedResponses.delete(event.target.value);
      });
    });
  }

  function renderComparisonTable(chart, card, result) {
    const mode = result.valueMode || chart.compareValueMode;
    const suffix = mode === 'counts' ? '' : '%';
    renderTableNote(card, result.labels.length, 'comparison rows');
    const rows = result.labels.map(primary => {
      const total = result.primaryRespondentTotals.get(primary) || 0;
      const cells = result.compareLabels.map(compare => {
        const count = result.matrix.get(primary).get(compare) || 0;
        let value = count;
        if (mode === 'primaryPercent') value = total ? roundOne((count / total) * 100) : 0;
        if (mode === 'comparePercent') value = result.compareRespondentTotals.get(compare) ? roundOne((count / result.compareRespondentTotals.get(compare)) * 100) : 0;
        if (mode === 'totalPercent') value = result.respondentTotal ? roundOne((count / result.respondentTotal) * 100) : 0;
        return `<td class="number">${escapeHtml(formatTableValue(value, suffix))}</td>`;
      }).join('');
      return `
        <tr>
          <td>${escapeHtml(primary)}</td>
          ${cells}
          <td class="number">${formatNumber(total)}</td>
        </tr>
      `;
    }).join('');

    card.querySelector('.summary-table').innerHTML = `
      <table>
        <thead>
          <tr>
            <th>${escapeHtml(getActiveColumnDisplayName(chart.primaryColumn))}</th>
            ${result.compareLabels.map(label => `<th class="number">${escapeHtml(label)}</th>`).join('')}
            <th class="number">Total</th>
          </tr>
        </thead>
        <tbody>${rows || `<tr><td colspan="${result.compareLabels.length + 2}">No data is available.</td></tr>`}</tbody>
      </table>
    `;
  }

  function renderTableNote(card, totalRows, label) {
    const note = card.querySelector('.table-note');
    if (!note) return;
    note.textContent = totalRows > TABLE_ROW_LIMIT
      ? `Showing first ${TABLE_ROW_LIMIT} ${label}.`
      : '';
  }

  function capLabels(labels, totals, limit) {
    const sorted = [...labels].sort((a, b) => (totals.get(b) || 0) - (totals.get(a) || 0) || a.localeCompare(b));
    if (sorted.length <= limit) return sorted;
    const topLabels = sorted.slice(0, limit - 1);
    return topLabels.includes('Other') ? topLabels : [...topLabels, 'Other'];
  }

  function sortItems(items, sortMode) {
    return [...items].sort((a, b) => {
      if (sortMode === 'asc') return a.count - b.count || a.response.localeCompare(b.response);
      if (sortMode === 'alpha') return a.response.localeCompare(b.response);
      return b.count - a.count || a.response.localeCompare(b.response);
    });
  }

  function applyTopGrouping(items, topMode, totalRows, nonBlankRows) {
    const topCount = Math.min(Number(topMode) || CHART_RESPONSE_DISPLAY_LIMIT, CHART_RESPONSE_DISPLAY_LIMIT);
    if (items.length <= topCount) return items;

    const topItems = items.slice(0, topCount - 1);
    const otherCount = items.slice(topCount - 1).reduce((sum, item) => sum + item.count, 0);
    if (otherCount > 0) {
      topItems.push({
        response: 'Other',
        count: otherCount,
        rowPercent: totalRows ? roundOne((otherCount / totalRows) * 100) : 0,
        nonBlankPercent: nonBlankRows ? roundOne((otherCount / nonBlankRows) * 100) : 0
      });
    }
    return topItems;
  }

  function getUniqueValues(rows, column) {
    const columnLabels = getColumnNonBlankLabels(rows, column);
    const useYesNoLabels = ChartRules.shouldUseYesNoLabels(columnLabels);
    return Array.from(new Set(rows.flatMap(row => getDisplayResponseLabelsForColumn(row[column], columnLabels, useYesNoLabels))))
      .sort((a, b) => a.localeCompare(b));
  }

  function getColumnNonBlankLabels(rows, column) {
    return rows.flatMap(row => getResponseLabels(row[column])).filter(label => label !== NO_RESPONSE);
  }

  function getDisplayResponseLabelsForColumn(value, nonBlankLabels, useYesNoLabels = ChartRules.shouldUseYesNoLabels(nonBlankLabels)) {
    const labels = getResponseLabels(value);
    return labels.map(label => ChartRules.getDisplayAnswerLabel(label, nonBlankLabels, useYesNoLabels));
  }

  function getResponseLabels(value) {
    const labels = ChartRules.splitResponseLabels(value)
      .map(getResponseLabel)
      .filter(label => label !== NO_RESPONSE);
    return labels.length ? labels : [NO_RESPONSE];
  }

  function getResponseLabel(value) {
    const normalized = normalizeValue(value);
    return normalized === '' ? NO_RESPONSE : normalized;
  }

  function getMergedLabel(label, merges) {
    const match = merges.find(merge => merge.sources.has(label));
    return match ? match.name : label;
  }

  function normalizeValue(value) {
    if (value === null || value === undefined) return '';
    if (Array.isArray(value)) return value.map(normalizeValue).filter(Boolean).join('; ');
    if (value instanceof Date) return value.toLocaleDateString();
    if (typeof value === 'string') return value.trim();
    return String(value).trim();
  }

  function displayCell(value) {
    if (Array.isArray(value)) return value.map(displayCell).filter(Boolean).join('; ');
    const normalized = normalizeValue(value);
    return normalized === '' ? '' : normalized;
  }

  function exportChartPng(chart) {
    if (!chart.chartInstance) {
      showToast('This chart is currently shown as a table only.', 'warning');
      return;
    }
    const link = document.createElement('a');
    link.download = `${safeFileName(chart.title)}.png`;
    link.href = chart.chartInstance.toBase64Image('image/png', 1);
    link.click();
    showToast('Chart PNG generated.');
  }

  function exportSummaryCsv(chart) {
    const rows = applyFilters(state.rows, chart.filters);
    const result = chart.compareColumn
      ? buildComparisonResult(rows, chart)
      : buildSingleColumnResult(rows, chart);
    const csvRows = [];

    if (result.type === 'comparison') {
      csvRows.push([getActiveColumnDisplayName(chart.primaryColumn), ...result.compareLabels, 'Total']);
      result.labels.forEach(primary => {
        csvRows.push([
          primary,
          ...result.compareLabels.map(compare => result.matrix.get(primary).get(compare) || 0),
          result.primaryTotals.get(primary) || 0
        ]);
      });
    } else {
      csvRows.push([getActiveColumnDisplayName(chart.primaryColumn), 'Count', 'Percentage']);
      result.items.forEach(item => csvRows.push([item.response, item.count, `${item.rowPercent}%`]));
    }

    downloadCsv(csvRows, `${safeFileName(chart.title)}-summary.csv`);
    showToast('Summary CSV generated.');
  }

  function exportFilteredDataCsv(chart) {
    const rows = applyFilters(state.rows, chart.filters);
    const csvRows = [state.columns.map(getActiveColumnDisplayName), ...rows.map(row => state.columns.map(column => displayCell(row[column])))];
    downloadCsv(csvRows, `${safeFileName(chart.title)}-filtered-data.csv`);
    showToast('Filtered data CSV generated.');
  }

  function toggleLinkedSurveySetup() {
    const opening = els.linkedSurveySetup.classList.contains('hidden');
    els.linkedSurveySetup.classList.toggle('hidden', !opening);
    els.toggleLinkedSurveyBtn.setAttribute('aria-expanded', String(opening));
    els.toggleLinkedSurveyBtn.textContent = opening ? 'Hide setup' : (state.linkedSurvey.active ? 'Edit link' : 'Set up link');
    if (opening) renderLinkedSurveyPanel();
  }

  function renderLinkedSurveyPanel() {
    if (!els.linkedSurveyPanel) return;
    const hasDataset = Boolean(state.workbook && state.allRows.length);
    els.linkedSurveyPanel.classList.toggle('hidden', !hasDataset);
    if (!hasDataset) return;

    els.linkPrimaryName.textContent = state.fileName || 'Active dataset';
    els.linkPrimarySheet.textContent = state.sheetName || 'No sheet selected';
    populateSelect(els.linkPrimaryField, state.allColumns, els.linkPrimaryField.value || pickLinkField(state.allColumns), false, 'None', column => getActiveColumnDisplayName(column));
    els.toggleLinkedSurveyBtn.textContent = els.linkedSurveySetup.classList.contains('hidden')
      ? (state.linkedSurvey.active ? 'Edit link' : 'Set up link')
      : 'Hide setup';
    renderLinkedSurveySource();
    renderLinkDiagnostics();
  }

  function getSecondaryWorkbook() {
    return els.linkSecondarySource.value === 'file' ? state.linkedSurvey.secondaryWorkbook : state.workbook;
  }

  function getLinkedQuestionDisplayName() {
    const linked = state.linkedSurvey;
    if (!linked.question) return '';
    return getDisplayColumnName(linked.secondaryWorkbook || state.workbook, linked.secondarySheet || els.linkSecondarySheet.value, linked.question);
  }

  function renderLinkedSurveySource() {
    if (!state.workbook) return;
    const fileMode = els.linkSecondarySource.value === 'file';
    els.linkSecondaryFileWrap.classList.toggle('hidden', !fileMode);
    els.linkSecondaryFileName.textContent = fileMode
      ? (state.linkedSurvey.secondaryFileName || 'Choose an Excel or CSV file to continue.')
      : 'Using another sheet from the active workbook.';
    const workbook = getSecondaryWorkbook();
    const sheetNames = workbook
      ? getSurveySheetNames(workbook.SheetNames).filter(name => fileMode || name !== state.sheetName)
      : [];
    populateSelect(els.linkSecondarySheet, sheetNames, els.linkSecondarySheet.value, false);
    if (!sheetNames.length) els.linkSecondarySheet.innerHTML = `<option value="">${fileMode ? 'Upload a secondary file' : 'No other sheets available'}</option>`;
    renderLinkedSurveyFields();
  }

  async function loadSecondarySurveyFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    const extension = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(extension)) {
      showLinkValidation('Please choose an .xlsx, .xls, or .csv file.', 'error');
      return;
    }
    showLinkValidation('Reading the secondary survey...', 'loading');
    try {
      const workbook = await parseWorkbook(await file.arrayBuffer(), extension);
      if (!workbook.SheetNames.length) throw new Error('No sheets were found in the secondary file.');
      clearSurveyLink(false);
      state.linkedSurvey.secondaryWorkbook = workbook;
      state.linkedSurvey.secondaryFileName = file.name;
      renderLinkedSurveySource();
      showLinkValidation('Secondary survey loaded. Select matching fields, then match the surveys.', '');
    } catch (error) {
      console.error(error);
      state.linkedSurvey.secondaryWorkbook = null;
      state.linkedSurvey.secondaryFileName = '';
      renderLinkedSurveySource();
      showLinkValidation(error.message || 'The secondary survey could not be opened.', 'error');
    }
  }

  function renderLinkedSurveyFields() {
    const workbook = getSecondaryWorkbook();
    const sheetName = els.linkSecondarySheet.value;
    const columns = workbook && sheetName ? getSheetColumns(workbook, sheetName) : [];
    populateSelect(els.linkSecondaryField, columns, els.linkSecondaryField.value || pickLinkField(columns), false, 'None', column => getDisplayColumnName(workbook, sheetName, column));
    const selectedQuestion = state.linkedSurvey.question;
    const questions = columns.filter(column => column !== els.linkSecondaryField.value);
    populateSelect(els.linkQuestion, questions, selectedQuestion, true, state.linkedSurvey.result ? 'Select a secondary question' : 'Match surveys first', column => getDisplayColumnName(workbook, sheetName, column));
    if (!state.linkedSurvey.active || !selectedQuestion) els.linkQuestion.value = '';
    els.linkQuestion.disabled = !state.linkedSurvey.active;
  }

  function createSurveyLink() {
    const workbook = getSecondaryWorkbook();
    const secondarySheet = els.linkSecondarySheet.value;
    const primaryField = els.linkPrimaryField.value;
    const secondaryField = els.linkSecondaryField.value;
    if (!primaryField || !secondaryField) {
      showLinkValidation('Select a matching field from both surveys.', 'error');
      return;
    }
    if (!workbook || !secondarySheet) {
      showLinkValidation('Select or upload a secondary survey and choose its sheet.', 'error');
      return;
    }

    try {
      const secondaryRows = getSheetRecords(workbook, secondarySheet);
      if (!secondaryRows.length) throw new Error('The selected secondary sheet has no usable rows.');
      const result = LinkedSurvey.analyzeLink(state.allRows, secondaryRows, primaryField, secondaryField);
      resetLinkedSurveyMatch();
      state.linkedSurvey.secondarySheet = secondarySheet;
      state.linkedSurvey.result = result;
      state.linkedSurvey.active = result.stats.matchedRows > 0;
      invalidateGeneratedReport();
      if (!result.stats.matchedRows) {
        showLinkValidation('No rows matched. Review the selected fields and unmatched records; the existing single-survey analysis remains active.', 'error');
      } else {
        showLinkValidation(`${formatNumber(result.stats.matchedRows)} primary rows matched. Select a secondary question to add the linked breakdown.`, result.stats.unmatchedRows ? 'warning' : '');
      }
      updateAnalysisColumns();
      renderLinkedSurveyFields();
      renderLinkDiagnostics();
      renderFileStats();
      renderAllCharts();
      renderReportColumns();
      els.clearSurveyLinkBtn.classList.toggle('hidden', !state.linkedSurvey.active && !state.linkedSurvey.result);
    } catch (error) {
      console.error(error);
      showLinkValidation(error.message || 'The surveys could not be matched.', 'error');
    }
  }

  function applyLinkedQuestion() {
    if (!state.linkedSurvey.active || !state.linkedSurvey.result) {
      showLinkValidation('Match the surveys before selecting a secondary question.', 'error');
      return;
    }
    const previousColumn = state.linkedSurvey.column;
    if (previousColumn) state.columnStats.delete(previousColumn);
    state.linkedSurvey.question = els.linkQuestion.value;
    state.linkedSurvey.column = '';
    try {
      updateAnalysisColumns();
      invalidateGeneratedReport();
      renderFileStats();
      renderAllCharts();
      renderReportColumns();
      if (state.linkedSurvey.question) {
        showLinkValidation(`Linked breakdown ready: ${getLinkedQuestionDisplayName()}`, '');
        showToast('Linked survey breakdown added to charts and reports.');
      } else {
        showLinkValidation('Surveys are matched. Select a secondary question to add a breakdown.', '');
      }
    } catch (error) {
      console.error(error);
      state.linkedSurvey.question = '';
      state.linkedSurvey.column = '';
      updateAnalysisColumns();
      showLinkValidation(error.message || 'The secondary question could not be applied.', 'error');
    }
  }

  function resetLinkedSurveyMatch() {
    const linked = state.linkedSurvey;
    if (!linked) return;
    if (linked.column) state.columnStats.delete(linked.column);
    linked.active = false;
    linked.result = null;
    linked.secondarySheet = '';
    linked.question = '';
    linked.column = '';
  }

  function clearSurveyLink(announce) {
    const hadLink = Boolean(state.linkedSurvey.result);
    resetLinkedSurveyMatch();
    if (hadLink) invalidateGeneratedReport();
    if (state.allRows.length) {
      updateAnalysisColumns();
      renderFileStats();
      renderAllCharts();
      renderReportColumns();
    }
    renderLinkedSurveyFields();
    renderLinkDiagnostics();
    showLinkValidation(announce ? 'Linked survey removed. The original single-survey analysis is active.' : '', '');
    if (announce) showToast('Linked survey removed.');
  }

  function renderLinkDiagnostics() {
    const result = state.linkedSurvey.result;
    els.clearSurveyLinkBtn.classList.toggle('hidden', !result);
    els.linkStatusSummary.classList.toggle('hidden', !result);
    els.linkDiagnosticActions.classList.toggle('hidden', !result);
    if (!result) {
      els.linkHeadlineStatus.classList.add('hidden');
      els.linkHeadlineStatus.textContent = '';
      els.linkStatusSummary.innerHTML = '';
      els.linkWarnings.classList.add('hidden');
      els.linkWarnings.innerHTML = '';
      return;
    }
    const stats = result.stats;
    els.linkHeadlineStatus.classList.remove('hidden');
    els.linkHeadlineStatus.classList.toggle('is-warning', stats.unmatchedRows > 0 || result.duplicates.secondary.length > 0);
    els.linkHeadlineStatus.textContent = `${formatNumber(stats.matchedRows)} matched · ${formatNumber(stats.unmatchedRows)} unmatched`;
    const values = [
      ['Primary rows', formatNumber(stats.totalPrimaryRows), false],
      ['Secondary rows', formatNumber(stats.totalSecondaryRows), false],
      ['Matched rows', formatNumber(stats.matchedRows), false],
      ['Unmatched rows', formatNumber(stats.unmatchedRows), stats.unmatchedRows > 0],
      ['Match rate', `${roundOne(stats.matchRate)}%`, false],
      ['Unmatched rate', `${roundOne(stats.unmatchedRate)}%`, stats.unmatchedRows > 0]
    ];
    els.linkStatusSummary.innerHTML = values.map(([label, value, warning]) => `<div class="link-status-stat${warning ? ' is-warning' : ''}"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`).join('');

    const warnings = [];
    if (stats.unmatchedRows) warnings.push(`${formatNumber(stats.unmatchedRows)} primary row${stats.unmatchedRows === 1 ? '' : 's'} will be excluded from linked analysis.`);
    if (result.duplicates.primary.length) warnings.push(`${formatNumber(result.duplicates.primary.length)} repeated primary matching value${result.duplicates.primary.length === 1 ? '' : 's'} found. This can be expected when multiple responses belong to one site.`);
    if (result.duplicates.secondary.length) warnings.push(`${formatNumber(result.duplicates.secondary.length)} duplicate secondary matching value${result.duplicates.secondary.length === 1 ? '' : 's'} found. Those ambiguous matches are excluded.`);
    els.linkWarnings.classList.toggle('hidden', !warnings.length);
    els.linkWarnings.innerHTML = warnings.length ? `<strong>Review matching issues</strong><ul>${warnings.map(warning => `<li>${escapeHtml(warning)}</li>`).join('')}</ul>` : '';
    els.viewUnmatchedBtn.disabled = !result.unmatched.length;
    els.downloadUnmatchedBtn.disabled = !result.unmatched.length;
    els.viewDuplicatesBtn.disabled = !result.duplicates.primary.length && !result.duplicates.secondary.length;
  }

  function showLinkValidation(message, type) {
    els.linkValidation.textContent = message;
    els.linkValidation.className = `status-message ${type || ''}`.trim();
  }

  function invalidateGeneratedReport() {
    state.reportResult = null;
    els.reportOutputTitle.textContent = 'No report generated yet';
    els.reportOutputMeta.textContent = 'Select questions and generate a breakdown report.';
    els.reportContextBar.classList.add('hidden');
    els.reportContextBar.innerHTML = '';
    els.distributionOutput.innerHTML = '<div class="report-empty"><span class="empty-state-icon" aria-hidden="true">▦</span><strong>Select questions and generate a breakdown report.</strong><span>Your report preview will appear here.</span></div>';
  }

  function showUnmatchedRecords() {
    const result = state.linkedSurvey.result;
    if (!result || !result.unmatched.length) return;
    els.linkDetailsTitle.textContent = 'Unmatched primary records';
    const columns = state.allColumns;
    els.linkDetailsBody.innerHTML = `<table><thead><tr><th>Source row</th><th>Issue</th>${columns.map(column => `<th>${escapeHtml(getActiveColumnDisplayName(column))}</th>`).join('')}</tr></thead><tbody>${result.unmatched.map(item => `<tr><td class="number">${item.rowNumber}</td><td>${escapeHtml(item.reason)}</td>${columns.map(column => `<td>${escapeHtml(displayCell(item.row[column]))}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
    els.linkDetailsDialog.showModal();
  }

  function downloadUnmatchedRecords() {
    const result = state.linkedSurvey.result;
    if (!result || !result.unmatched.length) return;
    const rows = [['Source row', 'Matching value', 'Issue', ...state.allColumns.map(getActiveColumnDisplayName)]];
    result.unmatched.forEach(item => rows.push([item.rowNumber, item.value, item.reason, ...state.allColumns.map(column => displayCell(item.row[column]))]));
    downloadCsv(rows, `${safeFileName(state.fileName)}-unmatched-records.csv`);
    showToast('Unmatched records CSV generated.');
  }

  function showDuplicateValues() {
    const result = state.linkedSurvey.result;
    if (!result) return;
    const groups = [
      ...result.duplicates.primary.map(item => ({ survey: 'Primary', ...item })),
      ...result.duplicates.secondary.map(item => ({ survey: 'Secondary', ...item }))
    ];
    if (!groups.length) return;
    els.linkDetailsTitle.textContent = 'Duplicate matching values';
    els.linkDetailsBody.innerHTML = `<table><thead><tr><th>Survey</th><th>Matching value</th><th>Occurrences</th><th>Source rows</th><th>Effect</th></tr></thead><tbody>${groups.map(item => `<tr><td>${escapeHtml(item.survey)}</td><td>${escapeHtml(item.value)}</td><td class="number">${item.count}</td><td>${escapeHtml(item.rowNumbers.join(', '))}</td><td>${item.survey === 'Secondary' ? 'Excluded as ambiguous' : 'Matched as repeated responses'}</td></tr>`).join('')}</tbody></table>`;
    els.linkDetailsDialog.showModal();
  }

  async function loadPublicGoogleSheet() {
    const url = normalizeValue(els.publicSheetUrl.value);
    const sheetId = extractGoogleSheetId(url);
    if (!sheetId) {
      showReportStatus('Paste a valid public Google Sheets link.', 'error');
      return;
    }

    showReportStatus('Loading public Google Sheet...', 'loading');
    setButtonLoading(els.loadPublicSheetBtn, true, 'Loading…');
    try {
      const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
      const response = await fetch(exportUrl);
      if (!response.ok) throw new Error(`Google Sheets returned ${response.status}`);
      const buffer = await response.arrayBuffer();
      const workbook = await parseWorkbook(buffer);
      if (!workbook.SheetNames.length) throw new Error('No sheets found');
      const sourceName = `Google Sheet: ${sheetId.slice(0, 8)}...`;
      upsertReportSource(`google-${sheetId}`, sourceName, workbook);
      activateWorkbook(workbook, sourceName, workbook.SheetNames[0]);
      els.reportSourceSelect.value = `google-${sheetId}`;
      renderReportControls();
      showStatus('Public Google Sheet loaded. Your charts now use this data source.', '');
      showReportStatus('Public Google Sheet loaded. It is used only in this browser session.', '');
      setActiveTab('charts');
      showToast('Google Sheet loaded successfully.');
    } catch (error) {
      console.error(error);
      showReportStatus('Could not load that Google Sheet. Make sure it is public or shared with anyone who has the link.', 'error');
      showToast('The Google Sheet could not be loaded.', 'error');
    } finally {
      setButtonLoading(els.loadPublicSheetBtn, false);
    }
  }

  function upsertReportSource(id, name, workbook) {
    const existing = state.sources.find(source => source.id === id);
    if (existing) {
      existing.name = name;
      existing.workbook = workbook;
    } else {
      state.sources.push({ id, name, workbook });
    }
    state.reportResult = null;
    state.previewSearch = '';
    if (els.previewSearch) els.previewSearch.value = '';
  }

  function getSelectedReportSource() {
    return state.sources.find(source => source.id === els.reportSourceSelect.value) || state.sources[0] || null;
  }

  function renderReportControls() {
    const previousSource = els.reportSourceSelect.value;
    els.reportSourceSelect.innerHTML = state.sources.length
      ? state.sources.map(source => `<option value="${escapeAttr(source.id)}">${escapeHtml(source.name)}</option>`).join('')
      : '<option value="">No source loaded</option>';

    if (state.sources.some(source => source.id === previousSource)) {
      els.reportSourceSelect.value = previousSource;
    } else if (state.sources.length) {
      els.reportSourceSelect.value = state.sources[0].id;
    }

    const source = getSelectedReportSource();
    const disabled = !source;
    [
      els.reportNameInput, els.reportDataSheetSelect,
      els.primaryBreakdownSelect, els.reportFilterColumnSelect,
      els.generateReportBtn, els.downloadReportCsvBtn, els.downloadReportXlsxBtn
    ].forEach(control => {
      control.disabled = disabled;
    });

    if (!source) {
      els.reportDataSheetSelect.innerHTML = '<option value="">No sheets</option>';
      els.primaryBreakdownSelect.innerHTML = '<option value="">No breakdown</option>';
      els.reportFilterColumnSelect.innerHTML = '<option value="">No filter</option>';
      renderCheckboxList(els.questionChecklist, [], { emptyText: 'No response columns found' });
      renderCheckboxList(els.reportFilterValues, [], { emptyText: 'No filter values found' });
      els.reportFilterValues.classList.add('hidden');
      els.reportFilterNote.textContent = '';
      updateReportColumnNote([]);
      return;
    }

    const sheetNames = getSurveySheetNames(source.workbook.SheetNames);
    populateSelect(els.reportDataSheetSelect, sheetNames, pickDataSheet(sheetNames), false);
    if (!normalizeValue(els.reportNameInput.value)) els.reportNameInput.value = 'Question breakdown';
    renderReportColumns();
  }

  function renderReportColumns() {
    const source = getSelectedReportSource();
    const sheetName = els.reportDataSheetSelect.value;
    const primaryColumns = source && sheetName ? getSheetColumns(source.workbook, sheetName) : [];
    const reportData = source && sheetName ? getLinkedReportData(source, sheetName) : { rows: [], column: '' };
    const dataRows = reportData.rows;
    const columns = dataRows.length ? Object.keys(dataRows[0]).filter(column => !column.startsWith('__')) : primaryColumns;
    const reportValuesByColumn = new Map(columns.map(column => [column, getReportUniqueValues(dataRows, column)]));
    const reportFilterValuesByColumn = new Map(columns.map(column => [column, getReportFilterValues(dataRows, column)]));
    const columnStats = columns.map(column => ({
      column,
      uniqueCount: reportValuesByColumn.get(column).length
    }));
    const filterColumnStats = columns.map(column => ({
      column,
      nonBlankUniqueCount: reportValuesByColumn.get(column).length,
      uniqueCount: reportFilterValuesByColumn.get(column).length
    }));
    const eligibleColumns = columnStats
      .filter(item => item.uniqueCount > 0 && item.uniqueCount <= REPORT_UNIQUE_VALUE_LIMIT)
      .map(item => item.column);
    const linkedColumn = reportData.column;
    if (isLinkedReportContext(source, sheetName)) state.linkedSurvey.column = linkedColumn;
    const eligibleBreakdownColumns = uniqueList([...eligibleColumns, linkedColumn].filter(Boolean));
    const eligibleQuestionColumns = eligibleColumns.filter(column => column !== linkedColumn);
    const eligibleFilterColumns = filterColumnStats
      .filter(item => item.nonBlankUniqueCount > 0 && item.uniqueCount > 0 && item.uniqueCount < REPORT_FILTER_UNIQUE_VALUE_LIMIT)
      .map(item => item.column);
    const emptyColumns = columnStats.filter(item => item.uniqueCount === 0);
    const ignoredColumns = columnStats.filter(item => item.uniqueCount > REPORT_UNIQUE_VALUE_LIMIT);
    const responseColumns = primaryColumns.filter(column => !isLikelyMetadataColumn(column));
    const defaultResponseColumns = responseColumns.filter(column => eligibleQuestionColumns.includes(column));
    const displayColumn = column => getDisplayColumnName(source.workbook, sheetName, column);
    renderCheckboxList(els.questionChecklist, eligibleQuestionColumns.map(column => ({ value: column, label: displayColumn(column) })), {
      checkedValues: defaultResponseColumns.length ? defaultResponseColumns : eligibleQuestionColumns,
      emptyText: 'No response columns found',
      onChange: updateReportSelectionCount
    });
    populateSelect(els.primaryBreakdownSelect, eligibleBreakdownColumns, linkedColumn, true, 'No main breakdown', displayColumn);
    els.primaryBreakdownSelect.value = linkedColumn || '';
    populateSelect(els.reportFilterColumnSelect, eligibleFilterColumns, '', true, 'No filter', displayColumn);
    els.reportFilterColumnSelect.value = '';
    renderReportFilterValues();
    updateReportColumnNote(ignoredColumns, emptyColumns, displayColumn);
    filterReportQuestions();
    updateReportSelectionCount();
  }

  function renderReportFilterValues() {
    const source = getSelectedReportSource();
    const sheetName = els.reportDataSheetSelect.value;
    const filterColumn = els.reportFilterColumnSelect.value;
    const dataRows = source && sheetName ? getReportDataRows(source, sheetName) : [];

    if (!filterColumn || !dataRows.length) {
      renderCheckboxList(els.reportFilterValues, [], { emptyText: 'No filter values found' });
      els.reportFilterValues.classList.add('hidden');
      els.reportFilterNote.textContent = `Filter columns must have fewer than ${REPORT_FILTER_UNIQUE_VALUE_LIMIT} unique values.`;
      return;
    }

    const values = getReportFilterValues(dataRows, filterColumn);
    renderCheckboxList(els.reportFilterValues, values.map(value => ({ value, label: value })), {
      checkedValues: values,
      emptyText: 'No filter values found',
      onChange: updateReportFilterSelectionNote
    });
    els.reportFilterValues.classList.toggle('hidden', !values.length);
    updateReportFilterSelectionNote();
  }

  function updateReportFilterSelectionNote() {
    const inputs = Array.from(els.reportFilterValues.querySelectorAll('input[type="checkbox"]'));
    if (!inputs.length) {
      els.reportFilterNote.textContent = 'No filter values found for this column.';
      return;
    }

    const selectedCount = inputs.filter(input => input.checked).length;
    els.reportFilterNote.textContent = `${selectedCount} of ${inputs.length} filter value${inputs.length === 1 ? '' : 's'} selected.`;
  }

  function syncBreakdownQuestionSelection() {
    const breakdownColumn = els.primaryBreakdownSelect.value;
    if (!breakdownColumn) return;
    Array.from(els.questionChecklist.querySelectorAll('input[type="checkbox"]'))
      .filter(input => input.value === breakdownColumn)
      .forEach(input => {
        input.checked = false;
      });
  }

  function updateReportColumnNote(ignoredColumns, emptyColumns = [], displayColumn = column => column) {
    if (!els.reportColumnNote) return;
    if (!ignoredColumns.length && !emptyColumns.length) {
      els.reportColumnNote.textContent = `Only columns with ${REPORT_UNIQUE_VALUE_LIMIT} or fewer unique responses are shown here.`;
      return;
    }

    const parts = [];
    if (ignoredColumns.length) {
      const names = ignoredColumns.slice(0, 4).map(item => displayColumn(item.column)).join(', ');
      const extra = ignoredColumns.length > 4 ? `, and ${ignoredColumns.length - 4} more` : '';
      parts.push(`${ignoredColumns.length} column${ignoredColumns.length === 1 ? '' : 's'} hidden because ${ignoredColumns.length === 1 ? 'it has' : 'they have'} more than ${REPORT_UNIQUE_VALUE_LIMIT} unique responses: ${names}${extra}`);
    }
    if (emptyColumns.length) {
      const names = emptyColumns.slice(0, 4).map(item => displayColumn(item.column)).join(', ');
      const extra = emptyColumns.length > 4 ? `, and ${emptyColumns.length - 4} more` : '';
      parts.push(`${emptyColumns.length} empty column${emptyColumns.length === 1 ? '' : 's'} hidden: ${names}${extra}`);
    }
    els.reportColumnNote.textContent = `${parts.join('. ')}.`;
  }

  function generateDistributionReport() {
    const source = getSelectedReportSource();
    if (!source) {
      showReportStatus('Load or upload a source first.', 'error');
      return;
    }

    setButtonLoading(els.generateReportBtn, true, 'Generating…');
    showReportStatus('Generating the breakdown report...', 'loading');
    try {
      const dataRows = getReportDataRows(source, els.reportDataSheetSelect.value);
      const reportName = normalizeValue(els.reportNameInput.value) || 'Question breakdown';
      if (!dataRows.length) throw new Error('The raw data sheet has no rows.');

      const reportFilter = getReportFilter(dataRows);
      const filteredRows = applyReportFilter(dataRows, reportFilter);
      const breakdownColumns = uniqueList([
        els.primaryBreakdownSelect.value
      ]).filter(column => column && column in (dataRows[0] || {}));
      const breakdownSet = new Set(breakdownColumns);
      const questions = getCheckedItems(els.questionChecklist).map(item => ({
        column: item.value,
        display: item.label
      })).filter(question => !breakdownSet.has(question.column));

      if (!filteredRows.length) throw new Error('No rows match the selected report filter.');
      if (!questions.length) throw new Error('Select at least one response column to include.');

      const output = buildDistributionOutput(filteredRows, questions, breakdownColumns, reportFilter.label);
      const linkedSummary = state.linkedSurvey.column && breakdownColumns.includes(state.linkedSurvey.column)
        ? LinkedSurvey.buildCategorySummary(filteredRows, state.linkedSurvey.column)
        : [];
      const linkedSummaryRows = linkedSummary.length
        ? [['Linked category coverage'], ['Category', 'Matched sites', 'Survey responses'], ...linkedSummary.map(item => [item.category, item.matchedSites, item.surveyResponses]), []]
        : [];
      state.reportResult = {
        title: reportName,
        aoa: [...linkedSummaryRows, ...output.aoa],
        rows: output.rows,
        skipped: output.skipped,
        questionCount: output.questionCount,
        breakdownLabel: output.breakdownLabel,
        filterLabel: output.filterLabel,
        sourceName: source.name,
        linkedSummary
      };
      renderDistributionOutput(state.reportResult);
      showReportStatus('Breakdown report generated.', '');
      showToast('Breakdown report generated.');
    } catch (error) {
      console.error(error);
      showReportStatus(error.message || 'Could not generate the report.', 'error');
      showToast(error.message || 'Could not generate the report.', 'error');
    } finally {
      setButtonLoading(els.generateReportBtn, false);
    }
  }

  function getReportFilter(dataRows) {
    const column = els.reportFilterColumnSelect.value;
    if (!column) return { column: '', values: new Set(), label: 'All rows' };
    const selectedValues = getCheckedItems(els.reportFilterValues).map(item => item.value);
    const totalValues = getReportFilterValues(dataRows, column).length;
    return {
      column,
      values: new Set(selectedValues.map(normalizeForMatch)),
      label: formatReportFilterLabel(column, selectedValues, totalValues, getCurrentReportColumnDisplayName(column))
    };
  }

  function getCurrentReportColumnDisplayName(column) {
    const source = getSelectedReportSource();
    const sheetName = els.reportDataSheetSelect.value;
    return source && sheetName ? getDisplayColumnName(source.workbook, sheetName, column) : (column || '');
  }

  function formatReportFilterLabel(column, selectedValues, totalValues, displayColumn = column) {
    if (!selectedValues.length) return `${displayColumn}: no values selected`;
    if (selectedValues.length === totalValues) return `${displayColumn}: all ${totalValues} value${totalValues === 1 ? '' : 's'}`;
    const preview = selectedValues.slice(0, 5).join(', ');
    const extra = selectedValues.length > 5 ? `, and ${selectedValues.length - 5} more` : '';
    return `${displayColumn}: ${preview}${extra}`;
  }

  function applyReportFilter(dataRows, filter) {
    if (!filter.column) return dataRows;
    if (!filter.values.size) return [];
    return dataRows.filter(row => getResponseLabels(row[filter.column]).some(value => filter.values.has(normalizeForMatch(value))));
  }

  function buildDistributionOutput(dataRows, questions, breakdownColumns, filterLabel = 'All rows') {
    const aoa = [];
    const rows = [];
    const skipped = [];
    let questionCount = 0;
    const breakdownLabel = breakdownColumns.length ? getCurrentReportColumnDisplayName(breakdownColumns[0]) : 'No main breakdown selected';
    questions.forEach(question => {
      if (!(question.column in (dataRows[0] || {}))) {
        skipped.push(question.column);
        return;
      }
      const section = buildQuestionSection(dataRows, question.display || question.column, question.column, breakdownColumns);
      aoa.push(...section.aoa);
      rows.push(...section.rows);
      questionCount += 1;
    });

    return { aoa, rows, skipped, questionCount, breakdownLabel, filterLabel };
  }

  function buildQuestionSection(dataRows, displayName, questionColumn, breakdownColumns) {
    const answers = sortReportAnswers(getReportUniqueValues(dataRows, questionColumn), questionColumn);
    const combos = createBreakdownCombos(dataRows, breakdownColumns);
    const headerRows = createReportHeaderRows(combos, breakdownColumns);
    const width = Math.max(1, headerRows[0].length) + 1;
    const aoa = [];
    const rows = [];

    const titleRow = Array(width).fill('');
    titleRow[0] = displayName;
    aoa.push(titleRow);
    rows.push(titleRow.map((value, index) => ({ value, type: index === 0 ? 'question' : 'blank' })));

    headerRows.forEach(row => {
      const fullRow = ['', ...row];
      aoa.push(fullRow);
      rows.push(fullRow.map(value => ({ value, type: 'header' })));
    });

    answers.forEach(answer => {
      const row = Array(width).fill('');
      const renderedRow = Array(width).fill(null).map(() => ({ value: '', type: 'plain' }));
      row[0] = answer;
      renderedRow[0] = { value: answer, type: 'answer' };

      combos.forEach((combo, comboIndex) => {
        const filtered = dataRows.filter(dataRow => combo.conditions.every(condition => getReportValues(dataRow[condition.column]).includes(condition.value)));
        const counts = countAnswers(filtered, questionColumn);
        const total = filtered.filter(dataRow => getReportValues(dataRow[questionColumn]).length > 0).length;
        const count = counts.get(normalizeForMatch(answer)) || 0;
        const percent = total ? count / total : 0;
        const countIndex = 1 + (comboIndex * 2);
        row[countIndex] = count;
        row[countIndex + 1] = percent;
        renderedRow[countIndex] = { value: count, type: 'count' };
        renderedRow[countIndex + 1] = { value: percent, type: 'percent' };
      });

      aoa.push(row);
      rows.push(renderedRow);
    });

    const spacer = Array(width).fill('');
    aoa.push(spacer);
    rows.push(spacer.map(value => ({ value, type: 'spacer' })));
    return { aoa, rows };
  }

  function getReportSectionWidth(dataRows, breakdownColumns) {
    const combos = createBreakdownCombos(dataRows, breakdownColumns);
    const headerRows = createReportHeaderRows(combos, breakdownColumns);
    return Math.max(3, headerRows[0].length + 1);
  }

  function createReportHeaderRows(combos, breakdownColumns) {
    if (!breakdownColumns.length) return [['Count', 'Percentage']];

    const rows = breakdownColumns.map(column => {
      const row = [];
      combos.forEach(combo => {
        const value = combo.conditions.find(condition => condition.column === column)?.value || '';
        row.push(value, '');
      });
      return row;
    });
    rows.push(combos.flatMap(() => ['Count', 'Percentage']));
    return rows;
  }

  function createBreakdownCombos(dataRows, breakdownColumns) {
    if (!breakdownColumns.length) return [{ label: 'All rows', conditions: [] }];
    const valuesByColumn = breakdownColumns.map(column => ({
      column,
      values: getReportUniqueValues(dataRows, column)
    })).filter(item => item.values.length);

    if (!valuesByColumn.length) return [{ label: 'All rows', conditions: [] }];

    return cartesianProduct(valuesByColumn.map(item => item.values)).map(values => ({
      label: values.join(' / '),
      conditions: values.map((value, index) => ({ column: valuesByColumn[index].column, value }))
    }));
  }

  function countAnswers(rows, questionColumn) {
    const counts = new Map();
    rows.forEach(row => {
      getReportValues(row[questionColumn]).forEach(value => {
        const key = normalizeForMatch(value);
        counts.set(key, (counts.get(key) || 0) + 1);
      });
    });
    return counts;
  }

  function renderDistributionOutput(result) {
    els.reportOutputTitle.textContent = result.title;
    const skippedText = result.skipped.length ? `, skipped ${result.skipped.length} missing columns` : '';
    els.reportOutputMeta.textContent = `${result.questionCount} response columns from ${result.sourceName}${skippedText}`;
    renderReportContextBar(result);
    const linkedSummary = result.linkedSummary && result.linkedSummary.length ? `
      <div class="linked-category-summary">
        <h3>Linked category coverage</h3>
        <table><thead><tr><th>Category</th><th class="number">Matched sites</th><th class="number">Survey responses</th></tr></thead><tbody>${result.linkedSummary.map(item => `<tr><td>${escapeHtml(item.category)}</td><td class="number">${formatNumber(item.matchedSites)}</td><td class="number">${formatNumber(item.surveyResponses)}</td></tr>`).join('')}</tbody></table>
      </div>` : '';
    els.distributionOutput.innerHTML = `${linkedSummary}
      <table>
        <tbody>
          ${result.rows.map(row => {
            const spacer = row.every(cell => cell.type === 'spacer' || normalizeValue(cell.value) === '');
            return `<tr class="${spacer ? 'report-spacer' : ''}">
              ${row.map(cell => renderReportCell(cell)).join('')}
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    `;
  }

  function renderReportContextBar(result) {
    const items = [];
    if (result.breakdownLabel && result.breakdownLabel !== 'No main breakdown selected') {
      items.push(['Breakdown', result.breakdownLabel]);
    }
    if (result.filterLabel && result.filterLabel !== 'All rows') {
      items.push(['Filter', result.filterLabel]);
    }

    els.reportContextBar.classList.toggle('hidden', !items.length);
    els.reportContextBar.innerHTML = items.map(([label, value]) => `
      <div class="report-context-chip">
        <span>${escapeHtml(label)}</span>
        <strong>${escapeHtml(value)}</strong>
      </div>
    `).join('');
  }

  function renderReportCell(cell) {
    const colspan = cell.colspan ? ` colspan="${cell.colspan}"` : '';
    if (cell.type === 'spacer') return '<td></td>';
    if (cell.type === 'meta-label') return `<td${colspan} class="report-meta-label">${escapeHtml(cell.value)}</td>`;
    if (cell.type === 'meta-value') return `<td${colspan} class="report-meta-value">${escapeHtml(cell.value)}</td>`;
    if (cell.type === 'question') return `<td${colspan} class="question-title">${escapeHtml(cell.value)}</td>`;
    if (cell.type === 'header') {
      const valueClass = cell.value === 'Count' ? ' report-count' : (cell.value === 'Percentage' ? ' report-percent' : '');
      return `<td${colspan} class="report-header${valueClass}">${escapeHtml(cell.value)}</td>`;
    }
    if (cell.type === 'count') return `<td${colspan} class="number report-count">${formatNumber(cell.value)}</td>`;
    if (cell.type === 'percent') {
      const percent = Number(cell.value) || 0;
      return `<td${colspan} class="number heat-cell report-percent" style="background:${getHeatColor(percent)}">${formatPercent(percent)}</td>`;
    }
    return `<td${colspan}>${escapeHtml(cell.value)}</td>`;
  }

  function downloadDistributionCsv() {
    if (!state.reportResult) {
      showToast('Generate a breakdown report first.', 'warning');
      return;
    }
    downloadCsv(state.reportResult.aoa.map(row => row.map(value => typeof value === 'number' ? value : displayCell(value))), `${safeFileName(state.reportResult.title)}.csv`);
    showToast('Report CSV generated.');
  }

  function downloadDistributionXlsx() {
    if (!state.reportResult) {
      showToast('Generate a breakdown report first.', 'warning');
      return;
    }
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet(state.reportResult.aoa);
    XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName(state.reportResult.title));
    XLSX.writeFile(workbook, `${safeFileName(state.reportResult.title)}.xlsx`);
    showToast('Excel report generated.');
  }

  function downloadCsv(rows, fileName) {
    const csv = rows.map(row => row.map(csvEscape).join(',')).join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  function csvEscape(value) {
    const text = String(value ?? '');
    return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  }

  function populateSelect(select, values, selectedValue, includeNone, noneLabel = 'None', labelForValue = value => value) {
    const previous = select.value;
    const options = includeNone ? [`<option value="">${escapeHtml(noneLabel)}</option>`] : [];
    options.push(...values.map(value => `<option value="${escapeAttr(value)}">${escapeHtml(labelForValue(value))}</option>`));
    select.innerHTML = options.join('');

    if (selectedValue && values.includes(selectedValue)) select.value = selectedValue;
    else if (previous && (values.includes(previous) || (includeNone && previous === ''))) select.value = previous;
    else if (includeNone) select.value = '';
    else if (values.length) select.value = values[0];
  }

  function pickSheet(sheetNames, pattern) {
    return sheetNames.find(name => pattern.test(name)) || sheetNames[0] || '';
  }

  function pickDataSheet(sheetNames) {
    return sheetNames.find(name => !isDictionarySheet(name) && !/question|config|input|generation|site|scs|lookup|mapping/i.test(name)) || sheetNames[0] || '';
  }

  function pickColumn(columns, pattern) {
    return columns.find(column => pattern.test(column)) || columns[0] || '';
  }

  function pickLinkField(columns) {
    const exactPatterns = [/^site(?: name| id)?$/i, /^school(?: name| id)?$/i, /^location(?: name| id)?$/i, /^center(?: name| id)?$/i, /^centre(?: name| id)?$/i, /^program id$/i, /^id$/i];
    return columns.find(column => exactPatterns.some(pattern => pattern.test(column))) || columns[0] || '';
  }

  function getSurveySheetNames(sheetNames) {
    const names = Array.isArray(sheetNames) ? sheetNames : [];
    return typeof DataDictionary === 'undefined'
      ? names
      : names.filter(name => !DataDictionary.isDictionarySheetName(name));
  }

  function isDictionarySheet(sheetName) {
    return typeof DataDictionary !== 'undefined' && DataDictionary.isDictionarySheetName(sheetName);
  }

  function getWorkbookDictionary(workbook) {
    if (!workbook || typeof DataDictionary === 'undefined') return null;
    if (workbookDictionaryCache.has(workbook)) return workbookDictionaryCache.get(workbook);

    const dictionarySheet = workbook.SheetNames.find(name => DataDictionary.isDictionarySheetName(name));
    const dictionary = dictionarySheet ? DataDictionary.create(getSheetMatrix(workbook, dictionarySheet)) : null;
    workbookDictionaryCache.set(workbook, dictionary);
    return dictionary;
  }

  function getDisplayColumnName(workbook, sheetName, column) {
    const fallback = normalizeValue(column);
    if (!fallback || !workbook || !sheetName || typeof DataDictionary === 'undefined') return fallback;

    const isActiveSheet = workbook === state.workbook && sheetName === state.sheetName;
    const originalHeader = isActiveSheet
      ? (state.columnOriginalHeaders.get(column) || fallback)
      : (getSheetHeaderDetails(workbook, sheetName).originalByColumn.get(column) || fallback);
    return DataDictionary.getDisplayQuestion(getWorkbookDictionary(workbook), sheetName, originalHeader) || fallback;
  }

  function getActiveColumnDisplayName(column) {
    return getDisplayColumnName(state.workbook, state.sheetName, column);
  }

  function getColumnStatsLabel(column) {
    const stats = state.columnStats.get(column);
    if (!stats) return '';
    return `${stats.type} · ${formatNumber(stats.answeredCount)} answered · ${formatNumber(stats.uniqueCount)} unique`;
  }

  function isLikelyMetadataColumn(column) {
    return ChartRules.isLikelyMetadataColumn(column);
  }

  function getSheetColumns(workbook, sheetName) {
    return getSheetHeaderDetails(workbook, sheetName).columns;
  }

  function getSheetHeaderDetails(workbook, sheetName, explicitHeaderRow = null) {
    if (!workbook) return { headerRow: 0, columns: [], originalByColumn: new Map() };
    let workbookHeaders = sheetHeaderCache.get(workbook);
    if (!workbookHeaders) {
      workbookHeaders = new Map();
      sheetHeaderCache.set(workbook, workbookHeaders);
    }
    const cacheKey = `${sheetName}\u0000${explicitHeaderRow === null ? 'auto' : explicitHeaderRow}`;
    if (workbookHeaders.has(cacheKey)) return workbookHeaders.get(cacheKey);

    const rows = getSheetMatrix(workbook, sheetName);
    const headerRow = explicitHeaderRow === null ? findHeaderRow(rows) : explicitHeaderRow;
    const rawHeaders = rows[headerRow] || [];
    const columns = makeUniqueHeaders(rawHeaders);
    const originalByColumn = new Map(columns.map((column, index) => [column, normalizeValue(rawHeaders[index]) || column]));
    const details = { headerRow, columns, originalByColumn };
    workbookHeaders.set(cacheKey, details);
    return details;
  }

  function getSheetRecords(workbook, sheetName) {
    let workbookRecords = sheetRecordsCache.get(workbook);
    if (!workbookRecords) {
      workbookRecords = new Map();
      sheetRecordsCache.set(workbook, workbookRecords);
    }
    if (workbookRecords.has(sheetName)) return workbookRecords.get(sheetName);

    const headerDetails = getSheetHeaderDetails(workbook, sheetName);
    const rows = getSheetMatrix(workbook, sheetName);
    const headerRow = headerDetails.headerRow;
    const headers = headerDetails.columns;
    const records = rows.slice(headerRow + 1)
      .map(row => {
        const record = {};
        headers.forEach((header, index) => {
          record[header] = row[index] === undefined ? '' : row[index];
        });
        return record;
      });
    workbookRecords.set(sheetName, records);
    return records;
  }

  function isLinkedReportContext(source, sheetName) {
    return Boolean(state.linkedSurvey.active
      && state.linkedSurvey.result
      && source
      && source.workbook === state.workbook
      && sheetName === state.sheetName);
  }

  function getReportDataRows(source, sheetName) {
    return getLinkedReportData(source, sheetName).rows;
  }

  function getLinkedReportData(source, sheetName) {
    if (!isLinkedReportContext(source, sheetName)) {
      return { rows: getSheetRecords(source.workbook, sheetName), column: '' };
    }
    if (state.linkedSurvey.question) {
      return LinkedSurvey.enrichMatchedRows(state.linkedSurvey.result, state.linkedSurvey.question, {
        displayQuestion: getLinkedQuestionDisplayName()
      });
    }
    return {
      rows: state.linkedSurvey.result.matched.map(match => ({ ...match.primary, __linkedSiteKey: match.key })),
      column: ''
    };
  }

  function getSheetMatrix(workbook, sheetName) {
    let workbookSheets = sheetMatrixCache.get(workbook);
    if (!workbookSheets) {
      workbookSheets = new Map();
      sheetMatrixCache.set(workbook, workbookSheets);
    }
    if (workbookSheets.has(sheetName)) return workbookSheets.get(sheetName);

    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return [];
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      blankrows: false
    }).filter(row => row.some(cell => normalizeValue(cell) !== ''));
    workbookSheets.set(sheetName, matrix);
    return matrix;
  }

  function findHeaderRow(rows) {
    if (!rows.length) return 0;
    const surveyHeaderIndex = rows.slice(0, 6).findIndex(row => row.some(cell => /^survey$/i.test(normalizeValue(cell))));
    if (surveyHeaderIndex >= 0) return surveyHeaderIndex;
    return 0;
  }

  function renderCheckboxList(container, items, options = {}) {
    const normalizedItems = items.map(item => typeof item === 'string' ? { value: item, label: item } : item);
    const checkedValues = options.checkedValues === undefined
      ? normalizedItems.map(item => item.value)
      : options.checkedValues;
    const checked = new Set(checkedValues.map(normalizeValue));

    if (!normalizedItems.length) {
      container.innerHTML = `<div class="checklist-empty">${escapeHtml(options.emptyText || 'No options found')}</div>`;
      return;
    }

    container.innerHTML = normalizedItems.map((item, index) => {
      const id = `${container.id}-${hashString(item.value)}-${index}`;
      const isChecked = checked.has(normalizeValue(item.value));
      return `
        <label for="${id}">
          <input id="${id}" type="checkbox" value="${escapeAttr(item.value)}" data-label="${escapeAttr(item.label || item.value)}" ${isChecked ? 'checked' : ''}>
          <span>${escapeHtml(item.label || item.value)}</span>
        </label>
      `;
    }).join('');

    if (options.onChange) {
      container.querySelectorAll('input[type="checkbox"]').forEach(input => {
        input.addEventListener('change', options.onChange);
      });
    }
  }

  function getCheckedItems(container) {
    return Array.from(container.querySelectorAll('input[type="checkbox"]:checked'))
      .map(input => ({
        value: input.value,
        label: input.dataset.label || input.value
      }))
      .filter(item => item.value);
  }

  function getReportUniqueValues(rows, column) {
    const seen = new Set();
    const values = [];
    rows.forEach(row => {
      getReportValues(row[column]).forEach(value => {
        const key = normalizeForMatch(value);
        if (!key || seen.has(key)) return;
        seen.add(key);
        values.push(value);
      });
    });
    return values;
  }

  function getReportFilterValues(rows, column) {
    return Array.from(new Set(rows.flatMap(row => getResponseLabels(row[column]))))
      .sort((a, b) => a.localeCompare(b));
  }

  function pickRecordColumns(row, columns) {
    const record = {};
    columns.forEach(column => {
      record[column] = row[column] === undefined ? '' : row[column];
    });
    return record;
  }

  function getReportValue(value) {
    const normalized = normalizeValue(value);
    return normalized === NO_RESPONSE ? '' : normalized;
  }

  function getReportValues(value) {
    return getResponseLabels(value).map(getReportValue).filter(Boolean);
  }

  function sortReportAnswers(values, questionText) {
    const observed = values.filter(Boolean);
    const normalizedObserved = new Set(observed.map(normalizeForMatch));
    const hardSet = getHardCodedAnswerSet(questionText);
    if (hardSet) return hardSet;

    const matchingSet = ANSWER_SETS.find(set => {
      const normalizedSet = new Set(set.map(normalizeForMatch));
      return Array.from(normalizedObserved).every(value => normalizedSet.has(value));
    });

    if (matchingSet) return matchingSet;
    return [...observed].sort((a, b) => a.localeCompare(b));
  }

  function getHardCodedAnswerSet(questionText) {
    if (/How frequently did you engage with your students.*families this summer/i.test(questionText)) {
      return ['Never', 'Rarely', 'Occasionally', 'A moderate amount', 'A great deal'];
    }
    if (/To what extent did your child\/children enjoy participating/i.test(questionText)) {
      return ['Not at all', 'A little bit', 'Somewhat well', 'Quite well', 'Extremely well'];
    }
    if (/If given the opportunity, would you return/i.test(questionText)) {
      return ['Yes', 'Maybe', 'No'];
    }
    return null;
  }

  function cartesianProduct(arrays) {
    return arrays.reduce((acc, values) => acc.flatMap(prefix => values.map(value => [...prefix, value])), [[]]);
  }

  function uniqueList(values) {
    return Array.from(new Set(values));
  }

  function normalizeForMatch(value) {
    return normalizeValue(value)
      .toLowerCase()
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u201c\u201d]/g, '"')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function getHeatColor(percent) {
    const value = Math.max(0, Math.min(1, Number(percent) || 0));
    const lightness = 97 - (value * 35);
    const saturation = 58 + (value * 16);
    return `hsl(151, ${saturation}%, ${lightness}%)`;
  }

  function formatPercent(value) {
    return `${roundOne((Number(value) || 0) * 100)}%`;
  }

  function extractGoogleSheetId(url) {
    const match = normalizeValue(url).match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : '';
  }

  function safeSheetName(value) {
    return (normalizeValue(value).replace(/[\\/?*:[\]]+/g, ' ').trim() || 'Distribution').slice(0, 31);
  }

  function buildColumnStats(rows, columns) {
    const metrics = new Map(columns.map(column => [column, {
      values: [],
      answeredCount: 0,
      uniqueValues: new Set()
    }]));
    rows.forEach(row => columns.forEach(column => {
      const metric = metrics.get(column);
      const value = row[column];
      const normalized = normalizeValue(value);
      metric.values.push(value);
      if (normalized) {
        metric.answeredCount += 1;
        metric.uniqueValues.add(normalized);
      }
    }));

    return new Map(Array.from(metrics, ([column, metric]) => [column, {
      type: detectDataType(metric.values),
      answeredCount: metric.answeredCount,
      uniqueCount: metric.uniqueValues.size,
      missingCount: rows.length - metric.answeredCount
    }]));
  }


  function detectDataType(values) {
    const populated = values.filter(value => normalizeValue(value) !== '');
    if (!populated.length) return 'Empty';
    if (populated.every(value => value instanceof Date || /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(normalizeValue(value)))) return 'Date';
    if (populated.every(value => !Number.isNaN(Number(normalizeValue(value))))) return 'Number';
    if (populated.every(value => /^(true|false|yes|no)$/i.test(normalizeValue(value)))) return 'Boolean';
    return 'Text';
  }

  function updateAnalysisColumns() {
    const eligible = ChartRules.getSelectableChartColumns(state.allColumns, {
      hiddenColumns: state.hiddenAnalysisColumns
    });
    let analysisRows = state.allRows;
    const linked = state.linkedSurvey;
    if (linked.active && linked.result) {
      analysisRows = linked.result.matched.map(match => ({ ...match.primary, __linkedSiteKey: match.key }));
      if (linked.question) {
        const enriched = LinkedSurvey.enrichMatchedRows(linked.result, linked.question, {
          displayQuestion: getLinkedQuestionDisplayName()
        });
        analysisRows = enriched.rows;
        linked.column = enriched.column;
        const values = analysisRows.flatMap(row => getResponseLabels(row[linked.column])).filter(value => value !== NO_RESPONSE);
        state.columnStats.set(linked.column, {
          type: 'Linked survey',
          answeredCount: analysisRows.filter(row => getResponseLabels(row[linked.column]).some(value => value !== NO_RESPONSE)).length,
          uniqueCount: new Set(values.map(normalizeForMatch)).size,
          missingCount: analysisRows.filter(row => !getResponseLabels(row[linked.column]).some(value => value !== NO_RESPONSE)).length
        });
      }
    }
    state.columns = linked.active && linked.column ? [...eligible, linked.column] : eligible;
    state.eligibleChartColumns = linked.active && linked.column
      ? ChartRules.getSelectableChartColumns(state.columns, { hiddenColumns: state.hiddenAnalysisColumns })
      : eligible;
    state.rows = analysisRows.map(row => ({ ...pickRecordColumns(row, state.columns), __linkedSiteKey: row.__linkedSiteKey || '' }));
    state.excludedChartColumns = state.allColumns.filter(column => !eligible.includes(column));
  }

  function getColumnOptionLabel(column, compact = false) {
    const stats = state.columnStats.get(column);
    const displayName = getActiveColumnDisplayName(column);
    if (!stats) return displayName || 'No question selected';
    const statsLabel = getColumnStatsLabel(column);
    if (compact && !getWorkbookDictionary(state.workbook)) return statsLabel;
    if (compact) return `${displayName} · ${statsLabel}`;
    return `${displayName} — ${statsLabel}`;
  }

  function filterColumnOptions(select, query, selectedValue, includeNone) {
    const searchText = normalizeValue(query).toLowerCase();
    const matching = state.columns.filter(column => {
      if (!searchText) return true;
      return column.toLowerCase().includes(searchText)
        || getActiveColumnDisplayName(column).toLowerCase().includes(searchText);
    });
    if (selectedValue && state.columns.includes(selectedValue) && !matching.includes(selectedValue)) matching.unshift(selectedValue);
    const options = includeNone ? ['<option value="">No comparison</option>'] : [];
    options.push(...matching.map(column => `<option value="${escapeAttr(column)}">${escapeHtml(getColumnOptionLabel(column))}</option>`));
    select.innerHTML = options.join('');
    select.value = selectedValue && matching.includes(selectedValue) ? selectedValue : (includeNone ? '' : (matching[0] || ''));
  }

  function setActiveTab(tabName) {
    if (!state.workbook || !state.allRows.length) return;
    const panels = {
      charts: els.dashboardSection,
      report: document.getElementById('distributionSection'),
      preview: els.dataPreviewSection
    };
    if (!panels[tabName]) return;
    state.activeTab = tabName;
    Object.entries(panels).forEach(([name, panel]) => panel.classList.toggle('hidden', name !== tabName));
    els.tabButtons.forEach(button => {
      const active = button.dataset.tab === tabName;
      button.classList.toggle('is-active', active);
      button.setAttribute('aria-selected', String(active));
      button.tabIndex = active ? 0 : -1;
    });
    if (tabName === 'preview') renderDataPreview();
  }

  function showSheetPicker() {
    if (!state.workbook || state.workbook.SheetNames.length <= 1) {
      showToast('This dataset has only one sheet.', 'warning');
      return;
    }
    els.sheetPickerWrap.classList.remove('hidden');
    els.sheetSelect.focus();
  }

  function loadSampleData() {
    const sampleRows = [
      ['Region', 'Grade', 'Completed', 'Program rating', 'Would recommend', 'Support was helpful'],
      ['North', '6', 'Yes', 'Excellent', 'Yes', 'Very Helpful'],
      ['North', '7', 'Yes', 'Very Good', 'Yes', 'Helpful'],
      ['South', '8', 'No', 'Good', 'Maybe', 'Somewhat helpful'],
      ['East', '6', 'Yes', 'Excellent', 'Yes', 'Very Helpful'],
      ['West', '7', 'Yes', 'Good', 'Yes', 'Helpful'],
      ['South', '8', 'Yes', 'Fair', 'No', 'Somewhat helpful'],
      ['East', '6', 'Yes', 'Very Good', 'Yes', 'Very Helpful'],
      ['North', '7', 'No', 'Good', 'Maybe', 'Helpful'],
      ['West', '8', 'Yes', 'Excellent', 'Yes', 'Very Helpful'],
      ['South', '6', 'Yes', 'Very Good', 'Yes', 'Helpful'],
      ['East', '7', 'Yes', 'Good', 'Yes', 'Somewhat helpful'],
      ['West', '8', 'No', '', 'No', 'Not helpful']
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(sampleRows), 'Sample responses');
    resetDataset();
    upsertReportSource(UPLOADED_SOURCE_ID, 'Sample survey data', workbook);
    activateWorkbook(workbook, 'Sample survey data', 'Sample responses');
    showStatus('Sample data loaded. Explore the chart, report, and preview tabs.', '');
    showToast('Sample survey data loaded.');
  }

  function renderDataPreview() {
    if (!els.dataPreviewTable) return;
    if (!state.allRows.length || !state.allColumns.length) {
      els.previewStats.innerHTML = '';
      els.columnProfiles.innerHTML = '<div class="checklist-empty">Load a dataset to inspect its columns.</div>';
      els.dataPreviewTable.innerHTML = '<div class="report-empty"><strong>No data to preview.</strong></div>';
      els.previewResultCount.textContent = 'Showing 0 rows';
      return;
    }

    const searchText = state.previewSearch.trim().toLowerCase();
    const matchingRows = state.allRows.filter(row => !searchText || state.allColumns.some(column => normalizeValue(row[column]).toLowerCase().includes(searchText)));
    const visibleRows = matchingRows.slice(0, 100);
    const missingCells = state.allColumns.reduce((sum, column) => sum + (state.columnStats.get(column)?.missingCount || 0), 0);
    els.previewStats.innerHTML = [
      ['Rows', formatNumber(state.allRows.length)],
      ['Columns', formatNumber(state.allColumns.length)],
      ['Usable questions', formatNumber(state.columns.length)],
      ['Missing cells', formatNumber(missingCells)]
    ].map(([label, value]) => `<div class="preview-stat"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`).join('');

    const eligibleChartColumns = new Set(state.allColumns.filter(column => !state.excludedChartColumns.includes(column)));
    els.columnProfiles.innerHTML = state.allColumns.map(column => {
      const stats = state.columnStats.get(column);
      const automaticallyExcluded = !eligibleChartColumns.has(column);
      const inAnalysis = state.columns.includes(column);
      const displayName = getActiveColumnDisplayName(column);
      const exclusionReason = isLikelyMetadataColumn(column)
        ? 'metadata column'
        : (state.hiddenAnalysisColumns.has(column) ? 'hidden from analysis' : 'not selectable');
      return `<div class="column-profile"><strong title="${escapeAttr(displayName)}">${escapeHtml(displayName)}</strong><label title="${automaticallyExcluded ? `Automatically excluded: ${exclusionReason}` : 'Show or hide this column in chart analysis'}"><input type="checkbox" data-analysis-column="${escapeAttr(column)}" aria-label="Analyze ${escapeAttr(displayName)}" ${inAnalysis ? 'checked' : ''} ${automaticallyExcluded ? 'disabled' : ''}> Analyze</label><small>${escapeHtml(stats.type)} · ${formatNumber(stats.uniqueCount)} unique · ${formatNumber(stats.missingCount)} missing${automaticallyExcluded ? ` · ${escapeHtml(exclusionReason)}` : ''}</small></div>`;
    }).join('');
    els.columnProfiles.querySelectorAll('input[data-analysis-column]').forEach(input => input.addEventListener('change', () => {
      if (input.checked) state.hiddenAnalysisColumns.delete(input.dataset.analysisColumn);
      else state.hiddenAnalysisColumns.add(input.dataset.analysisColumn);
      updateAnalysisColumns();
      renderFileStats();
      renderAllCharts();
      renderDataPreview();
      showToast(input.checked ? 'Column restored to chart analysis.' : 'Column hidden from chart analysis.');
    }));

    els.previewResultCount.textContent = `Showing ${formatNumber(visibleRows.length)} of ${formatNumber(matchingRows.length)} matching rows`;
    els.dataPreviewTable.innerHTML = `<table><thead><tr><th>#</th>${state.allColumns.map(column => { const displayName = getActiveColumnDisplayName(column); return `<th title="${escapeAttr(displayName)}">${escapeHtml(truncateLabel(displayName, 42))}</th>`; }).join('')}</tr></thead><tbody>${visibleRows.map((row, index) => `<tr><td class="number">${index + 1}</td>${state.allColumns.map(column => {
      const value = displayCell(row[column]);
      return `<td title="${escapeAttr(value)}">${value ? escapeHtml(value) : '<span class="missing-value">Blank</span>'}</td>`;
    }).join('')}</tr>`).join('') || `<tr><td colspan="${state.allColumns.length + 1}">No rows match your search.</td></tr>`}</tbody></table>`;
  }

  function filterReportQuestions() {
    const searchText = normalizeValue(els.reportQuestionSearch.value).toLowerCase();
    els.questionChecklist.querySelectorAll('label').forEach(label => {
      label.classList.toggle('is-filtered-out', Boolean(searchText) && !label.textContent.toLowerCase().includes(searchText));
    });
    updateReportSelectionCount();
  }

  function setVisibleReportQuestions(checked) {
    els.questionChecklist.querySelectorAll('label:not(.is-filtered-out) input[type="checkbox"]').forEach(input => {
      input.checked = checked;
    });
    updateReportSelectionCount();
  }

  function updateReportSelectionCount() {
    if (!els.selectedQuestionCount) return;
    const count = els.questionChecklist.querySelectorAll('input[type="checkbox"]:checked').length;
    els.selectedQuestionCount.textContent = `${count} question${count === 1 ? '' : 's'} selected`;
  }

  function setReportMode(mode) {
    ['count', 'percentage', 'both'].forEach(value => els.distributionOutput.classList.toggle(`report-mode-${value}`, value === mode));
    document.querySelectorAll('[data-report-mode]').forEach(button => button.classList.toggle('is-active', button.dataset.reportMode === mode));
  }

  function setReportDensity(density) {
    ['compact', 'comfortable'].forEach(value => els.distributionOutput.classList.toggle(`density-${value}`, value === density));
    document.querySelectorAll('[data-density]').forEach(button => button.classList.toggle('is-active', button.dataset.density === density));
  }

  function setReportZoom(value) {
    state.reportZoom = Math.max(0.7, Math.min(1.4, Math.round(value * 10) / 10));
    els.distributionOutput.style.setProperty('--report-zoom', state.reportZoom);
    els.zoomValue.textContent = `${Math.round(state.reportZoom * 100)}%`;
  }

  function toggleReportFullscreen() {
    const panel = els.distributionOutput.closest('.report-output-panel');
    panel.classList.toggle('is-expanded');
    document.body.classList.toggle('has-expanded-view', panel.classList.contains('is-expanded'));
  }

  function startTitleEdit(card) {
    card.querySelector('.chart-title-row').classList.add('hidden');
    const input = card.querySelector('.chart-title-input');
    input.classList.remove('hidden');
    input.focus();
    input.select();
  }

  function finishTitleEdit(card) {
    card.querySelector('.chart-title-input').classList.add('hidden');
    card.querySelector('.chart-title-row').classList.remove('hidden');
  }

  function toggleChartExpanded(card) {
    card.classList.toggle('is-expanded');
    document.body.classList.toggle('has-expanded-view', card.classList.contains('is-expanded'));
    card.querySelector('.chart-more-menu').removeAttribute('open');
  }

  function handleGlobalKeydown(event) {
    if (event.key === 'Escape') {
      const expandedChart = document.querySelector('.chart-card.is-expanded');
      const expandedReport = document.querySelector('.report-output-panel.is-expanded');
      if (expandedChart) toggleChartExpanded(expandedChart);
      else if (expandedReport) toggleReportFullscreen();
      document.querySelectorAll('.menu-details[open]').forEach(menu => menu.removeAttribute('open'));
    }
    if ((event.key === 'ArrowLeft' || event.key === 'ArrowRight') && event.target.classList.contains('tab-button')) {
      event.preventDefault();
      const currentIndex = els.tabButtons.indexOf(event.target);
      const direction = event.key === 'ArrowRight' ? 1 : -1;
      const next = els.tabButtons[(currentIndex + direction + els.tabButtons.length) % els.tabButtons.length];
      next.focus();
      setActiveTab(next.dataset.tab);
    }
  }

  function setButtonLoading(button, loading, loadingLabel = 'Working…') {
    if (!button) return;
    if (loading) {
      if (!button.dataset.originalLabel) button.dataset.originalLabel = button.textContent;
      button.textContent = loadingLabel;
      button.disabled = true;
      button.setAttribute('aria-busy', 'true');
    } else {
      if (button.dataset.originalLabel) button.textContent = button.dataset.originalLabel;
      button.disabled = false;
      button.removeAttribute('aria-busy');
    }
  }

  function showToast(message, type = '') {
    if (!message) return;
    const toast = document.createElement('div');
    toast.className = `toast ${type}`.trim();
    toast.setAttribute('role', type === 'error' ? 'alert' : 'status');
    toast.textContent = message;
    els.toastRegion.appendChild(toast);
    window.setTimeout(() => toast.remove(), 4200);
  }

  function requestConfirmation(title, message, confirmLabel) {
    if (!els.confirmDialog || typeof els.confirmDialog.showModal !== 'function') return Promise.resolve(window.confirm(`${title}\n\n${message}`));
    els.confirmTitle.textContent = title;
    els.confirmMessage.textContent = message;
    els.confirmActionBtn.textContent = confirmLabel;
    els.confirmDialog.returnValue = 'cancel';
    els.confirmDialog.showModal();
    return new Promise(resolve => {
      els.confirmDialog.addEventListener('close', () => resolve(els.confirmDialog.returnValue === 'confirm'), { once: true });
    });
  }

  function showReportStatus(message, type) {
    els.reportStatus.textContent = message;
    els.reportStatus.className = `status-message ${type || ''}`.trim();
  }

  function resetDataset() {
    state.workbook = null;
    state.fileName = '';
    state.sheetName = '';
    state.rows = [];
    state.columns = [];
    state.allRows = [];
    state.allColumns = [];
    state.columnOriginalHeaders = new Map();
    state.rawColumnCount = 0;
    state.excludedChartColumns = [];
    state.eligibleChartColumns = [];
    state.selectedChartColumns = new Set();
    state.chartGenerationInProgress = false;
    state.charts.forEach(chart => {
      if (chart.chartInstance) chart.chartInstance.destroy();
    });
    state.charts = [];
    state.nextChartNumber = 1;
    state.sources = state.sources.filter(source => source.id !== UPLOADED_SOURCE_ID);
    state.reportResult = null;
    state.linkedSurvey = {
      active: false,
      secondaryWorkbook: null,
      secondaryFileName: '',
      secondarySheet: '',
      result: null,
      question: '',
      column: ''
    };
    els.chartGrid.innerHTML = '';
    showChartGenerationStatus('', '');
    els.fileStats.innerHTML = '';
    els.reportOutputTitle.textContent = 'No report generated yet';
    els.reportOutputMeta.textContent = 'Select questions and generate a breakdown report.';
    els.reportContextBar.classList.add('hidden');
    els.reportContextBar.innerHTML = '';
    els.distributionOutput.innerHTML = '<div class="report-empty"><span class="empty-state-icon" aria-hidden="true">▦</span><strong>Select questions and generate a breakdown report.</strong><span>Your report preview will appear here.</span></div>';
    renderDataset();
  }

  function showStatus(message, type) {
    els.statusMessage.textContent = message;
    els.statusMessage.className = `status-message ${type || ''}`.trim();
  }

  function makeId() {
    return `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  }

  function hashString(value) {
    let hash = 0;
    for (let index = 0; index < value.length; index += 1) {
      hash = ((hash << 5) - hash) + value.charCodeAt(index);
      hash |= 0;
    }
    return Math.abs(hash).toString(36);
  }

  function formatNumber(value) {
    return Number(value || 0).toLocaleString();
  }

  function roundOne(value) {
    return Math.round(value * 10) / 10;
  }

  function formatTableValue(value, suffix) {
    return suffix ? `${value}${suffix}` : formatNumber(value);
  }

  function safeFileName(value) {
    return normalizeValue(value).replace(/[\\/:*?"<>|]+/g, '-').slice(0, 80) || 'chart';
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function escapeAttr(value) {
    return escapeHtml(value);
  }
})();
