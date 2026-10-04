(function () {
  'use strict';

  const missingLibraries = ['XLSX', 'Chart', 'ChartDataLabels', 'ChartRules', 'DataDictionary', 'LinkedSurvey', 'SurveyCore', 'DataIO'].filter(name => !globalThis[name]);
  if (missingLibraries.length) {
    document.getElementById('statusMessage').textContent = 'The analyzer could not start. Refresh the page or check that all application files are available.';
    document.querySelectorAll('button, input, select').forEach(control => { control.disabled = true; });
    return;
  }
  const NO_RESPONSE = '(No response)';
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
  const sheetRowIndexCache = new WeakMap();
  const sheetWorkspaces = new WeakMap();
  const reportConfigs = new Map();
  let additionalReportFilters = [];
  let reportContextKey = '';
  let loadOperation = 0;
  let secondaryLoadOperation = 0;
  let reportOperation = 0;
  let activeReportWorker = null;
  let cancelReportWork = null;
  let activeLoadController = null;
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
    reportStale: false,
    reportMode: 'both',
    responseDelimiter: 'semicolon',
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
    reportFreshness: document.getElementById('reportFreshness'),
    includeLargeReportColumns: document.getElementById('includeLargeReportColumns'),
    responseDelimiter: document.getElementById('responseDelimiter'),
    linkMatchMode: document.getElementById('linkMatchMode'),
    linkDelimiter: document.getElementById('linkDelimiter'),
    previewSheetSelect: document.getElementById('previewSheetSelect'),
    chartLinkShortcut: document.getElementById('chartLinkShortcut'),
    reportExportMode: document.getElementById('reportExportMode'),
    selectVisibleQuestionsBtn: document.getElementById('selectVisibleQuestionsBtn'),
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
    reportFilterTools: document.getElementById('reportFilterTools'),
    reportFilterSearch: document.getElementById('reportFilterSearch'),
    additionalReportFilters: document.getElementById('additionalReportFilters'),
    addReportFilterBtn: document.getElementById('addReportFilterBtn'),
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
  els.reportSourceSelect.addEventListener('change', () => { saveReportConfig(); resetLinkedSurveyMatch(); updateAnalysisColumns(); renderAllCharts(); markReportStale(); renderReportControls(); renderLinkedSurveyPanel(); });
  els.reportDataSheetSelect.addEventListener('change', () => { saveReportConfig(); resetLinkedSurveyMatch(); updateAnalysisColumns(); renderAllCharts(); markReportStale(); renderReportColumns(); renderLinkedSurveyPanel(); renderFileStats(); });
  els.primaryBreakdownSelect.addEventListener('change', syncBreakdownQuestionSelection);
  els.reportFilterColumnSelect.addEventListener('change', () => renderReportFilterValues());
  els.reportFilterSearch.addEventListener('input', () => searchChecklist(els.reportFilterValues, els.reportFilterSearch.value));
  ['selectFilterValuesBtn', 'clearFilterValuesBtn'].forEach((id, index) => document.getElementById(id).addEventListener('click', () => {
    els.reportFilterValues.querySelectorAll('input[type=checkbox]').forEach(input => { input.checked = index === 0; });
    updateReportFilterSelectionNote(); markReportStale(); saveReportConfig();
  }));
  els.addReportFilterBtn.addEventListener('click', () => { additionalReportFilters.push({ column: '', values: null }); renderAdditionalReportFilters(); });
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
  els.tabButtons.forEach((button, index) => {
    button.addEventListener('click', () => setActiveTab(button.dataset.tab));
    button.addEventListener('keydown', event => {
      const target = event.key === 'ArrowRight' ? (index + 1) % els.tabButtons.length : event.key === 'ArrowLeft' ? (index + els.tabButtons.length - 1) % els.tabButtons.length : event.key === 'Home' ? 0 : event.key === 'End' ? els.tabButtons.length - 1 : -1;
      if (target < 0) return;
      event.preventDefault(); setActiveTab(els.tabButtons[target].dataset.tab); els.tabButtons[target].focus();
    });
  });
  els.previewSearch.addEventListener('input', event => {
    state.previewSearch = event.target.value;
    renderDataPreview();
  });
  els.reportQuestionSearch.addEventListener('input', filterReportQuestions);
  els.selectAllQuestionsBtn.addEventListener('click', () => setVisibleReportQuestions(true));
  els.selectVisibleQuestionsBtn.addEventListener('click', () => setVisibleReportQuestions(true, true));
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
  els.previewSheetSelect.addEventListener('change', () => loadSheet(els.previewSheetSelect.value));
  els.chartLinkShortcut.addEventListener('click', () => {
    saveReportConfig();
    const source = state.sources.find(item => item.workbook === state.workbook);
    if (source) els.reportSourceSelect.value = source.id;
    renderReportControls();
    els.reportDataSheetSelect.value = state.sheetName;
    renderReportColumns();
    setActiveTab('report');
    if (els.linkedSurveySetup.classList.contains('hidden')) toggleLinkedSurveySetup();
    els.linkPrimaryField.focus();
  });
  els.includeLargeReportColumns.addEventListener('change', () => { saveReportConfig(); markReportStale(); renderReportColumns(); });
  els.responseDelimiter.addEventListener('change', () => {
    state.responseDelimiter = els.responseDelimiter.value;
    reportConfigs.clear(); additionalReportFilters = [];
    markReportStale(); renderReportColumns();
  });
  [els.linkMatchMode, els.linkDelimiter].forEach(control => control.addEventListener('change', () => clearSurveyLink(false)));
  document.querySelector('.report-settings-panel').addEventListener('change', event => {
    if (event.target.closest('#linkedSurveyPanel')) return;
    markReportStale(); saveReportConfig();
  });
  els.reportNameInput.addEventListener('input', () => markReportStale());
  window.addEventListener('beforeunload', event => {
    if (state.charts.length || state.reportResult) { event.preventDefault(); event.returnValue = ''; }
  });

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
    const rows = DataIO.parseCsv(buffer);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), 'Sheet1');
    return workbook;
  }

  function parseWorkbook(buffer, extension = '') {
    if (extension === 'csv') return Promise.resolve().then(() => parseCsvWorkbook(buffer));
    if (typeof Worker === 'undefined' || location.protocol === 'file:') return Promise.resolve().then(() => XLSX.read(buffer, WORKBOOK_PARSE_OPTIONS));
    return new Promise((resolve, reject) => {
      let worker;
      try { worker = new Worker('data-worker.js?v=20261004-1'); }
      catch { resolve(XLSX.read(buffer, WORKBOOK_PARSE_OPTIONS)); return; }
      const timer = window.setTimeout(() => finish(new Error('Reading the workbook timed out. Try a smaller file or CSV.')), 60000);
      let settled = false;
      const finish = (error, workbook) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer); worker.terminate();
        if (error) reject(error); else resolve(workbook);
      };
      worker.onmessage = event => finish(event.data.type === 'success' ? null : new Error(event.data.message), event.data.workbook);
      worker.onerror = () => finish(new Error('The workbook reader could not start. Reload the app and try again.'));
      worker.postMessage({ buffer }, [buffer]);
    });
  }

  async function loadFile(file) {
    const extension = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(extension)) { showStatus('Please choose an .xlsx, .xls, or .csv file.', 'error'); return; }
    const operation = ++loadOperation;
    activeLoadController?.abort();
    showStatus('Reading your file…', 'loading');
    setButtonLoading(els.replaceFileBtn, true, 'Reading…');
    try {
      const workbook = await parseWorkbook(await file.arrayBuffer(), extension);
      if (operation !== loadOperation) return;
      const sheetName = getSurveySheetNames(workbook.SheetNames).find(name => getSheetRecords(workbook, name).length);
      if (!sheetName) throw new Error('No survey sheet with a header and response rows was found.');
      await yieldToBrowser();
      if (operation !== loadOperation) return;
      resetDataset(false);
      upsertReportSource(UPLOADED_SOURCE_ID, file.name, workbook);
      activateWorkbook(workbook, file.name, sheetName);
      showStatus(state.allRows.length > 50000 ? 'Large dataset loaded. Generate only the questions you need.' : 'File loaded. Uploaded data is processed in this browser.', state.allRows.length > 50000 ? 'warning' : '');
      showToast(`${file.name} loaded successfully.`);
    } catch (error) {
      if (operation !== loadOperation) return;
      showStatus(`${error.message || 'The file could not be opened.'} Your current analysis has been kept.`, 'error');
      showToast('The new file could not be opened. Your current analysis is unchanged.', 'error');
    } finally {
      if (operation === loadOperation) setButtonLoading(els.replaceFileBtn, false);
      els.fileInput.value = '';
    }
  }

  function activateWorkbook(workbook, fileName, sheetName) {
    if (state.workbook !== workbook) {
      state.charts.forEach(chart => { chart.chartInstance?.destroy(); chart.chartInstance = null; });
      if (state.workbook && state.sheetName) {
        if (!sheetWorkspaces.has(state.workbook)) sheetWorkspaces.set(state.workbook, new Map());
        sheetWorkspaces.get(state.workbook).set(state.sheetName, { charts: state.charts, selected: state.selectedChartColumns, hidden: state.hiddenAnalysisColumns, next: state.nextChartNumber });
      }
      state.charts = []; state.sheetName = '';
    }
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
    saveReportConfig();
    if (!sheetWorkspaces.has(state.workbook)) sheetWorkspaces.set(state.workbook, new Map());
    const workspaces = sheetWorkspaces.get(state.workbook);
    if (state.sheetName) {
      state.charts.forEach(chart => { chart.chartInstance?.destroy(); chart.chartInstance = null; });
      workspaces.set(state.sheetName, { charts: state.charts, selected: state.selectedChartColumns, hidden: state.hiddenAnalysisColumns, next: state.nextChartNumber });
    }
    const restored = workspaces.get(sheetName);
    state.charts = restored?.charts || [];
    state.nextChartNumber = restored?.next || 1;
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

    if (!rawRows.length) showStatus('This sheet is empty. Choose another sheet to continue.', 'warning');

    const headerDetails = getSheetHeaderDetails(state.workbook, sheetName);
    const allColumns = headerDetails.columns;
    const allRows = getSheetRecords(state.workbook, sheetName);
    const columnStats = buildColumnStats(allRows, allColumns);
    const chartColumns = ChartRules.getSelectableChartColumns(allColumns);

    state.rawColumnCount = allColumns.length;
    state.excludedChartColumns = allColumns.filter(column => !chartColumns.includes(column));
    state.eligibleChartColumns = chartColumns;
    state.selectedChartColumns = new Set();
    state.allColumns = allColumns;
    state.columnOriginalHeaders = headerDetails.originalByColumn;
    state.allRows = allRows;
    state.hiddenAnalysisColumns = restored?.hidden || new Set();
    state.selectedChartColumns = restored?.selected || new Set();
    state.columnStats = columnStats;
    updateAnalysisColumns();

    els.sheetSelect.value = sheetName;
    if (rawRows.length) showStatus('File ready. Your data stays in this browser.', '');
    renderDataset();
    populateSelect(els.previewSheetSelect, getSurveySheetNames(state.workbook.SheetNames), sheetName, false);
  }

  function makeUniqueHeaders(headerRow) { return SurveyCore.uniqueHeaders(headerRow); }

  function renderDataset() {
    const hasData = state.rows.length > 0 && state.columns.length > 0;
    const hasDataset = Boolean(state.workbook);
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
    const source = state.activeTab === 'report' ? getSelectedReportSource() : null;
    const sheetName = source ? els.reportDataSheetSelect.value : state.sheetName;
    const rows = source && sheetName ? getSheetRecords(source.workbook, sheetName) : state.allRows;
    const columns = source && sheetName ? getSheetColumns(source.workbook, sheetName) : state.allColumns;
    els.datasetFileName.textContent = source?.name || state.fileName;
    const stats = [[state.activeTab === 'report' ? 'Primary sheet' : 'Sheet', sheetName || 'None'], ['Rows', formatNumber(rows.length)], ['Columns', formatNumber(columns.length)]];
    els.fileStats.innerHTML = stats.map(([label, value]) => `<div class="stat"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`).join('');
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
      binaryLabels: false,
      delimiter: state.responseDelimiter,
      summaryPage: 0,
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
    const expanded = document.querySelector('.expanded-dialog');
    if (expanded?.querySelector('.chart-card')) expanded.close();
    renderChartEligibilitySummary();
    state.charts.forEach(chart => { chart.chartInstance?.destroy(); chart.chartInstance = null; });
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
    bindSelect(card, '.response-delimiter', chart, 'delimiter');
    bindCheckbox(card, '.binary-labels', chart, 'binaryLabels');
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
      chart.summaryPage = 0;
      renderSummaryTable(chart, card, getSummaryResult(chart));
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

    ['prev', 'next'].forEach(direction => card.querySelector(`.summary-${direction}`).addEventListener('click', () => { chart.summaryPage = Math.max(0, (chart.summaryPage || 0) + (direction === 'next' ? 1 : -1)); renderSummaryTable(chart, card, getSummaryResult(chart)); }));
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
      chart.summaryPage = 0;
      chart.selectedResponses.clear();
      updateChartCard(chart, card);
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

    const filteredRows = applyFilters(state.rows, chart.filters, chart.delimiter);
    card.querySelector('.row-count').textContent = `Showing ${formatNumber(filteredRows.length)} of ${formatNumber(state.rows.length)} rows`;

    const result = isComparison
      ? buildComparisonResult(filteredRows, chart)
      : buildSingleColumnResult(filteredRows, chart);

    const validResponses = filteredRows.filter(row => SurveyCore.labels(row[chart.primaryColumn], { delimiter: chart.delimiter }).length > 0).length;
    const categoryTotal = result.type === 'single' ? result.items.length : result.labels.length;
    card.querySelector('.valid-response-count').textContent = `${formatNumber(validResponses)} valid response${validResponses === 1 ? '' : 's'}`;
    card.querySelector('.category-count').textContent = `${formatNumber(categoryTotal)} categor${categoryTotal === 1 ? 'y' : 'ies'}`;
    const activeFilterCount = chart.filters.filter(filter => filter.selected.size).length;
    card.querySelector('.analysis-summary').textContent = chart.compareColumn
      ? `Compared by ${getActiveColumnDisplayName(chart.compareColumn)}${activeFilterCount ? ` · ${activeFilterCount} filter${activeFilterCount === 1 ? '' : 's'}` : ''}`
      : (activeFilterCount ? `${activeFilterCount} active filter${activeFilterCount === 1 ? '' : 's'}` : 'No filters or comparison');
    card.querySelector('.question-detail').textContent = getColumnOptionLabel(chart.primaryColumn);
    const multiSelect = filteredRows.some(row => SurveyCore.labels(row[chart.primaryColumn], { delimiter: chart.delimiter }).length > 1);
    if (multiSelect && ['pie', 'doughnut'].includes(chart.chartType)) { chart.chartType = 'horizontalBar'; card.querySelector('.chart-type').value = chart.chartType; }
    card.querySelector('.chart-area canvas').setAttribute('aria-label', `${getActiveColumnDisplayName(chart.primaryColumn)}. Values are available in the summary table below.`);
    card.querySelector('.chart-area canvas').setAttribute('role', 'img');
    card.querySelector('.compare-type').value = chart.compareType;
    card.querySelector('.table-note').textContent = multiSelect || result.overlapping ? 'Multiple selections: percentages may total above 100%. Each respondent is counted once per category.' : `Percentages use ${chart.includeBlanks ? 'all included' : 'answered'} respondents.`;

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

      const values = getUniqueValues(state.rows, filter.column, chart.delimiter);
      const searchText = (filter.search || '').toLowerCase();
      const visibleValues = values.filter(value => value.toLowerCase().includes(searchText));
      const valuesWrap = filterCard.querySelector('.filter-values');
      valuesWrap.innerHTML = visibleValues.map(value => {
        const id = `${filter.id}-${hashString(value)}`;
        return `
          <label for="${id}">
            <input id="${id}" type="checkbox" value="${escapeAttr(value)}" ${Array.from(filter.selected).some(selected => normalizeForMatch(selected) === normalizeForMatch(value)) ? 'checked' : ''}>
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

  function applyFilters(rows, filters, delimiter = state.responseDelimiter) {
    const activeFilters = filters.filter(filter => filter.column && filter.selected.size);
    if (!activeFilters.length) return rows;

    return rows.filter(row => activeFilters.every(filter => {
      return getResponseLabels(row[filter.column]).some(value => Array.from(filter.selected).some(selected => normalizeForMatch(selected) === normalizeForMatch(value)));
    }));
  }

  function buildSingleColumnResult(rows, chart) {
    const transformed = chart.binaryLabels ? rows.map(row => ({ ...row, [chart.primaryColumn]: SurveyCore.labels(row[chart.primaryColumn], { delimiter: chart.delimiter }).map(label => ChartRules.getDisplayAnswerLabel(label, [], true)) })) : rows;
    return { type: 'single', rows, ...SurveyCore.summarize(transformed, chart.primaryColumn, { ...chart, delimiter: chart.delimiter, blankLabel: NO_RESPONSE }) };
  }

  function buildComparisonResult(rows, chart) {
    const transformed = chart.binaryLabels ? rows.map(row => ({ ...row, [chart.primaryColumn]: SurveyCore.labels(row[chart.primaryColumn], { delimiter: chart.delimiter }).map(label => ChartRules.getDisplayAnswerLabel(label, [], true)) })) : rows;
    let result = SurveyCore.comparison(transformed, chart.primaryColumn, chart.compareColumn, { ...chart, delimiter: chart.delimiter, blankLabel: NO_RESPONSE });
    const overlapping = rows.some(row => SurveyCore.labels(row[chart.primaryColumn], { delimiter: chart.delimiter }).length > 1 || SurveyCore.labels(row[chart.compareColumn], { delimiter: chart.delimiter }).length > 1);
    if (chart.compareType === 'stacked100' && overlapping) chart.compareType = 'grouped';
    const valueMode = chart.compareType === 'stacked100' ? 'primaryPercent' : chart.compareValueMode;
    result.datasets = result.compareLabels.map((compare, index) => ({
      label: compare,
      data: result.labels.map(primary => {
        const count = result.matrix.get(primary).get(compare) || 0;
        const denominator = valueMode === 'primaryPercent' ? result.primaryRespondentTotals.get(primary) : valueMode === 'comparePercent' ? result.compareRespondentTotals.get(compare) : result.respondentTotal;
        return valueMode === 'counts' ? count : denominator ? roundOne(count / denominator * 100) : 0;
      }),
      backgroundColor: COLORS[index % COLORS.length], borderColor: COLORS[index % COLORS.length], borderWidth: 1
    }));
    return { ...result, type: 'comparison', rows, valueMode, overlapping };
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
            max: horizontal && chart.compareType === 'stacked100' ? 100 : undefined,
            grid: { color: CHART_GRID }
          },
          y: {
            stacked,
            beginAtZero: true,
            max: !horizontal && chart.compareType === 'stacked100' ? 100 : undefined,
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

  function getSummaryResult(chart) {
    const rows = applyFilters(state.rows, chart.filters, chart.delimiter);
    return chart.compareColumn ? buildComparisonResult(rows, chart) : buildSingleColumnResult(rows, chart);
  }

  function paginateSummary(chart, card, items) {
    const pageCount = Math.max(1, Math.ceil(items.length / TABLE_ROW_LIMIT));
    chart.summaryPage = Math.min(chart.summaryPage || 0, pageCount - 1);
    card.querySelector('.summary-page').textContent = `Page ${chart.summaryPage + 1} of ${pageCount}`;
    card.querySelector('.summary-prev').disabled = chart.summaryPage === 0;
    card.querySelector('.summary-next').disabled = chart.summaryPage + 1 >= pageCount;
    return items.slice(chart.summaryPage * TABLE_ROW_LIMIT, (chart.summaryPage + 1) * TABLE_ROW_LIMIT);
  }

  function renderSummaryTable(chart, card, result) {

    if (result.type === 'comparison') {
      renderComparisonTable(chart, card, result);
      return;
    }

    const searchText = chart.summarySearch.trim().toLowerCase();
    const matchingItems = result.items.filter(item => item.response.toLowerCase().includes(searchText));
    const visibleItems = paginateSummary(chart, card, matchingItems);
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
                  <input type="checkbox" class="response-select" aria-label="Select ${escapeAttr(item.response)}" value="${escapeAttr(item.response)}" ${chart.selectedResponses.has(item.response) ? 'checked' : ''}>
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
    const matchingLabels = result.labels.filter(label => normalizeForMatch(label).includes(normalizeForMatch(chart.summarySearch)));
    renderTableNote(card, matchingLabels.length, 'comparison rows');
    const rows = paginateSummary(chart, card, matchingLabels).map(primary => {
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
    if (totalRows > TABLE_ROW_LIMIT) note.textContent += ` ${formatNumber(totalRows)} ${label}; use Previous / Next to view all.`;
  }

  function getUniqueValues(rows, column, delimiter = state.responseDelimiter) {
    const columnLabels = getColumnNonBlankLabels(rows, column);
    const useYesNoLabels = false;
    const values = new Map();
    rows.forEach(row => getDisplayResponseLabelsForColumn(row[column], columnLabels, useYesNoLabels, delimiter).forEach(label => { if (!values.has(normalizeForMatch(label))) values.set(normalizeForMatch(label), label); }));
    return [...values.values()].sort((a, b) => a.localeCompare(b));
  }

  function getColumnNonBlankLabels(rows, column) {
    return rows.flatMap(row => SurveyCore.labels(row[column], { delimiter: state.responseDelimiter }));
  }

  function getDisplayResponseLabelsForColumn(value, nonBlankLabels, useYesNoLabels = false, delimiter = state.responseDelimiter) {
    const labels = getResponseLabels(value, delimiter);
    return labels.map(label => ChartRules.getDisplayAnswerLabel(label, nonBlankLabels, useYesNoLabels));
  }

  function getResponseLabels(value, delimiter = state.responseDelimiter) {
    const labels = SurveyCore.labels(value, { delimiter });
    return labels.length ? labels : [NO_RESPONSE];
  }

  function normalizeValue(value) { return Array.isArray(value) ? value.map(SurveyCore.text).join('; ') : SurveyCore.text(value); }

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
    const rows = applyFilters(state.rows, chart.filters, chart.delimiter);
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
    const rows = applyFilters(state.rows, chart.filters, chart.delimiter);
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
    const hasDataset = Boolean(state.workbook);
    els.linkedSurveyPanel.classList.toggle('hidden', !hasDataset);
    if (!hasDataset) return;

    const source = getSelectedReportSource();
    const sheetName = els.reportDataSheetSelect.value;
    const columns = source && sheetName ? getSheetColumns(source.workbook, sheetName) : [];
    els.linkPrimaryName.textContent = source?.name || 'Report source';
    els.linkPrimarySheet.textContent = sheetName || 'No sheet selected';
    populateSelect(els.linkPrimaryField, columns, els.linkPrimaryField.value || pickLinkField(columns), false, 'None', column => getDisplayColumnName(source.workbook, sheetName, column));
    els.toggleLinkedSurveyBtn.textContent = els.linkedSurveySetup.classList.contains('hidden')
      ? (state.linkedSurvey.active ? 'Edit link' : 'Set up link')
      : 'Hide setup';
    renderLinkedSurveySource();
    renderLinkDiagnostics();
  }

  function getSecondaryWorkbook() {
    return els.linkSecondarySource.value === 'file' ? state.linkedSurvey.secondaryWorkbook : getSelectedReportSource()?.workbook;
  }

  function getLinkedQuestionDisplayName() {
    const linked = state.linkedSurvey;
    if (!linked.question) return '';
    return getDisplayColumnName(linked.matchedSecondaryWorkbook || getSecondaryWorkbook(), linked.secondarySheet || els.linkSecondarySheet.value, linked.question);
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
      ? getSurveySheetNames(workbook.SheetNames).filter(name => fileMode || name !== els.reportDataSheetSelect.value)
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
    const operation = ++secondaryLoadOperation;
    showLinkValidation('Reading the secondary survey...', 'loading');
    try {
      const workbook = await parseWorkbook(await file.arrayBuffer(), extension);
      if (operation !== secondaryLoadOperation) return;
      if (!workbook.SheetNames.length) throw new Error('No sheets were found in the secondary file.');
      clearSurveyLink(false);
      state.linkedSurvey.secondaryWorkbook = workbook;
      state.linkedSurvey.secondaryFileName = file.name;
      renderLinkedSurveySource();
      showLinkValidation('Secondary survey loaded. Select matching fields, then match the surveys.', '');
    } catch (error) {
      console.error(error);
      if (operation !== secondaryLoadOperation) return;
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
      const source = getSelectedReportSource();
      const primarySheet = els.reportDataSheetSelect.value;
      const primaryRows = getSheetRecords(source.workbook, primarySheet);
      const result = LinkedSurvey.analyzeLink(primaryRows, secondaryRows, primaryField, secondaryField, { mode: els.linkMatchMode.value });
      resetLinkedSurveyMatch();
      state.linkedSurvey.primaryWorkbook = source.workbook;
      state.linkedSurvey.primarySheet = primarySheet;
      state.linkedSurvey.secondarySheet = secondarySheet;
      state.linkedSurvey.matchedSecondaryWorkbook = workbook;
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
    linked.matchedSecondaryWorkbook = null;
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
    if (result.normalizationCollisions?.length) warnings.push('Different source identifiers normalized to the same value. Review duplicate details or switch to exact matching.');
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
    state.reportStale = false;
    markReportStale();
    els.reportFreshness.classList.add('hidden');
    updateReportExportState();
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
    const linked = state.linkedSurvey;
    const columns = getSheetColumns(linked.primaryWorkbook, linked.primarySheet);
    els.linkDetailsBody.innerHTML = `<table><thead><tr><th>Source row</th><th>Issue</th>${columns.map(column => `<th>${escapeHtml(getDisplayColumnName(linked.primaryWorkbook, linked.primarySheet, column))}</th>`).join('')}</tr></thead><tbody>${result.unmatched.map(item => `<tr><td class="number">${item.rowNumber}</td><td>${escapeHtml(item.reason)}</td>${columns.map(column => `<td>${escapeHtml(displayCell(item.row[column]))}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
    els.linkDetailsDialog.showModal();
  }

  function downloadUnmatchedRecords() {
    const result = state.linkedSurvey.result;
    if (!result || !result.unmatched.length) return;
    const linked = state.linkedSurvey;
    const columns = getSheetColumns(linked.primaryWorkbook, linked.primarySheet);
    const rows = [['Source row', 'Matching value', 'Issue', ...columns.map(column => getDisplayColumnName(linked.primaryWorkbook, linked.primarySheet, column))]];
    result.unmatched.forEach(item => rows.push([item.rowNumber, item.value, item.reason, ...columns.map(column => displayCell(item.row[column]))]));
    downloadCsv(rows, `${safeFileName(linked.primarySheet)}-unmatched-records.csv`);
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
    const sheetId = extractGoogleSheetId(els.publicSheetUrl.value);
    if (!sheetId) { showStatus('Paste a valid public Google Sheets link.', 'error'); return; }
    const operation = ++loadOperation;
    activeLoadController?.abort(); activeLoadController = new AbortController();
    const controller = activeLoadController;
    const timer = window.setTimeout(() => controller.abort(), 60000);
    showStatus('Loading public Google Sheet…', 'loading');
    setButtonLoading(els.loadPublicSheetBtn, true, 'Loading…');
    try {
      const response = await fetch(`https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`, { signal: controller.signal, credentials: 'omit' });
      if (!response.ok) throw new Error('The sheet must be public and downloadable without signing in.');
      const workbook = await parseWorkbook(await response.arrayBuffer());
      if (operation !== loadOperation) return;
      const sheetName = getSurveySheetNames(workbook.SheetNames).find(name => getSheetRecords(workbook, name).length);
      if (!sheetName) throw new Error('No survey sheet with response rows was found.');
      const sourceName = `Google Sheet ${sheetId.slice(0, 8)}`;
      resetLinkedSurveyMatch();
      upsertReportSource(`google-${sheetId}`, sourceName, workbook);
      activateWorkbook(workbook, sourceName, sheetName);
      els.reportSourceSelect.value = `google-${sheetId}`; renderReportControls();
      showStatus('Public Google Sheet loaded. Data is held in this browser session.', '');
    } catch (error) {
      if (operation === loadOperation) showStatus(error.name === 'AbortError' ? 'Loading timed out. Your current analysis was kept.' : `${error.message} Your current analysis was kept.`, 'error');
    } finally {
      clearTimeout(timer);
      if (operation === loadOperation) setButtonLoading(els.loadPublicSheetBtn, false);
    }
  }

  function upsertReportSource(id, name, workbook) {
    const existing = state.sources.find(source => source.id === id);
    if (existing) {
      if (existing.workbook !== workbook) {
        for (const key of reportConfigs.keys()) if (key.startsWith(`${id}\u0000`)) reportConfigs.delete(key);
      }
      existing.name = name;
      existing.workbook = workbook;
    } else {
      state.sources.push({ id, name, workbook });
    }
    invalidateGeneratedReport();
    state.previewSearch = '';
    if (els.previewSearch) els.previewSearch.value = '';
  }

  function getSelectedReportSource() {
    return state.sources.find(source => source.id === els.reportSourceSelect.value) || state.sources[0] || null;
  }

  function renderReportControls() {
    const previousSource = els.reportSourceSelect.value;
    els.reportSourceSelect.innerHTML = state.sources.length ? state.sources.map(source => `<option value="${escapeAttr(source.id)}">${escapeHtml(source.name)}</option>`).join('') : '<option value="">No source loaded</option>';
    els.reportSourceSelect.value = state.sources.some(source => source.id === previousSource) ? previousSource : state.sources.find(source => source.workbook === state.workbook)?.id || state.sources[0]?.id || '';
    const source = getSelectedReportSource();
    [els.reportNameInput, els.reportDataSheetSelect, els.primaryBreakdownSelect, els.reportFilterColumnSelect, els.generateReportBtn].forEach(control => { control.disabled = !source; });
    if (!source) {
      [els.reportDataSheetSelect, els.primaryBreakdownSelect, els.reportFilterColumnSelect].forEach(control => { control.innerHTML = '<option value="">No data loaded</option>'; });
      renderCheckboxList(els.questionChecklist, [], { emptyText: 'Load a dataset to choose report questions.' });
      updateReportSelectionCount(); updateReportExportState(); return;
    }
    const sheetNames = getSurveySheetNames(source.workbook.SheetNames);
    const sameSource = reportContextKey.startsWith(source.id + '\u0000');
    const preferred = sameSource ? reportContextKey.split('\u0000')[1] : source.workbook === state.workbook ? state.sheetName : sheetNames[0];
    populateSelect(els.reportDataSheetSelect, sheetNames, preferred || sheetNames[0], false);
    if (!normalizeValue(els.reportNameInput.value)) els.reportNameInput.value = 'Question breakdown';
    renderReportColumns(); updateReportExportState();
  }

  function saveReportConfig() {
    if (!reportContextKey) return;
    reportConfigs.set(reportContextKey, {
      questions: getCheckedItems(els.questionChecklist).map(item => item.value),
      breakdown: els.primaryBreakdownSelect.value,
      filter: els.reportFilterColumnSelect.value,
      values: getCheckedItems(els.reportFilterValues).map(item => item.value),
      name: els.reportNameInput.value,
      includeLarge: els.includeLargeReportColumns.checked,
      additionalFilters: additionalReportFilters.map(filter => ({ column: filter.column, values: filter.values ? [...filter.values] : null }))
    });
  }

  function updateReportExportState() {
    const disabled = !state.reportResult || state.reportStale || Boolean(activeReportWorker);
    els.downloadReportCsvBtn.disabled = disabled; els.downloadReportXlsxBtn.disabled = disabled;
  }

  function markReportStale() {
    ++reportOperation;
    cancelReportWork?.();
    setButtonLoading(els.generateReportBtn, false);
    if (state.reportResult) {
      state.reportStale = true;
      els.reportFreshness.textContent = 'Settings changed. Regenerate the report before exporting.';
      els.reportFreshness.classList.remove('hidden');
    }
    updateReportExportState();
  }

  async function calculateReportSections(rows, questions, breakdownColumns, operation) {
    const options = { delimiter: state.responseDelimiter };
    if (typeof Worker === 'undefined' || location.protocol === 'file:') {
      const sections = [];
      for (const question of questions) {
        await yieldToBrowser();
        if (operation !== reportOperation) throw new DOMException('Report cancelled', 'AbortError');
        sections.push(SurveyCore.reportSection(rows, question.display, question.column, breakdownColumns, { ...options, answerOrder: question.answerOrder }));
      }
      return sections;
    }
    return new Promise((resolve, reject) => {
      const worker = new Worker('report-worker.js?v=20261004-1');
      activeReportWorker = worker;
      let finished = false;
      const timer = window.setTimeout(() => finish(new Error('Report generation timed out. Select fewer questions or groups.')), 120000);
      const finish = (error, sections) => {
        if (finished) return;
        finished = true; clearTimeout(timer); worker.terminate();
        if (activeReportWorker === worker) { activeReportWorker = null; cancelReportWork = null; }
        if (error) reject(error); else resolve(sections);
      };
      cancelReportWork = () => finish(new DOMException('Report cancelled', 'AbortError'));
      worker.onmessage = event => {
        const data = event.data;
        if (data.id !== operation) return;
        if (data.error) finish(new Error(data.error));
        else if (data.sections) finish(null, data.sections);
        else if (data.progress) showReportStatus(`Generating question ${data.progress.completed} of ${data.progress.total}…`, 'loading');
      };
      worker.onerror = () => finish(new Error('The report worker could not start. Reload and try again.'));
      worker.postMessage({ id: operation, rows, questions, breakdownColumns, options });
    });
  }

  function renderReportColumns() {
    const source = getSelectedReportSource();
    const sheetName = els.reportDataSheetSelect.value;
    reportContextKey = source ? `${source.id}\u0000${sheetName}` : '';
    const saved = reportConfigs.get(reportContextKey);
    if (saved) { els.reportNameInput.value = saved.name; els.includeLargeReportColumns.checked = saved.includeLarge; }
    const primaryColumns = source && sheetName ? getSheetColumns(source.workbook, sheetName) : [];
    const reportData = source && sheetName ? getLinkedReportData(source, sheetName) : { rows: [], column: '' };
    const dataRows = reportData.rows;
    const columns = uniqueList([...primaryColumns, reportData.column].filter(Boolean)).filter(column => !isLinkedReportContext(source, sheetName) || column !== '__linkedSiteKey');
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
      .filter(item => item.uniqueCount > 0 && (els.includeLargeReportColumns.checked || item.uniqueCount <= REPORT_UNIQUE_VALUE_LIMIT))
      .map(item => item.column);
    const linkedColumn = reportData.column;
    if (isLinkedReportContext(source, sheetName)) state.linkedSurvey.column = linkedColumn;
    const eligibleBreakdownColumns = uniqueList([...eligibleColumns, linkedColumn].filter(Boolean));
    const eligibleQuestionColumns = eligibleColumns.filter(column => column !== linkedColumn && !(source?.workbook === state.workbook && sheetName === state.sheetName && state.hiddenAnalysisColumns.has(column)));
    const eligibleFilterColumns = filterColumnStats
      .filter(item => item.nonBlankUniqueCount > 0 && item.uniqueCount > 0 && item.uniqueCount < REPORT_FILTER_UNIQUE_VALUE_LIMIT)
      .map(item => item.column);
    const emptyColumns = columnStats.filter(item => item.uniqueCount === 0);
    const ignoredColumns = columnStats.filter(item => item.uniqueCount > REPORT_UNIQUE_VALUE_LIMIT);
    const responseColumns = primaryColumns.filter(column => !isLikelyMetadataColumn(column));
    const defaultResponseColumns = responseColumns.filter(column => eligibleQuestionColumns.includes(column));
    const displayColumn = column => getDisplayColumnName(source.workbook, sheetName, column);
    renderCheckboxList(els.questionChecklist, eligibleQuestionColumns.map(column => ({ value: column, label: displayColumn(column) })), {
      checkedValues: saved ? saved.questions : defaultResponseColumns.length ? defaultResponseColumns : eligibleQuestionColumns,
      emptyText: 'No response columns found',
      onChange: () => { updateReportSelectionCount(); markReportStale(); saveReportConfig(); }
    });
    populateSelect(els.primaryBreakdownSelect, eligibleBreakdownColumns, linkedColumn, true, 'No main breakdown', displayColumn);
    els.primaryBreakdownSelect.value = eligibleBreakdownColumns.includes(saved?.breakdown) ? saved.breakdown : linkedColumn || '';
    populateSelect(els.reportFilterColumnSelect, eligibleFilterColumns, '', true, 'No filter', displayColumn);
    els.reportFilterColumnSelect.value = eligibleFilterColumns.includes(saved?.filter) ? saved.filter : '';
    renderReportFilterValues(saved?.values);
    additionalReportFilters = (saved?.additionalFilters || []).map(filter => ({ ...filter }));
    renderAdditionalReportFilters();
    updateReportColumnNote(ignoredColumns, emptyColumns, displayColumn);
    filterReportQuestions();
    updateReportSelectionCount();
  }

  function renderReportFilterValues(selectedValues) {
    const source = getSelectedReportSource();
    const sheetName = els.reportDataSheetSelect.value;
    const filterColumn = els.reportFilterColumnSelect.value;
    els.reportFilterTools.classList.toggle('hidden', !filterColumn);
    els.reportFilterSearch.value = '';
    const dataRows = source && sheetName ? getReportDataRows(source, sheetName) : [];

    if (!filterColumn || !dataRows.length) {
      renderCheckboxList(els.reportFilterValues, [], { emptyText: 'No filter values found' });
      els.reportFilterValues.classList.add('hidden');
      els.reportFilterNote.textContent = `Filter columns must have fewer than ${REPORT_FILTER_UNIQUE_VALUE_LIMIT} unique values.`;
      return;
    }

    const values = getReportFilterValues(dataRows, filterColumn);
    renderCheckboxList(els.reportFilterValues, values.map(value => ({ value, label: value })), {
      checkedValues: Array.isArray(selectedValues) ? selectedValues : values,
      emptyText: 'No filter values found',
      onChange: () => { updateReportFilterSelectionNote(); markReportStale(); saveReportConfig(); }
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
      .forEach(input => { input.checked = false; });
    updateReportSelectionCount(); markReportStale();
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

  async function generateDistributionReport() {
    const source = getSelectedReportSource();
    if (!source) { showReportStatus('Load a source first.', 'error'); return; }
    markReportStale();
    const operation = ++reportOperation;
    setButtonLoading(els.generateReportBtn, true, 'Generating…');
    showReportStatus('Generating the breakdown report…', 'loading');
    try {
      const sheetName = els.reportDataSheetSelect.value;
      const dataRows = getReportDataRows(source, sheetName);
      const filter = getReportFilter(dataRows);
      const filteredRows = applyReportFilter(dataRows, filter);
      const breakdownColumns = [els.primaryBreakdownSelect.value].filter(Boolean);
      const questions = getCheckedItems(els.questionChecklist).filter(item => !breakdownColumns.includes(item.value)).map(item => ({ column: item.value, display: item.label, answerOrder: sortReportAnswers(getReportUniqueValues(filteredRows, item.value), item.value) }));
      if (!filteredRows.length) throw new Error('No rows match the report filter.');
      if (!questions.length) throw new Error('Select at least one question.');
      const sections = await calculateReportSections(filteredRows, questions, breakdownColumns, operation);
      if (operation !== reportOperation) return;
      const linkedSummary = isLinkedReportContext(source, sheetName) && state.linkedSurvey.column && breakdownColumns.includes(state.linkedSurvey.column) ? LinkedSurvey.buildCategorySummary(filteredRows, state.linkedSurvey.column) : [];
      state.reportResult = {
        title: normalizeValue(els.reportNameInput.value) || 'Question breakdown',
        aoa: sections.flatMap(section => section.aoa), rows: sections.flatMap(section => section.rows),
        skipped: [], questionCount: questions.length,
        breakdownLabel: breakdownColumns.length ? getCurrentReportColumnDisplayName(breakdownColumns[0]) : 'None',
        filterLabel: filter.label, sourceName: source.name, sheetName,
        generatedAt: new Date().toISOString(), linkedSummary,
        eligibleRows: filteredRows.length, includedRows: filteredRows.length,
        excludedRows: getSheetRecords(source.workbook, sheetName).length - filteredRows.length,
        filteredOut: dataRows.length - filteredRows.length,
        denominatorDescription: 'Answered respondents per question and group. Multiple choices may total above 100%.',
        responseSeparator: state.responseDelimiter
      };
      state.reportStale = false;
      els.reportFreshness.classList.add('hidden');
      renderDistributionOutput(state.reportResult);
      showReportStatus('Breakdown report generated. Percentages use answered respondents within each group; multiple choices may total above 100%.', '');
      saveReportConfig();
    } catch (error) {
      if (error.name === 'AbortError') return;
      showReportStatus(error.message || 'Could not generate the report.', 'error');
      if (state.reportResult) { state.reportStale = true; els.reportFreshness.textContent = 'Generation failed. The previous report is shown below and cannot be exported.'; els.reportFreshness.classList.remove('hidden'); }
    } finally {
      if (operation === reportOperation) setButtonLoading(els.generateReportBtn, false);
      updateReportExportState();
    }
  }

  function searchChecklist(container, query) {
    container.querySelectorAll('label').forEach(label => { label.hidden = !normalizeForMatch(label.textContent).includes(normalizeForMatch(query)); });
  }

  function renderAdditionalReportFilters() {
    els.additionalReportFilters.innerHTML = '';
    additionalReportFilters.forEach((filter, index) => {
      const panel = document.createElement('div'); panel.className = 'field additional-filter';
      panel.innerHTML = `<label>Additional filter ${index + 1}<select aria-label="Additional filter ${index + 1}">${els.reportFilterColumnSelect.innerHTML}</select></label><button type="button" class="text-btn">Remove filter</button><input type="search" placeholder="Search values" aria-label="Search additional filter ${index + 1} values"><div class="checklist-toolbar"><button type="button" class="text-btn select-values">Select all values</button><button type="button" class="text-btn clear-values">Clear values</button></div><div class="checklist-box"></div>`;
      const select = panel.querySelector('select'), list = panel.querySelector('.checklist-box');
      select.value = filter.column;
      if (select.value !== filter.column) filter.column = '';
      const source = getSelectedReportSource();
      const rows = source ? getReportDataRows(source, els.reportDataSheetSelect.value) : [];
      const values = filter.column ? getReportFilterValues(rows, filter.column) : [];
      renderCheckboxList(list, values.map(value => ({value, label: value})), { checkedValues: filter.values || values, onChange: () => {
        filter.values = getCheckedItems(list).map(item => item.value); markReportStale(); saveReportConfig();
      }});
      select.addEventListener('change', () => { filter.column = select.value; filter.values = null; markReportStale(); saveReportConfig(); renderAdditionalReportFilters(); });
      panel.querySelector('button').addEventListener('click', () => { additionalReportFilters.splice(index, 1); markReportStale(); saveReportConfig(); renderAdditionalReportFilters(); });
      panel.querySelector('input[type=search]').addEventListener('input', event => searchChecklist(list, event.target.value));
      ['select-values', 'clear-values'].forEach((className, buttonIndex) => panel.querySelector(`.${className}`).addEventListener('click', () => {
        list.querySelectorAll('input[type=checkbox]').forEach(input => { input.checked = buttonIndex === 0; });
        filter.values = getCheckedItems(list).map(item => item.value); markReportStale(); saveReportConfig();
      }));
      els.additionalReportFilters.appendChild(panel);
    });
    els.addReportFilterBtn.disabled = !getSelectedReportSource();
  }

  function getReportFilter(dataRows) {
    const filters = [];
    const column = els.reportFilterColumnSelect.value;
    if (column) filters.push({ column, values: getCheckedItems(els.reportFilterValues).map(item => item.value) });
    additionalReportFilters.filter(filter => filter.column).forEach(filter => filters.push({ column: filter.column, values: filter.values || getReportFilterValues(dataRows, filter.column) }));
    return {
      filters: filters.map(filter => ({ column: filter.column, values: new Set(filter.values.map(normalizeForMatch)) })),
      label: filters.length ? filters.map(filter => formatReportFilterLabel(filter.column, filter.values, getReportFilterValues(dataRows, filter.column).length, getCurrentReportColumnDisplayName(filter.column))).join(' AND ') : 'All rows'
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
    if (!filter.filters.length) return dataRows;
    return dataRows.filter(row => filter.filters.every(item => getResponseLabels(row[item.column]).some(value => item.values.has(normalizeForMatch(value)))));
  }

  function renderDistributionOutput(result) {
    els.reportOutputTitle.textContent = result.title;
    const skippedText = result.skipped.length ? `, skipped ${result.skipped.length} missing columns` : '';
    els.reportOutputMeta.textContent = `${result.questionCount} questions · ${result.sheetName} · ${formatNumber(result.eligibleRows)} included rows · ${formatNumber(result.excludedRows)} excluded (${formatNumber(result.filteredOut || 0)} filtered)`;
    renderReportContextBar(result);
    const linkedSummary = result.linkedSummary && result.linkedSummary.length ? `
      <div class="linked-category-summary">
        <h3>Linked category coverage</h3>
        <table><thead><tr><th>Category</th><th class="number">Matched sites</th><th class="number">Survey responses</th></tr></thead><tbody>${result.linkedSummary.map(item => `<tr><td>${escapeHtml(item.category)}</td><td class="number">${formatNumber(item.matchedSites)}</td><td class="number">${formatNumber(item.surveyResponses)}</td></tr>`).join('')}</tbody></table>
      </div>` : '';
    els.distributionOutput.innerHTML = `${linkedSummary}
      <table>
        <caption class="sr-only">${escapeHtml(result.title)}. ${escapeHtml(result.sheetName)}. Counts and percentages by question.</caption>
        <tbody>
          ${result.rows.map(row => {
            if (row[0]?.type === 'metadata') {
              const values = row.filter(cell => cell.type === 'count');
              return `<tr class="report-base-row"><th scope="row">${escapeHtml(row[0].value)}</th>${values.map(cell => `<td colspan="${state.reportMode === 'both' ? 2 : 1}" class="number">${formatNumber(cell.value)}</td>`).join('')}</tr>`;
            }
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
    if (result.breakdownLabel && !['No main breakdown selected', 'None'].includes(result.breakdownLabel)) {
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
    if (cell.type === 'group-header') {
      return `<th scope="colgroup" colspan="${state.reportMode === 'both' ? 2 : 1}" class="report-header">${escapeHtml(cell.value)}</th>`;
    }
    const span = cell.type === 'question' && state.reportMode !== 'both' ? 1 + (cell.colspan - 1) / 2 : cell.colspan;
    const colspan = span ? ` colspan="${span}"` : '';
    if (cell.type === 'spacer') return '<td></td>';
    if (cell.type === 'question') return `<th scope="row"${colspan} class="question-title">${escapeHtml(cell.value)}</th>`;
    if (cell.type === 'header') {
      const cls = cell.value === 'Count' ? ' report-count' : cell.value === 'Percentage' ? ' report-percent' : '';
      return `<th scope="col"${colspan} class="report-header${cls}">${escapeHtml(cell.value)}</th>`;
    }
    if (cell.type === 'answer' || cell.type === 'metadata') return `<th scope="row">${escapeHtml(cell.value)}</th>`;
    if (cell.type === 'count') return `<td${colspan} class="number report-count">${cell.value === '' ? '' : formatNumber(cell.value)}</td>`;
    if (cell.type === 'percent') {
      if (cell.value === null) return `<td class="number report-percent" title="No answered respondents in this group">—</td>`;
      return `<td${colspan} class="number heat-cell report-percent" style="background:${getHeatColor(cell.value)}">${formatPercent(cell.value)}</td>`;
    }
    return `<td${colspan}>${escapeHtml(cell.value)}</td>`;
  }

  function downloadDistributionCsv() {
    if (!state.reportResult || state.reportStale) { showToast('Generate a current report before exporting.', 'warning'); return; }
    downloadCsv(DataIO.reportRows(state.reportResult, { mode: els.reportExportMode.value, percentText: true }), `${safeFileName(state.reportResult.title)}.csv`);
  }

  function downloadDistributionXlsx() {
    if (!state.reportResult || state.reportStale) { showToast('Generate a current report before exporting.', 'warning'); return; }
    const rows = DataIO.reportRows(state.reportResult, { mode: els.reportExportMode.value });
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    DataIO.configureReportSheet(sheet, rows);
    XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName(state.reportResult.title));
    XLSX.writeFile(workbook, `${safeFileName(state.reportResult.title)}.xlsx`);
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
    const rawHeaders = [...(rows[headerRow] || [])];
    const columnCount = rows.reduce((width, row) => Math.max(width, row.length), 0);
    while (rawHeaders.length < columnCount) rawHeaders.push('');
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
    const sourceRows = sheetRowIndexCache.get(workbook)?.get(sheetName) || [];
    const records = rows.slice(headerRow + 1)
      .map((row, index) => {
        const record = Object.create(null);
        headers.forEach((header, index) => {
          record[header] = row[index] === undefined ? '' : row[index];
        });
        record.__sourceRowNumber = sourceRows[headerRow + 1 + index] || headerRow + 2 + index;
        return record;
      });
    workbookRecords.set(sheetName, records);
    return records;
  }

  function isLinkedReportContext(source, sheetName) {
    return Boolean(state.linkedSurvey.active
      && state.linkedSurvey.result
      && source
      && source.workbook === state.linkedSurvey.primaryWorkbook
      && sheetName === state.linkedSurvey.primarySheet);
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
        displayQuestion: getLinkedQuestionDisplayName(), delimiter: els.linkDelimiter.value
      });
    }
    return {
      rows: state.linkedSurvey.result.matched.map(match => ({ ...match.primary, __linkedSiteKey: match.key })),
      column: ''
    };
  }

  function getSheetMatrix(workbook, sheetName) {
    if (!sheetMatrixCache.has(workbook)) sheetMatrixCache.set(workbook, new Map());
    const cache = sheetMatrixCache.get(workbook);
    if (cache.has(sheetName)) return cache.get(sheetName);
    const sheet = workbook.Sheets[sheetName];
    if (!sheet?.['!ref']) return [];
    const start = XLSX.utils.decode_range(sheet['!ref']).s.r;
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: true });
    const matrix = [], sourceRows = [];
    raw.forEach((row, index) => { if (row.some(cell => normalizeValue(cell) !== '')) { matrix.push(row); sourceRows.push(start + index + 1); } });
    cache.set(sheetName, matrix);
    if (!sheetRowIndexCache.has(workbook)) sheetRowIndexCache.set(workbook, new Map());
    sheetRowIndexCache.get(workbook).set(sheetName, sourceRows);
    return matrix;
  }

  function findHeaderRow() { return 0; }

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
    const values = new Map();
    rows.forEach(row => getReportValues(row[column]).forEach(value => { if (!values.has(normalizeForMatch(value))) values.set(normalizeForMatch(value), value); }));
    return [...values.values()];
  }

  function getReportFilterValues(rows, column) {
    const values = new Map();
    rows.forEach(row => getResponseLabels(row[column]).forEach(value => { if (!values.has(normalizeForMatch(value))) values.set(normalizeForMatch(value), value); }));
    return [...values.values()].sort((a, b) => a.localeCompare(b));
  }

  function pickRecordColumns(row, columns) {
    const record = {};
    columns.forEach(column => {
      record[column] = row[column] === undefined ? '' : row[column];
    });
    return record;
  }

  function getReportValues(value) { return SurveyCore.labels(value, { delimiter: state.responseDelimiter }); }

  function sortReportAnswers(values, questionText) {
    const hardSet = getHardCodedAnswerSet(questionText);
    const normalizedObserved = new Set(values.map(normalizeForMatch));
    const matchingSet = hardSet || ANSWER_SETS.find(set => [...normalizedObserved].every(value => new Set(set.map(normalizeForMatch)).has(value)));
    if (!matchingSet) return [...values].sort((a, b) => a.localeCompare(b));
    const expected = new Set(matchingSet.map(normalizeForMatch));
    return [...matchingSet, ...values.filter(value => !expected.has(normalizeForMatch(value)))];
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

  function uniqueList(values) {
    return Array.from(new Set(values));
  }

  function normalizeForMatch(value) { return SurveyCore.key(value); }

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
        getReportValues(value).forEach(label => metric.uniqueValues.add(normalizeForMatch(label)));
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
    if (linked.active && linked.result && linked.primaryWorkbook === state.workbook && linked.primarySheet === state.sheetName) {
      analysisRows = linked.result.matched.map(match => ({ ...match.primary, __linkedSiteKey: match.key }));
      if (linked.question) {
        const enriched = LinkedSurvey.enrichMatchedRows(linked.result, linked.question, {
          displayQuestion: getLinkedQuestionDisplayName(), delimiter: els.linkDelimiter.value
        });
        analysisRows = enriched.rows;
        linked.column = enriched.column;
        const values = analysisRows.flatMap(row => SurveyCore.labels(row[linked.column]));
        state.columnStats.set(linked.column, {
          type: 'Linked survey',
          answeredCount: analysisRows.filter(row => SurveyCore.labels(row[linked.column]).length > 0).length,
          uniqueCount: new Set(values.map(normalizeForMatch)).size,
          missingCount: analysisRows.filter(row => SurveyCore.labels(row[linked.column]).length === 0).length
        });
      }
    }
    state.columns = linked.active && linked.primaryWorkbook === state.workbook && linked.primarySheet === state.sheetName && linked.column ? [...eligible, linked.column] : eligible;
    state.eligibleChartColumns = linked.active && linked.primaryWorkbook === state.workbook && linked.primarySheet === state.sheetName && linked.column
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
    if (!state.workbook) return;
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
    renderFileStats();
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
      markReportStale(); renderReportColumns();
      showToast(input.checked ? 'Column restored to chart analysis.' : 'Column hidden from chart analysis.');
    }));

    els.previewResultCount.textContent = `Showing ${formatNumber(visibleRows.length)} of ${formatNumber(matchingRows.length)} matching rows`;
    els.dataPreviewTable.innerHTML = `<table><thead><tr><th>#</th>${state.allColumns.map(column => { const displayName = getActiveColumnDisplayName(column); return `<th title="${escapeAttr(displayName)}">${escapeHtml(truncateLabel(displayName, 42))}</th>`; }).join('')}</tr></thead><tbody>${visibleRows.map((row, index) => `<tr><td class="number">${row.__sourceRowNumber || index + 2}</td>${state.allColumns.map(column => {
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

  function setVisibleReportQuestions(checked, visibleOnly = false) {
    els.questionChecklist.querySelectorAll(visibleOnly ? 'label:not(.is-filtered-out) input[type="checkbox"]' : 'input[type="checkbox"]').forEach(input => {
      input.checked = checked;
    });
    updateReportSelectionCount(); markReportStale(); saveReportConfig();
  }

  function updateReportSelectionCount() {
    if (!els.selectedQuestionCount) return;
    const count = els.questionChecklist.querySelectorAll('input[type="checkbox"]:checked').length;
    els.selectedQuestionCount.textContent = `${count} question${count === 1 ? '' : 's'} selected`;
  }

  function setReportMode(mode) {
    state.reportMode = mode;
    if (state.reportResult) renderDistributionOutput(state.reportResult);
    ['count', 'percentage', 'both'].forEach(value => els.distributionOutput.classList.toggle(`report-mode-${value}`, value === mode));
    document.querySelectorAll('[data-report-mode]').forEach(button => { button.classList.toggle('is-active', button.dataset.reportMode === mode); button.setAttribute('aria-pressed', String(button.dataset.reportMode === mode)); });
  }

  function setReportDensity(density) {
    ['compact', 'comfortable'].forEach(value => els.distributionOutput.classList.toggle(`density-${value}`, value === density));
    document.querySelectorAll('[data-density]').forEach(button => { button.classList.toggle('is-active', button.dataset.density === density); button.setAttribute('aria-pressed', String(button.dataset.density === density)); });
  }

  function setReportZoom(value) {
    state.reportZoom = Math.max(0.7, Math.min(1.4, Math.round(value * 10) / 10));
    els.distributionOutput.style.setProperty('--report-zoom', state.reportZoom);
    els.zoomValue.textContent = `${Math.round(state.reportZoom * 100)}%`;
  }

  function expandPanel(panel, expanded) {
    if (expanded) {
      const previous = document.querySelector('.expanded-dialog');
      if (previous) previous.close();
      const dialog = document.createElement('dialog');
      dialog.className = 'expanded-dialog'; dialog.setAttribute('aria-label', 'Expanded analysis view');
      const close = document.createElement('button'); close.className = 'secondary-btn expanded-close'; close.textContent = 'Close expanded view';
      const placeholder = document.createComment('Expanded analysis placeholder');
      const returnFocus = document.activeElement;
      panel.before(placeholder); dialog.append(close, panel); document.body.append(dialog);
      panel.classList.add('is-expanded');
      dialog.addEventListener('close', () => { panel.classList.remove('is-expanded'); placeholder.replaceWith(panel); dialog.remove(); els.fullscreenReportBtn.setAttribute('aria-pressed', 'false'); returnFocus?.focus(); }, { once: true });
      close.addEventListener('click', () => dialog.close());
      dialog.showModal(); close.focus();
    } else panel.closest('dialog')?.close();
  }

  function toggleReportFullscreen() {
    const panel = els.distributionOutput.closest('.report-output-panel');
    const expanded = !panel.classList.contains('is-expanded');
    expandPanel(panel, expanded); els.fullscreenReportBtn.setAttribute('aria-pressed', String(expanded));
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
    expandPanel(card, !card.classList.contains('is-expanded'));
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

  function resetDataset(cancelLoads = true) {
    document.querySelector('.expanded-dialog')?.close();
    if (cancelLoads) { ++loadOperation; activeLoadController?.abort(); }
    ++secondaryLoadOperation; ++reportOperation; cancelReportWork?.();
    reportConfigs.clear(); additionalReportFilters = []; reportContextKey = '';
    state.hiddenAnalysisColumns = new Set();
    state.columnStats = new Map();
    state.reportStale = false;
    state.activeTab = 'charts';
    els.reportFreshness.classList.add('hidden');
    els.previewSearch.value = ''; state.previewSearch = '';
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
    state.sources = [];
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
    showReportStatus('', ''); showLinkValidation('', '');
    els.reportFreshness.classList.add('hidden');
    setButtonLoading(els.generateReportBtn, false);
    renderDataset(); updateReportExportState();
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
