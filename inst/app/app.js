const STORAGE_KEY = 'rgrid.workbook.v1';
const PLOT_PANE_WIDTH_KEY = 'rgrid.plotPaneWidth.v1';
const PLOT_RESOLUTION_KEY = 'rgrid.plotResolution.v1';
const THEME_KEY = 'rgrid.theme.v1';
const FORMULA_HEIGHT_KEY = 'rgrid.formulaHeight.v1';
const FORMULA_COLLAPSED_HEIGHT = 44;
const FORMULA_FALLBACK_EXPANDED_HEIGHT = 116;
const FORMAT_VERSION = 6;
const LOCAL_R_API_BASE = './__rgrid_api';
const XLSX_MODULE_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs';
const DEFAULT_ROWS = 100;
const DEFAULT_COLS = 26;
const MAX_ROWS = 10000;
const MAX_COLS = 1000;
const MAX_RENDERED_CELLS = 500000;
const DEFAULT_PLOT_RESOLUTION = Object.freeze({ width: 1008, height: 672 });
const CORE_R_PACKAGES = new Set(['base', 'compiler', 'datasets', 'graphics', 'grDevices', 'grid', 'methods', 'parallel', 'splines', 'stats', 'stats4', 'tcltk', 'tools', 'utils']);

const $ = (selector) => document.querySelector(selector);
const els = {
  app: $('#app'),
  grid: $('#grid'),
  gridViewport: $('#gridViewport'),
  cellTooltip: $('#cellTooltip'),
  sheetTabs: $('#sheetTabs'),
  nameBox: $('#nameBox'),
  formulaEditor: $('#formulaEditor'),
  formulaHighlight: $('#formulaHighlight'),
  formulaInput: $('#formulaInput'),
  commitFormulaBtn: $('#commitFormulaBtn'),
  runtimeStatus: $('#runtimeStatus'),
  runtimeStatusText: $('#runtimeStatusText'),
  selectionSummary: $('#selectionSummary'),
  saveStatus: $('#saveStatus'),
  namesDialog: $('#namesDialog'),
  namesList: $('#namesList'),
  newNameInput: $('#newNameInput'),
  newNameRefInput: $('#newNameRefInput'),
  messageDialog: $('#messageDialog'),
  messageTitle: $('#messageTitle'),
  messageBody: $('#messageBody'),
  importScriptInput: $('#importScriptInput'),
  importDataInput: $('#importDataInput'),
  importDataBtn: $('#importDataBtn'),
  formulaMessage: $('#formulaMessage'),
  undoBtn: $('#undoBtn'),
  redoBtn: $('#redoBtn'),
  plotsBtn: $('#plotsBtn'),
  plotPane: $('#plotPane'),
  plotPaneResizer: $('#plotPaneResizer'),
  plotPaneCount: $('#plotPaneCount'),
  plotList: $('#plotList'),
  closePlotPaneBtn: $('#closePlotPaneBtn'),
  plotSettingsBtn: $('#plotSettingsBtn'),
  plotSettingsDialog: $('#plotSettingsDialog'),
  plotResolutionPreset: $('#plotResolutionPreset'),
  plotWidthInput: $('#plotWidthInput'),
  plotHeightInput: $('#plotHeightInput'),
  plotResolutionHelp: $('#plotResolutionHelp'),
  applyPlotResolutionBtn: $('#applyPlotResolutionBtn'),
  formulaEditorResizer: $('#formulaEditorResizer'),
  busyOverlay: $('#busyOverlay'),
  busyText: $('#busyText'),
  themeBtn: $('#themeBtn'),
  exampleWorkbookBtn: $('#exampleWorkbookBtn'),
  helpBtn: $('#helpBtn'),
  helpDialog: $('#helpDialog'),
  refHelpDialog: $('#refHelpDialog'),
  objectDialog: $('#objectDialog'),
  objectDialogTitle: $('#objectDialogTitle'),
  objectDialogSubtitle: $('#objectDialogSubtitle'),
  objectTree: $('#objectTree'),
  expandObjectTreeBtn: $('#expandObjectTreeBtn'),
  toggleElementNamesBtn: $('#toggleElementNamesBtn'),
  saveNameBtn: $('#saveNameBtn'),
  referenceStyleBtn: $('#referenceStyleBtn'),
  fxHelpLink: $('#fxHelpLink'),
};

const runtime = {
  r: null,
  ready: false,
  calculating: false,
  recalcRequested: false,
  recalcTimer: null,
  forceRecalcRequested: false,
  packageStatus: new Map(),
};

async function localRRequest(endpoint, payload = null) {
  const response = await fetch(`${LOCAL_R_API_BASE}/${endpoint}`, {
    method: payload === null ? 'GET' : 'POST',
    headers: payload === null ? {} : { 'Content-Type': 'application/json' },
    body: payload === null ? undefined : JSON.stringify(payload),
  });
  const text = await response.text();
  let data = null;
  try { data = text ? JSON.parse(text) : null; }
  catch { data = { ok: false, error: text || response.statusText }; }
  if (!response.ok || data?.ok === false) {
    const message = data?.error || data?.message || response.statusText || 'Local R request failed';
    throw new Error(message);
  }
  return data;
}

class LocalRResult {
  constructor(value) { this.value = value; }
  async toJs() { return this.value; }
}

class LocalRClient {
  constructor(base = LOCAL_R_API_BASE) {
    this.base = base;
    const client = this;
    this.Shelter = class {
      async captureR(code, options = {}) { return client.captureR(code, options); }
      async purge() {}
    };
  }
  async init() {
    const status = await localRRequest('status');
    this.version = status.r_version || '';
    this.platform = status.platform || '';
  }
  async evalRVoid(code) {
    await localRRequest('eval-void', { code });
  }
  async evalRBoolean(code) {
    const result = await localRRequest('eval-boolean', { code });
    return Boolean(result.value);
  }
  async evalRString(code) {
    const result = await localRRequest('eval-string', { code });
    return String(result.value ?? '');
  }
  async installPackages(packages, options = {}) {
    await localRRequest('install-packages', { packages: asArray(packages), quiet: Boolean(options.quiet) });
  }
  async captureR(code, options = {}) {
    const result = await localRRequest('capture', { code, options });
    return {
      result: new LocalRResult(result.result),
      images: result.images || [],
      output: result.output || [],
    };
  }
}

let workbook = loadWorkbook() ?? createBlankWorkbook();
let computed = new Map();
let spillOwners = new Map();
let displayValues = new Map();
let plotsByCell = new Map();
let calculationState = createCalculationState();
let cellElements = new Map();
let previousSelectedAddresses = new Set();
let previousActiveHeaderElements = [];
let tooltipCell = null;
let saveTimer = null;
let selection = { r1: 1, c1: 1, r2: 1, c2: 1 };
const history = { undo: [], redo: [], limit: 100, restoring: false };
const formulaEdit = { active: false, sheetId: null, row: null, col: null };
let formulaReferenceAddresses = new Set();
let plotPaneOpen = true;
let plotPaneWidth = (() => {
  try { return Number(localStorage.getItem(PLOT_PANE_WIDTH_KEY)) || 390; }
  catch { return 390; }
})();
let plotResolution = loadPlotResolution();
let resizedSheetIds = new Set();
let editingWorkbookName = null;
let formulaEditorHeight = (() => {
  try { return Number(localStorage.getItem(FORMULA_HEIGHT_KEY)) || FORMULA_COLLAPSED_HEIGHT; }
  catch { return FORMULA_COLLAPSED_HEIGHT; }
})();
let formulaExpandedHeight = formulaEditorHeight > FORMULA_COLLAPSED_HEIGHT
  ? formulaEditorHeight
  : FORMULA_FALLBACK_EXPANDED_HEIGHT;
let formulaCodeMirror = null;
const dragSelection = {
  active: false,
  mode: null,
  sheetId: null,
  startRow: null,
  startCol: null,
  formulaStart: null,
  formulaEnd: null,
};
const formulaKeyboardPick = {
  active: false,
  sheetId: null,
  start: null,
  end: null,
};

function initialiseFormulaCodeEditor() {
  if (!window.CodeMirror) return;
  formulaCodeMirror = window.CodeMirror.fromTextArea(els.formulaInput, {
    mode: 'r',
    lineNumbers: false,
    lineWrapping: false,
    matchBrackets: true,
    indentUnit: 2,
    tabSize: 2,
    indentWithTabs: false,
    electricChars: true,
    showCursorWhenSelecting: true,
    placeholder: els.formulaInput.placeholder,
  });
  formulaCodeMirror.getWrapperElement().classList.add('formula-codemirror');
  formulaCodeMirror.setSize('100%', '100%');
}

function formulaValue() {
  return formulaCodeMirror ? formulaCodeMirror.getValue() : els.formulaInput.value;
}

function setFormulaValue(value) {
  const text = String(value ?? '');
  if (formulaCodeMirror) {
    if (formulaCodeMirror.getValue() !== text) formulaCodeMirror.setValue(text);
  } else {
    els.formulaInput.value = text;
  }
}

function formulaSelectionOffsets() {
  if (!formulaCodeMirror) {
    const start = els.formulaInput.selectionStart ?? 0;
    const end = els.formulaInput.selectionEnd ?? start;
    return { start, end };
  }
  return {
    start: formulaCodeMirror.indexFromPos(formulaCodeMirror.getCursor('from')),
    end: formulaCodeMirror.indexFromPos(formulaCodeMirror.getCursor('to')),
  };
}

function setFormulaSelection(start, end = start) {
  const safeStart = Math.max(0, Number(start) || 0);
  const safeEnd = Math.max(safeStart, Number(end) || safeStart);
  if (formulaCodeMirror) {
    formulaCodeMirror.setSelection(formulaCodeMirror.posFromIndex(safeStart), formulaCodeMirror.posFromIndex(safeEnd));
    formulaCodeMirror.scrollIntoView(formulaCodeMirror.getCursor(), 24);
  } else {
    els.formulaInput.setSelectionRange(safeStart, safeEnd);
  }
}

function replaceFormulaRange(text, start, end = start, selectStart = null, selectEnd = null) {
  const insertion = String(text ?? '');
  if (formulaCodeMirror) {
    formulaCodeMirror.replaceRange(
      insertion,
      formulaCodeMirror.posFromIndex(start),
      formulaCodeMirror.posFromIndex(end),
      '+input',
    );
  } else {
    els.formulaInput.setRangeText(insertion, start, end, 'end');
  }
  const caret = start + insertion.length;
  setFormulaSelection(selectStart ?? caret, selectEnd ?? selectStart ?? caret);
}

function setFormulaPlaceholder(text) {
  els.formulaInput.placeholder = text;
  formulaCodeMirror?.setOption('placeholder', text);
}

function focusFormulaEditor() {
  if (formulaCodeMirror) formulaCodeMirror.focus();
  else els.formulaInput.focus({ preventScroll: true });
}

function selectAllFormula() {
  if (formulaCodeMirror) formulaCodeMirror.execCommand('selectAll');
  else els.formulaInput.select();
}

function formulaEditorContainsTarget(target) {
  if (target === els.formulaInput) return true;
  return Boolean(formulaCodeMirror?.getWrapperElement().contains(target));
}

function refreshFormulaCodeEditor() {
  if (!formulaCodeMirror) return;
  requestAnimationFrame(() => formulaCodeMirror.refresh());
}

function functionTokenAtFormulaCursor() {
  const text = formulaValue();
  const cursor = formulaSelectionOffsets().end;
  const identifier = /[A-Za-z.][A-Za-z0-9._]*/g;
  let match;
  while ((match = identifier.exec(text))) {
    const start = match.index;
    const end = start + match[0].length;
    if (cursor < start || cursor > end) continue;
    let packageName = null;
    let packageStart = start;
    const prefix = text.slice(0, start);
    const namespace = /([A-Za-z][A-Za-z0-9.]*)\s*:::{0,1}\s*$/.exec(prefix);
    if (namespace) {
      packageName = namespace[1];
      packageStart = prefix.length - namespace[0].length;
    }
    return { name: match[0], packageName, start: packageStart, end };
  }
  return null;
}

function libraryPackagesInFormula(text = formulaValue()) {
  const packages = [];
  const regex = /\b(?:library|require)\s*\(\s*(?:package\s*=\s*)?(?:["']([A-Za-z][A-Za-z0-9.]*)["']|([A-Za-z][A-Za-z0-9.]*))/g;
  let match;
  while ((match = regex.exec(text))) packages.push(match[1] || match[2]);
  return packages;
}

async function resolveFunctionHelpTarget(token) {
  const formulaPackages = libraryPackagesInFormula();
  const fallbackPackage = token.packageName || formulaPackages.at(-1) || '';
  const rFormulaPackages = `c(${formulaPackages.map(rString).join(', ')})`;
  if (!runtime.ready) return { packageName: fallbackPackage, topic: token.name };
  const result = await runtime.r.evalRString(`local({
    .name <- ${rString(token.name)}
    .explicit <- ${rString(token.packageName || '')}
    .core <- c("base", "compiler", "datasets", "graphics", "grDevices", "grid", "methods", "parallel", "splines", "stats", "stats4", "tcltk", "tools", "utils")
    .package <- .explicit
    if (!nzchar(.package)) {
      .search_hits <- sub("^package:", "", grep("^package:", find(.name, mode = "function"), value = TRUE))
      .loaded <- loadedNamespaces()
      .namespace_hits <- .loaded[vapply(.loaded, function(.pkg) {
        tryCatch(exists(.name, envir = asNamespace(.pkg), mode = "function", inherits = FALSE), error = function(e) FALSE)
      }, logical(1L))]
      .candidates <- unique(c(.search_hits, .namespace_hits, ${rFormulaPackages}))
      if (length(.candidates)) .package <- .candidates[[1L]]
    }
    .topic <- .name
    if (nzchar(.package)) {
      .help <- suppressWarnings(try(utils::help(.name, package = .package), silent = TRUE))
      if (!inherits(.help, "try-error") && length(.help)) {
        .path <- as.character(.help)[[1L]]
        .candidate <- basename(.path)
        .candidate <- sub("\\\\.[^.]+$", "", .candidate)
        if (nzchar(.candidate)) .topic <- .candidate
      }
    }
    paste(.package, .topic, sep = "\\t")
  })`);
  const [packageName = '', topic = token.name] = String(result).split('\t');
  return { packageName: packageName || fallbackPackage, topic: topic || token.name };
}

function functionHelpUrl(packageName, topic, functionName) {
  if (!packageName) return 'https://rdrr.io/r/';
  if (CORE_R_PACKAGES.has(packageName)) {
    return `https://stat.ethz.ch/R-manual/R-devel/library/${encodeURIComponent(packageName)}/html/${encodeURIComponent(topic || functionName)}.html`;
  }
  return `https://cran.r-project.org/web/packages/${encodeURIComponent(packageName)}/refman/${encodeURIComponent(packageName)}.html#${encodeURIComponent(functionName)}`;
}

function openRefFunctionHelp() {
  if (!els.refHelpDialog.open) els.refHelpDialog.showModal();
  setFormulaMessage('Showing RGrid help for ref().', 'info');
}

async function openFunctionHelpAtCursor() {
  const token = functionTokenAtFormulaCursor();
  if (!token) {
    setFormulaMessage('Place the formula cursor on an R function name, then press F1.', 'info');
    return;
  }
  if (!token.packageName && token.name.toLowerCase() === 'ref') {
    openRefFunctionHelp();
    return;
  }
  const popup = window.open('about:blank', '_blank');
  if (popup) {
    popup.opener = null;
    popup.document.title = `R help: ${token.name}`;
    popup.document.body.textContent = `Looking up documentation for ${token.name}…`;
  }
  try {
    const target = await resolveFunctionHelpTarget(token);
    const url = functionHelpUrl(target.packageName, target.topic, token.name);
    if (popup) popup.location.replace(url);
    else window.open(url, '_blank', 'noopener');
    setFormulaMessage(`Opened help for ${target.packageName ? `${target.packageName}::` : ''}${token.name}`, 'info');
  } catch (error) {
    popup?.close();
    setFormulaMessage(`Could not resolve help for ${token.name}: ${error.message || error}`, 'error');
  }
}

function uid(prefix = 'id') {
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;
}

function createBlankWorkbook() {
  const sheetId = uid('sheet');
  return {
    formatVersion: FORMAT_VERSION,
    title: 'RGrid Workbook',
    activeSheetId: sheetId,
    seq: 0,
    names: {},
    view: { showElementNames: true, referenceStyle: 'A1' },
    sheets: [{
      id: sheetId,
      name: 'Sheet1',
      rows: DEFAULT_ROWS,
      cols: DEFAULT_COLS,
      cells: {},
    }],
  };
}

function createExampleWorkbook() {
  const sheetIds = {
    start: uid('sheet'),
    references: uid('sheet'),
    data: uid('sheet'),
    objects: uid('sheet'),
    plots: uid('sheet'),
  };
  let seq = 0;
  const makeCells = (entries) => Object.fromEntries(entries.map(([address, input, literal = false]) => [
    address,
    { input, seq: ++seq, ...((literal || isPlainTextInput(input)) ? { literal: true } : {}) },
  ]));

  const sheets = [
    {
      id: sheetIds.start,
      name: 'Start Here',
      rows: DEFAULT_ROWS,
      cols: DEFAULT_COLS,
      cells: makeCells([
        ['A1', 'RGrid example workbook'],
        ['A2', 'Five focused worksheets demonstrate formulas, references, imported data, R objects, pivots, and plots.'],
        ['A3', 'Select any formula cell to read its commented R code in the formula bar. F1 opens help for the function under the cursor.'],
        ['A5', 'Plain value'],
        ['B5', '42'],
        ['A6', 'R formula using ref()'],
        ['B6', '=sqrt(ref("B5"))'],
        ['A8', 'Dynamic spill'],
        ['B8', `={
  # A vector returned by R spills down from this anchor cell.
  seq(10, 50, by = 10)
}`],
        ['D5', 'Sum of the spill'],
        ['E5', '=sum(ref("B8#"))'],
        ['D6', 'Named range mean'],
        ['E6', '=mean(ref("example_numbers"))'],
        ['D7', 'Named expression'],
        ['E7', '=ref("vat_rate")'],
        ['D8', 'Named function'],
        ['E8', '=ref("double_it")(ref("B5"))'],
        ['A15', '**Workbook tips**'],
        ['A16', '• Use the Name manager for named ranges, constants, and functions.'],
        ['A17', '• A trailing # in ref("B8#") means the complete dynamic spill.'],
        ['A18', '• Double-click a list cell to inspect its tree; click a plot cell to locate its plot.'],
        ['A19', '• Imported formula-looking text is kept literal. The Objects & Import sheet shows this case.'],
        ['A21', 'Some examples install CRAN packages or fetch a very small public dataset the first time they calculate.'],
        ['A23', 'Text formatting: **bold**, __also bold__, *italic*, and _also italic_.'],
      ]),
    },
    {
      id: sheetIds.references,
      name: 'References',
      rows: DEFAULT_ROWS,
      cols: DEFAULT_COLS,
      cells: makeCells([
        ['A1', 'ref(): addresses, ranges, spills, names, and preserved objects'],
        ['A2', 'References are case-insensitive and may use A1 or R1C1 notation. Sheet names containing spaces are quoted.'],
        ['A4', 'Month'], ['B4', 'Sales'],
        ['A5', 'Jan'], ['B5', '120'],
        ['A6', 'Feb'], ['B6', '150'],
        ['A7', 'Mar'], ['B7', '=NA_real_'],
        ['A8', 'Apr'], ['B8', '180'],
        ['D4', 'Single A1 cell'], ['E4', '=ref("B5")'],
        ['D5', 'Case-insensitive address'], ['E5', '=ref("b6")'],
        ['D6', 'R1C1 address'], ['E6', '=ref("R8C2")'],
        ['D7', 'Rectangular range'], ['E7', '=ref("A4:B8")'],
        ['H4', 'Cross-sheet reference'], ['I4', '=ref("\'Start Here\'!B5")'],
        ['H5', 'Cross-sheet dynamic spill'], ['I5', '=ref("\'Start Here\'!B8#")'],
        ['H11', 'Named range'], ['I11', '=ref("example_numbers")'],
        ['J11', 'Named expression'], ['K11', '=ref("vat_rate")'],
        ['J12', 'Named function'], ['K12', '=ref("double_it")(5)'],
        ['A12', 'Function stored in a cell'],
        ['B12', '=function(x) x + 3'],
        ['A13', 'Call the stored function'],
        ['B13', '=ref("B12")(10)'],
        ['A16', 'Preserved data-frame object'],
        ['B16', `={
  # Data frames spill into cells, while ref() can still recover the original object.
  data.frame(item = c("A", "B", "C"), amount = c(4, 7, 2), active = c(TRUE, FALSE, TRUE))
}`],
        ['F16', 'Subset of preserved object'],
        ['G16', '=ref("B16#")[ref("B16#")$active, c("item", "amount"), drop = FALSE]'],
        ['A22', 'Reference to a blank cell'], ['B22', '=ref("C22")'],
        ['A23', 'Literal formula-looking text'], ['B23', '=mean(1:3)', true],
        ['A25', 'Corner cases to try safely'],
        ['A26', 'Put a value inside a spill destination to see #SPILL!, then remove it.'],
        ['A27', 'Circular references are reported as #CYCLE!; they are described here rather than activated in the example workbook.'],
        ['A28', 'Invalid addresses, missing names, and referenced errors are reported as #REF! with a detailed message.'],
      ]),
    },
    {
      id: sheetIds.data,
      name: 'Data & Pivot',
      rows: DEFAULT_ROWS,
      cols: DEFAULT_COLS,
      cells: makeCells([
        ['A1', 'Data tables, a dcast pivot table, and a small Eurostat import'],
        ['A2', 'The formulas below keep their R object classes while displaying rectangular results in the grid.'],
        ['A4', 'Long sales data (data.table)'],
        ['A5', `={
  # This long table is the input for the pivot formulas to the right.
  data.table::data.table(
    region = rep(c("North", "South"), each = 6),
    product = rep(rep(c("Hardware", "Services"), each = 3), 2),
    quarter = rep(paste0("Q", 1:3), 4),
    sales = c(24, 29, 31, 18, 22, 27, 20, 25, 28, 16, 21, 24)
  )
}`],
        ['F4', 'Pivot table: region + product by quarter'],
        ['F5', `={
  # data.table::dcast() is RGrid's pivot-table pattern.
  long <- data.table::as.data.table(ref("long_sales"))
  data.table::dcast(
    long,
    region + product ~ quarter,
    value.var = "sales",
    fun.aggregate = sum,
    fill = 0
  )
}`],
        ['F12', 'Pivot with two aggregate functions'],
        ['F13', `={
  # Multiple functions create clearly named value columns.
  long <- data.table::as.data.table(ref("A5#"))
  data.table::dcast(
    long,
    region ~ quarter,
    value.var = "sales",
    fun.aggregate = list(sum, mean),
    fill = 0
  )
}`],
        ['A21', 'Small Eurostat data request'],
        ['A22', 'This formula requests only two countries, two years, and one demographic slice. It needs network access.'],
        ['A23', `={
  # eurodata::importData() accepts a dataset code and a named filter list.
  # TIME_PERIOD is deliberately narrow so the returned data stays small.
  eurodata::importData(
    "demo_pjan",
    filters = list(
      freq = "A",
      age = "TOTAL",
      sex = "T",
      unit = "NR",
      geo = c("LU", "DE"),
      TIME_PERIOD = 2022:2023
    )
  )
}`],
        ['K21', 'Use imported data in another formula'],
        ['K22', 'After the Eurostat result is available, this selects a compact set of columns when they exist.'],
        ['K23', `={
  # ref() recovers the imported object, not merely its formatted cell text.
  x <- ref("A23#")
  wanted <- intersect(c("geo", "TIME_PERIOD", "values", "status"), names(x))
  x[, wanted, drop = FALSE]
}`],
      ]),
    },
    {
      id: sheetIds.objects,
      name: 'Objects & Import',
      rows: DEFAULT_ROWS,
      cols: DEFAULT_COLS,
      cells: makeCells([
        ['A1', 'R objects, list inspection, and file-import behavior'],
        ['A2', 'Lists remain single-cell objects. Double-click the list cell to open RGrid\'s expandable object viewer.'],
        ['A4', 'Nested R list'],
        ['B4', `=list(
  title = "Quarterly assumptions",
  scalars = list(vat = 0.21, active = TRUE),
  scenarios = setNames(c(10, 20, 30), c("low", "base", "high")),
  table = data.frame(code = c("A", "B"), value = c(1.5, 2.75)),
  missing = NA_real_
)`],
        ['A5', 'Extract a named nested value'], ['B5', '=ref("B4")$scalars$vat'],
        ['A6', 'Extract by position and name'], ['B6', '=ref("B4")[[3]][["base"]]'],
        ['A8', 'Named vector'], ['B8', '=setNames(c(12, 15, 9), c("Alpha", "Beta", "Gamma"))'],
        ['D8', 'Matrix with dimnames'],
        ['E8', `={
  m <- matrix(1:6, nrow = 2, byrow = TRUE)
  dimnames(m) <- list(c("first", "second"), c("x", "y", "z"))
  m
}`],
        ['A13', 'Function object capturing ref()'],
        ['B13', '=function(x, rate = ref("vat_rate")) x * (1 + rate)'],
        ['A14', 'Call the function object'], ['B14', '=ref("B13")(100)'],
        ['A17', 'rio::import() from a tiny temporary CSV'],
        ['A18', 'This self-contained example writes three rows to a temporary file and imports them by extension.'],
        ['B19', `={
  # rio::import() chooses an importer from the file extension.
  path <- tempfile(fileext = ".csv")
  writeLines(c("city,value", "Luxembourg,12", "Esch-sur-Alzette,7", "Differdange,5"), path)
  rio::import(path)
}`],
        ['A25', 'RGrid file import'],
        ['A26', 'Use Import data for CSV, TSV, TXT, XLS, or XLSX files. Each imported table becomes a new worksheet.'],
        ['A27', 'Leading = text imported from a data file remains text rather than silently becoming a formula.'],
        ['B28', '=1 + 2', true],
        ['A30', 'List structure as ordinary text'],
        ['B30', '=capture.output(str(ref("B4")))'],
        ['E25', 'Empty-object corner cases'],
        ['E26', 'Zero-length vector'], ['F26', '=character()'],
        ['E27', 'NULL'], ['F27', '=NULL'],
        ['E28', 'Missing scalar'], ['F28', '=NA_real_'],
      ]),
    },
    {
      id: sheetIds.plots,
      name: 'Plots',
      rows: DEFAULT_ROWS,
      cols: DEFAULT_COLS,
      cells: makeCells([
        ['A1', 'Base graphics, multiple plots, ggplot2, and lattice'],
        ['A2', 'Plot-producing cells get a corner marker. Select a cell to highlight its plots in the plot pane.'],
        ['A4', 'Base R line plot'],
        ['B4', `={
  # Base plots are captured from the graphics device.
  x <- 1:8
  plot(x, x^1.5, type = "b", pch = 19, xlab = "Period", ylab = "Index", main = "Base R plot")
  # Base plotting functions return an invisible helper object, so end with NULL.
  invisible(NULL)
}`],
        ['A8', 'Two base plots from one formula'],
        ['B8', `={
  # Every completed page is attached to this cell.
  old <- par(mfrow = c(1, 2))
  hist(mtcars$mpg, main = "Miles per gallon", xlab = "mpg")
  boxplot(mpg ~ factor(cyl), data = mtcars, xlab = "cylinders", ylab = "mpg")
  par(old)
  invisible(NULL)
}`],
        ['A12', 'Returned ggplot2 object'],
        ['B12', `={
  # Returning a ggplot object lets RGrid print and preserve it automatically.
  d <- data.frame(month = factor(month.abb[1:6], levels = month.abb[1:6]), sales = c(12, 15, 13, 19, 22, 25))
  ggplot2::ggplot(d, ggplot2::aes(month, sales, group = 1)) +
    ggplot2::geom_line(linewidth = 0.8) +
    ggplot2::geom_point(size = 2) +
    ggplot2::labs(title = "ggplot2 object", x = NULL, y = "Sales") +
    ggplot2::theme_minimal()
}`],
        ['A17', 'Plot values referenced from another worksheet'],
        ['B17', `={
  # The numeric series comes directly from the References worksheet.
  values <- as.numeric(ref("\'References\'!B5:B8"))
  values[is.na(values)] <- 0
  barplot(values, names.arg = c("Jan", "Feb", "Mar", "Apr"), ylab = "Sales", main = "Cross-sheet data")
  invisible(NULL)
}`],
        ['A22', 'Returned lattice object'],
        ['B22', `={
  # Trellis objects are printed and captured in the same way as ggplot objects.
  lattice::xyplot(mpg ~ wt | factor(cyl), data = mtcars, type = c("p", "r"), xlab = "Weight", ylab = "MPG")
}`],
        ['A27', 'Plot settings'],
        ['A28', 'Use Set plot size to change the device resolution, then recalculate the workbook.'],
      ]),
    },
  ];

  return {
    formatVersion: FORMAT_VERSION,
    title: 'RGrid Example Workbook',
    activeSheetId: sheetIds.start,
    seq,
    names: {
      example_numbers: { kind: 'range', sheetId: sheetIds.start, ref: 'B8:B12' },
      vat_rate: { kind: 'expression', expression: '0.21' },
      double_it: { kind: 'expression', expression: 'function(x) x * 2' },
      long_sales: { kind: 'range', sheetId: sheetIds.data, ref: 'A5#' },
    },
    view: { showElementNames: true, referenceStyle: 'A1' },
    sheets,
  };
}

function loadWorkbook() {
  try {
    const text = localStorage.getItem(STORAGE_KEY);
    if (!text) return null;
    return normalizeWorkbook(JSON.parse(text));
  } catch (error) {
    console.warn('Could not restore autosave:', error);
    return null;
  }
}

function normalizeWorkbook(value) {
  if (!value || !Array.isArray(value.sheets) || value.sheets.length === 0) {
    throw new Error('The workbook has no worksheets.');
  }
  const normalized = {
    formatVersion: FORMAT_VERSION,
    title: String(value.title || 'RGrid Workbook'),
    activeSheetId: String(value.activeSheetId || value.sheets[0].id),
    seq: Number.isFinite(value.seq) ? value.seq : 0,
    names: value.names && typeof value.names === 'object' ? value.names : {},
    view: {
      showElementNames: value.view?.showElementNames === undefined ? true : Boolean(value.view.showElementNames),
      referenceStyle: String(value.view?.referenceStyle || 'A1').toUpperCase() === 'R1C1' ? 'R1C1' : 'A1',
    },
    sheets: value.sheets.map((sheet, index) => ({
      id: String(sheet.id || uid('sheet')),
      name: String(sheet.name || `Sheet${index + 1}`),
      rows: clampInt(sheet.rows, 1, MAX_ROWS, DEFAULT_ROWS),
      cols: clampInt(sheet.cols, 1, MAX_COLS, DEFAULT_COLS),
      cells: normalizeCells(sheet.cells),
    })),
  };
  if (!normalized.sheets.some((sheet) => sheet.id === normalized.activeSheetId)) {
    normalized.activeSheetId = normalized.sheets[0].id;
  }
  const validSheetIds = new Set(normalized.sheets.map((sheet) => sheet.id));
  normalized.names = Object.fromEntries(Object.entries(normalized.names)
    .map(([name, def]) => [name, normalizeNameDefinition(def, validSheetIds)])
    .filter(([, def]) => Boolean(def)));
  normalized.seq = Math.max(normalized.seq, ...normalized.sheets.flatMap((sheet) =>
    Object.values(sheet.cells).map((cell) => Number(cell.seq) || 0)));
  return normalized;
}


function normalizeNameDefinition(definition, validSheetIds = new Set(workbook?.sheets?.map((sheet) => sheet.id) || [])) {
  if (!definition || typeof definition !== 'object') return null;
  if (definition.kind === 'expression' || typeof definition.expression === 'string') {
    const expression = String(definition.expression || '').trim();
    return expression ? { kind: 'expression', expression } : null;
  }
  if (validSheetIds.has(String(definition.sheetId)) && typeof definition.ref === 'string') {
    return { kind: 'range', sheetId: String(definition.sheetId), ref: String(definition.ref) };
  }
  return null;
}

function normalizeCells(cells) {
  if (!cells || typeof cells !== 'object') return {};
  const out = {};
  for (const [address, cell] of Object.entries(cells)) {
    if (!cell || typeof cell.input !== 'string' || cell.input === '') continue;
    try {
      const parsed = parseA1Address(address);
      out[toA1(parsed.row, parsed.col)] = {
        input: cell.input,
        seq: Number.isFinite(cell.seq) ? cell.seq : 0,
        literal: Boolean(cell.literal),
      };
    } catch {
      // Ignore invalid stored cell addresses.
    }
  }
  return out;
}

function clampInt(value, min, max, fallback) {
  const number = Number.parseInt(value, 10);
  return Number.isFinite(number) ? Math.min(max, Math.max(min, number)) : fallback;
}

function captureHistoryState(label = '') {
  return {
    workbookJson: JSON.stringify(workbook),
    selection: { ...selection },
    label,
  };
}

function recordHistory(label) {
  if (history.restoring) return;
  const state = captureHistoryState(label);
  const previous = history.undo.at(-1);
  if (!previous || previous.workbookJson !== state.workbookJson) {
    history.undo.push(state);
    if (history.undo.length > history.limit) history.undo.shift();
  }
  history.redo.length = 0;
  updateHistoryButtons();
}

function restoreHistoryState(state) {
  history.restoring = true;
  try {
    workbook = normalizeWorkbook(JSON.parse(state.workbookJson));
    const sheet = activeSheet();
    selection = {
      r1: Math.max(1, Math.min(sheet.rows, state.selection?.r1 || 1)),
      c1: Math.max(1, Math.min(sheet.cols, state.selection?.c1 || 1)),
      r2: Math.max(1, Math.min(sheet.rows, state.selection?.r2 || state.selection?.r1 || 1)),
      c2: Math.max(1, Math.min(sheet.cols, state.selection?.c2 || state.selection?.c1 || 1)),
    };
    resetCalculationState();
    endFormulaEdit();
    renderSheetTabs();
    updateViewToggleButtons();
    buildGrid();
    scheduleSave();
    scheduleRecalculation(0);
  } finally {
    history.restoring = false;
    updateHistoryButtons();
  }
}

function undoWorkbook() {
  const state = history.undo.pop();
  if (!state) return;
  history.redo.push(captureHistoryState(state.label));
  restoreHistoryState(state);
  els.saveStatus.textContent = state.label ? `Undid ${state.label}` : 'Undone';
}

function redoWorkbook() {
  const state = history.redo.pop();
  if (!state) return;
  history.undo.push(captureHistoryState(state.label));
  restoreHistoryState(state);
  els.saveStatus.textContent = state.label ? `Redid ${state.label}` : 'Redone';
}

function updateHistoryButtons() {
  els.undoBtn.disabled = history.undo.length === 0;
  els.redoBtn.disabled = history.redo.length === 0;
  els.undoBtn.title = history.undo.length ? `Undo ${history.undo.at(-1).label || ''} (Ctrl+Z)`.trim() : 'Undo (Ctrl+Z)';
  els.redoBtn.title = history.redo.length ? `Redo ${history.redo.at(-1).label || ''} (Ctrl+Y or Ctrl+Shift+Z)`.trim() : 'Redo (Ctrl+Y or Ctrl+Shift+Z)';
}

function activeSheet() {
  return workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) || workbook.sheets[0];
}

function sheetById(id) {
  return workbook.sheets.find((sheet) => sheet.id === id) || null;
}

function sheetByName(name) {
  return workbook.sheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase()) || null;
}

function cellKey(sheetId, address) {
  return `${sheetId}!${address.toUpperCase()}`;
}

function coordKey(sheetId, row, col) {
  return cellKey(sheetId, toA1(row, col));
}

function columnLabel(col) {
  let n = col;
  let label = '';
  while (n > 0) {
    n -= 1;
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26);
  }
  return label;
}

function columnNumber(label) {
  let value = 0;
  for (const char of label.toUpperCase()) {
    if (char < 'A' || char > 'Z') throw new Error(`Invalid column: ${label}`);
    value = value * 26 + char.charCodeAt(0) - 64;
  }
  return value;
}

function toA1(row, col) {
  return `${columnLabel(col)}${row}`;
}

function toR1C1(row, col) {
  return `R${row}C${col}`;
}

function currentReferenceStyle() {
  return workbook.view?.referenceStyle === 'R1C1' ? 'R1C1' : 'A1';
}

function referenceToken(row, col, style = currentReferenceStyle()) {
  return style === 'R1C1' ? toR1C1(row, col) : toA1(row, col);
}

function rangeAddressForStyle(range = selection, style = currentReferenceStyle()) {
  const a = referenceToken(Math.min(range.r1, range.r2), Math.min(range.c1, range.c2), style);
  const b = referenceToken(Math.max(range.r1, range.r2), Math.max(range.c1, range.c2), style);
  return a === b ? a : `${a}:${b}`;
}

function parseA1Address(address) {
  const match = /^([A-Za-z]+)([1-9]\d*)$/.exec(String(address).trim());
  if (!match) throw new Error(`Invalid A1 address: ${address}`);
  return { row: Number(match[2]), col: columnNumber(match[1]) };
}

function parseCellToken(token) {
  const text = String(token).trim();
  const a1 = /^([A-Za-z]+)([1-9]\d*)$/.exec(text);
  if (a1) return { row: Number(a1[2]), col: columnNumber(a1[1]) };
  const r1c1 = /^R([1-9]\d*)C([1-9]\d*)$/i.exec(text);
  if (r1c1) return { row: Number(r1c1[1]), col: Number(r1c1[2]) };
  throw new Error(`#REF!: invalid reference token "${token}"`);
}

function splitQualifiedReference(text) {
  const trimmed = text.trim();
  const quoted = /^'((?:[^']|'')+)'!(.+)$/.exec(trimmed);
  if (quoted) return { sheetName: quoted[1].replaceAll("''", "'"), local: quoted[2] };
  const bang = trimmed.indexOf('!');
  if (bang > 0) return { sheetName: trimmed.slice(0, bang), local: trimmed.slice(bang + 1) };
  return { sheetName: null, local: trimmed };
}


function setWorkbookName(name, definition) {
  for (const key of Object.keys(workbook.names)) {
    if (key.toLowerCase() === name.toLowerCase()) delete workbook.names[key];
  }
  workbook.names[name] = definition;
}

function lookupName(name) {
  const entry = Object.entries(workbook.names).find(([key]) => key.toLowerCase() === name.toLowerCase());
  return entry ? { name: entry[0], ...entry[1] } : null;
}

function parseReference(text, currentSheetId, depth = 0) {
  if (depth > 4) throw new Error('#REF!: named range recursion is too deep');
  const original = String(text).trim();
  if (!original) throw new Error('#REF!: empty reference');

  const qualified = splitQualifiedReference(original);
  let targetSheet = qualified.sheetName ? sheetByName(qualified.sheetName) : sheetById(currentSheetId);
  let local = qualified.local.trim();
  let dynamic = false;
  if (local.endsWith('#')) {
    dynamic = true;
    local = local.slice(0, -1).trim();
  }

  const directPattern = /^(?:[A-Za-z]+[1-9]\d*|R[1-9]\d*C[1-9]\d*)(?::(?:[A-Za-z]+[1-9]\d*|R[1-9]\d*C[1-9]\d*))?$/i;
  if (!directPattern.test(local)) {
    if (qualified.sheetName) throw new Error(`#REF!: invalid range "${original}"`);
    const named = lookupName(local);
    if (!named) throw new Error(`#REF!: unknown name or range "${original}"`);
    if (named.kind === 'expression') throw new Error(`#REF!: name "${named.name}" contains an R value, not a worksheet range`);
    const namedSheet = sheetById(named.sheetId);
    if (!namedSheet) throw new Error(`#REF!: named range "${named.name}" points to a missing sheet`);
    const parsed = parseReference(`${quoteSheetName(namedSheet.name)}!${named.ref}`, named.sheetId, depth + 1);
    if (dynamic) parsed.dynamic = true;
    parsed.sourceName = named.name;
    return parsed;
  }

  if (!targetSheet) throw new Error(`#REF!: unknown sheet "${qualified.sheetName}"`);
  const parts = local.split(':');
  if (dynamic && parts.length !== 1) throw new Error('#REF!: the # operator requires one spill anchor');
  const start = parseCellToken(parts[0]);
  const end = parts.length === 2 ? parseCellToken(parts[1]) : start;
  const r1 = Math.min(start.row, end.row);
  const c1 = Math.min(start.col, end.col);
  const r2 = Math.max(start.row, end.row);
  const c2 = Math.max(start.col, end.col);
  if (r1 < 1 || c1 < 1 || r2 > targetSheet.rows || c2 > targetSheet.cols) {
    throw new Error(`#REF!: "${original}" is outside worksheet bounds`);
  }
  return { sheetId: targetSheet.id, r1, c1, r2, c2, dynamic, original };
}

function quoteSheetName(name) {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replaceAll("'", "''")}'`;
}

function rangeAddress(range = selection) {
  const a = toA1(Math.min(range.r1, range.r2), Math.min(range.c1, range.c2));
  const b = toA1(Math.max(range.r1, range.r2), Math.max(range.c1, range.c2));
  return a === b ? a : `${a}:${b}`;
}

function setFormulaMessage(message = '', state = '') {
  els.formulaMessage.textContent = message;
  els.formulaMessage.className = `formula-message${state ? ` ${state}` : ''}`;
  els.formulaMessage.title = message;
}


const R_KEYWORDS = new Set(['if', 'else', 'repeat', 'while', 'function', 'for', 'in', 'next', 'break']);
const R_CONSTANTS = new Set([
  'TRUE', 'FALSE', 'NULL', 'Inf', 'NaN', 'NA', 'NA_integer_', 'NA_real_', 'NA_complex_', 'NA_character_',
]);

function escapeFormulaHtml(text) {
  return text.replace(/[&<>]/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[character]));
}

function delimiterMatchPositions(source, caret) {
  const openToClose = { '(': ')', '[': ']', '{': '}' };
  const closeToOpen = { ')': '(', ']': '[', '}': '{' };
  const pairs = new Map();
  const unmatched = new Set();
  const stack = [];
  let quote = null;
  let comment = false;
  let escaped = false;
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (comment) {
      if (character === '\n') comment = false;
      continue;
    }
    if (quote) {
      if (escaped) { escaped = false; continue; }
      if (character === '\\') { escaped = true; continue; }
      if (character === quote) quote = null;
      continue;
    }
    if (character === '#') { comment = true; continue; }
    if (character === '"' || character === "'" || character === '`') { quote = character; continue; }
    if (openToClose[character]) {
      stack.push({ character, index });
      continue;
    }
    if (closeToOpen[character]) {
      const top = stack.at(-1);
      if (top?.character === closeToOpen[character]) {
        stack.pop();
        pairs.set(index, top.index);
        pairs.set(top.index, index);
      } else {
        unmatched.add(index);
      }
    }
  }
  for (const item of stack) unmatched.add(item.index);
  const candidates = [caret, caret - 1].filter((position) => position >= 0 && position < source.length);
  const position = candidates.find((candidate) => openToClose[source[candidate]] || closeToOpen[source[candidate]]);
  if (position === undefined) return { matches: new Set(), mismatches: new Set() };
  if (pairs.has(position)) return { matches: new Set([position, pairs.get(position)]), mismatches: new Set() };
  return { matches: new Set(), mismatches: new Set([position]) };
}

function rSyntaxHtml(source, marks = { matches: new Set(), mismatches: new Set() }) {
  let output = '';
  let index = 0;
  const append = (className, value, startPosition = index) => {
    let plainStart = 0;
    const flush = (end) => {
      if (end <= plainStart) return;
      const text = escapeFormulaHtml(value.slice(plainStart, end));
      output += className ? `<span class="${className}">${text}</span>` : text;
    };
    for (let offset = 0; offset < value.length; offset += 1) {
      const absolute = startPosition + offset;
      const markClass = marks.matches.has(absolute) ? 'rtok-match' : marks.mismatches.has(absolute) ? 'rtok-mismatch' : '';
      if (!markClass) continue;
      flush(offset);
      const classes = [className, markClass].filter(Boolean).join(' ');
      output += `<span class="${classes}">${escapeFormulaHtml(value[offset])}</span>`;
      plainStart = offset + 1;
    }
    flush(value.length);
  };
  const identifierStart = (character) => /[A-Za-z_.]/.test(character || '');
  const identifierPart = (character) => /[A-Za-z0-9._]/.test(character || '');

  while (index < source.length) {
    const character = source[index];
    if (index === 0 && character === '=') {
      append('rtok-formula', character, index);
      index += 1;
      continue;
    }
    if (character === '#') {
      let end = index + 1;
      while (end < source.length && source[end] !== '\n') end += 1;
      append('rtok-comment', source.slice(index, end), index);
      index = end;
      continue;
    }
    if (character === '"' || character === "'" || character === '`') {
      const quote = character;
      let end = index + 1;
      while (end < source.length) {
        if (source[end] === '\\') { end += 2; continue; }
        if (source[end] === quote) { end += 1; break; }
        end += 1;
      }
      append('rtok-string', source.slice(index, end), index);
      index = end;
      continue;
    }
    const numberMatch = source.slice(index).match(/^(?:0[xX][0-9A-Fa-f]+(?:p[+-]?\d+)?|(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)[Li]?/);
    if (numberMatch) {
      append('rtok-number', numberMatch[0], index);
      index += numberMatch[0].length;
      continue;
    }
    if (character === '%') {
      const end = source.indexOf('%', index + 1);
      if (end >= 0) {
        append('rtok-operator', source.slice(index, end + 1), index);
        index = end + 1;
        continue;
      }
    }
    const operatorMatch = source.slice(index).match(/^(?:<<-|->>|<-|->|\|>|:::|::|===|==|!=|<=|>=|&&|\|\||\+|-|\*|\/|\^|%%|%\/%|=|<|>|!|&|\||~|\?|:|\$|@)/);
    if (operatorMatch) {
      append('rtok-operator', operatorMatch[0], index);
      index += operatorMatch[0].length;
      continue;
    }
    if (identifierStart(character)) {
      let end = index + 1;
      while (end < source.length && identifierPart(source[end])) end += 1;
      const identifier = source.slice(index, end);
      let lookahead = end;
      while (/\s/.test(source[lookahead] || '')) lookahead += 1;
      let className = '';
      if (R_KEYWORDS.has(identifier)) className = 'rtok-keyword';
      else if (R_CONSTANTS.has(identifier)) className = 'rtok-constant';
      else if (source.slice(lookahead, lookahead + 3) === ':::' || source.slice(lookahead, lookahead + 2) === '::') className = 'rtok-namespace';
      else if (source[lookahead] === '(') className = identifier === 'ref' ? 'rtok-ref' : 'rtok-function';
      append(className, identifier, index);
      index = end;
      continue;
    }
    append('', character, index);
    index += 1;
  }
  return output || ' ';
}

function syncFormulaEditorScroll() {
  if (formulaCodeMirror || !els.formulaHighlight) return;
  els.formulaHighlight.scrollTop = els.formulaInput.scrollTop;
  els.formulaHighlight.scrollLeft = els.formulaInput.scrollLeft;
}

function updateFormulaSyntaxHighlight() {
  if (formulaCodeMirror) {
    refreshFormulaCodeEditor();
    return;
  }
  if (!els.formulaHighlight) return;
  const code = els.formulaHighlight.querySelector('code');
  if (code) {
    const caret = els.formulaInput.selectionStart ?? 0;
    code.innerHTML = `${rSyntaxHtml(els.formulaInput.value, delimiterMatchPositions(els.formulaInput.value, caret))}<span class="formula-sentinel">&#8203;</span>`;
  }
  syncFormulaEditorScroll();
}

function insertFormulaLineBreak() {
  const { start, end } = formulaSelectionOffsets();
  replaceFormulaRange('\n', start, end);
  resetFormulaKeyboardPick();
  beginFormulaEdit();
  updateFormulaSyntaxHighlight();
  updateFormulaReferenceHighlight();
}

function wrapSelectedFormulaText(opening) {
  const pairs = { "'": "'", '"': '"', '(': ')', '[': ']', '{': '}' };
  const closing = pairs[opening];
  const { start, end } = formulaSelectionOffsets();
  if (!closing || end <= start) return false;
  const selected = formulaValue().slice(start, end);
  replaceFormulaRange(`${opening}${selected}${closing}`, start, end, start + 1, end + 1);
  resetFormulaKeyboardPick();
  beginFormulaEdit();
  updateFormulaSyntaxHighlight();
  updateFormulaReferenceHighlight();
  return true;
}

function loadPlotResolution() {
  try {
    const saved = JSON.parse(localStorage.getItem(PLOT_RESOLUTION_KEY) || 'null');
    return normalizePlotResolution(saved?.width, saved?.height);
  } catch {
    return { ...DEFAULT_PLOT_RESOLUTION };
  }
}

function normalizePlotResolution(width, height) {
  const normalizedWidth = Math.round(Math.min(4096, Math.max(320, Number(width) || DEFAULT_PLOT_RESOLUTION.width)) / 2) * 2;
  const normalizedHeight = Math.round(Math.min(4096, Math.max(240, Number(height) || DEFAULT_PLOT_RESOLUTION.height)) / 2) * 2;
  return { width: normalizedWidth, height: normalizedHeight };
}

function plotDeviceOptions() {
  return {
    width: Math.max(1, Math.round(plotResolution.width)),
    height: Math.max(1, Math.round(plotResolution.height)),
    bg: 'white',
    capture: true,
  };
}

function plotResolutionPresetValue(width = plotResolution.width, height = plotResolution.height) {
  const value = `${width}x${height}`;
  return [...els.plotResolutionPreset.options].some((option) => option.value === value) ? value : 'custom';
}

function updatePlotResolutionEditor() {
  els.plotWidthInput.value = String(plotResolution.width);
  els.plotHeightInput.value = String(plotResolution.height);
  els.plotResolutionPreset.value = plotResolutionPresetValue();
  updatePlotResolutionHelp();
}

function updatePlotResolutionHelp() {
  const candidate = normalizePlotResolution(els.plotWidthInput.value, els.plotHeightInput.value);
  els.plotResolutionPreset.value = plotResolutionPresetValue(candidate.width, candidate.height);
  els.plotResolutionHelp.textContent = `Final image: ${candidate.width} × ${candidate.height} px. local R PNG device: ${candidate.width} × ${candidate.height} px.`;
}

function applyPlotResolution() {
  plotResolution = normalizePlotResolution(els.plotWidthInput.value, els.plotHeightInput.value);
  try { localStorage.setItem(PLOT_RESOLUTION_KEY, JSON.stringify(plotResolution)); }
  catch { /* Resolution persistence is optional. */ }
  updatePlotResolutionEditor();
  els.plotSettingsDialog.close();
  els.saveStatus.textContent = `Plot resolution: ${plotResolution.width} × ${plotResolution.height} px`;
  scheduleRecalculation(0, { force: true });
}

function clampPlotPaneWidth(width) {
  const viewportWidth = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
  const maximum = Math.max(280, Math.min(1800, viewportWidth - 140));
  return Math.round(Math.min(maximum, Math.max(280, Number(width) || 390)));
}

function setPlotPaneWidth(width, { persist = true } = {}) {
  plotPaneWidth = clampPlotPaneWidth(width);
  document.documentElement.style.setProperty('--plot-pane-width', `${plotPaneWidth}px`);
  els.plotPaneResizer?.setAttribute('aria-valuenow', String(plotPaneWidth));
  els.plotPaneResizer?.setAttribute('aria-valuemax', String(clampPlotPaneWidth(100000)));
  if (persist) {
    try { localStorage.setItem(PLOT_PANE_WIDTH_KEY, String(plotPaneWidth)); }
    catch { /* Width persistence is optional. */ }
  }
}

function beginPlotPaneResize(event) {
  if (event.button !== 0 && event.pointerType !== 'touch') return;
  event.preventDefault();
  const startX = event.clientX;
  const startWidth = plotPaneWidth;
  els.plotPaneResizer.classList.add('dragging');
  els.plotPaneResizer.setPointerCapture?.(event.pointerId);
  const move = (moveEvent) => setPlotPaneWidth(startWidth + startX - moveEvent.clientX, { persist: false });
  const finish = () => {
    els.plotPaneResizer.classList.remove('dragging');
    try { localStorage.setItem(PLOT_PANE_WIDTH_KEY, String(plotPaneWidth)); }
    catch { /* Width persistence is optional. */ }
    window.removeEventListener('pointermove', move);
    window.removeEventListener('pointerup', finish);
    window.removeEventListener('pointercancel', finish);
  };
  window.addEventListener('pointermove', move);
  window.addEventListener('pointerup', finish);
  window.addEventListener('pointercancel', finish);
}


function clampFormulaEditorHeight(height) {
  const maximum = Math.max(80, Math.min(window.innerHeight * 0.65, 720));
  return Math.round(Math.min(maximum, Math.max(FORMULA_COLLAPSED_HEIGHT, Number(height) || FORMULA_COLLAPSED_HEIGHT)));
}

function setFormulaEditorHeight(height, { persist = true, rememberExpanded = true } = {}) {
  formulaEditorHeight = clampFormulaEditorHeight(height);
  if (rememberExpanded && formulaEditorHeight > FORMULA_COLLAPSED_HEIGHT) formulaExpandedHeight = formulaEditorHeight;
  els.formulaEditor.style.height = `${formulaEditorHeight}px`;
  if (formulaCodeMirror) formulaCodeMirror.setSize('100%', '100%');
  else {
    els.formulaInput.style.height = '100%';
    els.formulaHighlight.style.height = '100%';
  }
  refreshFormulaCodeEditor();
  if (persist) {
    try { localStorage.setItem(FORMULA_HEIGHT_KEY, String(formulaEditorHeight)); }
    catch { /* Formula editor height persistence is optional. */ }
  }
}

function beginFormulaEditorResize(event) {
  if (event.button !== 0 && event.pointerType !== 'touch') return;
  event.preventDefault();
  const startY = event.clientY;
  const startHeight = els.formulaEditor.getBoundingClientRect().height;
  els.formulaEditorResizer.classList.add('dragging');
  els.formulaEditorResizer.setPointerCapture?.(event.pointerId);
  const move = (moveEvent) => setFormulaEditorHeight(startHeight + moveEvent.clientY - startY, { persist: false, rememberExpanded: false });
  const finish = () => {
    els.formulaEditorResizer.classList.remove('dragging');
    setFormulaEditorHeight(formulaEditorHeight, { persist: true, rememberExpanded: formulaEditorHeight > FORMULA_COLLAPSED_HEIGHT });
    window.removeEventListener('pointermove', move);
    window.removeEventListener('pointerup', finish);
    window.removeEventListener('pointercancel', finish);
  };
  window.addEventListener('pointermove', move);
  window.addEventListener('pointerup', finish);
  window.addEventListener('pointercancel', finish);
}

function toggleFormulaEditorHeight() {
  if (formulaEditorHeight <= FORMULA_COLLAPSED_HEIGHT) {
    setFormulaEditorHeight(formulaExpandedHeight > FORMULA_COLLAPSED_HEIGHT
      ? formulaExpandedHeight
      : FORMULA_FALLBACK_EXPANDED_HEIGHT);
  } else {
    formulaExpandedHeight = formulaEditorHeight;
    setFormulaEditorHeight(FORMULA_COLLAPSED_HEIGHT, { rememberExpanded: false });
  }
}

function currentTheme() {
  try {
    const stored = localStorage.getItem(THEME_KEY);
    if (stored === 'light' || stored === 'dark') return stored;
  } catch { /* Ignore theme storage failures. */ }
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function applyTheme(theme, { persist = true } = {}) {
  const normalized = theme === 'dark' ? 'dark' : 'light';
  document.documentElement.dataset.theme = normalized;
  document.querySelector('meta[name="color-scheme"]')?.setAttribute('content', normalized);
  els.themeBtn.textContent = normalized === 'dark' ? 'Light mode' : 'Dark mode';
  els.themeBtn.setAttribute('aria-pressed', normalized === 'dark' ? 'true' : 'false');
  if (persist) {
    try { localStorage.setItem(THEME_KEY, normalized); }
    catch { /* Theme persistence is optional. */ }
  }
  refreshFormulaCodeEditor();
}

function updateViewToggleButtons() {
  const showNames = Boolean(workbook.view?.showElementNames);
  const referenceStyle = currentReferenceStyle();
  els.toggleElementNamesBtn.setAttribute('aria-pressed', String(showNames));
  els.referenceStyleBtn?.setAttribute('aria-pressed', String(referenceStyle === 'R1C1'));
  els.toggleElementNamesBtn.title = `${showNames ? 'Hide' : 'Show'} vector, row, and column names`;
  const styleIcon = els.referenceStyleBtn?.querySelector('.ribbon-icon');
  if (styleIcon) styleIcon.textContent = referenceStyle;
  els.referenceStyleBtn.title = `Currently using ${referenceStyle}; switch to ${referenceStyle === 'A1' ? 'R1C1' : 'A1'} references`;
}

function toggleReferenceStyle() {
  recordHistory('toggle A1/R1C1 reference style');
  workbook.view ??= { showElementNames: true, referenceStyle: 'A1' };
  workbook.view.referenceStyle = currentReferenceStyle() === 'A1' ? 'R1C1' : 'A1';
  updateViewToggleButtons();
  buildGrid();
  scheduleSave();
}

function toggleWorkbookView(setting, label) {
  recordHistory(`toggle ${label}`);
  workbook.view ??= { showElementNames: true, referenceStyle: 'A1' };
  workbook.view[setting] = !workbook.view[setting];
  updateViewToggleButtons();
  scheduleSave();
  scheduleRecalculation(0, { force: true });
}

function clearFormulaReferenceHighlight() {
  for (const address of formulaReferenceAddresses) cellElements.get(address)?.classList.remove('formula-ref-target');
  formulaReferenceAddresses.clear();
}

function beginFormulaEdit() {
  if (formulaEdit.active) return;
  formulaEdit.active = true;
  formulaEdit.sheetId = activeSheet().id;
  formulaEdit.row = selection.r1;
  formulaEdit.col = selection.c1;
}

function endFormulaEdit() {
  resetFormulaKeyboardPick();
  formulaEdit.active = false;
  formulaEdit.sheetId = null;
  formulaEdit.row = null;
  formulaEdit.col = null;
  clearFormulaReferenceHighlight();
}

function isFormulaReferencePicking() {
  return formulaEdit.active && formulaValue().trimStart().startsWith('=');
}

function resetFormulaKeyboardPick() {
  formulaKeyboardPick.active = false;
  formulaKeyboardPick.sheetId = null;
  formulaKeyboardPick.start = null;
  formulaKeyboardPick.end = null;
}

function normalizedRange(r1, c1, r2 = r1, c2 = c1) {
  return {
    r1: Math.min(r1, r2),
    c1: Math.min(c1, c2),
    r2: Math.max(r1, r2),
    c2: Math.max(c1, c2),
  };
}

function expandDynamicReference(parsed) {
  const expanded = { ...parsed };
  if (!expanded.dynamic) return expanded;
  const result = computed.get(coordKey(expanded.sheetId, expanded.r1, expanded.c1));
  if (result?.matrix && !result.error) {
    expanded.r2 = expanded.r1 + result.matrix.length - 1;
    expanded.c2 = expanded.c1 + (result.matrix[0]?.length || 1) - 1;
  }
  return expanded;
}

function preferredNameForRange(sheetId, r1, c1, r2 = r1, c2 = c1) {
  const target = normalizedRange(r1, c1, r2, c2);
  const matches = [];
  for (const [name, definition] of Object.entries(workbook.names)) {
    if (definition.kind === 'expression') continue;
    try {
      const parsed = expandDynamicReference(parseReference(definition.ref, definition.sheetId));
      if (parsed.sheetId === sheetId && parsed.r1 === target.r1 && parsed.c1 === target.c1 && parsed.r2 === target.r2 && parsed.c2 === target.c2) {
        matches.push(name);
      }
    } catch { /* Ignore broken names while editing. */ }
  }
  return matches.sort((a, b) => a.localeCompare(b))[0] || null;
}

function spillReferenceForRange(sheetId, r1, c1, r2 = r1, c2 = c1) {
  const target = normalizedRange(r1, c1, r2, c2);
  const anchorKey = coordKey(sheetId, target.r1, target.c1);
  const owner = spillOwners.get(anchorKey);
  if (!owner || owner !== anchorKey) return null;
  const result = computed.get(owner);
  if (!result?.matrix || result.error || matrixArea(result.matrix) <= 1) return null;
  const rows = result.matrix.length;
  const cols = result.matrix[0]?.length || 1;
  if (target.r2 !== target.r1 + rows - 1 || target.c2 !== target.c1 + cols - 1) return null;
  for (let row = target.r1; row <= target.r2; row += 1) {
    for (let col = target.c1; col <= target.c2; col += 1) {
      if (spillOwners.get(coordKey(sheetId, row, col)) !== owner) return null;
    }
  }
  return `${referenceToken(target.r1, target.c1)}#`;
}

function formulaReferenceForRange(sheetId, r1, c1, r2 = r1, c2 = c1) {
  const target = normalizedRange(r1, c1, r2, c2);
  const named = preferredNameForRange(sheetId, target.r1, target.c1, target.r2, target.c2);
  if (named) return named;
  const local = spillReferenceForRange(sheetId, target.r1, target.c1, target.r2, target.c2)
    || rangeAddressForStyle(target);
  if (sheetId === formulaEdit.sheetId) return local;
  return `${quoteSheetName(sheetById(sheetId)?.name || '')}!${local}`;
}

function findRefSpans(text) {
  const spans = [];
  const regex = /\bref\s*\(\s*(["'])(.*?)\1\s*\)/gis;
  let match;
  while ((match = regex.exec(text))) spans.push({ start: match.index, end: regex.lastIndex, ref: match[2] });
  return spans;
}

function highlightFormulaReference(refText) {
  clearFormulaReferenceHighlight();
  let parsed;
  try {
    parsed = expandDynamicReference(parseReference(refText, formulaEdit.sheetId || activeSheet().id));
  } catch (error) {
    setFormulaMessage(String(error.message || error), 'error');
    return;
  }
  const sheet = sheetById(parsed.sheetId);
  const label = parsed.dynamic
    ? `${sheet?.name || '?'}!${toA1(parsed.r1, parsed.c1)}# (spill ${rangeFromParsed(parsed)})`
    : `${sheet?.name || '?'}!${rangeFromParsed(parsed)}`;
  if (parsed.sheetId !== activeSheet().id) {
    setFormulaMessage(`Reference ${refText} points to ${label}`, 'info');
    return;
  }
  for (let row = parsed.r1; row <= parsed.r2; row += 1) {
    for (let col = parsed.c1; col <= parsed.c2; col += 1) {
      const address = toA1(row, col);
      const cell = cellElements.get(address);
      if (cell) {
        cell.classList.add('formula-ref-target');
        formulaReferenceAddresses.add(address);
      }
    }
  }
  setFormulaMessage(`Reference ${refText} → ${label}`, 'info');
}

function updateFormulaReferenceHighlight() {
  if (!formulaEdit.active) return clearFormulaReferenceHighlight();
  const position = formulaSelectionOffsets().end;
  const span = findRefSpans(formulaValue()).find((item) => position >= item.start && position <= item.end);
  if (span) highlightFormulaReference(span.ref);
  else {
    clearFormulaReferenceHighlight();
    updateFormulaMessage();
  }
}

function insertReferenceFromRange(sheetId, r1, c1, r2 = r1, c2 = c1, start = null, end = null) {
  resetFormulaKeyboardPick();
  const reference = formulaReferenceForRange(sheetId, r1, c1, r2, c2);
  const insertion = `ref(${rString(reference)})`;
  const selectionOffsets = formulaSelectionOffsets();
  const insertionStart = start ?? selectionOffsets.start ?? formulaValue().length;
  const insertionEnd = end ?? selectionOffsets.end ?? insertionStart;
  replaceFormulaRange(insertion, insertionStart, insertionEnd);
  updateFormulaSyntaxHighlight();
  focusFormulaEditor();
  highlightFormulaReference(reference);
}

function insertReferenceFromCell(row, col) {
  insertReferenceFromRange(activeSheet().id, row, col);
}

function selectionContains(row, col) {
  const r1 = Math.min(selection.r1, selection.r2);
  const r2 = Math.max(selection.r1, selection.r2);
  const c1 = Math.min(selection.c1, selection.c2);
  const c2 = Math.max(selection.c1, selection.c2);
  return row >= r1 && row <= r2 && col >= c1 && col <= c2;
}

function buildGrid() {
  hideCellTooltip();
  const sheet = activeSheet();
  const fragment = document.createDocumentFragment();
  const headerRow = document.createElement('tr');
  const corner = document.createElement('th');
  corner.className = 'corner';
  corner.setAttribute('aria-hidden', 'true');
  headerRow.appendChild(corner);
  for (let col = 1; col <= sheet.cols; col += 1) {
    const th = document.createElement('th');
    th.className = 'col-header';
    th.textContent = currentReferenceStyle() === 'R1C1' ? String(col) : columnLabel(col);
    th.dataset.col = String(col);
    headerRow.appendChild(th);
  }
  fragment.appendChild(headerRow);
  cellElements = new Map();

  for (let row = 1; row <= sheet.rows; row += 1) {
    const tr = document.createElement('tr');
    const rowHeader = document.createElement('th');
    rowHeader.className = 'row-header';
    rowHeader.textContent = String(row);
    rowHeader.dataset.row = String(row);
    tr.appendChild(rowHeader);
    for (let col = 1; col <= sheet.cols; col += 1) {
      const address = toA1(row, col);
      const td = document.createElement('td');
      td.className = 'cell';
      td.dataset.row = String(row);
      td.dataset.col = String(col);
      td.dataset.address = address;
      td.setAttribute('role', 'gridcell');
      td.setAttribute('aria-label', referenceToken(row, col));
      tr.appendChild(td);
      cellElements.set(address, td);
    }
    fragment.appendChild(tr);
  }
  els.grid.replaceChildren(fragment);
  document.body.scrollLeft = 0;
  document.documentElement.scrollLeft = 0;
  previousSelectedAddresses.clear();
  previousActiveHeaderElements = [];
  refreshGridValues();
  refreshSelection(true);
}

function appendInlineMarkdown(target, text) {
  const source = String(text ?? '');
  const pattern = /(\*\*|__)(?=\S)([\s\S]*?\S)\1|(\*|_)(?=\S)([\s\S]*?\S)\3/g;
  let cursor = 0;
  let match;
  while ((match = pattern.exec(source))) {
    const marker = match[1] || match[3];
    const before = match.index > 0 ? source[match.index - 1] : '';
    const afterIndex = match.index + match[0].length;
    const after = afterIndex < source.length ? source[afterIndex] : '';
    if (marker.includes('_') && /[A-Za-z0-9]/.test(before) && /[A-Za-z0-9]/.test(after)) continue;
    if (match.index > cursor) target.append(document.createTextNode(source.slice(cursor, match.index)));
    const formatted = document.createElement(match[1] ? 'strong' : 'em');
    formatted.textContent = match[2] || match[4] || '';
    target.append(formatted);
    cursor = afterIndex;
  }
  if (cursor < source.length) target.append(document.createTextNode(source.slice(cursor)));
}

function positionCellTooltip(cell) {
  if (!cell || els.cellTooltip.hidden) return;
  const gap = 7;
  const margin = 8;
  const cellRect = cell.getBoundingClientRect();
  const tooltipRect = els.cellTooltip.getBoundingClientRect();
  let left = cellRect.left;
  let top = cellRect.bottom + gap;
  if (left + tooltipRect.width > window.innerWidth - margin) left = window.innerWidth - tooltipRect.width - margin;
  if (top + tooltipRect.height > window.innerHeight - margin) top = cellRect.top - tooltipRect.height - gap;
  els.cellTooltip.style.left = `${Math.max(margin, left)}px`;
  els.cellTooltip.style.top = `${Math.max(margin, top)}px`;
}

function showCellTooltip(cell) {
  const text = cell?.dataset.tooltipText || '';
  if (!text) { hideCellTooltip(); return; }
  if (tooltipCell && tooltipCell !== cell) tooltipCell.removeAttribute('aria-describedby');
  tooltipCell = cell;
  els.cellTooltip.replaceChildren();
  if (cell.dataset.tooltipMarkdown === 'true') appendInlineMarkdown(els.cellTooltip, text);
  else els.cellTooltip.textContent = text;
  els.cellTooltip.hidden = false;
  cell.setAttribute('aria-describedby', 'cellTooltip');
  positionCellTooltip(cell);
}

function hideCellTooltip() {
  if (tooltipCell) tooltipCell.removeAttribute('aria-describedby');
  tooltipCell = null;
  els.cellTooltip.hidden = true;
  els.cellTooltip.replaceChildren();
}

function refreshGridValues() {
  const sheet = activeSheet();
  for (let row = 1; row <= sheet.rows; row += 1) {
    for (let col = 1; col <= sheet.cols; col += 1) {
      const address = toA1(row, col);
      const td = cellElements.get(address);
      if (!td) continue;
      const key = cellKey(sheet.id, address);
      const value = displayValues.has(key) ? displayValues.get(key) : null;
      const ownerKey = spillOwners.get(key);
      const ownCell = sheet.cells[address];
      const isPureTextCell = Boolean(ownCell && (ownCell.literal || isPlainTextInput(ownCell.input)));
      const ownComputed = computed.get(key);
      const plots = plotsByCell.get(key) || [];
      const displayText = formatDisplay(value);
      const plotKind = ownComputed?.plotKind || '';
      const isGgplot = plotKind === 'ggplot';
      const isLattice = plotKind === 'lattice';
      const isPlotObject = isGgplot || isLattice;
      const hasPlotIndicator = plots.length > 0 || isPlotObject;
      const showPlotLabel = isPlotObject || (plots.length > 0 && (!displayText || displayText === '<object>'));
      const renderedText = showPlotLabel
        ? (plots.length > 1 ? `${plots.length} plots` : 'Plot')
        : (displayText || (plots.length ? (plots.length === 1 ? 'Plot' : `${plots.length} plots`) : ''));

      const valueElement = document.createElement('span');
      valueElement.className = 'cell-value';
      if (showPlotLabel) {
        valueElement.append(document.createTextNode(renderedText));
        if (isPlotObject) {
          const kind = document.createElement('sup');
          kind.className = 'plot-kind-superscript';
          kind.textContent = isGgplot ? 'ggplot2' : 'lattice';
          valueElement.appendChild(kind);
        }
      } else if (typeof value === 'string' && !isErrorValue(value)) {
        appendInlineMarkdown(valueElement, renderedText);
      } else {
        valueElement.textContent = renderedText;
      }
      td.replaceChildren(valueElement);

      const isList = ownComputed?.objectKind === 'list' && !isPlotObject;
      const isGenericObject = ['function', 'environment', 'other'].includes(ownComputed?.objectKind) && !isPlotObject;
      td.classList.toggle('number', !showPlotLabel && typeof value === 'number' && Number.isFinite(value));
      td.classList.toggle('error', isErrorValue(value));
      td.classList.toggle('spill', Boolean(ownerKey && ownerKey !== key));
      td.classList.toggle('spill-anchor', Boolean(ownComputed && !ownComputed.error && matrixArea(ownComputed.matrix) > 1));
      td.classList.toggle('pending', Boolean(ownCell && !runtime.ready && !displayValues.has(key)));
      td.classList.toggle('has-plot', hasPlotIndicator);
      td.classList.toggle('plot-only', showPlotLabel);
      td.classList.toggle('ggplot-cell', isGgplot && showPlotLabel);
      td.classList.toggle('lattice-cell', isLattice && showPlotLabel);
      td.classList.toggle('has-object', isList);
      td.classList.toggle('has-generic-object', isGenericObject && !hasPlotIndicator);
      td.classList.toggle('has-corner-marker', hasPlotIndicator || isList || isGenericObject);
      const pieces = [];
      if (ownCell) pieces.push(`Input: ${ownCell.input}`);
      if (ownComputed?.error && ownComputed.message) pieces.push(`Error: ${ownComputed.message}`);
      if (ownerKey && ownerKey !== key) pieces.push(`Spill from ${ownerKey.split('!')[1]}`);
      if (value !== null && value !== '') pieces.push(`Value: ${formatDisplay(value)}`);
      if (plots.length) pieces.push(`${plots.length} plot${plots.length === 1 ? '' : 's'} generated by this formula`);
      if (isGgplot) pieces.push('ggplot2 plot object');
      if (isLattice) pieces.push('lattice plot object');
      if (isList) pieces.push('Double-click to inspect this R list');
      const tooltipText = isPureTextCell ? ownCell.input : pieces.join('\n');
      if (tooltipText) {
        td.dataset.tooltipText = tooltipText;
        td.dataset.tooltipMarkdown = isPureTextCell ? 'true' : 'false';
      } else {
        delete td.dataset.tooltipText;
        delete td.dataset.tooltipMarkdown;
      }
      td.removeAttribute('title');
      td.setAttribute('aria-selected', selectionContains(row, col) ? 'true' : 'false');
    }
  }
  renderPlotPane();
}

function formatDisplay(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return String(value);
    const abs = Math.abs(value);
    if ((abs !== 0 && abs < 1e-7) || abs >= 1e12) return value.toExponential(6);
    return new Intl.NumberFormat('en-US', { useGrouping: false, maximumSignificantDigits: 12 }).format(value);
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

function isErrorValue(value) {
  return typeof value === 'string' && value.startsWith('#');
}

function matrixArea(matrix) {
  return matrix ? matrix.length * (matrix[0]?.length || 0) : 0;
}

function plotEntries() {
  const sheetOrder = new Map(workbook.sheets.map((sheet, index) => [sheet.id, index]));
  return [...plotsByCell.entries()]
    .map(([key, plots]) => {
      const separator = key.indexOf('!');
      const sheetId = key.slice(0, separator);
      const address = key.slice(separator + 1);
      const sheet = sheetById(sheetId);
      return { key, sheetId, address, sheet, plots };
    })
    .filter((entry) => entry.sheet && entry.plots.length)
    .sort((a, b) => (sheetOrder.get(a.sheetId) - sheetOrder.get(b.sheetId))
      || ((a.sheet.cells[a.address]?.seq || 0) - (b.sheet.cells[b.address]?.seq || 0))
      || a.address.localeCompare(b.address));
}

function showPlotPane(open = true) {
  plotPaneOpen = open;
  renderPlotPane();
}

function renderPlotPane() {
  if (!els.plotPane || !els.plotList) return;
  const entries = plotEntries();
  const total = entries.reduce((sum, entry) => sum + entry.plots.length, 0);
  els.plotsBtn.disabled = total === 0;
  const plotsLabel = els.plotsBtn.querySelector('.ribbon-button-label');
  if (plotsLabel) plotsLabel.innerHTML = total ? `Show plots<br>(${total})` : 'Show<br>plots';
  els.plotPaneCount.textContent = total ? `${total} plot${total === 1 ? '' : 's'} · ${plotResolution.width}×${plotResolution.height}px` : '';
  const hidden = !plotPaneOpen || total === 0;
  els.plotPane.hidden = hidden;
  els.plotPaneResizer.hidden = hidden;
  if (!total) {
    els.plotList.replaceChildren();
    return;
  }
  const activeKey = spillOwners.get(coordKey(activeSheet().id, selection.r1, selection.c1))
    || coordKey(activeSheet().id, selection.r1, selection.c1);
  const fragment = document.createDocumentFragment();
  for (const entry of entries) {
    entry.plots.forEach((plot, index) => {
      const card = document.createElement('article');
      card.className = `plot-card${entry.key === activeKey ? ' selected' : ''}`;
      card.dataset.plotCellKey = entry.key;
      const header = document.createElement('div');
      header.className = 'plot-card-header';
      const label = document.createElement('code');
      label.textContent = `${entry.sheet.name}!${entry.address}`;
      const counter = document.createElement('span');
      counter.className = 'plot-card-index';
      const kind = computed.get(entry.key)?.plotKind;
      counter.textContent = entry.plots.length > 1
        ? `${index + 1}/${entry.plots.length}`
        : (kind === 'ggplot' ? 'ggplot2 plot' : kind === 'lattice' ? 'lattice plot' : 'R plot');
      header.append(label, counter);
      const image = document.createElement('img');
      image.className = 'plot-image';
      image.src = plot.dataUrl;
      image.width = plot.width;
      image.height = plot.height;
      image.alt = `Plot from ${entry.sheet.name}!${entry.address}`;
      image.loading = 'eager';
      card.style.setProperty('--plot-aspect-ratio', `${plot.width} / ${plot.height}`);
      card.append(header, image);
      card.addEventListener('click', () => {
        if (activeSheet().id !== entry.sheetId) activateSheet(entry.sheetId);
        setSelection(parseA1Address(entry.address).row, parseA1Address(entry.address).col);
      });
      fragment.appendChild(card);
    });
  }
  els.plotList.replaceChildren(fragment);
  const selectedCard = els.plotList.querySelector('.plot-card.selected');
  if (selectedCard) selectedCard.scrollIntoView({ block: 'nearest' });
}

function revealPlotForCell(key) {
  const plots = plotsByCell.get(key);
  if (!plots?.length) return false;
  showPlotPane(true);
  requestAnimationFrame(() => {
    const card = [...els.plotList.querySelectorAll('.plot-card')].find((item) => item.dataset.plotCellKey === key);
    if (!card) return;
    card.scrollIntoView({ block: 'start', behavior: 'smooth' });
    card.classList.add('reveal');
    setTimeout(() => card.classList.remove('reveal'), 1200);
  });
  return true;
}

function objectTreeDetails() {
  return [...els.objectTree.querySelectorAll('details.object-node')];
}

function updateObjectTreeExpandButton() {
  const details = objectTreeDetails();
  const allOpen = details.length > 0 && details.every((node) => node.open);
  els.expandObjectTreeBtn.disabled = details.length === 0;
  els.expandObjectTreeBtn.textContent = allOpen ? 'Collapse all' : 'Expand all';
  els.expandObjectTreeBtn.setAttribute('aria-expanded', String(allOpen));
}

function toggleObjectTreeExpansion() {
  const details = objectTreeDetails();
  const expand = details.some((node) => !node.open);
  for (const node of details) node.open = expand;
  updateObjectTreeExpandButton();
  els.objectTree.focus({ preventScroll: true });
}

function renderObjectTree(result) {
  const rows = result?.tree || [];
  if (!rows.length) {
    els.objectTree.textContent = 'No list structure was returned.';
    updateObjectTreeExpandButton();
    return;
  }
  const byParent = new Map();
  for (const row of rows) {
    const group = byParent.get(row.parent) || [];
    group.push(row);
    byParent.set(row.parent, group);
  }
  const build = (row) => {
    const children = byParent.get(row.index) || [];
    if (children.length) {
      const details = document.createElement('details');
      details.className = 'object-node';
      details.open = row.parent === 0 || row.index === 1;
      const summary = document.createElement('summary');
      const label = document.createElement('span');
      label.className = 'object-label';
      label.textContent = row.label;
      const type = document.createElement('span');
      type.className = 'object-type';
      type.textContent = `  ${row.type}`;
      const value = document.createElement('span');
      value.className = 'object-summary';
      value.textContent = `  ${row.summary}`;
      summary.append(label, type, value);
      details.appendChild(summary);
      for (const child of children) details.appendChild(build(child));
      return details;
    }
    const leaf = document.createElement('div');
    leaf.className = 'object-leaf';
    const label = document.createElement('span');
    label.className = 'object-label';
    label.textContent = row.label;
    const type = document.createElement('span');
    type.className = 'object-type';
    type.textContent = row.type;
    const value = document.createElement('span');
    value.className = 'object-summary';
    value.textContent = row.summary;
    leaf.append(label, type, value);
    return leaf;
  };
  const roots = byParent.get(0) || [];
  els.objectTree.replaceChildren(...roots.map(build));
  updateObjectTreeExpandButton();
}

function showObjectViewer(key) {
  const result = computed.get(key);
  if (!result || result.objectKind !== 'list') return false;
  const separator = key.indexOf('!');
  const sheet = sheetById(key.slice(0, separator));
  const address = key.slice(separator + 1);
  els.objectDialogTitle.textContent = `${sheet?.name || '?'}!${address}`;
  els.objectDialogSubtitle.textContent = 'R list tree';
  renderObjectTree(result);
  els.objectDialog.showModal();
  els.objectTree.focus({ preventScroll: true });
  return true;
}

async function bitmapToPlotRecord(bitmap) {
  if (bitmap?.dataUrl) return bitmap;
  const canvas = document.createElement('canvas');
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  const context = canvas.getContext('2d');
  if (!context) throw new Error('The browser could not create a 2D canvas for the R plot.');
  context.drawImage(bitmap, 0, 0, bitmap.width, bitmap.height);
  bitmap.close?.();
  return { width: canvas.width, height: canvas.height, dataUrl: canvas.toDataURL('image/png') };
}

function refreshActiveHeaders() {
  for (const header of previousActiveHeaderElements) {
    header.classList.remove('active-header');
    header.removeAttribute('aria-current');
  }
  previousActiveHeaderElements = [];
  const rowHeader = els.grid.querySelector(`.row-header[data-row="${selection.r1}"]`);
  const columnHeader = els.grid.querySelector(`.col-header[data-col="${selection.c1}"]`);
  for (const header of [rowHeader, columnHeader]) {
    if (!header) continue;
    header.classList.add('active-header');
    header.setAttribute('aria-current', 'true');
    previousActiveHeaderElements.push(header);
  }
}

function refreshSelection(force = false) {
  if (!force) {
    for (const address of previousSelectedAddresses) {
      const td = cellElements.get(address);
      if (td) td.classList.remove('selected', 'active');
    }
  }
  const current = new Set();
  const r1 = Math.min(selection.r1, selection.r2);
  const r2 = Math.max(selection.r1, selection.r2);
  const c1 = Math.min(selection.c1, selection.c2);
  const c2 = Math.max(selection.c1, selection.c2);
  for (let row = r1; row <= r2; row += 1) {
    for (let col = c1; col <= c2; col += 1) {
      const address = toA1(row, col);
      const td = cellElements.get(address);
      if (!td) continue;
      td.classList.add('selected');
      td.setAttribute('aria-selected', 'true');
      current.add(address);
    }
  }
  const activeAddress = toA1(selection.r1, selection.c1);
  const activeTd = cellElements.get(activeAddress);
  if (activeTd) activeTd.classList.add('active');
  refreshActiveHeaders();
  previousSelectedAddresses = current;
  updateFormulaBar();
  updateSelectionSummary();
  renderPlotPane();
}

function updateFormulaBar(force = false) {
  const sheet = activeSheet();
  const address = toA1(selection.r1, selection.c1);
  els.nameBox.value = rangeAddressForStyle();
  if (formulaEdit.active && !force) {
    updateFormulaMessage();
    return;
  }
  const ownCell = sheet.cells[address];
  if (ownCell) {
    setFormulaValue(ownCell.input);
    setFormulaPlaceholder('Enter an R expression');
  } else {
    setFormulaValue('');
    const owner = spillOwners.get(cellKey(sheet.id, address));
    setFormulaPlaceholder(owner && owner !== cellKey(sheet.id, address)
      ? `Spill value from ${owner.split('!')[1]}`
      : `Enter an R expression, for example mean(ref("${currentReferenceStyle() === 'R1C1' ? 'R1C1:R5C1' : 'A1:A5'}"))`);
  }
  updateFormulaSyntaxHighlight();
  updateFormulaMessage();
}

function updateFormulaMessage() {
  if (formulaEdit.active && formulaReferenceAddresses.size) return;
  const sheet = activeSheet();
  const key = coordKey(sheet.id, selection.r1, selection.c1);
  const owner = spillOwners.get(key) || key;
  const result = computed.get(owner);
  if (result?.error) {
    const ownerAddress = owner.split('!')[1];
    const prefix = owner !== key ? `${ownerAddress}: ` : '';
    setFormulaMessage(`${prefix}${result.error} ${result.message || 'R evaluation failed'}`, 'error');
  } else {
    setFormulaMessage('');
  }
}

function updateSelectionSummary() {
  const sheet = activeSheet();
  const values = [];
  const r1 = Math.min(selection.r1, selection.r2);
  const r2 = Math.max(selection.r1, selection.r2);
  const c1 = Math.min(selection.c1, selection.c2);
  const c2 = Math.max(selection.c1, selection.c2);
  let nonblank = 0;
  for (let row = r1; row <= r2; row += 1) {
    for (let col = c1; col <= c2; col += 1) {
      const value = displayValues.get(coordKey(sheet.id, row, col));
      if (value !== null && value !== undefined && value !== '') nonblank += 1;
      if (typeof value === 'number' && Number.isFinite(value)) values.push(value);
    }
  }
  if (values.length) {
    const sum = values.reduce((a, b) => a + b, 0);
    els.selectionSummary.textContent = `Count ${nonblank}   Sum ${formatDisplay(sum)}   Average ${formatDisplay(sum / values.length)}`;
  } else {
    els.selectionSummary.textContent = nonblank ? `Count ${nonblank}` : 'Ready';
  }
}

function setSelection(row, col, extend = false) {
  const sheet = activeSheet();
  const nextRow = Math.max(1, Math.min(sheet.rows, row));
  const nextCol = Math.max(1, Math.min(sheet.cols, col));
  if (extend) {
    selection.r2 = nextRow;
    selection.c2 = nextCol;
  } else {
    selection = { r1: nextRow, c1: nextCol, r2: nextRow, c2: nextCol };
  }
  refreshSelection();
  scrollCellIntoView(nextRow, nextCol);
}

function scrollCellIntoView(row = selection.r2, col = selection.c2) {
  const cell = cellElements.get(toA1(row, col));
  if (!cell) return;
  const viewportRect = els.gridViewport.getBoundingClientRect();
  const cellRect = cell.getBoundingClientRect();
  const leftEdge = viewportRect.left + 47;
  const topEdge = viewportRect.top + 29;
  if (cellRect.left < leftEdge) els.gridViewport.scrollLeft -= leftEdge - cellRect.left;
  if (cellRect.right > viewportRect.right) els.gridViewport.scrollLeft += cellRect.right - viewportRect.right;
  if (cellRect.top < topEdge) els.gridViewport.scrollTop -= topEdge - cellRect.top;
  if (cellRect.bottom > viewportRect.bottom) els.gridViewport.scrollTop += cellRect.bottom - viewportRect.bottom;
  document.body.scrollLeft = 0;
  document.documentElement.scrollLeft = 0;
}

function scrollActiveCellIntoView() {
  scrollCellIntoView(selection.r2, selection.c2);
}

function cellHasNavigationContent(row, col) {
  const sheet = activeSheet();
  const address = toA1(row, col);
  if (sheet.cells[address]?.input) return true;
  const value = displayValues.get(coordKey(sheet.id, row, col));
  return value !== null && value !== undefined && value !== '';
}

function ctrlArrowDestination(row, col, rowStep, colStep) {
  const sheet = activeSheet();
  const inBounds = (r, c) => r >= 1 && r <= sheet.rows && c >= 1 && c <= sheet.cols;
  let currentRow = row;
  let currentCol = col;
  let nextRow = currentRow + rowStep;
  let nextCol = currentCol + colStep;
  if (!inBounds(nextRow, nextCol)) return { row: currentRow, col: currentCol };

  const currentFilled = cellHasNavigationContent(currentRow, currentCol);
  const nextFilled = cellHasNavigationContent(nextRow, nextCol);
  if (currentFilled && nextFilled) {
    while (inBounds(nextRow, nextCol) && cellHasNavigationContent(nextRow, nextCol)) {
      currentRow = nextRow;
      currentCol = nextCol;
      nextRow += rowStep;
      nextCol += colStep;
    }
    return { row: currentRow, col: currentCol };
  }

  while (inBounds(nextRow, nextCol) && !cellHasNavigationContent(nextRow, nextCol)) {
    currentRow = nextRow;
    currentCol = nextCol;
    nextRow += rowStep;
    nextCol += colStep;
  }
  if (inBounds(nextRow, nextCol)) return { row: nextRow, col: nextCol };
  return { row: currentRow, col: currentCol };
}

function moveSelection(rowStep, colStep, { extend = false, jump = false } = {}) {
  const startRow = selection.r2;
  const startCol = selection.c2;
  const destination = jump
    ? ctrlArrowDestination(startRow, startCol, rowStep, colStep)
    : { row: startRow + rowStep, col: startCol + colStep };
  setSelection(destination.row, destination.col, extend);
}

function commitFormula(move = 0) {
  const input = formulaValue();
  const sheetId = formulaEdit.sheetId || activeSheet().id;
  const row = formulaEdit.row || selection.r1;
  const col = formulaEdit.col || selection.c1;
  endFormulaEdit();
  setCellInput(sheetId, row, col, input);
  if (sheetId === activeSheet().id) {
    selection = { r1: row, c1: col, r2: row, c2: col };
    if (move) setSelection(row + move, col);
    else refreshSelection();
  }
  updateFormulaBar(true);
}

function setCellInput(sheetId, row, col, input, { defer = false, record = true, literal = false } = {}) {
  const sheet = sheetById(sheetId);
  if (!sheet || row < 1 || col < 1 || row > sheet.rows || col > sheet.cols) return false;
  const address = toA1(row, col);
  const normalized = String(input).replace(/\r\n/g, '\n');
  const previousCell = sheet.cells[address];
  const previous = previousCell?.input || '';
  if (previous === normalized && Boolean(previousCell?.literal) === Boolean(literal)) return false;
  if (record) recordHistory(`edit ${sheet.name}!${address}`);
  if (normalized === '') {
    delete sheet.cells[address];
  } else {
    workbook.seq += 1;
    sheet.cells[address] = { input: normalized, seq: workbook.seq, ...(literal ? { literal: true } : {}) };
  }
  scheduleSave();
  if (!defer) scheduleRecalculation();
  if (sheetId === activeSheet().id && !formulaEdit.active) updateFormulaBar(true);
  return true;
}

function clearSelection() {
  const sheet = activeSheet();
  const r1 = Math.min(selection.r1, selection.r2);
  const r2 = Math.max(selection.r1, selection.r2);
  const c1 = Math.min(selection.c1, selection.c2);
  const c2 = Math.max(selection.c1, selection.c2);
  let hasInput = false;
  for (let row = r1; row <= r2 && !hasInput; row += 1) {
    for (let col = c1; col <= c2; col += 1) if (sheet.cells[toA1(row, col)]) { hasInput = true; break; }
  }
  if (!hasInput) return;
  recordHistory(`clear ${sheet.name}!${rangeAddress()}`);
  for (let row = r1; row <= r2; row += 1) {
    for (let col = c1; col <= c2; col += 1) delete sheet.cells[toA1(row, col)];
  }
  scheduleSave();
  scheduleRecalculation();
  updateFormulaBar(true);
}

function scheduleSave() {
  clearTimeout(saveTimer);
  els.saveStatus.textContent = 'Saving locally…';
  saveTimer = setTimeout(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(workbook));
      els.saveStatus.textContent = `Saved locally ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;
    } catch (error) {
      els.saveStatus.textContent = 'Autosave failed';
      console.error(error);
    }
  }, 250);
}

function scheduleRecalculation(delay = 80, options = {}) {
  if (options.force) runtime.forceRecalcRequested = true;
  clearTimeout(runtime.recalcTimer);
  runtime.recalcTimer = setTimeout(() => {
    const force = runtime.forceRecalcRequested;
    runtime.forceRecalcRequested = false;
    recalculateWorkbook({ force });
  }, delay);
}

function allInputCells() {
  const cells = [];
  for (const sheet of workbook.sheets) {
    for (const [address, cell] of Object.entries(sheet.cells)) {
      cells.push({
        key: cellKey(sheet.id, address),
        sheetId: sheet.id,
        address,
        input: cell.input,
        seq: Number(cell.seq) || 0,
        literal: Boolean(cell.literal),
      });
    }
  }
  return cells;
}

function expressionFromInput(input) {
  const text = String(input);
  return text.startsWith('=') ? text.slice(1).trim() : text.trim();
}

function matchingCallClose(text, openingIndex) {
  let depth = 0;
  let quote = '';
  let escaped = false;
  for (let index = openingIndex; index < text.length; index += 1) {
    const character = text[index];
    if (quote) {
      if (escaped) escaped = false;
      else if (character === '\\') escaped = true;
      else if (character === quote) quote = '';
      continue;
    }
    if (character === '"' || character === "'" || character === '`') {
      quote = character;
      continue;
    }
    if (character === '(') depth += 1;
    else if (character === ')') {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}

function startsWithRCallExpression(text) {
  const call = text.match(/^(?:(?:[A-Za-z.][A-Za-z0-9._]*):::{0,1})?[A-Za-z.][A-Za-z0-9._]*\s*\(/);
  if (!call) return false;
  const openingIndex = text.indexOf('(', call[0].length - 1);
  const closingIndex = matchingCallClose(text, openingIndex);
  if (closingIndex < 0) return true;
  const suffix = text.slice(closingIndex + 1).trimStart();
  if (!suffix) return true;
  if (suffix.startsWith('#')) return true;
  if (suffix.startsWith(': ')) return false;
  return /^(?:\[|\$|@|\(|\|>|%[^%]*%|\+|-|\*|\/|\^|:{1,2}(?!\s)|==|!=|<=|>=|<|>|&&?|\|\|?|~|\?)/.test(suffix);
}

function isPlainTextInput(input) {
  const source = String(input ?? '');
  if (source.startsWith('=')) return false;
  const text = source.trim();
  if (text === '') return true;
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) return false;
  if (/^(?:TRUE|FALSE|NA(?:_(?:integer|real|complex|character)_)?|NaN|Inf|NULL)$/i.test(text)) return false;
  if (/^(['"`])(?:\\.|(?!\1)[\s\S])*\1$/.test(text)) return false;
  if (/^[{[(]/.test(text)) return false;
  if (/^function\s*\(/i.test(text)) return false;
  if (/^[A-Za-z.][A-Za-z0-9._]*\s*(?:<-|<<-|->|->>)\s*/.test(text)) return false;
  if (/^(?:[+-]?\d+(?:\.\d+)?|[A-Za-z.][A-Za-z0-9._]*)\s*:\s*(?:[+-]?\d+(?:\.\d+)?|[A-Za-z.][A-Za-z0-9._]*)$/.test(text)) return false;
  if (/^[A-Za-z.][A-Za-z0-9._]*(?:\s*(?:\$|@)\s*[A-Za-z.][A-Za-z0-9._]*|\s*\[[\s\S]*\])+$/.test(text)) return false;
  if (/^(?:[+-]?\d+(?:\.\d+)?|[A-Za-z.][A-Za-z0-9._]*)\s*(?:\+|\*|\/|\^|%%|%\/%|==|!=|<=|>=|<|>|&&?|\|\|?)\s*(?:[+-]?\d+(?:\.\d+)?|[A-Za-z.][A-Za-z0-9._]*)/.test(text)) return false;
  if (/^(?:(?:[+-]?\d+(?:\.\d+)?)\s*-\s*(?:[+-]?\d+(?:\.\d+)?|[A-Za-z.][A-Za-z0-9._]*)|[A-Za-z.][A-Za-z0-9._]*\s+-\s+(?:[+-]?\d+(?:\.\d+)?|[A-Za-z.][A-Za-z0-9._]*))/.test(text)) return false;
  if (startsWithRCallExpression(text)) return false;
  return true;
}

function directRefCallsFromExpression(expression) {
  const refs = [];
  const regex = /\bref\s*\(\s*(["'])(.*?)\1\s*\)/gis;
  let match;
  while ((match = regex.exec(expression))) refs.push(match[2]);
  return refs;
}

function extractRefCalls(input, seenNames = new Set()) {
  if (isPlainTextInput(input)) return [];
  const expression = expressionFromInput(input);
  const refs = directRefCallsFromExpression(expression);
  const expanded = [...refs];
  for (const refText of refs) {
    const qualified = splitQualifiedReference(String(refText));
    if (qualified.sheetName) continue;
    const named = lookupName(qualified.local.replace(/#\s*$/, '').trim());
    if (!named || named.kind !== 'expression' || seenNames.has(named.name.toLowerCase())) continue;
    const nextSeen = new Set(seenNames).add(named.name.toLowerCase());
    expanded.push(...extractRefCalls(`=${named.expression}`, nextSeen));
  }
  return [...new Set(expanded.map(String))];
}

function extractPackageNames(input) {
  if (isPlainTextInput(input)) return [];
  const expressions = [expressionFromInput(input)];
  for (const refText of extractRefCalls(input)) {
    const qualified = splitQualifiedReference(String(refText));
    const named = qualified.sheetName ? null : lookupName(qualified.local.replace(/#\s*$/, '').trim());
    if (named?.kind === 'expression') expressions.push(named.expression);
  }
  const expression = expressions.join('\n');
  const packages = new Set();
  const libraryRegex = /\b(?:library|require)\s*\(\s*(?:package\s*=\s*)?(?:"([A-Za-z][A-Za-z0-9.]*)"|'([A-Za-z][A-Za-z0-9.]*)'|([A-Za-z][A-Za-z0-9.]*))/g;
  const namespaceRegex = /(?:^|[^A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)\s*:::{0,1}(?=[A-Za-z.])/g;
  let match;
  while ((match = libraryRegex.exec(expression))) packages.add(match[1] || match[2] || match[3]);
  while ((match = namespaceRegex.exec(expression))) packages.add(match[1]);
  return [...packages];
}

async function ensurePackagesForExpression(input) {
  const packages = extractPackageNames(input);
  if (!packages.length) return;
  const missing = [];
  for (const packageName of packages) {
    if (runtime.packageStatus.get(packageName) === true) continue;
    let installed = false;
    try {
      installed = await runtime.r.evalRBoolean(`requireNamespace(${rString(packageName)}, quietly = TRUE)`);
    } catch { /* Treat lookup failure as missing. */ }
    if (installed) runtime.packageStatus.set(packageName, true);
    else missing.push(packageName);
  }
  if (!missing.length) return;
  setRuntimeStatus(`Installing ${missing.join(', ')}…`, 'loading');
  try {
    await runtime.r.installPackages(missing, { quiet: true });
    for (const packageName of missing) {
      const installed = await runtime.r.evalRBoolean(`requireNamespace(${rString(packageName)}, quietly = TRUE)`);
      if (!installed) throw new Error(`Package ${packageName} was downloaded but is still unavailable.`);
      runtime.packageStatus.set(packageName, true);
    }
  } catch (error) {
    throw new Error(`#PKG!: Could not install ${missing.join(', ')}. ${error?.message || error}`);
  }
}


function createCalculationState() {
  return {
    initialized: false,
    cellSignatures: new Map(),
    dependencies: new Map(),
    graphSignature: '',
  };
}

function clearCalculatedMaps() {
  computed = new Map();
  spillOwners = new Map();
  displayValues = new Map();
  plotsByCell = new Map();
}

function resetCalculationState() {
  clearCalculatedMaps();
  calculationState = createCalculationState();
}

function rRuntime() {
  return runtime.r || runtime.webR;
}

async function runtimeEvalVoid(code) {
  const engine = rRuntime();
  if (!engine) throw new Error('R runtime is unavailable.');
  await engine.evalRVoid(code);
}

function runtimeReadyStatusText() {
  if (runtime.r) return 'Local R ready';
  return globalThis.crossOriginIsolated ? 'webR ready' : 'webR ready (GitHub Pages mode)';
}

function dependencySetSignature(dependencies) {
  return [...(dependencies || new Set())].sort().join('\u001f');
}

function cloneDependencyMap(source) {
  return new Map([...source.entries()].map(([key, value]) => [key, new Set(value)]));
}

function collectReferencedBindings(input, names, sheets, seenNames = new Set()) {
  if (isPlainTextInput(input)) return;
  const expression = expressionFromInput(input);
  for (const refText of directRefCallsFromExpression(expression)) {
    const qualified = splitQualifiedReference(String(refText));
    if (qualified.sheetName) {
      const sheet = sheetByName(qualified.sheetName);
      sheets.add(`${qualified.sheetName.toLowerCase()}\u001f${sheet?.id || ''}`);
      continue;
    }
    const local = qualified.local.replace(/#\s*$/, '').trim();
    const named = lookupName(local);
    if (!named) continue;
    const lower = named.name.toLowerCase();
    names.add(lower);
    if (named.kind === 'expression' && !seenNames.has(lower)) {
      seenNames.add(lower);
      collectReferencedBindings(`=${named.expression}`, names, sheets, seenNames);
    }
  }
}

function workbookNameSignature(lowerName) {
  const entry = Object.entries(workbook.names).find(([key]) => key.toLowerCase() === lowerName);
  if (!entry) return [lowerName, null];
  const [actualName, definition] = entry;
  return [
    lowerName,
    actualName,
    definition.kind || 'range',
    definition.sheetId || '',
    definition.ref || '',
    definition.expression || '',
  ];
}

function cellEvaluationSignature(node) {
  const names = new Set();
  const sheets = new Set();
  if (!node.literal && !isPlainTextInput(node.input)) collectReferencedBindings(node.input, names, sheets);
  return JSON.stringify({
    input: String(node.input ?? ''),
    literal: Boolean(node.literal),
    names: [...names].sort().map(workbookNameSignature),
    sheets: [...sheets].sort(),
  });
}

function cellSignaturesForNodes(nodes) {
  return new Map(nodes.map((node) => [node.key, cellEvaluationSignature(node)]));
}

function fingerprintValue(value) {
  if (value === null || value === undefined) return ['null'];
  if (typeof value === 'number') {
    if (Number.isNaN(value)) return ['number', 'NaN'];
    if (value === Infinity) return ['number', 'Inf'];
    if (value === -Infinity) return ['number', '-Inf'];
  }
  return [typeof value, value];
}

function fingerprintMatrix(matrix) {
  const rows = Array.isArray(matrix) && matrix.length ? matrix : [[null]];
  return rows.map((row) => (Array.isArray(row) ? row : [row]).map(fingerprintValue));
}

function resultFingerprint(result) {
  if (!result) return 'missing';
  return JSON.stringify({
    error: result.error || null,
    message: result.message || '',
    matrix: fingerprintMatrix(result.matrix),
    objectKind: result.objectKind || '',
    plotKind: result.plotKind || '',
    objectClass: result.objectClass || '',
    preserveObject: Boolean(result.preserveObject),
    tree: (result.tree || []).map((item) => [item.index, item.label, item.type, item.summary, item.parent]),
  });
}

function resultMayHideReferenceRelevantState(result) {
  if (!result || result.error) return false;
  if (result.preserveObject) return true;
  return ['vector', 'matrix', 'array', 'data.frame', 'list', 'function', 'environment', 'plot', 'other'].includes(result.objectKind || '');
}

function cellResultChanged(previous, current, evaluated) {
  if (resultFingerprint(previous) !== resultFingerprint(current)) return true;
  return Boolean(evaluated && resultMayHideReferenceRelevantState(current));
}

function replayCachedCell(node, cached) {
  if (!cached) {
    setCellError(node.key, '#REF!', 'Cached value is unavailable; recalculate the cell.');
    return;
  }
  if (node.literal || isPlainTextInput(node.input)) {
    placeMatrix(node, [[node.input]]);
    return;
  }
  if (cached.error) {
    setCellError(node.key, cached.error, cached.message || '');
    if (cached.plots?.length) attachPlots(node.key, cached.plots);
    return;
  }
  placeMatrix(node, cached.matrix || [[null]], cached);
  if (cached.plots?.length) attachPlots(node.key, cached.plots);
}

async function removeRGridObjects(keys) {
  const list = [...keys].filter(Boolean);
  if (!list.length || !runtime.ready) return;
  const chunkSize = 400;
  for (let index = 0; index < list.length; index += chunkSize) {
    const chunk = list.slice(index, index + chunkSize).map(rString).join(', ');
    await runtimeEvalVoid(`if (exists(".rgrid_cells", envir = .GlobalEnv, inherits = FALSE)) {\n  rm(list = intersect(ls(.GlobalEnv$.rgrid_cells, all.names = TRUE), c(${chunk})), envir = .GlobalEnv$.rgrid_cells)\n}`);
  }
}

function dependencyGraph(previousSpills) {
  const nodes = allInputCells();
  const nodeByKey = new Map(nodes.map((node) => [node.key, node]));
  const deps = new Map(nodes.map((node) => [node.key, new Set()]));

  for (const node of nodes) {
    if (node.literal) continue;
    for (const refText of extractRefCalls(node.input)) {
      let parsed;
      try {
        parsed = parseReference(refText, node.sheetId);
      } catch {
        continue;
      }
      if (parsed.dynamic) {
        const key = coordKey(parsed.sheetId, parsed.r1, parsed.c1);
        if (nodeByKey.has(key)) deps.get(node.key).add(key);
        continue;
      }
      for (let row = parsed.r1; row <= parsed.r2; row += 1) {
        for (let col = parsed.c1; col <= parsed.c2; col += 1) {
          const key = coordKey(parsed.sheetId, row, col);
          if (nodeByKey.has(key)) deps.get(node.key).add(key);
          const spillOwner = previousSpills.get(key);
          if (spillOwner && nodeByKey.has(spillOwner)) deps.get(node.key).add(spillOwner);
        }
      }
    }
  }
  return { nodes, nodeByKey, deps };
}

function topologicalOrder(graph) {
  const indegree = new Map();
  const outgoing = new Map(graph.nodes.map((node) => [node.key, new Set()]));
  for (const node of graph.nodes) {
    const validDeps = [...graph.deps.get(node.key)].filter((key) => graph.nodeByKey.has(key));
    indegree.set(node.key, validDeps.length);
    for (const dependency of validDeps) outgoing.get(dependency).add(node.key);
  }
  const ready = graph.nodes.filter((node) => indegree.get(node.key) === 0)
    .sort((a, b) => a.seq - b.seq || a.key.localeCompare(b.key));
  const order = [];
  while (ready.length) {
    const node = ready.shift();
    order.push(node);
    for (const dependentKey of outgoing.get(node.key)) {
      indegree.set(dependentKey, indegree.get(dependentKey) - 1);
      if (indegree.get(dependentKey) === 0) {
        ready.push(graph.nodeByKey.get(dependentKey));
        ready.sort((a, b) => a.seq - b.seq || a.key.localeCompare(b.key));
      }
    }
  }
  const cycleKeys = new Set(graph.nodes.filter((node) => indegree.get(node.key) > 0).map((node) => node.key));
  return { order, cycleKeys };
}

async function recalculateWorkbook(options = {}) {
  const force = Boolean(options.force);
  if (!runtime.ready) {
    refreshGridValues();
    return;
  }
  if (runtime.calculating) {
    runtime.recalcRequested = true;
    if (force) runtime.forceRecalcRequested = true;
    return;
  }
  runtime.calculating = true;
  runtime.recalcRequested = false;
  resizedSheetIds = new Set();
  setRuntimeStatus('Calculating…', 'loading');

  const previousComputed = new Map(computed);
  const previousSignatures = new Map(calculationState.cellSignatures);
  const previousDependencies = cloneDependencyMap(calculationState.dependencies);
  const forceFull = force || !calculationState.initialized;
  const currentNodes = allInputCells();
  const currentSignatures = cellSignaturesForNodes(currentNodes);
  const currentKeys = new Set(currentNodes.map((node) => node.key));
  const previousKeys = new Set([...previousSignatures.keys(), ...previousComputed.keys()]);
  const removedKeys = new Set([...previousKeys].filter((key) => !currentKeys.has(key)));
  const baseDirtyKeys = new Set();

  for (const node of currentNodes) {
    if (forceFull || previousSignatures.get(node.key) !== currentSignatures.get(node.key)) baseDirtyKeys.add(node.key);
  }

  try {
    if (forceFull) await runtimeEvalVoid('.rgrid_cells <- new.env(hash = TRUE, parent = emptyenv())');
    else await removeRGridObjects(new Set([...removedKeys, ...baseDirtyKeys]));

    let priorSpills = new Map(spillOwners);
    let finalGraph = null;
    for (let pass = 0; pass < 3; pass += 1) {
      const graph = dependencyGraph(priorSpills);
      finalGraph = graph;
      const result = topologicalOrder(graph);
      const passDirtyKeys = new Set(baseDirtyKeys);

      for (const node of graph.nodes) {
        const deps = graph.deps.get(node.key) || new Set();
        if (forceFull || dependencySetSignature(previousDependencies.get(node.key)) !== dependencySetSignature(deps)) passDirtyKeys.add(node.key);
        if (!previousComputed.has(node.key) || previousComputed.get(node.key)?.error === '#SPILL!') passDirtyKeys.add(node.key);
      }
      if (!forceFull) await removeRGridObjects(passDirtyKeys);

      clearCalculatedMaps();
      const changedValueKeys = new Set();

      for (const key of result.cycleKeys) {
        setCellError(key, '#CYCLE!', 'Circular reference');
        if (cellResultChanged(previousComputed.get(key), computed.get(key), passDirtyKeys.has(key))) changedValueKeys.add(key);
      }

      for (const node of result.order) {
        const deps = graph.deps.get(node.key) || new Set();
        const dependencyValueChanged = [...deps].some((key) => changedValueKeys.has(key));
        const previous = previousComputed.get(node.key);
        const mustEvaluate = forceFull || passDirtyKeys.has(node.key) || dependencyValueChanged || !previous;
        if (mustEvaluate) await evaluateAndPlaceCell(node);
        else replayCachedCell(node, previous);
        if (cellResultChanged(previous, computed.get(node.key), mustEvaluate)) changedValueKeys.add(node.key);
      }

      const nextGraph = dependencyGraph(spillOwners);
      finalGraph = nextGraph;
      if (graphSignature(nextGraph) === graphSignature(graph)) break;
      priorSpills = new Map(spillOwners);
    }

    calculationState.initialized = true;
    calculationState.cellSignatures = currentSignatures;
    calculationState.dependencies = cloneDependencyMap(finalGraph?.deps || new Map());
    calculationState.graphSignature = finalGraph ? graphSignature(finalGraph) : '';

    if (resizedSheetIds.has(activeSheet().id)) buildGrid();
    else {
      refreshGridValues();
      refreshSelection(true);
    }
    if (resizedSheetIds.size) scheduleSave();
    setRuntimeStatus(runtimeReadyStatusText(), 'ready');
  } catch (error) {
    console.error(error);
    setRuntimeStatus('Calculation failed', 'error');
    showMessage('Calculation failed', String(error?.stack || error));
  } finally {
    runtime.calculating = false;
    if (runtime.recalcRequested) scheduleRecalculation(0);
  }
}

function graphSignature(graph) {
  return graph.nodes
    .map((node) => `${node.key}<-${[...graph.deps.get(node.key)].sort().join(',')}`)
    .sort()
    .join('|');
}

async function evaluateAndPlaceCell(node) {
  if (node.literal || isPlainTextInput(node.input)) {
    placeMatrix(node, [[node.input]]);
    return;
  }
  let refs;
  try {
    await ensurePackagesForExpression(node.input);
    refs = new Map();
    for (const refText of extractRefCalls(node.input)) {
      const normalizedKey = refText.toLowerCase();
      if (!refs.has(normalizedKey)) refs.set(normalizedKey, resolveReferenceBinding(refText, node.sheetId));
    }
  } catch (error) {
    const normalized = normalizeFormulaError(error);
    setCellError(node.key, normalized.code, normalized.message);
    return;
  }

  const rCode = buildRCode(expressionFromInput(node.input), refs, node.key);
  let shelter;
  try {
    shelter = await new runtime.r.Shelter();
    const capture = await shelter.captureR(rCode, {
      captureGraphics: plotDeviceOptions(),
      captureStreams: true,
      captureConditions: true,
      withAutoprint: false,
    });
    const js = await capture.result.toJs();
    const packed = parsePackedResult(js);
    const plotRecords = await Promise.all((capture.images || []).map(bitmapToPlotRecord));
    if (!packed.ok) {
      const normalized = normalizeFormulaError(packed.error || 'Unknown R error');
      setCellError(node.key, normalized.code, normalized.message);
      attachPlots(node.key, plotRecords);
      return;
    }
    placeMatrix(node, packed.matrix, packed);
    attachPlots(node.key, plotRecords);
  } catch (error) {
    const normalized = normalizeFormulaError(error);
    setCellError(node.key, normalized.code, normalized.message);
  } finally {
    if (shelter) await shelter.purge();
  }
}

function attachPlots(key, plots) {
  if (!plots?.length) return;
  plotsByCell.set(key, plots);
  const result = computed.get(key);
  if (result) result.plots = plots;
}

function buildRCode(expression, refs, resultKey) {
  const branches = [...refs.entries()].map(([key, binding]) =>
    `if (identical(key, ${rString(key)})) return(${binding.code})`).join('\n      ');
  const showNames = workbook.view?.showElementNames ? 'TRUE' : 'FALSE';
  return `local({
    ref <- function(x) {
      key <- tolower(trimws(as.character(x)[1]))
      ${branches}
      stop(sprintf("#REF!: unknown reference '%s'", x), call. = FALSE)
    }
    .rgrid_one_line <- function(x, limit = 240L) {
      text <- paste(deparse(x, width.cutoff = 120L), collapse = " ")
      if (nchar(text) > limit) paste0(substr(text, 1L, limit - 1L), "…") else text
    }
    .rgrid_summary <- function(x) {
      if (inherits(x, c("gg", "ggplot", "trellis"))) return("Plot")
      if (is.function(x)) return(paste0("<function(", paste(names(formals(x)), collapse = ", "), ")>"))
      if (is.environment(x)) return("<environment>")
      if (is.list(x) && !is.data.frame(x)) return(sprintf("<list [%d]>", length(x)))
      if (!is.null(dim(x))) return(sprintf("<%s %s>", paste(class(x), collapse = "/"), paste(dim(x), collapse = "×")))
      if (length(x) == 0L) return(sprintf("<%s [0]>", paste(class(x), collapse = "/")))
      .rgrid_one_line(x)
    }
    .rgrid_tree <- function(x, max_depth = 8L, max_nodes = 800L) {
      labels <- types <- summaries <- character()
      parents <- integer()
      add <- function(value, label, parent = 0L, depth = 0L) {
        if (length(labels) >= max_nodes) return(invisible(NULL))
        labels <<- c(labels, as.character(label))
        types <<- c(types, paste(class(value), collapse = "/"))
        summaries <<- c(summaries, .rgrid_summary(value))
        parents <<- c(parents, as.integer(parent))
        current <- length(labels)
        if (depth >= max_depth) return(invisible(current))
        if (is.list(value) && !is.data.frame(value)) {
          item_names <- names(value)
          for (i in seq_along(value)) {
            child_label <- if (!is.null(item_names) && nzchar(item_names[i])) paste0("$", item_names[i]) else paste0("[[", i, "]]")
            add(value[[i]], child_label, current, depth + 1L)
          }
        }
        invisible(current)
      }
      add(x, "value")
      list(label = labels, type = types, summary = summaries, parent = parents)
    }
    .rgrid_display_matrix <- function(x) {
      show_names <- ${showNames}
      original <- x
      if (is.null(x)) return(matrix(NA, nrow = 1L, ncol = 1L))
      if (is.function(x) || is.environment(x) || (is.list(x) && !is.data.frame(x)) || (!is.atomic(x) && !is.data.frame(x))) {
        out <- matrix(.rgrid_summary(x), nrow = 1L, ncol = 1L)
      } else {
        if (inherits(x, c("Date", "POSIXct", "POSIXlt")) || is.factor(x)) x <- as.character(x)
        if (is.raw(x)) x <- as.integer(x)
        if (is.complex(x)) x <- as.character(x)
        if (is.data.frame(x)) {
          out <- as.matrix(x)
          if (show_names) {
            column_names <- names(original)
            row_names <- rownames(original)
            default_rows <- identical(row_names, as.character(seq_len(nrow(original))))
            if (!default_rows) {
              out <- cbind(row_names, out)
              column_names <- c("", column_names)
            }
            out <- rbind(column_names, out)
          }
        } else {
          dimensions <- dim(x)
          if (is.null(dimensions)) {
            out <- matrix(x, nrow = max(1L, length(x)), ncol = 1L)
            if (show_names && !is.null(names(original))) out <- cbind(names(original), out)
          } else {
            rows <- as.integer(dimensions[1L])
            cols <- as.integer(if (length(dimensions) > 1L) prod(dimensions[-1L]) else 1L)
            out <- array(x, dim = c(rows, cols))
            if (show_names) {
              dimension_names <- dimnames(original)
              column_names <- if (length(dimension_names) >= 2L) dimension_names[[2L]] else NULL
              row_names <- if (length(dimension_names) >= 1L) dimension_names[[1L]] else NULL
              if (!is.null(row_names)) out <- cbind(row_names, out)
              if (!is.null(column_names)) out <- rbind(c(if (!is.null(row_names)) "" else character(), column_names), out)
            }
          }
        }
      }
      out
    }
    .rgrid_pack <- function(x) {
      object_kind <- if (is.null(x)) "null" else if (inherits(x, c("gg", "ggplot", "trellis"))) "plot" else if (is.function(x)) "function" else if (is.data.frame(x)) "data.frame" else if (is.list(x)) "list" else if (is.environment(x)) "environment" else if (is.matrix(x)) "matrix" else if (is.array(x)) "array" else if (length(x) == 1L && is.null(dim(x))) "scalar" else if (is.atomic(x)) "vector" else "other"
      plot_kind <- if (inherits(x, c("gg", "ggplot"))) "ggplot" else if (inherits(x, "trellis")) "lattice" else ""
      object_class <- paste(class(x), collapse = " ")
      preserve_object <- is.function(x) || is.environment(x) || (is.list(x) && !is.data.frame(x)) || (length(x) <= 1L && is.null(dim(x)))
      tree <- if (is.list(x) && !is.data.frame(x)) .rgrid_tree(x) else list(label = character(), type = character(), summary = character(), parent = integer())
      displayed <- .rgrid_display_matrix(x)
      list(
        ok = TRUE,
        nrow = as.integer(nrow(displayed)),
        ncol = as.integer(ncol(displayed)),
        kind = typeof(displayed),
        object_kind = object_kind,
        plot_kind = plot_kind,
        object_class = object_class,
        preserve_object = preserve_object,
        values = as.vector(displayed),
        tree_label = tree$label,
        tree_type = tree$type,
        tree_summary = tree$summary,
        tree_parent = tree$parent
      )
    }
    tryCatch({
      .rgrid_value <- {\n${expression}\n}
      if (inherits(.rgrid_value, c("gg", "ggplot", "trellis"))) print(.rgrid_value)
      assign(${rString(resultKey)}, .rgrid_value, envir = .GlobalEnv$.rgrid_cells)
      .rgrid_pack(.rgrid_value)
    }, error = function(e) {
      call <- conditionCall(e)
      where <- if (is.null(call)) "" else paste0("\\nIn: ", paste(deparse(call, width.cutoff = 500L), collapse = " "))
      list(ok = FALSE, error = paste0(conditionMessage(e), where))
    })
  })`;
}

function rString(value) {
  return JSON.stringify(String(value)).replaceAll('\u2028', '\\u2028').replaceAll('\u2029', '\\u2029');
}

function rScalar(value) {
  if (value === null || value === undefined) return 'NA';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') {
    if (Number.isNaN(value)) return 'NaN';
    if (value === Infinity) return 'Inf';
    if (value === -Infinity) return '-Inf';
    return Number.isInteger(value) ? `${value}` : `${value}`;
  }
  return rString(value);
}

function matrixToR(matrix) {
  const rows = matrix.length || 1;
  const cols = matrix[0]?.length || 1;
  if (rows === 1 && cols === 1) return rScalar(matrix[0]?.[0] ?? null);
  const flat = [];
  for (const row of matrix) for (const value of row) flat.push(rScalar(value));
  return `matrix(c(${flat.join(',')}), nrow = ${rows}L, ncol = ${cols}L, byrow = TRUE)`;
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value === null || value === undefined) return [];
  return [value];
}

function parseDirectPackedResult(js) {
  const okValue = Array.isArray(js.ok) ? js.ok[0] : js.ok;
  if (!okValue) {
    const error = Array.isArray(js.error) ? js.error[0] : js.error;
    return { ok: false, error: String(error || 'R evaluation failed') };
  }
  const rows = Math.max(1, Number((Array.isArray(js.nrow) ? js.nrow[0] : js.nrow) || 1));
  const cols = Math.max(1, Number((Array.isArray(js.ncol) ? js.ncol[0] : js.ncol) || 1));
  const values = asArray(js.values);
  const matrix = Array.isArray(js.matrix)
    ? js.matrix
    : Array.from({ length: rows }, () => Array(cols).fill(null));
  if (!Array.isArray(js.matrix)) {
    for (let col = 0; col < cols; col += 1) {
      for (let row = 0; row < rows; row += 1) {
        matrix[row][col] = values[col * rows + row] ?? null;
      }
    }
  }
  let tree = [];
  if (Array.isArray(js.tree)) {
    tree = js.tree.map((row, index) => ({
      index: Number(row.index || index + 1),
      label: String(row.label || ''),
      type: String(row.type || ''),
      summary: String(row.summary || ''),
      parent: Number(row.parent || 0),
    }));
  } else {
    const labels = asArray(js.tree_label).map(String);
    const types = asArray(js.tree_type).map(String);
    const summaries = asArray(js.tree_summary).map(String);
    const parents = asArray(js.tree_parent).map((value) => Number(value) || 0);
    tree = labels.map((label, index) => ({
      index: index + 1,
      label,
      type: types[index] || '',
      summary: summaries[index] || '',
      parent: parents[index] || 0,
    }));
  }
  const field = (snake, camel, fallback = '') => {
    const value = js[snake] ?? js[camel] ?? fallback;
    return Array.isArray(value) ? value[0] : value;
  };
  return {
    ok: true,
    matrix,
    objectKind: String(field('object_kind', 'objectKind', 'other') || 'other'),
    plotKind: String(field('plot_kind', 'plotKind', '') || ''),
    objectClass: String(field('object_class', 'objectClass', '') || ''),
    preserveObject: Boolean(field('preserve_object', 'preserveObject', false)),
    tree,
  };
}

function parsePackedResult(js) {
  if (js && Object.prototype.hasOwnProperty.call(js, 'ok')) return parseDirectPackedResult(js);
  if (!js || js.type !== 'list' || !Array.isArray(js.names)) {
    throw new Error('The R runtime returned an unsupported result.');
  }
  const nodes = Object.fromEntries(js.names.map((name, index) => [name, js.values[index]]));
  const valuesOf = (name) => nodes[name]?.values ?? [];
  const ok = Boolean(valuesOf('ok')[0]);
  if (!ok) return { ok: false, error: String(valuesOf('error')[0] || 'R evaluation failed') };
  const rows = Math.max(1, Number(valuesOf('nrow')[0] || 1));
  const cols = Math.max(1, Number(valuesOf('ncol')[0] || 1));
  const values = valuesOf('values');
  const matrix = Array.from({ length: rows }, () => Array(cols).fill(null));
  for (let col = 0; col < cols; col += 1) {
    for (let row = 0; row < rows; row += 1) {
      matrix[row][col] = values[col * rows + row] ?? null;
    }
  }
  const labels = valuesOf('tree_label').map(String);
  const types = valuesOf('tree_type').map(String);
  const summaries = valuesOf('tree_summary').map(String);
  const parents = valuesOf('tree_parent').map((value) => Number(value) || 0);
  const tree = labels.map((label, index) => ({
    index: index + 1,
    label,
    type: types[index] || '',
    summary: summaries[index] || '',
    parent: parents[index] || 0,
  }));
  return {
    ok: true,
    matrix,
    objectKind: String(valuesOf('object_kind')[0] || 'other'),
    plotKind: String(valuesOf('plot_kind')[0] || ''),
    objectClass: String(valuesOf('object_class')[0] || ''),
    preserveObject: Boolean(valuesOf('preserve_object')[0]),
    tree,
  };
}

function normalizeFormulaError(error) {
  const message = String(error?.message || error || 'R evaluation failed').replace(/^Error:\s*/i, '').trim();
  const known = ['#SPILL!', '#CYCLE!', '#REF!', '#VALUE!', '#PKG!', '#PARSE!', '#R!'];
  let code = known.find((item) => message.includes(item));
  if (!code && /(?:unexpected|parse error|unexpected end of input|<text>:\d+:)/i.test(message)) code = '#PARSE!';
  return { code: code || '#R!', message };
}

function setCellError(key, code, message) {
  computed.set(key, { error: code, message, matrix: [[code]] });
  displayValues.set(key, code);
  spillOwners.set(key, key);
}

function ensureSheetCapacity(sheet, requiredRows, requiredCols) {
  if (requiredRows > MAX_ROWS || requiredCols > MAX_COLS) {
    throw new Error(`#SPILL!: the result needs ${requiredRows} rows × ${requiredCols} columns, above the worksheet limit of ${MAX_ROWS} × ${MAX_COLS}.`);
  }
  const nextRows = Math.max(sheet.rows, requiredRows);
  const nextCols = Math.max(sheet.cols, requiredCols);
  if (nextRows * nextCols > MAX_RENDERED_CELLS) {
    throw new Error(`#SPILL!: displaying this result would create ${nextRows * nextCols} grid cells; the current browser-safe limit is ${MAX_RENDERED_CELLS}.`);
  }
  if (nextRows !== sheet.rows || nextCols !== sheet.cols) {
    sheet.rows = nextRows;
    sheet.cols = nextCols;
    resizedSheetIds.add(sheet.id);
  }
}

function placeMatrix(node, matrix, metadata = {}) {
  const sheet = sheetById(node.sheetId);
  const anchor = parseA1Address(node.address);
  const rows = Math.max(1, matrix.length);
  const cols = Math.max(1, matrix[0]?.length || 1);
  try {
    ensureSheetCapacity(sheet, anchor.row + rows - 1, anchor.col + cols - 1);
  } catch (error) {
    const normalized = normalizeFormulaError(error);
    setCellError(node.key, normalized.code, normalized.message);
    return;
  }
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      if (r === 0 && c === 0) continue;
      const targetAddress = toA1(anchor.row + r, anchor.col + c);
      const targetKey = cellKey(node.sheetId, targetAddress);
      if (sheet.cells[targetAddress]?.input || spillOwners.has(targetKey)) {
        setCellError(node.key, '#SPILL!', `Spill range is blocked at ${targetAddress}.`);
        return;
      }
    }
  }
  computed.set(node.key, { error: null, matrix, objectKind: metadata.objectKind || 'display', plotKind: metadata.plotKind || '', objectClass: metadata.objectClass || '', preserveObject: Boolean(metadata.preserveObject), tree: metadata.tree || [] });
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const targetKey = coordKey(node.sheetId, anchor.row + r, anchor.col + c);
      spillOwners.set(targetKey, node.key);
      displayValues.set(targetKey, matrix[r]?.[c] ?? null);
    }
  }
}

function resolveReferenceBinding(refText, currentSheetId) {
  const qualified = splitQualifiedReference(String(refText));
  const localWithoutSpill = qualified.local.replace(/#\s*$/, '').trim();
  if (!qualified.sheetName) {
    const named = lookupName(localWithoutSpill);
    if (named?.kind === 'expression') return { code: `({${named.expression}})`, kind: 'expression' };
  }
  const parsed = parseReference(refText, currentSheetId);
  if (parsed.dynamic) {
    const anchorKey = coordKey(parsed.sheetId, parsed.r1, parsed.c1);
    const result = computed.get(anchorKey);
    if (!result) {
      const sheet = sheetById(parsed.sheetId);
      if (!sheet.cells[toA1(parsed.r1, parsed.c1)]) return { code: 'NA', kind: 'matrix' };
      throw new Error(`#REF!: ${refText} has not been calculated`);
    }
    if (result.error) throw new Error(`${result.error}: ${result.message || 'referenced error'}`);
    return { code: `get(${rString(anchorKey)}, envir = .GlobalEnv$.rgrid_cells, inherits = FALSE)`, kind: 'object' };
  }
  if (parsed.r1 === parsed.r2 && parsed.c1 === parsed.c2) {
    const key = coordKey(parsed.sheetId, parsed.r1, parsed.c1);
    const owner = spillOwners.get(key) || key;
    const result = computed.get(owner);
    if (owner === key && result?.preserveObject && !result.error) {
      return { code: `get(${rString(key)}, envir = .GlobalEnv$.rgrid_cells, inherits = FALSE)`, kind: 'object' };
    }
  }
  return { code: matrixToR(resolveReferenceMatrix(refText, currentSheetId)), kind: 'matrix' };
}

function resolveReferenceMatrix(refText, currentSheetId) {
  const parsed = parseReference(refText, currentSheetId);
  if (parsed.dynamic) {
    const anchorKey = coordKey(parsed.sheetId, parsed.r1, parsed.c1);
    const result = computed.get(anchorKey);
    if (!result) {
      const sheet = sheetById(parsed.sheetId);
      if (!sheet.cells[toA1(parsed.r1, parsed.c1)]) return [[null]];
      throw new Error(`#REF!: ${refText} has not been calculated`);
    }
    if (result.error) throw new Error(`${result.error}: ${result.message || 'referenced error'}`);
    return result.matrix;
  }

  const matrix = [];
  for (let row = parsed.r1; row <= parsed.r2; row += 1) {
    const outputRow = [];
    for (let col = parsed.c1; col <= parsed.c2; col += 1) {
      const key = coordKey(parsed.sheetId, row, col);
      const value = displayValues.has(key) ? displayValues.get(key) : null;
      if (isErrorValue(value)) {
        const owner = spillOwners.get(key) || key;
        const result = computed.get(owner);
        throw new Error(`${value}: ${result?.message || `referenced cell ${toA1(row, col)} contains an error`}`);
      }
      outputRow.push(value ?? null);
    }
    matrix.push(outputRow);
  }
  return matrix;
}

function renderSheetTabs() {
  const fragment = document.createDocumentFragment();
  for (const sheet of workbook.sheets) {
    const button = document.createElement('button');
    button.className = `sheet-tab${sheet.id === workbook.activeSheetId ? ' active' : ''}`;
    button.textContent = sheet.name;
    button.dataset.sheetId = sheet.id;
    button.setAttribute('role', 'tab');
    button.setAttribute('aria-selected', sheet.id === workbook.activeSheetId ? 'true' : 'false');
    fragment.appendChild(button);
  }
  els.sheetTabs.replaceChildren(fragment);
}

function activateSheet(sheetId) {
  if (!sheetById(sheetId)) return;
  endFormulaEdit();
  workbook.activeSheetId = sheetId;
  selection = { r1: 1, c1: 1, r2: 1, c2: 1 };
  renderSheetTabs();
  updateViewToggleButtons();
  buildGrid();
  scheduleSave();
}

function uniqueSheetName(base = 'Sheet') {
  const used = new Set(workbook.sheets.map((sheet) => sheet.name.toLowerCase()));
  let index = 1;
  let candidate = base;
  while (used.has(candidate.toLowerCase())) candidate = `${base}${index++}`;
  return candidate;
}

function addSheet() {
  recordHistory('add worksheet');
  const sheet = { id: uid('sheet'), name: uniqueSheetName(`Sheet${workbook.sheets.length + 1}`), rows: DEFAULT_ROWS, cols: DEFAULT_COLS, cells: {} };
  workbook.sheets.push(sheet);
  activateSheet(sheet.id);
  scheduleRecalculation();
}

function renameActiveSheet() {
  const sheet = activeSheet();
  const proposed = prompt('Worksheet name:', sheet.name);
  if (proposed === null) return;
  const name = proposed.trim();
  if (!name) return showMessage('Invalid worksheet name', 'A worksheet name cannot be empty.');
  if (workbook.sheets.some((item) => item.id !== sheet.id && item.name.toLowerCase() === name.toLowerCase())) {
    return showMessage('Invalid worksheet name', `A worksheet named "${name}" already exists.`);
  }
  if (name === sheet.name) return;
  recordHistory(`rename worksheet ${sheet.name}`);
  sheet.name = name.slice(0, 80);
  renderSheetTabs();
  scheduleSave();
  renderNamesList();
  scheduleRecalculation();
}

function deleteActiveSheet() {
  if (workbook.sheets.length === 1) return showMessage('Cannot delete worksheet', 'A workbook must contain at least one worksheet.');
  const sheet = activeSheet();
  if (!confirm(`Delete worksheet "${sheet.name}"?`)) return;
  recordHistory(`delete worksheet ${sheet.name}`);
  const index = workbook.sheets.findIndex((item) => item.id === sheet.id);
  workbook.sheets.splice(index, 1);
  for (const [name, def] of Object.entries(workbook.names)) if (def.sheetId === sheet.id) delete workbook.names[name];
  workbook.activeSheetId = workbook.sheets[Math.max(0, index - 1)].id;
  renderSheetTabs();
  buildGrid();
  scheduleSave();
  scheduleRecalculation();
}

function resizeActiveSheet(addRows, addCols) {
  const sheet = activeSheet();
  const nextRows = Math.min(MAX_ROWS, sheet.rows + addRows);
  const nextCols = Math.min(MAX_COLS, sheet.cols + addCols);
  if (nextRows === sheet.rows && nextCols === sheet.cols) return;
  recordHistory(`resize worksheet ${sheet.name}`);
  sheet.rows = nextRows;
  sheet.cols = nextCols;
  buildGrid();
  scheduleSave();
  scheduleRecalculation();
}

function displayNameDefinition(definition) {
  if (definition.kind === 'expression') return definition.expression;
  const sheet = sheetById(definition.sheetId);
  return `${quoteSheetName(sheet?.name || '?')}!${definition.ref}`;
}

function resetNameEditor() {
  editingWorkbookName = null;
  els.newNameInput.value = '';
  els.newNameRefInput.value = `${quoteSheetName(activeSheet().name)}!${rangeAddress()}`;
  els.saveNameBtn.textContent = 'Add';
}

function editWorkbookName(name) {
  const definition = workbook.names[name];
  if (!definition) return;
  editingWorkbookName = name;
  els.newNameInput.value = name;
  els.newNameRefInput.value = displayNameDefinition(definition);
  els.saveNameBtn.textContent = 'Update';
  els.newNameInput.focus();
  els.newNameInput.select();
}

function renderNamesList() {
  const entries = Object.entries(workbook.names).sort(([a], [b]) => a.localeCompare(b));
  if (!entries.length) {
    els.namesList.innerHTML = '<div class="empty-state">No workbook names have been defined.</div>';
    return;
  }
  const fragment = document.createDocumentFragment();
  for (const [name, def] of entries) {
    const row = document.createElement('div');
    row.className = 'name-row';
    const nameCode = document.createElement('code');
    nameCode.textContent = name;
    const refCode = document.createElement('code');
    refCode.textContent = displayNameDefinition(def);
    refCode.title = refCode.textContent;
    const edit = document.createElement('button');
    edit.type = 'button';
    edit.textContent = 'Edit';
    edit.className = 'edit-name-button';
    edit.dataset.editName = name;
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.textContent = 'Remove';
    remove.dataset.removeName = name;
    row.append(nameCode, refCode, edit, remove);
    fragment.appendChild(row);
  }
  els.namesList.replaceChildren(fragment);
}

async function saveNameFromDialog() {
  const name = els.newNameInput.value.trim();
  const targetText = els.newNameRefInput.value.trim();
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) || /^(?:[A-Za-z]+[1-9]\d*|R[1-9]\d*C[1-9]\d*)$/i.test(name)) {
    return showMessage('Invalid name', 'Use letters, digits, underscores, and periods. A name cannot look like a cell address.');
  }
  if (!targetText) return showMessage('Invalid name value', 'Enter a worksheet reference or an R value.');
  let definition;
  try {
    const parsed = parseReference(targetText, activeSheet().id);
    definition = {
      kind: 'range',
      sheetId: parsed.sheetId,
      ref: parsed.dynamic ? `${toA1(parsed.r1, parsed.c1)}#` : rangeFromParsed(parsed),
    };
  } catch (rangeError) {
    const expression = expressionFromInput(targetText);
    try {
      if (runtime.ready) await runtime.r.evalRVoid(`invisible(parse(text = ${rString(expression)}))`);
      definition = { kind: 'expression', expression };
    } catch (error) {
      return showMessage('Invalid name value', `${rangeError.message || rangeError}\n\nR parse error: ${error.message || error}`);
    }
  }
  recordHistory(`${editingWorkbookName ? 'edit' : 'define'} name ${name}`);
  if (editingWorkbookName && editingWorkbookName.toLowerCase() !== name.toLowerCase()) delete workbook.names[editingWorkbookName];
  setWorkbookName(name, definition);
  resetNameEditor();
  renderNamesList();
  scheduleSave();
  scheduleRecalculation();
  els.saveStatus.textContent = `Defined ${name}`;
}

function rangeFromParsed(parsed) {
  const a = toA1(parsed.r1, parsed.c1);
  const b = toA1(parsed.r2, parsed.c2);
  return a === b ? a : `${a}:${b}`;
}

function handleNameBoxCommit() {
  const text = els.nameBox.value.trim();
  if (!text) return updateFormulaBar();
  try {
    const parsed = parseReference(text, activeSheet().id);
    if (parsed.sheetId !== activeSheet().id) activateSheet(parsed.sheetId);
    selection = { r1: parsed.r1, c1: parsed.c1, r2: parsed.r2, c2: parsed.c2 };
    refreshSelection();
    scrollActiveCellIntoView();
    return;
  } catch {
    // It may be a new range name.
  }
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(text) || /^(?:[A-Za-z]+[1-9]\d*|R[1-9]\d*C[1-9]\d*)$/i.test(text)) {
    showMessage('Invalid name or address', `"${text}" is neither a valid range nor a valid workbook name.`);
    return updateFormulaBar();
  }
  recordHistory(`define name ${text}`);
  setWorkbookName(text, { sheetId: activeSheet().id, ref: rangeAddress() });
  scheduleSave();
  scheduleRecalculation();
  els.saveStatus.textContent = `Defined ${text}`;
  updateFormulaBar();
}

function workbookCalculatedMatrix(sheet) {
  let maxRow = 0;
  let maxCol = 0;
  for (let row = 1; row <= sheet.rows; row += 1) {
    for (let col = 1; col <= sheet.cols; col += 1) {
      const value = displayValues.get(coordKey(sheet.id, row, col));
      if (value !== null && value !== undefined && value !== '') {
        maxRow = Math.max(maxRow, row);
        maxCol = Math.max(maxCol, col);
      }
    }
  }
  if (maxRow === 0 || maxCol === 0) return [['']];
  return Array.from({ length: maxRow }, (_, rowIndex) =>
    Array.from({ length: maxCol }, (_, colIndex) => {
      const value = displayValues.get(coordKey(sheet.id, rowIndex + 1, colIndex + 1));
      return value === null || value === undefined ? '' : value;
    }));
}

function csvEscape(value) {
  const text = value === null || value === undefined ? '' : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function matrixToCsv(matrix) {
  return matrix.map((row) => row.map(csvEscape).join(',')).join('\r\n') + '\r\n';
}

async function exportCsvZip() {
  await ensureCalculated();
  const files = workbook.sheets.map((sheet) => ({
    name: `${safeFilename(sheet.name)}.csv`,
    data: matrixToCsv(workbookCalculatedMatrix(sheet)),
  }));
  const blob = createStoredZip(files);
  downloadBlob(blob, `${safeFilename(workbook.title)}-csv.zip`);
}

async function exportExcel() {
  await ensureCalculated();
  setRuntimeStatus('Preparing Excel export…', 'loading');
  try {
    const XLSX = await import(XLSX_MODULE_URL);
    const wb = XLSX.utils.book_new();
    const usedNames = new Set();
    for (const sheet of workbook.sheets) {
      const ws = XLSX.utils.aoa_to_sheet(workbookCalculatedMatrix(sheet));
      XLSX.utils.book_append_sheet(wb, ws, excelSheetName(sheet.name, usedNames));
    }
    const bytes = XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: true });
    downloadBlob(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${safeFilename(workbook.title)}.xlsx`);
  } catch (error) {
    showMessage('Excel export failed', `${error.message || error}\n\nThe SheetJS module is loaded from its official CDN, so the first export needs network access.`);
  } finally {
    setRuntimeStatus(runtime.ready ? 'Local R ready' : 'Local R unavailable', runtime.ready ? 'ready' : 'error');
  }
}

function excelSheetName(name, used) {
  const base = name.replace(/[\\/?*\[\]:]/g, '_').slice(0, 31) || 'Sheet';
  let candidate = base;
  let index = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = `_${index++}`;
    candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

async function ensureCalculated() {
  if (runtime.ready) await recalculateWorkbook();
}

function fileStem(filename) {
  return String(filename || 'Imported data').replace(/\.[^.]+$/, '') || 'Imported data';
}

function uniqueImportedSheetName(preferred, reserved = new Set(workbook.sheets.map((sheet) => sheet.name.toLowerCase()))) {
  const clean = String(preferred || 'Imported data').replace(/[\/?*\[\]:]/g, '_').trim().slice(0, 80) || 'Imported data';
  let candidate = clean;
  let index = 2;
  while (reserved.has(candidate.toLowerCase())) {
    const suffix = ` ${index++}`;
    candidate = `${clean.slice(0, Math.max(1, 80 - suffix.length))}${suffix}`;
  }
  reserved.add(candidate.toLowerCase());
  return candidate;
}

function trimImportedMatrix(matrix) {
  const rows = Array.isArray(matrix) ? matrix.map((row) => Array.isArray(row) ? [...row] : [row]) : [];
  while (rows.length && rows.at(-1).every((value) => value === null || value === undefined || value === '')) rows.pop();
  let lastColumn = 0;
  for (const row of rows) {
    for (let col = row.length - 1; col >= 0; col -= 1) {
      if (row[col] !== null && row[col] !== undefined && row[col] !== '') {
        lastColumn = Math.max(lastColumn, col + 1);
        break;
      }
    }
  }
  return rows.map((row) => row.slice(0, lastColumn));
}

function coerceDelimitedValue(value) {
  const original = String(value ?? '');
  if (original === '') return null;
  const text = original.trim();
  if (/^(TRUE|FALSE)$/i.test(text)) return text.toUpperCase() === 'TRUE';
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)
      && !/^[+-]?0\d/.test(text)) {
    const number = Number(text);
    if (Number.isFinite(number)) return number;
  }
  return original;
}

function parseDelimitedText(text, delimiter) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  const source = String(text || '').replace(/^\uFEFF/, '');
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (quoted) {
      if (char === '"' && source[index + 1] === '"') { field += '"'; index += 1; }
      else if (char === '"') quoted = false;
      else field += char;
      continue;
    }
    if (char === '"' && field === '') { quoted = true; continue; }
    if (char === delimiter) { row.push(coerceDelimitedValue(field)); field = ''; continue; }
    if (char === '\n' || char === '\r') {
      if (char === '\r' && source[index + 1] === '\n') index += 1;
      row.push(coerceDelimitedValue(field));
      rows.push(row);
      row = [];
      field = '';
      continue;
    }
    field += char;
  }
  if (field !== '' || row.length || !rows.length) {
    row.push(coerceDelimitedValue(field));
    rows.push(row);
  }
  return trimImportedMatrix(rows);
}

function importedValueToCell(value) {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date) return { input: value.toISOString(), literal: true };
  if (typeof value === 'string') return { input: value, literal: true };
  if (typeof value === 'boolean') return { input: value ? 'TRUE' : 'FALSE', literal: false };
  if (typeof value === 'number') {
    if (Number.isNaN(value)) return { input: 'NaN', literal: false };
    if (value === Infinity) return { input: 'Inf', literal: false };
    if (value === -Infinity) return { input: '-Inf', literal: false };
    return { input: String(value), literal: false };
  }
  return { input: String(value), literal: true };
}

function createImportedSheet(name, sourceMatrix, reservedNames) {
  const matrix = trimImportedMatrix(sourceMatrix);
  const usedRows = Math.max(1, matrix.length);
  const usedCols = Math.max(1, ...matrix.map((row) => row.length));
  if (usedRows > MAX_ROWS || usedCols > MAX_COLS || usedRows * usedCols > MAX_RENDERED_CELLS) {
    throw new Error(`${name}: ${usedRows} rows × ${usedCols} columns exceeds the worksheet safety limits.`);
  }
  let rows = Math.max(DEFAULT_ROWS, usedRows);
  let cols = Math.max(DEFAULT_COLS, usedCols);
  if (rows * cols > MAX_RENDERED_CELLS) { rows = usedRows; cols = usedCols; }
  const sheet = {
    id: uid('sheet'),
    name: uniqueImportedSheetName(name, reservedNames),
    rows,
    cols,
    cells: {},
  };
  matrix.forEach((row, rowIndex) => row.forEach((value, colIndex) => {
    const imported = importedValueToCell(value);
    if (!imported) return;
    workbook.seq += 1;
    sheet.cells[toA1(rowIndex + 1, colIndex + 1)] = {
      input: imported.input,
      seq: workbook.seq,
      ...(imported.literal ? { literal: true } : {}),
    };
  }));
  return sheet;
}

async function matricesFromDataFile(file, XLSX) {
  const extension = file.name.split('.').pop()?.toLowerCase() || '';
  if (extension === 'csv' || extension === 'tsv' || extension === 'txt') {
    const text = await file.text();
    const delimiter = extension === 'tsv' ? '\t' : (extension === 'csv' ? ',' : (text.includes('\t') ? '\t' : ','));
    return [{ name: fileStem(file.name), matrix: parseDelimitedText(text, delimiter) }];
  }
  const bytes = await file.arrayBuffer();
  const source = XLSX.read(bytes, { type: 'array', raw: true, cellDates: true, cellFormula: false });
  return source.SheetNames.map((sheetName) => ({
    name: source.SheetNames.length === 1 ? fileStem(file.name) : `${fileStem(file.name)} – ${sheetName}`,
    matrix: XLSX.utils.sheet_to_json(source.Sheets[sheetName], {
      header: 1,
      raw: true,
      defval: null,
      blankrows: true,
    }),
  }));
}

async function importDataFiles(files) {
  const selected = [...files];
  if (!selected.length) return;
  recordHistory(`import data from ${selected.map((file) => file.name).join(', ')}`);
  setRuntimeStatus(`Importing ${selected.length} data file${selected.length === 1 ? '' : 's'}…`, 'loading');
  const reservedNames = new Set(workbook.sheets.map((sheet) => sheet.name.toLowerCase()));
  const importedSheets = [];
  const failures = [];
  const excelExtensions = new Set(['xlsx', 'xls', 'xlsb', 'xlsm']);
  const needsXlsx = selected.some((file) => excelExtensions.has(file.name.split('.').pop()?.toLowerCase() || ''));
  let XLSX = null;
  try {
    if (needsXlsx) {
      try {
        XLSX = await import(XLSX_MODULE_URL);
      } catch (error) {
        throw new Error(`Could not load the SheetJS browser module needed for Excel import: ${error.message || error}`);
      }
    }
    for (const file of selected) {
      try {
        const matrices = await matricesFromDataFile(file, XLSX);
        for (const item of matrices) importedSheets.push(createImportedSheet(item.name, item.matrix, reservedNames));
      } catch (error) {
        failures.push(`${file.name}: ${error.message || error}`);
      }
    }
    if (importedSheets.length) {
      workbook.sheets.push(...importedSheets);
      workbook.activeSheetId = importedSheets[0].id;
      selection = { r1: 1, c1: 1, r2: 1, c2: 1 };
      renderSheetTabs();
      buildGrid();
      scheduleSave();
      scheduleRecalculation(0);
      els.saveStatus.textContent = `Imported ${importedSheets.length} worksheet${importedSheets.length === 1 ? '' : 's'}`;
    } else {
      history.undo.pop();
      updateHistoryButtons();
    }
    if (failures.length) showMessage(importedSheets.length ? 'Data import partially completed' : 'Data import failed', failures.join('\n'));
  } catch (error) {
    history.undo.pop();
    updateHistoryButtons();
    showMessage('Data import failed', `${error.message || error}${needsXlsx ? '\n\nExcel import loads SheetJS on demand, so the first Excel import needs network access.' : ''}`);
  } finally {
    setRuntimeStatus(runtime.ready ? 'Local R ready' : 'Local R unavailable', runtime.ready ? 'ready' : 'error');
  }
}

const RGRID_EXPORT_RUNTIME = String.raw`# ---- Executable RGrid runtime -------------------------------------------------
.rgrid <- new.env(parent = emptyenv())
.rgrid$sheets <- list()
.rgrid$names <- list()
.rgrid$active_sheet <- NULL
.rgrid$current_sheet <- NULL
.rgrid$results <- new.env(hash = TRUE, parent = emptyenv())
.rgrid$objects <- new.env(hash = TRUE, parent = emptyenv())
.rgrid$display_values <- new.env(hash = TRUE, parent = emptyenv())
.rgrid$owners <- new.env(hash = TRUE, parent = emptyenv())
.rgrid$error_values <- new.env(hash = TRUE, parent = emptyenv())

rgrid_define_sheet <- function(name, rows, cols) {
  .rgrid$sheets[[name]] <- list(rows = as.integer(rows), cols = as.integer(cols), cells = list())
  invisible(name)
}

rgrid_use_sheet <- function(name) {
  stopifnot(name %in% names(.rgrid$sheets))
  .rgrid$current_sheet <- name
  invisible(name)
}

rgrid_cell <- function(address, input, order, is_text = FALSE) {
  stopifnot(!is.null(.rgrid$current_sheet))
  address <- toupper(address)
  .rgrid$sheets[[.rgrid$current_sheet]]$cells[[address]] <- list(
    input = input,
    order = as.integer(order),
    is_text = isTRUE(is_text)
  )
  invisible(address)
}

rgrid_name <- function(name, sheet = NULL, ref = NULL, expression = NULL) {
  if (!is.null(expression)) {
    .rgrid$names[[name]] <- list(kind = "expression", expression = expression)
  } else {
    .rgrid$names[[name]] <- list(kind = "range", sheet = sheet, ref = ref)
  }
  invisible(name)
}

rgrid_set_view <- function(show_element_names = FALSE, reference_style = "A1", ...) {
  .rgrid$view <- list(
    show_element_names = isTRUE(show_element_names),
    reference_style = if (toupper(reference_style) == "R1C1") "R1C1" else "A1"
  )
  invisible(.rgrid$view)
}

rgrid_activate <- function(name) {
  .rgrid$active_sheet <- name
  invisible(name)
}

rgrid_key <- function(sheet, address) paste0(sheet, "!", toupper(address))

rgrid_col_number <- function(label) {
  chars <- utf8ToInt(toupper(label)) - utf8ToInt("A") + 1L
  if (!length(chars) || any(chars < 1L | chars > 26L)) stop("#REF!: invalid column ", label, call. = FALSE)
  Reduce(function(total, value) total * 26L + value, chars, init = 0L)
}

rgrid_col_label <- function(column) {
  column <- as.integer(column)
  out <- character()
  while (column > 0L) {
    column <- column - 1L
    out <- c(intToUtf8(utf8ToInt("A") + column %% 26L), out)
    column <- column %/% 26L
  }
  paste0(out, collapse = "")
}

rgrid_address <- function(row, col) paste0(rgrid_col_label(col), as.integer(row))

rgrid_parse_token <- function(token) {
  token <- trimws(token)
  m <- regexec("^([A-Za-z]+)([1-9][0-9]*)$", token)
  parts <- regmatches(token, m)[[1L]]
  if (length(parts)) return(list(row = as.integer(parts[3L]), col = rgrid_col_number(parts[2L])))
  m <- regexec("^[Rr]([1-9][0-9]*)[Cc]([1-9][0-9]*)$", token)
  parts <- regmatches(token, m)[[1L]]
  if (length(parts)) return(list(row = as.integer(parts[2L]), col = as.integer(parts[3L])))
  stop(sprintf("#REF!: invalid reference token '%s'", token), call. = FALSE)
}

rgrid_split_qualified <- function(text, current_sheet) {
  text <- trimws(text)
  quoted <- regexec("^'((?:[^']|'')+)'!(.+)$", text, perl = TRUE)
  parts <- regmatches(text, quoted)[[1L]]
  if (length(parts)) return(list(sheet = gsub("''", "'", parts[2L], fixed = TRUE), local = parts[3L]))
  bang <- regexpr("!", text, fixed = TRUE)[1L]
  if (bang > 1L) return(list(sheet = substr(text, 1L, bang - 1L), local = substr(text, bang + 1L, nchar(text))))
  list(sheet = current_sheet, local = text)
}

rgrid_parse_reference <- function(text, current_sheet, depth = 0L) {
  if (depth > 8L) stop("#REF!: named range recursion is too deep", call. = FALSE)
  original <- trimws(as.character(text)[1L])
  if (!nzchar(original)) stop("#REF!: empty reference", call. = FALSE)
  qualified <- rgrid_split_qualified(original, current_sheet)
  local <- trimws(qualified$local)
  dynamic <- endsWith(local, "#")
  if (dynamic) local <- trimws(substr(local, 1L, nchar(local) - 1L))
  direct <- grepl("^(?:[A-Za-z]+[1-9][0-9]*|[Rr][1-9][0-9]*[Cc][1-9][0-9]*)(?::(?:[A-Za-z]+[1-9][0-9]*|[Rr][1-9][0-9]*[Cc][1-9][0-9]*))?$", local, perl = TRUE)
  if (!direct) {
    if (!identical(qualified$sheet, current_sheet)) stop(sprintf("#REF!: invalid range '%s'", original), call. = FALSE)
    name_index <- which(tolower(names(.rgrid$names)) == tolower(local))[1L]
    if (is.na(name_index)) stop(sprintf("#REF!: unknown name or range '%s'", original), call. = FALSE)
    definition <- .rgrid$names[[name_index]]
    if (identical(definition$kind, "expression")) stop(sprintf("#REF!: name '%s' contains an R value, not a worksheet range", local), call. = FALSE)
    parsed <- rgrid_parse_reference(definition$ref, definition$sheet, depth + 1L)
    if (dynamic) parsed$dynamic <- TRUE
    parsed$source_name <- names(.rgrid$names)[name_index]
    return(parsed)
  }
  if (!qualified$sheet %in% names(.rgrid$sheets)) stop(sprintf("#REF!: unknown sheet '%s'", qualified$sheet), call. = FALSE)
  tokens <- strsplit(local, ":", fixed = TRUE)[[1L]]
  if (dynamic && length(tokens) != 1L) stop("#REF!: the # operator requires one spill anchor", call. = FALSE)
  start <- rgrid_parse_token(tokens[1L])
  end <- if (length(tokens) == 2L) rgrid_parse_token(tokens[2L]) else start
  parsed <- list(
    sheet = qualified$sheet,
    r1 = min(start$row, end$row),
    c1 = min(start$col, end$col),
    r2 = max(start$row, end$row),
    c2 = max(start$col, end$col),
    dynamic = dynamic,
    original = original
  )
  dimensions <- .rgrid$sheets[[parsed$sheet]]
  if (parsed$r1 < 1L || parsed$c1 < 1L || parsed$r2 > dimensions$rows || parsed$c2 > dimensions$cols) {
    stop(sprintf("#REF!: '%s' is outside worksheet bounds", original), call. = FALSE)
  }
  parsed
}

rgrid_summary <- function(x) {
  if (inherits(x, c("gg", "ggplot", "trellis"))) return("Plot")
  if (is.function(x)) return(paste0("<function(", paste(names(formals(x)), collapse = ", "), ")>"))
  if (is.environment(x)) return("<environment>")
  if (is.list(x) && !is.data.frame(x)) return(sprintf("<list [%d]>", length(x)))
  text <- paste(deparse(x, width.cutoff = 120L), collapse = " ")
  if (nchar(text) > 240L) paste0(substr(text, 1L, 239L), "…") else text
}

rgrid_preserve_object <- function(x) {
  is.function(x) || is.environment(x) || (is.list(x) && !is.data.frame(x)) || (length(x) <= 1L && is.null(dim(x)))
}

rgrid_normalize_result <- function(x) {
  original <- x
  show_names <- isTRUE(.rgrid$view$show_element_names)
  if (is.null(x)) return(matrix(NA, nrow = 1L, ncol = 1L))
  if (is.function(x) || is.environment(x) || (is.list(x) && !is.data.frame(x)) || (!is.atomic(x) && !is.data.frame(x))) {
    out <- matrix(rgrid_summary(x), nrow = 1L, ncol = 1L)
  } else {
    if (inherits(x, c("Date", "POSIXct", "POSIXlt")) || is.factor(x)) x <- as.character(x)
    if (is.raw(x)) x <- as.integer(x)
    if (is.complex(x)) x <- as.character(x)
    if (is.data.frame(x)) {
      out <- as.matrix(x)
      if (show_names) {
        column_names <- names(original)
        row_names <- rownames(original)
        default_rows <- identical(row_names, as.character(seq_len(nrow(original))))
        if (!default_rows) {
          out <- cbind(row_names, out)
          column_names <- c("", column_names)
        }
        out <- rbind(column_names, out)
      }
    } else {
      dimensions <- dim(x)
      if (is.null(dimensions)) {
        out <- matrix(x, nrow = max(1L, length(x)), ncol = 1L)
        if (show_names && !is.null(names(original))) out <- cbind(names(original), out)
      } else {
        rows <- as.integer(dimensions[1L])
        cols <- as.integer(if (length(dimensions) > 1L) prod(dimensions[-1L]) else 1L)
        out <- array(x, dim = c(rows, cols))
        if (show_names) {
          dimension_names <- dimnames(original)
          column_names <- if (length(dimension_names) >= 2L) dimension_names[[2L]] else NULL
          row_names <- if (length(dimension_names) >= 1L) dimension_names[[1L]] else NULL
          if (!is.null(row_names)) out <- cbind(row_names, out)
          if (!is.null(column_names)) out <- rbind(c(if (!is.null(row_names)) "" else character(), column_names), out)
        }
      }
    }
  }
  out
}

rgrid_error_code <- function(message) {
  hit <- regmatches(message, regexpr("#(?:SPILL|CYCLE|REF|VALUE|PKG|PARSE|R)!", message, perl = TRUE))
  if (length(hit) && nzchar(hit)) hit else "#R!"
}

rgrid_set_error <- function(sheet, address, code, message) {
  key <- rgrid_key(sheet, address)
  result <- list(error = code, message = message, matrix = matrix(code, nrow = 1L, ncol = 1L))
  assign(key, result, envir = .rgrid$results)
  assign(key, code, envir = .rgrid$display_values)
  assign(key, key, envir = .rgrid$owners)
  assign(key, message, envir = .rgrid$error_values)
  result
}

rgrid_place <- function(sheet, address, value) {
  address <- toupper(address)
  key <- rgrid_key(sheet, address)
  anchor <- rgrid_parse_token(address)
  original_value <- value
  value <- rgrid_normalize_result(value)
  rows <- nrow(value)
  cols <- ncol(value)
  required_rows <- anchor$row + rows - 1L
  required_cols <- anchor$col + cols - 1L
  dimensions <- .rgrid$sheets[[sheet]]
  if (required_rows > dimensions$rows || required_cols > dimensions$cols) {
    .rgrid$sheets[[sheet]]$rows <- max(dimensions$rows, required_rows)
    .rgrid$sheets[[sheet]]$cols <- max(dimensions$cols, required_cols)
  }
  for (row_offset in seq_len(rows) - 1L) {
    for (col_offset in seq_len(cols) - 1L) {
      if (row_offset == 0L && col_offset == 0L) next
      target <- rgrid_address(anchor$row + row_offset, anchor$col + col_offset)
      target_key <- rgrid_key(sheet, target)
      if (!is.null(.rgrid$sheets[[sheet]]$cells[[target]]) || exists(target_key, envir = .rgrid$owners, inherits = FALSE)) {
        return(rgrid_set_error(sheet, address, "#SPILL!", sprintf("Spill range is blocked at %s.", target)))
      }
    }
  }
  result <- list(error = NULL, message = NULL, matrix = value, object = original_value, preserve_object = rgrid_preserve_object(original_value))
  assign(key, result, envir = .rgrid$results)
  assign(key, original_value, envir = .rgrid$objects)
  for (row_offset in seq_len(rows) - 1L) {
    for (col_offset in seq_len(cols) - 1L) {
      target <- rgrid_address(anchor$row + row_offset, anchor$col + col_offset)
      target_key <- rgrid_key(sheet, target)
      assign(target_key, value[row_offset + 1L, col_offset + 1L], envir = .rgrid$display_values)
      assign(target_key, key, envir = .rgrid$owners)
    }
  }
  result
}

rgrid_scalar_or_matrix <- function(value) {
  if (length(value) == 1L && nrow(value) == 1L && ncol(value) == 1L) value[1L, 1L] else value
}

rgrid_ref <- function(text, current_sheet, stack = character()) {
  original <- trimws(as.character(text)[1L])
  qualified <- rgrid_split_qualified(original, current_sheet)
  local <- sub("#\\s*$", "", trimws(qualified$local))
  if (identical(qualified$sheet, current_sheet)) {
    name_index <- which(tolower(names(.rgrid$names)) == tolower(local))[1L]
    if (!is.na(name_index)) {
      definition <- .rgrid$names[[name_index]]
      if (identical(definition$kind, "expression")) {
        environment <- new.env(parent = .GlobalEnv)
        environment$ref <- function(x) rgrid_ref(x, current_sheet, stack)
        return(eval(parse(text = definition$expression, keep.source = TRUE), envir = environment))
      }
    }
  }
  parsed <- rgrid_parse_reference(text, current_sheet)
  if (parsed$dynamic) {
    result <- rgrid_eval_cell(parsed$sheet, rgrid_address(parsed$r1, parsed$c1), stack)
    if (!is.null(result$error)) stop(paste0(result$error, ": ", result$message), call. = FALSE)
    return(result$object)
  }
  if (parsed$r1 == parsed$r2 && parsed$c1 == parsed$c2) {
    address <- rgrid_address(parsed$r1, parsed$c1)
    key <- rgrid_key(parsed$sheet, address)
    if (!is.null(.rgrid$sheets[[parsed$sheet]]$cells[[address]]) && !exists(key, envir = .rgrid$results, inherits = FALSE)) {
      rgrid_eval_cell(parsed$sheet, address, stack)
    }
    if (exists(key, envir = .rgrid$results, inherits = FALSE)) {
      result <- get(key, envir = .rgrid$results, inherits = FALSE)
      if (isTRUE(result$preserve_object) && is.null(result$error)) return(result$object)
    }
  }
  values <- vector("list", (parsed$r2 - parsed$r1 + 1L) * (parsed$c2 - parsed$c1 + 1L))
  index <- 1L
  for (row in parsed$r1:parsed$r2) {
    for (col in parsed$c1:parsed$c2) {
      address <- rgrid_address(row, col)
      key <- rgrid_key(parsed$sheet, address)
      if (!is.null(.rgrid$sheets[[parsed$sheet]]$cells[[address]]) && !exists(key, envir = .rgrid$results, inherits = FALSE)) {
        rgrid_eval_cell(parsed$sheet, address, stack)
      }
      if (exists(key, envir = .rgrid$display_values, inherits = FALSE)) {
        value <- get(key, envir = .rgrid$display_values, inherits = FALSE)
        if (is.character(value) && length(value) == 1L && startsWith(value, "#")) {
          owner <- if (exists(key, envir = .rgrid$owners, inherits = FALSE)) get(key, envir = .rgrid$owners) else key
          owner_result <- if (exists(owner, envir = .rgrid$results, inherits = FALSE)) get(owner, envir = .rgrid$results) else NULL
          stop(paste0(value, ": ", rgrid_or(owner_result$message, sprintf("referenced cell %s contains an error", address))), call. = FALSE)
        }
        values[[index]] <- value
      } else {
        values[[index]] <- NA
      }
      index <- index + 1L
    }
  }
  combined <- do.call(c, values)
  matrix_value <- matrix(combined, nrow = parsed$r2 - parsed$r1 + 1L, ncol = parsed$c2 - parsed$c1 + 1L, byrow = TRUE)
  rgrid_scalar_or_matrix(matrix_value)
}

rgrid_or <- function(x, y) if (is.null(x) || !length(x)) y else x

rgrid_eval_cell <- function(sheet, address, stack = character()) {
  address <- toupper(address)
  key <- rgrid_key(sheet, address)
  if (exists(key, envir = .rgrid$results, inherits = FALSE)) return(get(key, envir = .rgrid$results, inherits = FALSE))
  if (key %in% stack) return(rgrid_set_error(sheet, address, "#CYCLE!", paste("Circular reference:", paste(c(stack, key), collapse = " -> "))))
  cell <- .rgrid$sheets[[sheet]]$cells[[address]]
  if (is.null(cell)) return(list(error = NULL, message = NULL, matrix = matrix(NA, nrow = 1L, ncol = 1L)))
  if (isTRUE(cell$is_text)) return(rgrid_place(sheet, address, cell$input))
  expression <- trimws(cell$input)
  if (startsWith(expression, "=")) expression <- trimws(substring(expression, 2L))
  environment <- new.env(parent = .GlobalEnv)
  environment$ref <- function(x) rgrid_ref(x, sheet, c(stack, key))
  tryCatch({
    parsed <- parse(text = expression, keep.source = TRUE)
    value <- eval(parsed, envir = environment)
    if (inherits(value, c("gg", "ggplot", "trellis"))) print(value)
    rgrid_place(sheet, address, value)
  }, error = function(error) {
    call <- conditionCall(error)
    where <- if (is.null(call)) "" else paste0("\nIn: ", paste(deparse(call, width.cutoff = 500L), collapse = " "))
    message <- paste0(conditionMessage(error), where)
    rgrid_set_error(sheet, address, rgrid_error_code(message), message)
  })
}

rgrid_reset_results <- function() {
  .rgrid$results <- new.env(hash = TRUE, parent = emptyenv())
  .rgrid$objects <- new.env(hash = TRUE, parent = emptyenv())
  .rgrid$display_values <- new.env(hash = TRUE, parent = emptyenv())
  .rgrid$owners <- new.env(hash = TRUE, parent = emptyenv())
  .rgrid$error_values <- new.env(hash = TRUE, parent = emptyenv())
}

rgrid_format_value <- function(value) {
  if (length(value) == 0L || is.null(value) || is.na(value)) return("")
  if (is.logical(value)) return(if (isTRUE(value)) "TRUE" else "FALSE")
  if (is.numeric(value)) {
    absolute <- abs(value)
    if ((absolute != 0 && absolute < 1e-7) || absolute >= 1e12) return(formatC(value, format = "e", digits = 6L, decimal.mark = "."))
    return(format(value, digits = 12L, trim = TRUE, scientific = FALSE, decimal.mark = ".", big.mark = ""))
  }
  as.character(value)
}

rgrid_sheet_values <- function(sheet, formatted = FALSE) {
  dimensions <- .rgrid$sheets[[sheet]]
  max_row <- 0L
  max_col <- 0L
  for (row in seq_len(dimensions$rows)) {
    for (col in seq_len(dimensions$cols)) {
      key <- rgrid_key(sheet, rgrid_address(row, col))
      if (exists(key, envir = .rgrid$display_values, inherits = FALSE)) {
        value <- get(key, envir = .rgrid$display_values, inherits = FALSE)
        if (!(length(value) == 0L || is.null(value) || is.na(value) || identical(value, ""))) {
          max_row <- max(max_row, row)
          max_col <- max(max_col, col)
        }
      }
    }
  }
  if (max_row == 0L || max_col == 0L) return(matrix(if (formatted) "" else list(""), nrow = 1L, ncol = 1L))
  if (formatted) {
    out <- matrix("", nrow = max_row, ncol = max_col)
  } else {
    out <- matrix(vector("list", max_row * max_col), nrow = max_row, ncol = max_col)
  }
  for (row in seq_len(max_row)) {
    for (col in seq_len(max_col)) {
      key <- rgrid_key(sheet, rgrid_address(row, col))
      value <- if (exists(key, envir = .rgrid$display_values, inherits = FALSE)) get(key, envir = .rgrid$display_values, inherits = FALSE) else NA
      if (formatted) out[row, col] <- rgrid_format_value(value) else out[row, col] <- list(value)
    }
  }
  dimnames(out) <- list(as.character(seq_len(max_row)), vapply(seq_len(max_col), rgrid_col_label, character(1L)))
  out
}

rgrid_calculate <- function(evaluation_order) {
  rgrid_reset_results()
  for (item in evaluation_order) rgrid_eval_cell(item[[1L]], item[[2L]])
  exact <- setNames(lapply(names(.rgrid$sheets), rgrid_sheet_values, formatted = FALSE), names(.rgrid$sheets))
  display <- setNames(lapply(names(.rgrid$sheets), rgrid_sheet_values, formatted = TRUE), names(.rgrid$sheets))
  errors <- as.list(.rgrid$error_values)
  objects <- as.list(.rgrid$objects)
  list(workbook = .rgrid, values = exact, display = display, objects = objects, errors = errors)
}

rgrid_install_required_packages <- function(packages) {
  packages <- unique(packages[nzchar(packages)])
  missing <- packages[!vapply(packages, requireNamespace, logical(1L), quietly = TRUE)]
  if (!length(missing)) return(invisible(character()))
  message("Installing required packages: ", paste(missing, collapse = ", "))
  repositories <- getOption("repos")
  if (!length(repositories) || identical(unname(repositories["CRAN"]), "@CRAN@")) repositories["CRAN"] <- "https://cloud.r-project.org"
  install.packages(missing, repos = repositories)
  still_missing <- missing[!vapply(missing, requireNamespace, logical(1L), quietly = TRUE)]
  if (length(still_missing)) stop("Packages remain unavailable: ", paste(still_missing, collapse = ", "), call. = FALSE)
  invisible(missing)
}
# ---- End executable RGrid runtime ---------------------------------------------`;

function executableEvaluationOrder() {
  const graph = dependencyGraph(spillOwners);
  const result = topologicalOrder(graph);
  const cycleNodes = graph.nodes.filter((node) => result.cycleKeys.has(node.key))
    .sort((a, b) => a.seq - b.seq || a.key.localeCompare(b.key));
  return [...result.order, ...cycleNodes];
}

function exportAnnotatedR() {
  const exactJson = JSON.stringify(workbook);
  const payload = utf8ToBase64(exactJson);
  const payloadLines = payload.match(/.{1,96}/g) || [''];
  const requiredPackages = [...new Set([
    ...allInputCells().flatMap((cell) => extractPackageNames(cell.input)),
    ...Object.values(workbook.names).filter((definition) => definition.kind === 'expression').flatMap((definition) => extractPackageNames(`=${definition.expression}`)),
  ])].sort();
  const evaluationOrder = executableEvaluationOrder();
  const lines = [
    '# RGrid annotated and executable workbook',
    '# Source this file in a standard R interpreter.',
    '# Required packages are installed with install.packages() when unavailable.',
    '# The rgrid object contains exact values, formatted displays, errors, formulas, names, and workbook metadata.',
    `# rgrid-format: ${FORMAT_VERSION}`,
    ...payloadLines.map((line) => `# rgrid-payload-b64: ${line}`),
    '',
    RGRID_EXPORT_RUNTIME,
    '',
    `rgrid_required_packages <- c(${requiredPackages.map(rString).join(', ')})`,
    'rgrid_install_required_packages(rgrid_required_packages)',
    '',
  ];
  for (const sheet of workbook.sheets) {
    lines.push(`rgrid_define_sheet(${rString(sheet.name)}, ${sheet.rows}L, ${sheet.cols}L)`);
  }
  lines.push(`rgrid_set_view(show_element_names = ${workbook.view?.showElementNames ? 'TRUE' : 'FALSE'}, reference_style = ${rString(currentReferenceStyle())})`);
  for (const [name, def] of Object.entries(workbook.names).sort(([a], [b]) => a.localeCompare(b))) {
    if (def.kind === 'expression') lines.push(`rgrid_name(${rString(name)}, expression = ${rString(def.expression)})`);
    else lines.push(`rgrid_name(${rString(name)}, sheet = ${rString(sheetById(def.sheetId)?.name || '')}, ref = ${rString(def.ref)})`);
  }
  lines.push('', '# Cells are registered below in workbook call order.');
  let currentSheetName = null;
  const orderedCells = allInputCells().sort((a, b) => a.seq - b.seq || a.key.localeCompare(b.key));
  for (const cell of orderedCells) {
    const sheet = sheetById(cell.sheetId);
    if (!sheet) continue;
    if (sheet.name !== currentSheetName) {
      currentSheetName = sheet.name;
      lines.push(`rgrid_use_sheet(${rString(sheet.name)})`);
    }
    lines.push(`rgrid_cell(${rString(cell.address)}, ${rString(cell.input)}, ${cell.seq}L, is_text = ${(cell.literal || isPlainTextInput(cell.input)) ? 'TRUE' : 'FALSE'})`);
  }
  lines.push(`rgrid_activate(${rString(activeSheet().name)})`, '');
  lines.push('# Calculation order is dependency-aware, with call order used as the stable tie-breaker.');
  lines.push('rgrid_evaluation_order <- list(');
  evaluationOrder.forEach((cell, index) => {
    const sheet = sheetById(cell.sheetId);
    lines.push(`  c(${rString(sheet?.name || '')}, ${rString(cell.address)})${index + 1 < evaluationOrder.length ? ',' : ''}`);
  });
  lines.push(')', '');
  lines.push('rgrid <- rgrid_calculate(rgrid_evaluation_order)');
  lines.push('# Exact per-cell R objects: rgrid$objects');
  lines.push('# Calculated grid contents: rgrid$values');
  lines.push('# Formula and spill errors: rgrid$errors');
  lines.push('# Console-friendly calculated sheets:');
  lines.push('rgrid$display', '');
  downloadBlob(new Blob([lines.join('\n')], { type: 'text/plain;charset=utf-8' }), `${safeFilename(workbook.title)}.R`);
}

function importAnnotatedR(text) {
  const chunks = [...text.matchAll(/^# rgrid-payload-b64:\s*(\S+)\s*$/gm)].map((match) => match[1]);
  if (!chunks.length) throw new Error('No RGrid payload annotation was found in the script.');
  const json = base64ToUtf8(chunks.join(''));
  const restored = normalizeWorkbook(JSON.parse(json));
  workbook = restored;
  resetCalculationState();
  selection = { r1: 1, c1: 1, r2: 1, c2: 1 };
  renderSheetTabs();
  updateViewToggleButtons();
  buildGrid();
  scheduleSave();
  scheduleRecalculation(0);
}

function utf8ToBase64(text) {
  const bytes = new TextEncoder().encode(text);
  let binary = '';
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  return btoa(binary);
}

function base64ToUtf8(base64) {
  const binary = atob(base64);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function safeFilename(name) {
  return String(name || 'workbook').replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').trim().slice(0, 120) || 'workbook';
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function createStoredZip(files) {
  const encoder = new TextEncoder();
  const now = new Date();
  const dosTime = ((now.getHours() & 0x1F) << 11) | ((now.getMinutes() & 0x3F) << 5) | ((Math.floor(now.getSeconds() / 2)) & 0x1F);
  const dosDate = (((now.getFullYear() - 1980) & 0x7F) << 9) | (((now.getMonth() + 1) & 0x0F) << 5) | (now.getDate() & 0x1F);
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const file of files) {
    const name = encoder.encode(file.name);
    const data = typeof file.data === 'string' ? encoder.encode(file.data) : file.data;
    const crc = crc32(data);
    const localHeader = new Uint8Array(30);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0x0800, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, dosTime, true);
    localView.setUint16(12, dosDate, true);
    localView.setUint32(14, crc, true);
    localView.setUint32(18, data.length, true);
    localView.setUint32(22, data.length, true);
    localView.setUint16(26, name.length, true);
    localView.setUint16(28, 0, true);
    localParts.push(localHeader, name, data);

    const centralHeader = new Uint8Array(46);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0x0800, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, dosTime, true);
    centralView.setUint16(14, dosDate, true);
    centralView.setUint32(16, crc, true);
    centralView.setUint32(20, data.length, true);
    centralView.setUint32(24, data.length, true);
    centralView.setUint16(28, name.length, true);
    centralView.setUint16(30, 0, true);
    centralView.setUint16(32, 0, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, offset, true);
    centralParts.push(centralHeader, name);
    offset += localHeader.length + name.length + data.length;
  }
  const centralOffset = offset;
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(4, 0, true);
  endView.setUint16(6, 0, true);
  endView.setUint16(8, files.length, true);
  endView.setUint16(10, files.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, centralOffset, true);
  endView.setUint16(20, 0, true);
  return new Blob([...localParts, ...centralParts, end], { type: 'application/zip' });
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xFFFFFFFF;
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function showMessage(title, body) {
  els.messageTitle.textContent = title;
  els.messageBody.textContent = body;
  els.messageDialog.showModal();
}

function setRuntimeStatus(text, state) {
  els.runtimeStatusText.textContent = text;
  const dot = els.runtimeStatus.querySelector('.status-dot');
  dot.className = `status-dot ${state}`;
  const busy = state === 'loading';
  els.busyOverlay.hidden = !busy;
  els.busyText.textContent = text;
  els.app.setAttribute('aria-busy', busy ? 'true' : 'false');
}

const RGRID_CLIPBOARD_TYPE = 'application/x-rgrid-cells+json';

function clipboardCellValue(sheet, row, col) {
  const address = toA1(row, col);
  const stored = sheet.cells[address];
  if (stored?.input !== undefined) return { input: stored.input, literal: Boolean(stored.literal), value: displayValues.get(coordKey(sheet.id, row, col)) ?? null };
  const value = displayValues.get(coordKey(sheet.id, row, col));
  return { input: null, literal: false, value: value ?? null };
}

function clipboardValueToInput(cell) {
  if (cell?.input !== null && cell?.input !== undefined) return String(cell.input);
  if (cell?.value === null || cell?.value === undefined) return '';
  if (typeof cell.value === 'boolean') return cell.value ? 'TRUE' : 'FALSE';
  return String(cell.value);
}

function handlePaste(event) {
  let rows;
  let customCells = false;
  const custom = event.clipboardData?.getData(RGRID_CLIPBOARD_TYPE);
  if (custom) {
    try {
      const parsed = JSON.parse(custom);
      if (Array.isArray(parsed.cells)) { rows = parsed.cells; customCells = true; }
    } catch { /* Fall back to tab-separated text. */ }
  }
  if (!rows) {
    const text = event.clipboardData?.getData('text/plain');
    if (text === undefined) return;
    rows = text.replace(/\r\n/g, '\n').replace(/\n$/, '').split('\n').map((line) => line.split('\t'));
  }
  event.preventDefault();
  const sheet = activeSheet();
  const operations = [];
  for (let r = 0; r < rows.length; r += 1) {
    for (let c = 0; c < rows[r].length; c += 1) {
      const row = selection.r1 + r;
      const col = selection.c1 + c;
      if (row > sheet.rows || col > sheet.cols) continue;
      const address = toA1(row, col);
      const source = customCells ? rows[r][c] : { input: rows[r][c], literal: false };
      const input = clipboardValueToInput(source);
      const literal = Boolean(customCells && source?.literal);
      const existing = sheet.cells[address];
      if ((existing?.input || '') !== input || Boolean(existing?.literal) !== literal) operations.push({ row, col, input, literal });
    }
  }
  if (!operations.length) return;
  recordHistory(`paste into ${sheet.name}!${toA1(selection.r1, selection.c1)}`);
  for (const operation of operations) setCellInput(sheet.id, operation.row, operation.col, operation.input, { defer: true, record: false, literal: operation.literal });
  selection.r2 = Math.min(sheet.rows, selection.r1 + rows.length - 1);
  selection.c2 = Math.min(sheet.cols, selection.c1 + Math.max(...rows.map((row) => row.length)) - 1);
  scheduleRecalculation(0);
  refreshSelection();
}

function handleCopy(event) {
  const sheet = activeSheet();
  const r1 = Math.min(selection.r1, selection.r2);
  const r2 = Math.max(selection.r1, selection.r2);
  const c1 = Math.min(selection.c1, selection.c2);
  const c2 = Math.max(selection.c1, selection.c2);
  const cells = [];
  const lines = [];
  for (let row = r1; row <= r2; row += 1) {
    const cellRow = [];
    const textRow = [];
    for (let col = c1; col <= c2; col += 1) {
      const cell = clipboardCellValue(sheet, row, col);
      cellRow.push(cell);
      textRow.push(clipboardValueToInput(cell));
    }
    cells.push(cellRow);
    lines.push(textRow.join('	'));
  }
  const payload = JSON.stringify({ version: 1, rows: r2 - r1 + 1, cols: c2 - c1 + 1, cells });
  try { event.clipboardData?.setData(RGRID_CLIPBOARD_TYPE, payload); } catch { /* Some browsers allow text only. */ }
  event.clipboardData?.setData('text/plain', lines.join('\n'));
  event.preventDefault();
}

function handleCut(event) {
  handleCopy(event);
  clearSelection();
}

function handleGridKeydown(event) {
  if (formulaEditorContainsTarget(event.target) || event.target === els.nameBox) return;
  const { key, shiftKey, ctrlKey, metaKey } = event;
  const command = ctrlKey || metaKey;
  if (command && key.toLowerCase() === 'z') { event.preventDefault(); shiftKey ? redoWorkbook() : undoWorkbook(); return; }
  if (command && key.toLowerCase() === 'y') { event.preventDefault(); redoWorkbook(); return; }
  if (key === 'ArrowUp') { event.preventDefault(); moveSelection(-1, 0, { extend: shiftKey, jump: command }); return; }
  if (key === 'ArrowDown') { event.preventDefault(); moveSelection(1, 0, { extend: shiftKey, jump: command }); return; }
  if (key === 'ArrowLeft') { event.preventDefault(); moveSelection(0, -1, { extend: shiftKey, jump: command }); return; }
  if (key === 'ArrowRight') { event.preventDefault(); moveSelection(0, 1, { extend: shiftKey, jump: command }); return; }
  if (key === 'Enter') { event.preventDefault(); setSelection(selection.r2 + (shiftKey ? -1 : 1), selection.c2); return; }
  if (key === 'Tab') { event.preventDefault(); setSelection(selection.r2, selection.c2 + (shiftKey ? -1 : 1)); return; }
  if (key === 'Delete' || key === 'Backspace') { event.preventDefault(); clearSelection(); return; }
  if (key === 'F9') { event.preventDefault(); scheduleRecalculation(0, { force: true }); return; }
  if (key === 'F2') {
    event.preventDefault();
    focusFormulaEditor();
    const end = formulaValue().length;
    setFormulaSelection(end, end);
    return;
  }
  if (command && ['c', 'v', 'x', 'a'].includes(key.toLowerCase())) return;
  if (!ctrlKey && !metaKey && key.length === 1) {
    event.preventDefault();
    setFormulaValue(key);
    updateFormulaSyntaxHighlight();
    focusFormulaEditor();
    setFormulaSelection(1, 1);
  }
}

function handleFormulaReferenceArrow(event) {
  if (!isFormulaReferencePicking() || !event.shiftKey || !['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) return false;
  event.preventDefault();
  const direction = {
    ArrowUp: [-1, 0],
    ArrowDown: [1, 0],
    ArrowLeft: [0, -1],
    ArrowRight: [0, 1],
  }[event.key];
  if (!formulaKeyboardPick.active) {
    const offsets = formulaSelectionOffsets();
    const caret = offsets.end ?? formulaValue().length;
    const span = findRefSpans(formulaValue()).find((item) => caret >= item.start && caret <= item.end);
    formulaKeyboardPick.active = true;
    formulaKeyboardPick.sheetId = activeSheet().id;
    formulaKeyboardPick.start = span?.start ?? caret;
    formulaKeyboardPick.end = span?.end ?? (offsets.end ?? caret);
    if (span) {
      try {
        const parsed = expandDynamicReference(parseReference(span.ref, formulaEdit.sheetId || activeSheet().id));
        if (parsed.sheetId === activeSheet().id) {
          selection = { r1: parsed.r1, c1: parsed.c1, r2: parsed.r2, c2: parsed.c2 };
        }
      } catch { /* Use the current grid selection if the existing ref is invalid. */ }
    }
  }
  moveSelection(direction[0], direction[1], { extend: true, jump: event.ctrlKey || event.metaKey });
  const reference = formulaReferenceForRange(activeSheet().id, selection.r1, selection.c1, selection.r2, selection.c2);
  const insertion = `ref(${rString(reference)})`;
  const start = formulaKeyboardPick.start;
  const end = formulaKeyboardPick.end;
  replaceFormulaRange(insertion, start, end);
  updateFormulaSyntaxHighlight();
  formulaKeyboardPick.end = start + insertion.length;
  setFormulaSelection(formulaKeyboardPick.end, formulaKeyboardPick.end);
  highlightFormulaReference(reference);
  return true;
}

function updateFormulaDragHighlight() {
  if (!dragSelection.active || dragSelection.mode !== 'formula') return;
  const reference = formulaReferenceForRange(
    dragSelection.sheetId,
    selection.r1,
    selection.c1,
    selection.r2,
    selection.c2,
  );
  highlightFormulaReference(reference);
  setFormulaMessage(`Release to insert ref(${rString(reference)})`, 'info');
}

function beginCellDrag(event, cell) {
  if (event.button !== 0) return;
  const row = Number(cell.dataset.row);
  const col = Number(cell.dataset.col);
  const formulaMode = isFormulaReferencePicking();
  resetFormulaKeyboardPick();
  event.preventDefault();
  dragSelection.active = true;
  dragSelection.mode = formulaMode ? 'formula' : 'select';
  dragSelection.sheetId = activeSheet().id;
  dragSelection.startRow = formulaMode || !event.shiftKey ? row : selection.r1;
  dragSelection.startCol = formulaMode || !event.shiftKey ? col : selection.c1;
  const formulaOffsets = formulaSelectionOffsets();
  dragSelection.formulaStart = formulaMode ? (formulaOffsets.start ?? formulaValue().length) : null;
  dragSelection.formulaEnd = formulaMode ? (formulaOffsets.end ?? dragSelection.formulaStart) : null;

  if (!formulaMode && formulaEdit.active) commitFormula(0);
  selection = {
    r1: dragSelection.startRow,
    c1: dragSelection.startCol,
    r2: row,
    c2: col,
  };
  refreshSelection();
  scrollCellIntoView(row, col);
  if (formulaMode) updateFormulaDragHighlight();
  else els.gridViewport.focus({ preventScroll: true });
}

function updateCellDrag(event) {
  if (!dragSelection.active || !(event.buttons & 1)) return;
  const target = document.elementFromPoint(event.clientX, event.clientY);
  const cell = target?.closest?.('.cell');
  if (!cell || !els.grid.contains(cell) || activeSheet().id !== dragSelection.sheetId) return;
  const row = Number(cell.dataset.row);
  const col = Number(cell.dataset.col);
  if (selection.r2 === row && selection.c2 === col) return;
  selection.r2 = row;
  selection.c2 = col;
  refreshSelection();
  scrollCellIntoView(row, col);
  if (dragSelection.mode === 'formula') updateFormulaDragHighlight();
}

function finishCellDrag() {
  if (!dragSelection.active) return;
  const mode = dragSelection.mode;
  const sheetId = dragSelection.sheetId;
  const formulaStart = dragSelection.formulaStart;
  const formulaEnd = dragSelection.formulaEnd;
  dragSelection.active = false;
  dragSelection.mode = null;
  dragSelection.sheetId = null;
  dragSelection.startRow = null;
  dragSelection.startCol = null;
  dragSelection.formulaStart = null;
  dragSelection.formulaEnd = null;
  if (mode === 'formula') {
    insertReferenceFromRange(sheetId, selection.r1, selection.c1, selection.r2, selection.c2, formulaStart, formulaEnd);
  } else {
    els.gridViewport.focus({ preventScroll: true });
  }
}

function bindEvents() {
  els.grid.addEventListener('mouseover', (event) => {
    const cell = event.target.closest('.cell');
    if (!cell || cell.contains(event.relatedTarget)) return;
    showCellTooltip(cell);
  });
  els.grid.addEventListener('mouseout', (event) => {
    const cell = event.target.closest('.cell');
    if (!cell || cell.contains(event.relatedTarget)) return;
    hideCellTooltip();
  });
  els.gridViewport.addEventListener('scroll', hideCellTooltip, { passive: true });
  els.grid.addEventListener('mousedown', (event) => {
    const cell = event.target.closest('.cell');
    if (cell) {
      beginCellDrag(event, cell);
      return;
    }
    const columnHeader = event.target.closest('.col-header');
    if (columnHeader) {
      const col = Number(columnHeader.dataset.col);
      selection = { r1: 1, c1: col, r2: activeSheet().rows, c2: col };
      refreshSelection();
      els.gridViewport.focus({ preventScroll: true });
      return;
    }
    const rowHeader = event.target.closest('.row-header');
    if (rowHeader) {
      const row = Number(rowHeader.dataset.row);
      selection = { r1: row, c1: 1, r2: row, c2: activeSheet().cols };
      refreshSelection();
      els.gridViewport.focus({ preventScroll: true });
      return;
    }
    if (event.target.closest('.corner')) {
      selection = { r1: 1, c1: 1, r2: activeSheet().rows, c2: activeSheet().cols };
      refreshSelection();
      els.gridViewport.focus({ preventScroll: true });
    }
  });
  document.addEventListener('mousemove', updateCellDrag);
  document.addEventListener('mouseup', finishCellDrag);
  window.addEventListener('blur', finishCellDrag);

  els.grid.addEventListener('dblclick', (event) => {
    const cell = event.target.closest('.cell');
    if (!cell) return;
    const key = coordKey(activeSheet().id, Number(cell.dataset.row), Number(cell.dataset.col));
    const owner = spillOwners.get(key) || key;
    if (revealPlotForCell(owner)) return;
    if (showObjectViewer(owner)) return;
    focusFormulaEditor();
    selectAllFormula();
  });
  els.gridViewport.addEventListener('keydown', handleGridKeydown);
  els.gridViewport.addEventListener('copy', handleCopy);
  els.gridViewport.addEventListener('paste', handlePaste);
  els.gridViewport.addEventListener('cut', handleCut);

  const handleFormulaEditorKeydown = (event) => {
    if (event.key === 'F1') {
      event.preventDefault();
      openFunctionHelpAtCursor();
      return;
    }
    if (!event.ctrlKey && !event.metaKey && !event.altKey && wrapSelectedFormulaText(event.key)) {
      event.preventDefault();
      return;
    }
    if (event.key === 'Enter' && event.altKey) {
      event.preventDefault();
      insertFormulaLineBreak();
      return;
    }
    if (handleFormulaReferenceArrow(event)) return;
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      commitFormula(1);
      els.gridViewport.focus();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      endFormulaEdit();
      updateFormulaBar(true);
      els.gridViewport.focus();
    }
  };

  if (formulaCodeMirror) {
    formulaCodeMirror.on('focus', beginFormulaEdit);
    formulaCodeMirror.on('change', (_editor, change) => {
      if (change.origin === 'setValue') return;
      resetFormulaKeyboardPick();
      beginFormulaEdit();
      updateFormulaReferenceHighlight();
    });
    formulaCodeMirror.on('cursorActivity', () => updateFormulaReferenceHighlight());
    formulaCodeMirror.on('keydown', (_editor, event) => handleFormulaEditorKeydown(event));
  } else {
    els.formulaInput.addEventListener('focus', beginFormulaEdit);
    els.formulaInput.addEventListener('input', () => {
      resetFormulaKeyboardPick();
      beginFormulaEdit();
      updateFormulaSyntaxHighlight();
      updateFormulaReferenceHighlight();
    });
    els.formulaInput.addEventListener('scroll', syncFormulaEditorScroll);
    for (const eventName of ['click', 'keyup', 'select']) {
      els.formulaInput.addEventListener(eventName, () => {
        updateFormulaSyntaxHighlight();
        updateFormulaReferenceHighlight();
      });
    }
    els.formulaInput.addEventListener('keydown', handleFormulaEditorKeydown);
  }
  els.commitFormulaBtn.addEventListener('click', () => commitFormula(0));
  els.nameBox.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') { event.preventDefault(); handleNameBoxCommit(); els.gridViewport.focus(); }
    if (event.key === 'Escape') { event.preventDefault(); updateFormulaBar(); els.gridViewport.focus(); }
  });

  els.sheetTabs.addEventListener('click', (event) => {
    const tab = event.target.closest('.sheet-tab');
    if (tab) activateSheet(tab.dataset.sheetId);
  });
  const replaceWorkbook = (nextWorkbook, historyLabel) => {
    recordHistory(historyLabel);
    endFormulaEdit();
    workbook = nextWorkbook;
    resetCalculationState();
    selection = { r1: 1, c1: 1, r2: 1, c2: 1 };
    renderSheetTabs();
    updateViewToggleButtons();
    buildGrid();
    scheduleSave();
    scheduleRecalculation(0);
  };
  $('#newWorkbookBtn').addEventListener('click', () => {
    if (!confirm('Create a blank workbook? The current workbook remains available only if you exported it.')) return;
    replaceWorkbook(createBlankWorkbook(), 'create new workbook');
  });
  els.exampleWorkbookBtn.addEventListener('click', () => {
    if (!confirm('Replace the current workbook with the five-sheet example workbook?')) return;
    replaceWorkbook(createExampleWorkbook(), 'load example workbook');
  });
  els.undoBtn.addEventListener('click', undoWorkbook);
  els.redoBtn.addEventListener('click', redoWorkbook);
  els.plotsBtn.addEventListener('click', () => showPlotPane(true));
  els.toggleElementNamesBtn.addEventListener('click', () => toggleWorkbookView('showElementNames', 'element names'));
  els.referenceStyleBtn.addEventListener('click', toggleReferenceStyle);
  els.themeBtn.addEventListener('click', () => applyTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'));
  els.expandObjectTreeBtn.addEventListener('click', toggleObjectTreeExpansion);
  els.objectTree.addEventListener('toggle', () => requestAnimationFrame(updateObjectTreeExpandButton), true);
  els.helpBtn.addEventListener('click', (event) => {
    if (event.ctrlKey || event.metaKey || event.shiftKey || event.altKey || event.button === 1) return;
    event.preventDefault();
    els.helpDialog.showModal();
  });
  els.formulaEditorResizer.addEventListener('pointerdown', beginFormulaEditorResize);
  els.formulaEditorResizer.addEventListener('dblclick', toggleFormulaEditorHeight);
  els.formulaEditorResizer.addEventListener('keydown', (event) => {
    if (!['ArrowUp', 'ArrowDown', 'Home'].includes(event.key)) return;
    event.preventDefault();
    if (event.key === 'Home') setFormulaEditorHeight(FORMULA_COLLAPSED_HEIGHT, { rememberExpanded: false });
    else setFormulaEditorHeight(formulaEditorHeight + (event.key === 'ArrowDown' ? 24 : -24));
  });
  els.plotSettingsBtn.addEventListener('click', () => {
    updatePlotResolutionEditor();
    els.plotSettingsDialog.showModal();
  });
  els.plotResolutionPreset.addEventListener('change', () => {
    if (els.plotResolutionPreset.value !== 'custom') {
      const [width, height] = els.plotResolutionPreset.value.split('x').map(Number);
      els.plotWidthInput.value = String(width);
      els.plotHeightInput.value = String(height);
    }
    updatePlotResolutionHelp();
  });
  for (const input of [els.plotWidthInput, els.plotHeightInput]) {
    input.addEventListener('input', updatePlotResolutionHelp);
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        applyPlotResolution();
      }
    });
  }
  els.applyPlotResolutionBtn.addEventListener('click', applyPlotResolution);
  els.closePlotPaneBtn.addEventListener('click', () => showPlotPane(false));
  els.plotPaneResizer.addEventListener('pointerdown', beginPlotPaneResize);
  els.plotPaneResizer.addEventListener('dblclick', () => setPlotPaneWidth(390));
  els.plotPaneResizer.addEventListener('keydown', (event) => {
    if (!['ArrowLeft', 'ArrowRight', 'Home'].includes(event.key)) return;
    event.preventDefault();
    if (event.key === 'Home') setPlotPaneWidth(390);
    else setPlotPaneWidth(plotPaneWidth + (event.key === 'ArrowLeft' ? 24 : -24));
  });
  window.addEventListener('resize', () => {
    hideCellTooltip();
    setPlotPaneWidth(plotPaneWidth, { persist: false });
    setFormulaEditorHeight(formulaEditorHeight, { persist: false, rememberExpanded: false });
  });
  $('#addSheetBtn').addEventListener('click', addSheet);
  $('#renameSheetBtn').addEventListener('click', renameActiveSheet);
  $('#deleteSheetBtn').addEventListener('click', deleteActiveSheet);
  $('#addRowBtn').addEventListener('click', () => resizeActiveSheet(25, 0));
  $('#addColBtn').addEventListener('click', () => resizeActiveSheet(0, 5));
  $('#exportRBtn').addEventListener('click', exportAnnotatedR);
  $('#exportCsvBtn').addEventListener('click', () => exportCsvZip().catch((error) => showMessage('CSV export failed', String(error))));
  $('#exportXlsxBtn').addEventListener('click', exportExcel);

  els.importDataBtn.addEventListener('click', () => els.importDataInput.click());
  els.importDataInput.addEventListener('change', async () => {
    const files = [...(els.importDataInput.files || [])];
    els.importDataInput.value = '';
    if (files.length) await importDataFiles(files);
  });

  $('#importScriptBtn').addEventListener('click', () => els.importScriptInput.click());
  els.importScriptInput.addEventListener('change', async () => {
    const file = els.importScriptInput.files?.[0];
    els.importScriptInput.value = '';
    if (!file) return;
    try {
      recordHistory(`import ${file.name}`);
      importAnnotatedR(await file.text());
      els.saveStatus.textContent = `Restored ${file.name}`;
    } catch (error) {
      showMessage('Import failed', String(error.message || error));
    }
  });

  $('#namesBtn').addEventListener('click', () => {
    renderNamesList();
    resetNameEditor();
    els.namesDialog.showModal();
  });
  els.saveNameBtn.addEventListener('click', () => saveNameFromDialog());
  els.namesList.addEventListener('click', (event) => {
    const edit = event.target.closest('[data-edit-name]');
    if (edit) { editWorkbookName(edit.dataset.editName); return; }
    const remove = event.target.closest('[data-remove-name]');
    if (!remove) return;
    recordHistory(`remove name ${remove.dataset.removeName}`);
    delete workbook.names[remove.dataset.removeName];
    if (editingWorkbookName === remove.dataset.removeName) resetNameEditor();
    renderNamesList(); scheduleSave(); scheduleRecalculation();
  });

  window.addEventListener('beforeunload', () => {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(workbook)); } catch { /* Ignore unload storage errors. */ }
  });
}

async function initialiseLocalR() {
  try {
    runtime.r = new LocalRClient();
    await runtime.r.init();
    runtime.ready = true;
    els.app.setAttribute('aria-busy', 'false');
    const suffix = runtime.r.version ? ` (${runtime.r.version})` : '';
    setRuntimeStatus(`Local R ready${suffix}`, 'ready');
    await recalculateWorkbook();
  } catch (error) {
    runtime.ready = false;
    els.app.setAttribute('aria-busy', 'false');
    setRuntimeStatus('Local R failed to connect', 'error');
    refreshGridValues();
    showMessage('Local R failed to connect', `${error.message || error}\n\nStart RGrid-local from R with openRGrid(), which serves the app and its local R API from localhost.`);
  }
}

initialiseFormulaCodeEditor();
applyTheme(currentTheme(), { persist: false });
setPlotPaneWidth(plotPaneWidth, { persist: false });
setFormulaEditorHeight(formulaEditorHeight, { persist: false });
updatePlotResolutionEditor();
updateViewToggleButtons();
renderSheetTabs();
buildGrid();
bindEvents();
updateFormulaSyntaxHighlight();
updateHistoryButtons();
scheduleSave();
initialiseLocalR();
