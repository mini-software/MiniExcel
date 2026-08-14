import init, { create_demo_xlsx, inspect_xlsx, version } from "./pkg/miniexcel_wasm.js";

const MAX_FILE_SIZE = 64 * 1024 * 1024;
const state = {
  bytes: null,
  demoBytes: null,
  fileName: "",
  result: null,
  activeTab: "grid",
  toastTimer: null,
};

const elements = {
  runtimeStatus: byId("runtimeStatus"),
  fileInput: byId("fileInput"),
  openFileButton: byId("openFileButton"),
  loadDemoButton: byId("loadDemoButton"),
  downloadDemoButton: byId("downloadDemoButton"),
  downloadJsonButton: byId("downloadJsonButton"),
  dropZone: byId("dropZone"),
  fileName: byId("fileName"),
  fileSize: byId("fileSize"),
  sheetSelect: byId("sheetSelect"),
  startCellInput: byId("startCellInput"),
  rowLimitInput: byId("rowLimitInput"),
  headerToggle: byId("headerToggle"),
  emptyRowsToggle: byId("emptyRowsToggle"),
  refreshButton: byId("refreshButton"),
  previewTitle: byId("previewTitle"),
  metricRows: byId("metricRows"),
  metricColumns: byId("metricColumns"),
  metricTime: byId("metricTime"),
  resultNotice: byId("resultNotice"),
  loadingState: byId("loadingState"),
  emptyState: byId("emptyState"),
  gridView: byId("gridView"),
  jsonView: byId("jsonView"),
  previewTable: byId("previewTable"),
  previewTab: byId("previewTab"),
  jsonTab: byId("jsonTab"),
  toast: byId("toast"),
};

bindEvents();
boot();

async function boot() {
  try {
    await init();
    elements.runtimeStatus.textContent = `Rust WASM v${version()}`;
    elements.runtimeStatus.classList.add("is-ready");
    await loadDemo();
  } catch (error) {
    fail(error, "WebAssembly could not start");
  }
}

function bindEvents() {
  elements.openFileButton.addEventListener("click", () => elements.fileInput.click());
  elements.dropZone.addEventListener("click", () => elements.fileInput.click());
  elements.dropZone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      elements.fileInput.click();
    }
  });
  elements.fileInput.addEventListener("change", async () => {
    const [file] = elements.fileInput.files;
    if (file) await loadFile(file);
    elements.fileInput.value = "";
  });

  for (const eventName of ["dragenter", "dragover"]) {
    elements.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropZone.classList.add("is-dragging");
    });
  }
  for (const eventName of ["dragleave", "drop"]) {
    elements.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropZone.classList.remove("is-dragging");
    });
  }
  elements.dropZone.addEventListener("drop", async (event) => {
    const [file] = event.dataTransfer.files;
    if (file) await loadFile(file);
  });

  elements.loadDemoButton.addEventListener("click", loadDemo);
  elements.downloadDemoButton.addEventListener("click", downloadDemo);
  elements.downloadJsonButton.addEventListener("click", downloadJson);
  elements.refreshButton.addEventListener("click", refreshPreview);
  elements.sheetSelect.addEventListener("change", refreshPreview);
  elements.headerToggle.addEventListener("change", refreshPreview);
  elements.emptyRowsToggle.addEventListener("change", refreshPreview);
  elements.startCellInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") refreshPreview();
  });
  elements.rowLimitInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") refreshPreview();
  });
  elements.previewTab.addEventListener("click", () => setTab("grid"));
  elements.jsonTab.addEventListener("click", () => setTab("json"));
}

async function loadFile(file) {
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    showToast("Choose an .xlsx workbook.", true);
    return;
  }
  if (file.size > MAX_FILE_SIZE) {
    showToast(`File exceeds the ${formatBytes(MAX_FILE_SIZE)} browser limit.`, true);
    return;
  }

  setLoading(`Reading ${file.name}…`);
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    await setWorkbook(bytes, file.name, file.size);
  } catch (error) {
    fail(error, "Workbook could not be read");
  }
}

async function loadDemo() {
  setLoading("Generating workbook in Rust…");
  try {
    state.demoBytes ??= new Uint8Array(create_demo_xlsx());
    await setWorkbook(
      new Uint8Array(state.demoBytes),
      "miniexcel-browser-demo.xlsx",
      state.demoBytes.byteLength,
    );
  } catch (error) {
    fail(error, "Demo workbook could not be generated");
  }
}

async function setWorkbook(bytes, name, size) {
  state.bytes = bytes;
  state.fileName = name;
  state.result = null;
  elements.fileName.textContent = name;
  elements.fileSize.textContent = formatBytes(size);
  elements.sheetSelect.replaceChildren(new Option("First worksheet", ""));
  elements.sheetSelect.disabled = true;
  elements.downloadJsonButton.disabled = true;
  await refreshPreview(true);
}

async function refreshPreview(resetSheet = false) {
  if (!state.bytes) return;
  const startCell = elements.startCellInput.value.trim().toUpperCase();
  if (!/^\$?[A-Z]{1,3}\$?[1-9]\d*$/.test(startCell)) {
    showToast("Start cell must use A1 notation, for example B2.", true);
    elements.startCellInput.focus();
    return;
  }
  const rowLimit = Number.parseInt(elements.rowLimitInput.value, 10);
  if (!Number.isInteger(rowLimit) || rowLimit < 1 || rowLimit > 2000) {
    showToast("Row limit must be between 1 and 2000.", true);
    elements.rowLimitInput.focus();
    return;
  }

  setLoading("Parsing workbook in WebAssembly…");
  await nextFrame();
  const started = performance.now();
  try {
    const options = {
      sheetName: resetSheet ? null : elements.sheetSelect.value || null,
      hasHeader: elements.headerToggle.checked,
      startCell,
      ignoreEmptyRows: elements.emptyRowsToggle.checked,
      limit: rowLimit,
    };
    const result = JSON.parse(inspect_xlsx(state.bytes, JSON.stringify(options)));
    state.result = result;
    const elapsed = performance.now() - started;
    renderResult(result, elapsed);
  } catch (error) {
    fail(error, "Query failed");
  }
}

function renderResult(result, elapsed) {
  updateSheets(result.sheetNames, result.selectedSheet);
  elements.previewTitle.textContent = `${state.fileName} · ${result.selectedSheet || "No worksheet"}`;
  elements.metricRows.textContent = String(result.totalRows);
  elements.metricColumns.textContent = String(result.columns.length);
  elements.metricTime.textContent = `${elapsed.toFixed(elapsed < 10 ? 1 : 0)} ms`;
  elements.resultNotice.textContent = result.truncated
    ? `Showing ${result.displayedRows} of ${result.totalRows}`
    : `${result.displayedRows} rows`;
  elements.downloadJsonButton.disabled = false;

  renderTable(result);
  renderJson(result);
  elements.loadingState.hidden = true;
  elements.emptyState.hidden = result.rows.length !== 0;
  setTab(state.activeTab);
}

function updateSheets(sheetNames, selectedSheet) {
  const current = elements.sheetSelect.value;
  elements.sheetSelect.replaceChildren();
  for (const sheetName of sheetNames) {
    elements.sheetSelect.add(new Option(sheetName, sheetName));
  }
  elements.sheetSelect.value = selectedSheet || current || sheetNames[0] || "";
  elements.sheetSelect.disabled = sheetNames.length === 0;
}

function renderTable(result) {
  const headRow = document.createElement("tr");
  headRow.append(createCell("th", "#", "row-index"));
  for (const column of result.columns) {
    headRow.append(createCell("th", column));
  }
  elements.previewTable.tHead.replaceChildren(headRow);

  const body = document.createDocumentFragment();
  result.rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    tr.append(createCell("td", String(rowIndex + 1), "row-index"));
    row.forEach((value, columnIndex) => {
      const cell = createCell("td", displayValue(value));
      cell.dataset.type = result.cellTypes[rowIndex][columnIndex];
      cell.title = `${result.columns[columnIndex]} · ${result.cellTypes[rowIndex][columnIndex]}`;
      tr.append(cell);
    });
    body.append(tr);
  });
  elements.previewTable.tBodies[0].replaceChildren(body);
}

function renderJson(result) {
  const rows = result.rows.map((values) =>
    Object.fromEntries(result.columns.map((column, index) => [column, values[index]])),
  );
  elements.jsonView.textContent = JSON.stringify(rows, null, 2);
}

function setTab(tab) {
  state.activeTab = tab;
  const hasRows = Boolean(state.result?.rows.length);
  const isGrid = tab === "grid";
  elements.previewTab.classList.toggle("is-active", isGrid);
  elements.previewTab.setAttribute("aria-selected", String(isGrid));
  elements.jsonTab.classList.toggle("is-active", !isGrid);
  elements.jsonTab.setAttribute("aria-selected", String(!isGrid));
  elements.gridView.hidden = !hasRows || !isGrid;
  elements.jsonView.hidden = !hasRows || isGrid;
  elements.emptyState.hidden = hasRows || !state.result;
}

function setLoading(message) {
  elements.loadingState.hidden = false;
  elements.loadingState.lastElementChild.textContent = message;
  elements.emptyState.hidden = true;
  elements.gridView.hidden = true;
  elements.jsonView.hidden = true;
  elements.resultNotice.textContent = "Working";
}

function downloadDemo() {
  try {
    state.demoBytes ??= new Uint8Array(create_demo_xlsx());
    downloadBlob(
      new Blob([state.demoBytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      "miniexcel-browser-demo.xlsx",
    );
  } catch (error) {
    fail(error, "Demo workbook could not be downloaded");
  }
}

function downloadJson() {
  if (!state.result) return;
  const rows = state.result.rows.map((values) =>
    Object.fromEntries(state.result.columns.map((column, index) => [column, values[index]])),
  );
  downloadBlob(
    new Blob([JSON.stringify(rows, null, 2)], { type: "application/json" }),
    `${state.fileName.replace(/\.xlsx$/i, "") || "miniexcel"}.json`,
  );
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function fail(error, prefix) {
  const message = error instanceof Error ? error.message : String(error);
  elements.loadingState.hidden = true;
  elements.resultNotice.textContent = "Error";
  showToast(`${prefix}: ${message}`, true);
  console.error(error);
}

function showToast(message, isError = false) {
  clearTimeout(state.toastTimer);
  elements.toast.textContent = message;
  elements.toast.classList.toggle("is-error", isError);
  elements.toast.hidden = false;
  state.toastTimer = setTimeout(() => {
    elements.toast.hidden = true;
  }, 5000);
}

function createCell(tagName, text, className) {
  const cell = document.createElement(tagName);
  cell.textContent = text;
  if (className) cell.className = className;
  return cell;
}

function displayValue(value) {
  if (value === null || value === undefined) return "—";
  if (typeof value === "boolean") return value ? "true" : "false";
  return String(value);
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes)) return "—";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
}

function nextFrame() {
  return new Promise((resolve) => requestAnimationFrame(() => resolve()));
}

function byId(id) {
  return document.getElementById(id);
}
