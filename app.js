const STORAGE_KEY = "budget-2025-card-view-v2";

const JOURNAL_SHEET_NAME = "Journalier";
const RECAP_SHEET_NAME = "Recapitulatif";
const TCD_SHEET_NAME = "TCD";

const DATE_COL = "D";
const CATEGORY_COL = "E";
const VALUE_COL = "F";
const CATEGORY_LIST_COL = "B";
const HEADER_ROW = 2;
const START_ROW = 3;

const state = {
  workbookName: "",
  workbook: null,
  mode: "idle",
  activeView: JOURNAL_SHEET_NAME,
  search: "",
  editingIndex: null,
  editorMode: "create",
  lastAction: "En attente",
  budget: createEmptyBudgetModel(),
  recap: createEmptyRecapModel(),
  recapFilters: createEmptyRecapFilters(),
};

const refs = {};
let deferredInstallPrompt = null;
let appShellReady = false;

document.addEventListener("DOMContentLoaded", () => {
  cacheDom();
  bindEvents();
  restoreDraft();
  syncLibraryState();
  setupAppShell();
  renderAll();
});

function createEmptyBudgetModel() {
  return {
    headers: ["Date", "Categories", "Value"],
    rows: [],
    categories: [],
    clearEndRow: START_ROW,
  };
}

function createEmptyRecapModel() {
  return {
    available: false,
    snapshotDate: "",
    planTemplate: [],
  };
}

function createEmptyRecapFilters() {
  return {
    year: "all",
    month: "all",
  };
}

function cacheDom() {
  refs.fileInput = document.getElementById("excel-file");
  refs.sheetSelect = document.getElementById("sheet-select");
  refs.recapYearField = document.getElementById("recap-year-field");
  refs.recapYearSelect = document.getElementById("recap-year-select");
  refs.recapMonthField = document.getElementById("recap-month-field");
  refs.recapMonthSelect = document.getElementById("recap-month-select");
  refs.searchInput = document.getElementById("search-input");
  refs.addButton = document.getElementById("add-record");
  refs.exportButton = document.getElementById("export-workbook");
  refs.cardsGrid = document.getElementById("cards-grid");
  refs.cardsEmpty = document.getElementById("cards-empty");
  refs.recapView = document.getElementById("recap-view");
  refs.form = document.getElementById("record-form");
  refs.formFields = document.getElementById("form-fields");
  refs.formTitle = document.getElementById("form-title");
  refs.formSubtitle = document.getElementById("form-subtitle");
  refs.saveButton = document.getElementById("save-record");
  refs.cancelButton = document.getElementById("cancel-edit");
  refs.recordsLabel = document.getElementById("records-label");
  refs.recordsCount = document.getElementById("records-count");
  refs.columnsLabel = document.getElementById("columns-label");
  refs.columnsCount = document.getElementById("columns-count");
  refs.activeSheet = document.getElementById("active-sheet");
  refs.lastAction = document.getElementById("last-action");
  refs.metricMode = document.getElementById("metric-mode");
  refs.metricFile = document.getElementById("metric-file");
  refs.metricSave = document.getElementById("metric-save");
  refs.installButton = document.getElementById("install-app");
  refs.appShellStatus = document.getElementById("app-shell-status");
  refs.libraryWarning = document.getElementById("library-warning");
  refs.cardsKicker = document.getElementById("cards-kicker");
  refs.cardsTitle = document.getElementById("cards-title");
  refs.cardsCaption = document.getElementById("cards-caption");
  refs.defaultEmptyMarkup = refs.cardsEmpty.innerHTML;
}

function bindEvents() {
  refs.fileInput.addEventListener("change", onFileSelected);
  refs.sheetSelect.addEventListener("change", onViewChanged);
  refs.recapYearSelect.addEventListener("change", onRecapYearChanged);
  refs.recapMonthSelect.addEventListener("change", onRecapMonthChanged);
  refs.searchInput.addEventListener("input", onSearchChanged);
  refs.addButton.addEventListener("click", startCreateMode);
  refs.exportButton.addEventListener("click", exportWorkbook);
  refs.installButton.addEventListener("click", onInstallApp);
  refs.form.addEventListener("submit", onSaveRecord);
  refs.cancelButton.addEventListener("click", resetEditor);
  refs.cardsGrid.addEventListener("click", onCardAction);
}

function syncLibraryState() {
  const ready = Boolean(window.XLSX);
  refs.libraryWarning.classList.toggle("hidden", ready);

  if (!ready) {
    setLastAction("Bibliotheque Excel indisponible");
  }
}

function setupAppShell() {
  renderAppShellState();

  window.addEventListener("online", renderAppShellState);
  window.addEventListener("offline", renderAppShellState);
  window.addEventListener("beforeinstallprompt", onBeforeInstallPrompt);
  window.addEventListener("appinstalled", onAppInstalled);

  if (window.matchMedia) {
    const standaloneMedia = window.matchMedia("(display-mode: standalone)");
    if (typeof standaloneMedia.addEventListener === "function") {
      standaloneMedia.addEventListener("change", renderAppShellState);
    } else if (typeof standaloneMedia.addListener === "function") {
      standaloneMedia.addListener(renderAppShellState);
    }
  }

  registerAppShell();
}

function onBeforeInstallPrompt(event) {
  event.preventDefault();
  deferredInstallPrompt = event;
  renderAppShellState();
}

function onAppInstalled() {
  deferredInstallPrompt = null;
  renderAppShellState();
}

async function onInstallApp() {
  if (!deferredInstallPrompt) {
    renderAppShellState();
    return;
  }

  deferredInstallPrompt.prompt();

  try {
    await deferredInstallPrompt.userChoice;
  } catch (error) {
    console.error(error);
  } finally {
    deferredInstallPrompt = null;
    renderAppShellState();
  }
}

async function registerAppShell() {
  if (!("serviceWorker" in navigator) || !window.isSecureContext) {
    renderAppShellState();
    return;
  }

  try {
    await navigator.serviceWorker.register("service-worker.js");
    await navigator.serviceWorker.ready;
    appShellReady = true;
  } catch (error) {
    console.error(error);
    appShellReady = false;
  }

  renderAppShellState();
}

function renderAppShellState() {
  if (!refs.appShellStatus || !refs.installButton) {
    return;
  }

  const standalone = isStandaloneMode();
  const onlineLabel = navigator.onLine ? "En ligne" : "Hors ligne";
  const installVisible = standalone || Boolean(deferredInstallPrompt);

  refs.installButton.classList.toggle("hidden", !installVisible);
  refs.installButton.disabled = !deferredInstallPrompt;
  refs.installButton.textContent = standalone ? "App installee" : "Installer l'app";

  if (standalone) {
    refs.appShellStatus.textContent = `${onlineLabel} - Mode app actif${appShellReady ? " - Cache hors ligne pret" : ""}`;
    return;
  }

  if (location.protocol === "file:") {
    refs.appShellStatus.textContent = `${onlineLabel} - Mode local - Publiez l'app en HTTPS pour l'installer`;
    return;
  }

  if (!window.isSecureContext) {
    refs.appShellStatus.textContent = `${onlineLabel} - HTTPS requis pour l'installation`;
    return;
  }

  if (deferredInstallPrompt) {
    refs.appShellStatus.textContent = `${onlineLabel} - Installation disponible${appShellReady ? " - Cache hors ligne pret" : ""}`;
    return;
  }

  if (isAppleMobileDevice()) {
    refs.appShellStatus.textContent = `${onlineLabel} - iPhone/iPad: utilisez Partager puis Sur l'ecran d'accueil`;
    return;
  }

  refs.appShellStatus.textContent = `${onlineLabel} - L'installation sera proposee quand l'app sera prete${appShellReady ? " - Cache hors ligne pret" : ""}`;
}

function isStandaloneMode() {
  return window.matchMedia?.("(display-mode: standalone)")?.matches || window.navigator.standalone === true;
}

function isAppleMobileDevice() {
  return /iphone|ipad|ipod/i.test(window.navigator.userAgent || "");
}

function restoreDraft() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return;
    }

    const draft = JSON.parse(raw);
    if (draft.mode !== "budget" || !Array.isArray(draft.rows)) {
      return;
    }

    state.mode = "budget";
    state.workbookName = draft.workbookName || "";
    state.activeView = draft.activeView === RECAP_SHEET_NAME ? RECAP_SHEET_NAME : JOURNAL_SHEET_NAME;
    state.budget = {
      headers: ["Date", "Categories", "Value"],
      categories: Array.isArray(draft.categories) ? draft.categories : [],
      rows: draft.rows.map((row) => sanitizeBudgetRow(row)),
      clearEndRow: Number.isFinite(draft.clearEndRow) ? draft.clearEndRow : START_ROW,
    };
    state.recap = {
      available: Boolean(draft.recap?.available),
      snapshotDate: String(draft.recap?.snapshotDate || ""),
      planTemplate: Array.isArray(draft.recap?.planTemplate)
        ? draft.recap.planTemplate.map((row) => ({
            label: String(row?.label || ""),
            plan: normalizeAmountValue(row?.plan),
          }))
        : [],
    };
    state.recapFilters = {
      year: String(draft.recapFilters?.year || "all"),
      month: String(draft.recapFilters?.month || "all"),
    };
    state.lastAction = "Brouillon restaure. Rechargez le fichier pour exporter.";
  } catch (error) {
    console.error(error);
  }
}

function persistDraft() {
  if (state.mode !== "budget") {
    return;
  }

  const payload = {
    mode: state.mode,
    workbookName: state.workbookName,
    activeView: state.activeView,
    categories: state.budget.categories,
    rows: state.budget.rows,
    clearEndRow: state.budget.clearEndRow,
    recap: {
      available: state.recap.available,
      snapshotDate: state.recap.snapshotDate,
      planTemplate: state.recap.planTemplate,
    },
    recapFilters: state.recapFilters,
    savedAt: new Date().toISOString(),
  };

  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

async function onFileSelected(event) {
  const [file] = event.target.files || [];

  if (!file) {
    return;
  }

  if (!window.XLSX) {
    setLastAction("Import impossible: bibliotheque Excel absente");
    renderStats();
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      cellFormula: true,
      cellNF: true,
    });

    if (!isBudgetWorkbook(workbook)) {
      state.workbookName = file.name;
      state.workbook = null;
      state.mode = "idle";
      state.activeView = JOURNAL_SHEET_NAME;
      state.search = "";
      state.editingIndex = null;
      state.editorMode = "create";
      state.budget = createEmptyBudgetModel();
      state.recap = createEmptyRecapModel();
      state.recapFilters = createEmptyRecapFilters();
      refs.searchInput.value = "";
      refs.fileInput.value = "";
      setLastAction("Ce prototype attend Budget_2025 Final.xlsx");
      renderAll();
      return;
    }

    state.workbookName = file.name;
    state.workbook = workbook;
    state.mode = "budget";
    state.activeView = JOURNAL_SHEET_NAME;
    state.search = "";
    state.editingIndex = null;
    state.editorMode = "create";
    state.budget = parseBudgetWorkbook(workbook);
    state.recap = parseRecapWorkbook(workbook);

    refs.searchInput.value = "";
    refs.fileInput.value = "";

    persistDraft();
    setLastAction(`Classeur charge: ${file.name}`);
    renderAll();
  } catch (error) {
    console.error(error);
    state.workbook = null;
    setLastAction("Le fichier n'a pas pu etre lu");
    renderAll();
  }
}

function isBudgetWorkbook(workbook) {
  const sheet = workbook?.Sheets?.[JOURNAL_SHEET_NAME];
  if (!sheet) {
    return false;
  }

  const d2 = normalizeHeaderName(readCellText(sheet[`${DATE_COL}${HEADER_ROW}`]));
  const e2 = normalizeHeaderName(readCellText(sheet[`${CATEGORY_COL}${HEADER_ROW}`]));
  const f2 = normalizeHeaderName(readCellText(sheet[`${VALUE_COL}${HEADER_ROW}`]));
  const b2 = normalizeHeaderName(readCellText(sheet[`${CATEGORY_LIST_COL}${HEADER_ROW}`]));

  return d2 === "date" && e2.startsWith("categorie") && f2 === "value" && b2.startsWith("categorie");
}

function parseBudgetWorkbook(workbook) {
  const sheet = workbook.Sheets[JOURNAL_SHEET_NAME];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const clearEndRow = range.e.r + 1;

  const categories = parseBudgetCategories(sheet, clearEndRow);
  const rows = parseBudgetRows(sheet, clearEndRow);
  sortBudgetRowsInPlace(rows);

  return {
    headers: [
      readCellText(sheet[`${DATE_COL}${HEADER_ROW}`]) || "Date",
      readCellText(sheet[`${CATEGORY_COL}${HEADER_ROW}`]) || "Categories",
      readCellText(sheet[`${VALUE_COL}${HEADER_ROW}`]) || "Value",
    ],
    rows,
    categories,
    clearEndRow,
  };
}

function parseBudgetCategories(sheet, maxRow) {
  const categories = [];

  for (let row = START_ROW; row <= maxRow; row += 1) {
    const value = String(readCellText(sheet[`${CATEGORY_LIST_COL}${row}`]) || "").trim();

    if (!value) {
      if (categories.length) {
        break;
      }
      continue;
    }

    if (!categories.includes(value)) {
      categories.push(value);
    }
  }

  return categories;
}

function parseBudgetRows(sheet, maxRow) {
  const rows = [];

  for (let row = START_ROW; row <= maxRow; row += 1) {
    const dateCell = sheet[`${DATE_COL}${row}`];
    const categoryCell = sheet[`${CATEGORY_COL}${row}`];
    const valueCell = sheet[`${VALUE_COL}${row}`];

    if (!dateCell && !categoryCell && !valueCell) {
      continue;
    }

    const dateValue = normalizeDateValue(readCellRawValue(dateCell));
    const categoryValue = String(readCellText(categoryCell) || "").trim();
    const amountValue = normalizeAmountValue(readCellRawValue(valueCell));

    if (!dateValue && !categoryValue && !amountValue) {
      continue;
    }

    rows.push({
      __id: createId(),
      Date: dateValue,
      Categories: categoryValue,
      Value: amountValue,
    });
  }

  return rows;
}

function parseRecapWorkbook(workbook) {
  const recapSheet = workbook?.Sheets?.[RECAP_SHEET_NAME];
  const tcdSheet = workbook?.Sheets?.[TCD_SHEET_NAME];

  if (!tcdSheet) {
    return createEmptyRecapModel();
  }

  return {
    available: true,
    snapshotDate: formatDateForDisplay(readCellRawValue(recapSheet?.BJ2)),
    planTemplate: parseRecapPlanTemplate(tcdSheet),
  };
}

function parseRecapPlanTemplate(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const templateMap = new Map();

  for (let row = 4; row <= range.e.r + 1; row += 1) {
    const label = String(readCellText(sheet[`H${row}`]) || "").trim();
    if (!label || isIgnoredRecapLabel(label)) {
      continue;
    }

    const normalized = normalizeHeaderName(label);
    const plan = normalizeAmountValue(readCellRawValue(sheet[`I${row}`]));
    const hasPlanValue = plan !== "";
    const existing = templateMap.get(normalized);

    if (!existing || (!existing.plan && hasPlanValue)) {
      templateMap.set(normalized, { label, plan });
    }
  }

  return Array.from(templateMap.values());
}

function isIgnoredRecapLabel(label) {
  const normalized = normalizeHeaderName(label);
  return normalized === "sol" || normalized === "expenses";
}

function onViewChanged(event) {
  state.activeView = event.target.value === RECAP_SHEET_NAME ? RECAP_SHEET_NAME : JOURNAL_SHEET_NAME;
  state.search = "";
  refs.searchInput.value = "";
  state.editorMode = "create";
  state.editingIndex = null;
  persistDraft();
  renderAll();
}

function onRecapYearChanged(event) {
  state.recapFilters.year = String(event.target.value || "all");

  const availableMonths = getAvailableRecapMonths(state.recapFilters.year);
  if (state.recapFilters.month !== "all" && !availableMonths.includes(state.recapFilters.month)) {
    state.recapFilters.month = "all";
  }

  persistDraft();
  renderAll();
}

function onRecapMonthChanged(event) {
  state.recapFilters.month = String(event.target.value || "all");
  persistDraft();
  renderAll();
}

function onSearchChanged(event) {
  state.search = event.target.value.trim().toLowerCase();
  renderCards();
  renderStats();
}

function startCreateMode() {
  state.editorMode = "create";
  state.editingIndex = null;
  renderAll();
}

function resetEditor() {
  state.editorMode = "create";
  state.editingIndex = null;
  renderAll();
}

function onCardAction(event) {
  if (state.activeView !== JOURNAL_SHEET_NAME) {
    return;
  }

  const deleteButton = event.target.closest("[data-action='delete']");
  const editButton = event.target.closest("[data-action='edit']");
  const card = event.target.closest(".data-card");

  if (!card) {
    return;
  }

  const index = Number(card.dataset.index);
  if (Number.isNaN(index)) {
    return;
  }

  if (deleteButton) {
    deleteRecord(index);
    return;
  }

  if (editButton || card) {
    openEditor(index);
  }
}

function openEditor(index) {
  if (!state.budget.rows[index]) {
    return;
  }

  state.editorMode = "edit";
  state.editingIndex = index;
  renderAll();
}

function deleteRecord(index) {
  const target = state.budget.rows[index];
  if (!target) {
    return;
  }

  const title = target.Categories || `transaction ${index + 1}`;
  const confirmed = window.confirm(`Supprimer ${title} ?`);

  if (!confirmed) {
    return;
  }

  state.budget.rows.splice(index, 1);

  if (state.editingIndex === index) {
    state.editingIndex = null;
    state.editorMode = "create";
  } else if (state.editingIndex !== null && state.editingIndex > index) {
    state.editingIndex -= 1;
  }

  persistDraft();
  setLastAction(`Transaction supprimee: ${title}`);
  renderAll();
}

function onSaveRecord(event) {
  event.preventDefault();

  if (state.mode !== "budget") {
    return;
  }

  const formData = new FormData(refs.form);
  const nextRecord = {
    __id: createId(),
    Date: normalizeDateValue(formData.get("Date")),
    Categories: String(formData.get("Categories") ?? "").trim(),
    Value: normalizeAmountValue(formData.get("Value")),
  };

  if (!nextRecord.Date && !nextRecord.Categories && !nextRecord.Value) {
    setLastAction("Transaction vide ignoree");
    renderStats();
    return;
  }

  if (state.editorMode === "edit" && state.editingIndex !== null && state.budget.rows[state.editingIndex]) {
    nextRecord.__id = state.budget.rows[state.editingIndex].__id;
    state.budget.rows[state.editingIndex] = nextRecord;
    setLastAction("Transaction mise a jour");
  } else {
    state.budget.rows.push(nextRecord);
    setLastAction("Nouvelle transaction ajoutee");
  }

  sortBudgetRowsInPlace(state.budget.rows);
  state.editingIndex = state.budget.rows.findIndex((row) => row.__id === nextRecord.__id);
  state.editorMode = "edit";

  persistDraft();
  renderAll();
}

async function exportWorkbook() {
  if (!window.XLSX) {
    setLastAction("Export impossible: bibliotheque Excel absente");
    renderStats();
    return;
  }

  if (!state.workbook || state.mode !== "budget") {
    setLastAction("Rechargez Budget_2025 Final.xlsx pour exporter le vrai classeur");
    renderStats();
    return;
  }

  try {
    applyBudgetRowsToWorkbook(state.workbook, state.budget);
    const exportFileName = buildExportFileName();

    if (canUseNativeExcelExport()) {
      setLastAction("Preparation du fichier Excel pour l'app mobile");
      renderStats();
      await exportWorkbookWithNativeShare(state.workbook, exportFileName);
      setLastAction("Classeur exporte et partage depuis l'app mobile");
    } else {
      XLSX.writeFile(state.workbook, exportFileName);
      setLastAction("Classeur exporte avec Journalier mis a jour");
    }

    renderStats();
  } catch (error) {
    console.error(error);
    setLastAction(buildExportErrorMessage(error));
    renderStats();
  }
}

function buildExportFileName() {
  const baseName = state.workbookName
    ? state.workbookName.replace(/\.(xlsx|xls)$/i, "")
    : "Budget_2025 Final";

  return `${sanitizeExportFileName(baseName)}-card-view.xlsx`;
}

function sanitizeExportFileName(value) {
  const normalized = String(value || "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-")
    .replace(/\s+/g, " ")
    .trim();

  return normalized || "Budget_2025 Final";
}

function getCapacitorRuntime() {
  return window.Capacitor || null;
}

function getFilesystemPlugin() {
  return window.capacitorFilesystemPluginCapacitor?.Filesystem || null;
}

function getFilesystemDirectory() {
  return window.capacitorFilesystemPluginCapacitor?.FilesystemDirectory || null;
}

function getSharePlugin() {
  return window.capacitorShare?.Share || null;
}

function isNativeAppRuntime() {
  return Boolean(getCapacitorRuntime()?.isNativePlatform?.());
}

function canUseNativeExcelExport() {
  return Boolean(
    isNativeAppRuntime() &&
      getFilesystemPlugin() &&
      getSharePlugin() &&
      getFilesystemDirectory()?.Cache
  );
}

async function exportWorkbookWithNativeShare(workbook, fileName) {
  const Filesystem = getFilesystemPlugin();
  const Share = getSharePlugin();
  const Directory = getFilesystemDirectory();
  const shareSupport = typeof Share.canShare === "function" ? await Share.canShare() : { value: true };

  if (!shareSupport?.value) {
    throw new Error("Le partage natif n'est pas disponible sur cet appareil");
  }

  const relativePath = `exports/${fileName}`;
  const workbookData = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "base64",
  });

  await Filesystem.writeFile({
    path: relativePath,
    data: workbookData,
    directory: Directory.Cache,
    recursive: true,
  });

  const fileUri = await Filesystem.getUri({
    path: relativePath,
    directory: Directory.Cache,
  });

  await Share.share({
    title: "Budget 2025",
    text: "Classeur Budget 2025 exporte depuis l'app mobile.",
    files: [fileUri.uri],
    dialogTitle: "Partager le classeur Excel",
  });
}

function buildExportErrorMessage(error) {
  const rawMessage = String(error?.message || "").toLowerCase();
  if (rawMessage.includes("share")) {
    return "L'export a echoue: partage natif indisponible";
  }

  return "L'export a echoue";
}

function applyBudgetRowsToWorkbook(workbook, budgetModel) {
  const sheet = workbook.Sheets[JOURNAL_SHEET_NAME];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const clearEndRow = Math.max(budgetModel.clearEndRow, START_ROW + budgetModel.rows.length + 4);

  for (let row = START_ROW; row <= clearEndRow; row += 1) {
    delete sheet[`${DATE_COL}${row}`];
    delete sheet[`${CATEGORY_COL}${row}`];
    delete sheet[`${VALUE_COL}${row}`];
  }

  budgetModel.rows.forEach((row, index) => {
    const sheetRow = START_ROW + index;
    const isoDate = normalizeDateValue(row.Date);
    const amountNumber = parseAmount(row.Value);

    if (isoDate) {
      sheet[`${DATE_COL}${sheetRow}`] = {
        t: "n",
        v: isoDateToExcelSerial(isoDate),
        z: "m/d/yyyy",
      };
    }

    if (row.Categories) {
      sheet[`${CATEGORY_COL}${sheetRow}`] = {
        t: "s",
        v: row.Categories,
      };
    }

    if (Number.isFinite(amountNumber)) {
      sheet[`${VALUE_COL}${sheetRow}`] = {
        t: "n",
        v: amountNumber,
      };
    } else if (row.Value) {
      sheet[`${VALUE_COL}${sheetRow}`] = {
        t: "s",
        v: row.Value,
      };
    }
  });

  range.e.r = Math.max(range.e.r, START_ROW + budgetModel.rows.length + 2);
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function renderAll() {
  renderSheetOptions();
  syncRecapFilters();
  renderSectionHeading();
  renderStats();
  renderCards();
  renderEditor();
  renderControls();
  renderAppShellState();
}

function syncRecapFilters() {
  const availableYears = getAvailableRecapYears();

  if (!availableYears.length) {
    state.recapFilters = createEmptyRecapFilters();
    renderRecapFilterOptions([], []);
    return;
  }

  if (state.recapFilters.year !== "all" && !availableYears.includes(state.recapFilters.year)) {
    state.recapFilters.year = "all";
  }

  const availableMonths = getAvailableRecapMonths(state.recapFilters.year);
  if (state.recapFilters.month !== "all" && !availableMonths.includes(state.recapFilters.month)) {
    state.recapFilters.month = "all";
  }

  renderRecapFilterOptions(availableYears, availableMonths);
}

function renderSheetOptions() {
  refs.sheetSelect.innerHTML = "";

  if (state.mode !== "budget") {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "Aucune feuille";
    option.selected = true;
    refs.sheetSelect.appendChild(option);
    return;
  }

  [JOURNAL_SHEET_NAME, RECAP_SHEET_NAME].forEach((viewName) => {
    const option = document.createElement("option");
    option.value = viewName;
    option.textContent = viewName;
    option.selected = viewName === state.activeView;
    refs.sheetSelect.appendChild(option);
  });
}

function renderRecapFilterOptions(availableYears, availableMonths) {
  refs.recapYearSelect.innerHTML = "";
  appendSelectOption(refs.recapYearSelect, "all", "Toutes les annees");
  availableYears.forEach((year) => appendSelectOption(refs.recapYearSelect, year, year));
  refs.recapYearSelect.value = state.recapFilters.year;

  refs.recapMonthSelect.innerHTML = "";
  appendSelectOption(refs.recapMonthSelect, "all", "Tous les mois");
  availableMonths.forEach((month) => {
    appendSelectOption(refs.recapMonthSelect, month, formatMonthLabel(month));
  });
  refs.recapMonthSelect.value = state.recapFilters.month;
}

function appendSelectOption(select, value, label) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = label;
  select.appendChild(option);
}

function renderSectionHeading() {
  if (state.activeView === RECAP_SHEET_NAME) {
    refs.cardsKicker.textContent = "Recapitulatif";
    refs.cardsTitle.textContent = "Vue recap du budget";
    refs.cardsCaption.textContent =
      "Synthese reconstruite depuis TCD et vos transactions Journalier, avec un filtre par annee et par mois pour comparer les periodes disponibles.";
    return;
  }

  refs.cardsKicker.textContent = "Journalier";
  refs.cardsTitle.textContent = "Les ecritures deviennent des fiches";
  refs.cardsCaption.textContent =
    "Cette vue utilise Journalier!D:F, garde la liste de categories de Journalier!B et reprend le meme filtre annee/mois que le recapitulatif.";
}

function renderControls() {
  const hasBudget = state.mode === "budget";
  const journalActive = hasBudget && state.activeView === JOURNAL_SHEET_NAME;
  const recapActive = hasBudget && state.activeView === RECAP_SHEET_NAME;
  const availableYears = getAvailableRecapYears();
  const availableMonths = getAvailableRecapMonths(state.recapFilters.year);

  refs.sheetSelect.disabled = !hasBudget;
  refs.searchInput.disabled = !hasBudget;
  refs.searchInput.placeholder = recapActive
    ? "Chercher un poste ou une categorie du recap..."
    : "Categorie, date, valeur...";
  refs.recapYearField.classList.toggle("hidden", !hasBudget);
  refs.recapMonthField.classList.toggle("hidden", !hasBudget);
  refs.recapYearSelect.disabled = !hasBudget || !availableYears.length;
  refs.recapMonthSelect.disabled = !hasBudget || !availableMonths.length;
  refs.addButton.disabled = !journalActive;
  refs.exportButton.disabled = !hasBudget || !state.workbook || !window.XLSX;
  refs.saveButton.disabled = !journalActive;
  refs.cancelButton.disabled = !journalActive;
}

function renderStats() {
  if (state.mode !== "budget") {
    refs.recordsLabel.textContent = "Transactions";
    refs.recordsCount.textContent = "0";
    refs.columnsLabel.textContent = "Champs";
    refs.columnsCount.textContent = "0";
    refs.activeSheet.textContent = "Aucune";
    refs.lastAction.textContent = state.lastAction;
    refs.metricMode.textContent = "Pret pour l'import";
    refs.metricFile.textContent = state.workbookName || "Aucun fichier";
    refs.metricSave.textContent = "Chargez le fichier";
    return;
  }

  if (state.activeView === RECAP_SHEET_NAME) {
    const recapView = buildLiveRecapView();
    refs.recordsLabel.textContent = "Transactions";
    refs.recordsCount.textContent = String(recapView.transactionCount);
    refs.columnsLabel.textContent = "Mois dispo";
    refs.columnsCount.textContent = String(recapView.availableMonthCount);
    refs.activeSheet.textContent = RECAP_SHEET_NAME;
    refs.lastAction.textContent = state.lastAction;
    refs.metricMode.textContent = `Vue recap - ${recapView.periodLabel}`;
    refs.metricFile.textContent = state.workbookName || "Aucun fichier";
    refs.metricSave.textContent = state.workbook
      ? getExportCapabilityLabel()
      : "Rechargez le fichier pour exporter";
    return;
  }

  const filteredRows = getFilteredJournalRows();
  refs.recordsLabel.textContent = "Transactions";
  refs.recordsCount.textContent = String(filteredRows.length);
  refs.columnsLabel.textContent = "Champs";
  refs.columnsCount.textContent = "3";
  refs.activeSheet.textContent = JOURNAL_SHEET_NAME;
  refs.lastAction.textContent = state.lastAction;
  refs.metricMode.textContent = hasActiveRecapPeriodFilter()
    ? `Journalier card view - ${buildRecapPeriodLabel()}`
    : "Journalier card view";
  refs.metricFile.textContent = state.workbookName || "Aucun fichier";
  refs.metricSave.textContent = state.workbook
    ? getExportCapabilityLabel()
    : "Rechargez le fichier pour exporter";
}

function getExportCapabilityLabel() {
  return canUseNativeExcelExport() ? "Partage natif Excel" : "Export vers le classeur";
}

function getFilteredJournalRows() {
  const query = state.search;

  return state.budget.rows
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => {
      if (hasActiveRecapPeriodFilter() && !matchesRecapPeriod(row)) {
        return false;
      }

      if (!query) {
        return true;
      }

      const haystack = [
        row.Date,
        formatDateForDisplay(row.Date),
        row.Categories,
        row.Value,
        formatCurrency(row.Value),
      ].join(" ").toLowerCase();

      return haystack.includes(query);
    })
    .sort((left, right) => compareBudgetRowsForDisplay(left.row, right.row, left.index, right.index));
}

function renderCards() {
  refs.cardsGrid.innerHTML = "";
  refs.recapView.innerHTML = "";

  if (state.mode !== "budget") {
    refs.cardsGrid.classList.remove("hidden");
    refs.recapView.classList.add("hidden");
    refs.cardsEmpty.innerHTML = refs.defaultEmptyMarkup;
    refs.cardsEmpty.classList.remove("hidden");
    return;
  }

  if (state.activeView === RECAP_SHEET_NAME) {
    renderRecapView();
    return;
  }

  renderJournalCards();
}

function renderJournalCards() {
  refs.cardsGrid.classList.remove("hidden");
  refs.recapView.classList.add("hidden");

  if (!state.budget.rows.length) {
    refs.cardsEmpty.classList.remove("hidden");
    refs.cardsEmpty.innerHTML = [
      "<strong>Chargez Budget_2025 Final.xlsx pour demarrer.</strong>",
      "<p>L'app utilisera Journalier!D:F et la liste de categories de Journalier!B.</p>",
    ].join("");
    return;
  }

  const filteredRows = getFilteredJournalRows();

  if (!filteredRows.length) {
    refs.cardsEmpty.classList.remove("hidden");
    refs.cardsEmpty.innerHTML = buildJournalEmptyStateMarkup();
    return;
  }

  refs.cardsEmpty.classList.add("hidden");

  filteredRows.forEach(({ row, index }) => {
    const card = document.createElement("article");
    card.className = `data-card${index === state.editingIndex ? " active" : ""}`;
    card.dataset.index = String(index);

    const amountLabel = formatCurrency(row.Value) || row.Value || "-";
    const dateLabel = formatDateForDisplay(row.Date) || "Sans date";

    card.innerHTML = `
      <div class="card-topline">
        <span class="card-index">${escapeHtml(dateLabel)}</span>
        <div class="card-actions">
          <button class="card-action" type="button" data-action="edit" aria-label="Modifier">Edit</button>
          <button class="card-action delete" type="button" data-action="delete" aria-label="Supprimer">X</button>
        </div>
      </div>
      <div>
        <h3 class="card-title">${escapeHtml(row.Categories || "Categorie non definie")}</h3>
        <p class="card-subtitle">Feuille ${JOURNAL_SHEET_NAME}</p>
        <p class="card-amount">${escapeHtml(amountLabel)}</p>
      </div>
      <div class="card-details">
        ${createDetailMarkup("Date", dateLabel)}
        ${createDetailMarkup("Categorie", row.Categories || "-")}
        ${createDetailMarkup("Value", amountLabel)}
      </div>
    `;

    refs.cardsGrid.appendChild(card);
  });
}

function renderRecapView() {
  refs.cardsGrid.classList.add("hidden");
  refs.recapView.classList.remove("hidden");

  const recapView = buildLiveRecapView();

  if (!recapView.available) {
    refs.cardsEmpty.classList.remove("hidden");
    refs.recapView.classList.add("hidden");
    refs.cardsEmpty.innerHTML = [
      "<strong>La vue Recapitulatif n'est pas disponible.</strong>",
      "<p>Rechargez Budget_2025 Final.xlsx pour reconstruire la synthese.</p>",
    ].join("");
    return;
  }

  refs.cardsEmpty.classList.add("hidden");
  refs.recapView.innerHTML = createRecapMarkup(recapView);
}

function buildJournalEmptyStateMarkup() {
  if (hasActiveRecapPeriodFilter() && state.search) {
    return [
      "<strong>Aucune transaction ne correspond a la recherche pour cette periode.</strong>",
      "<p>Essayez un autre mot-cle ou elargissez le filtre annee/mois.</p>",
    ].join("");
  }

  if (hasActiveRecapPeriodFilter()) {
    return [
      "<strong>Aucune transaction pour la periode choisie.</strong>",
      "<p>Changez le filtre annee/mois ou ajoutez une nouvelle fiche datee dans cette periode.</p>",
    ].join("");
  }

  return [
    "<strong>Aucune transaction ne correspond a la recherche.</strong>",
    "<p>Essayez un autre mot-cle ou ajoutez une nouvelle fiche.</p>",
  ].join("");
}

function buildLiveRecapView() {
  if (state.mode !== "budget" || !state.recap.available) {
    return {
      available: false,
      snapshotDate: "",
      metrics: [],
      detailRows: [],
      planRows: [],
      transactionCount: 0,
      availableMonthCount: 0,
      availableMonthLabels: [],
      availabilityLabel: "",
      availabilityScopeLabel: "",
      periodLabel: "Toutes les donnees",
      filteredUndatedCount: 0,
    };
  }

  const recapRows = getFilteredRecapSourceRows();
  const actualMap = buildActualAmountMap(recapRows);
  const metrics = buildRecapMetrics(actualMap);
  const detailRows = buildRecapDetailRows(actualMap);
  const planRows = buildRecapPlanRows(actualMap, metrics);
  const availableYears = getAvailableRecapYears();
  const availableMonths = getAvailableRecapMonths(state.recapFilters.year);
  const filtersActive = hasActiveRecapPeriodFilter();

  return {
    available: true,
    snapshotDate: state.recap.snapshotDate,
    metrics,
    detailRows,
    planRows,
    transactionCount: recapRows.length,
    availableMonthCount: availableMonths.length,
    availableMonthLabels: availableMonths.map((month) => formatMonthLabel(month)),
    availabilityLabel: buildRecapAvailabilityLabel(availableYears, availableMonths),
    availabilityScopeLabel: buildRecapAvailabilityScopeLabel(),
    periodLabel: buildRecapPeriodLabel(),
    filteredUndatedCount: filtersActive ? countUndatedBudgetRows() : 0,
  };
}

function getFilteredRecapSourceRows() {
  if (!hasActiveRecapPeriodFilter()) {
    return state.budget.rows.slice();
  }

  return state.budget.rows.filter((row) => matchesRecapPeriod(row));
}

function matchesRecapPeriod(row) {
  const dateParts = getBudgetRowDateParts(row);
  if (!dateParts) {
    return false;
  }

  if (state.recapFilters.year !== "all" && dateParts.year !== state.recapFilters.year) {
    return false;
  }

  if (state.recapFilters.month !== "all" && dateParts.month !== state.recapFilters.month) {
    return false;
  }

  return true;
}

function hasActiveRecapPeriodFilter() {
  return state.recapFilters.year !== "all" || state.recapFilters.month !== "all";
}

function getAvailableRecapYears() {
  const years = new Set();

  state.budget.rows.forEach((row) => {
    const dateParts = getBudgetRowDateParts(row);
    if (dateParts) {
      years.add(dateParts.year);
    }
  });

  return Array.from(years).sort((left, right) => Number(right) - Number(left));
}

function getAvailableRecapMonths(year) {
  const months = new Set();

  state.budget.rows.forEach((row) => {
    const dateParts = getBudgetRowDateParts(row);
    if (!dateParts) {
      return;
    }

    if (year !== "all" && dateParts.year !== year) {
      return;
    }

    months.add(dateParts.month);
  });

  return Array.from(months).sort((left, right) => Number(left) - Number(right));
}

function getBudgetRowDateParts(row) {
  const iso = normalizeDateValue(row?.Date);
  if (!iso) {
    return null;
  }

  return {
    year: iso.slice(0, 4),
    month: iso.slice(5, 7),
  };
}

function countUndatedBudgetRows() {
  return state.budget.rows.filter((row) => !getBudgetRowDateParts(row)).length;
}

function buildRecapPeriodLabel() {
  const { year, month } = state.recapFilters;

  if (year === "all" && month === "all") {
    return "Toutes les donnees";
  }

  if (year !== "all" && month === "all") {
    return `Annee ${year}`;
  }

  if (year === "all" && month !== "all") {
    return `${formatMonthLabel(month)} - toutes les annees`;
  }

  return `${formatMonthLabel(month)} ${year}`;
}

function buildRecapAvailabilityLabel(availableYears, availableMonths) {
  if (!availableYears.length) {
    return "Aucune date exploitable dans Journalier.";
  }

  if (state.recapFilters.year === "all") {
    return `${availableYears.length} annee(s) avec donnees et ${availableMonths.length} mois couverts.`;
  }

  return `${availableMonths.length} mois avec donnees en ${state.recapFilters.year}.`;
}

function buildRecapAvailabilityScopeLabel() {
  if (state.recapFilters.year === "all") {
    return "Mois disponibles toutes annees confondues";
  }

  return `Mois disponibles en ${state.recapFilters.year}`;
}

function formatMonthLabel(monthValue) {
  const monthNumber = Number(monthValue);
  if (!Number.isInteger(monthNumber) || monthNumber < 1 || monthNumber > 12) {
    return String(monthValue || "");
  }

  const label = new Intl.DateTimeFormat("fr-CA", {
    month: "long",
    timeZone: "UTC",
  }).format(
    new Date(Date.UTC(2025, monthNumber - 1, 1, 12))
  );

  return label.charAt(0).toUpperCase() + label.slice(1);
}

function buildActualAmountMap(rows) {
  const amounts = new Map();

  rows.forEach((row) => {
    const category = String(row.Categories || "").trim();
    const amount = parseAmount(row.Value);

    if (!category || !Number.isFinite(amount)) {
      return;
    }

    const key = normalizeHeaderName(category);
    amounts.set(key, (amounts.get(key) || 0) + amount);
  });

  return amounts;
}

function buildRecapMetrics(actualMap) {
  const income = Math.abs(getActualAmount(actualMap, "Income"));
  const savings = Math.abs(getActualAmount(actualMap, "Savings"));
  const seasonalSavings = Math.abs(getActualAmount(actualMap, "Savings for seasonal exp."));
  const totalExpenses = computeTotalExpenses(actualMap);
  const cash = income - savings - totalExpenses;

  return [
    { label: "Income", value: income, tone: "positive" },
    { label: "Expenses", value: totalExpenses, tone: "negative" },
    { label: "Savings", value: savings, tone: "neutral" },
    { label: "Cash", value: cash, tone: cash >= 0 ? "positive" : "negative" },
    { label: "Seasonal Savings", value: seasonalSavings, tone: "neutral" },
  ];
}

function buildRecapDetailRows(actualMap) {
  const orderedRows = [];
  const seen = new Set();
  const query = state.search;

  state.budget.categories.forEach((category) => {
    const key = normalizeHeaderName(category);
    const amount = actualMap.get(key);

    if (!Number.isFinite(amount) || Math.abs(amount) < 0.005) {
      return;
    }

    const row = {
      label: category,
      amount,
      isTotal: false,
    };

    if (!matchesRecapSearch(row, query)) {
      return;
    }

    seen.add(key);
    orderedRows.push(row);
  });

  actualMap.forEach((amount, key) => {
    if (seen.has(key) || Math.abs(amount) < 0.005) {
      return;
    }

    const label = findOriginalCategoryLabel(key);
    const row = {
      label,
      amount,
      isTotal: false,
    };

    if (!matchesRecapSearch(row, query)) {
      return;
    }

    orderedRows.push(row);
  });

  const totalAmount = Array.from(actualMap.values()).reduce((sum, value) => sum + value, 0);
  const totalRow = {
    label: "Total general",
    amount: totalAmount,
    isTotal: true,
  };

  if (!query || matchesRecapSearch(totalRow, query)) {
    orderedRows.push(totalRow);
  }

  return orderedRows;
}

function buildRecapPlanRows(actualMap, metrics) {
  const query = state.search;

  return state.recap.planTemplate
    .map((row) => ({
      label: row.label,
      plan: row.plan,
      actual: normalizeAmountValue(computeActualForPlanLabel(row.label, actualMap, metrics)),
    }))
    .filter((row) => matchesRecapSearch(row, query));
}

function computeActualForPlanLabel(label, actualMap, metrics) {
  const normalized = normalizeHeaderName(label);

  if (normalized === "total savings") {
    const savings = metrics.find((metric) => metric.label === "Savings")?.value || 0;
    const seasonal = metrics.find((metric) => metric.label === "Seasonal Savings")?.value || 0;
    return savings + seasonal;
  }

  if (normalized === "total expenses") {
    return metrics.find((metric) => metric.label === "Expenses")?.value || 0;
  }

  if (normalized === "cash short/extra") {
    return metrics.find((metric) => metric.label === "Cash")?.value || 0;
  }

  return Math.abs(getActualAmount(actualMap, label));
}

function matchesRecapSearch(row, query) {
  if (!query) {
    return true;
  }

  const haystack = [
    row.label,
    row.plan,
    row.actual,
    formatSignedCurrency(row.amount),
    formatCurrency(row.actual),
  ].join(" ").toLowerCase();

  return haystack.includes(query);
}

function computeTotalExpenses(actualMap) {
  let total = 0;

  actualMap.forEach((amount, key) => {
    if (isIncomeOrSavingsKey(key)) {
      return;
    }

    if (amount < 0) {
      total += Math.abs(amount);
    }
  });

  return total;
}

function isIncomeOrSavingsKey(key) {
  return key === normalizeHeaderName("Income") ||
    key === normalizeHeaderName("Savings") ||
    key === normalizeHeaderName("Savings for seasonal exp.");
}

function getActualAmount(actualMap, label) {
  return actualMap.get(normalizeHeaderName(label)) || 0;
}

function findOriginalCategoryLabel(key) {
  const match = state.budget.categories.find((category) => normalizeHeaderName(category) === key);
  return match || key;
}

function createRecapMarkup(recapView) {
  const transactionLabel = `${recapView.transactionCount} transaction${recapView.transactionCount > 1 ? "s" : ""} retenue${recapView.transactionCount > 1 ? "s" : ""}.`;
  const undatedNote = recapView.filteredUndatedCount
    ? ` ${recapView.filteredUndatedCount} ligne(s) sans date sont exclues quand un filtre de periode est actif.`
    : "";
  const monthChips = recapView.availableMonthLabels
    .map((monthLabel) => `<span class="recap-chip">${escapeHtml(monthLabel)}</span>`)
    .join("");

  return `
    <div class="recap-shell">
      <div class="recap-metrics">
        ${recapView.metrics.map((metric) => createRecapMetricMarkup(metric)).join("")}
      </div>
      <div class="recap-note">
        <strong>Source:</strong> vue reconstruite depuis <code>${RECAP_SHEET_NAME}</code> et <code>${TCD_SHEET_NAME}</code>.
        ${recapView.snapshotDate ? ` Snapshot date: ${escapeHtml(recapView.snapshotDate)}.` : ""}
        <br><strong>Periode analysee:</strong> ${escapeHtml(recapView.periodLabel)}. ${escapeHtml(transactionLabel)}
        <br><strong>Disponibilite:</strong> ${escapeHtml(recapView.availabilityLabel)}${escapeHtml(undatedNote)}
        <div class="recap-availability">
          <span class="recap-chip">${escapeHtml(recapView.availabilityScopeLabel)}</span>
          ${monthChips || '<span class="recap-chip">Aucun mois date</span>'}
        </div>
      </div>
      ${createRecapTableMarkup(
        "Transactions par categorie",
        "Synthese live calculee a partir de Journalier selon la periode choisie.",
        ["Categorie", "Montant"],
        recapView.detailRows.map((row) => ({
          cells: [
            { value: row.label, numeric: false },
            { value: formatSignedCurrency(row.amount), numeric: true },
          ],
          total: row.isTotal,
        }))
      )}
      ${createRecapTableMarkup(
        "Plan vs reel",
        "Le plan vient de TCD, la colonne Reel suit vos transactions dans l'app pour la periode filtree.",
        ["Categorie", "Plan", "Reel"],
        recapView.planRows.map((row) => ({
          cells: [
            { value: row.label, numeric: false },
            { value: formatCurrency(row.plan), numeric: true },
            { value: formatCurrency(row.actual), numeric: true },
          ],
          total: /total|cash short\/extra/i.test(row.label),
        }))
      )}
    </div>
  `;
}

function createRecapMetricMarkup(metric) {
  return `
    <article class="recap-metric recap-metric-${metric.tone}">
      <span class="recap-metric-label">${escapeHtml(metric.label)}</span>
      <strong>${escapeHtml(formatCurrency(metric.value))}</strong>
    </article>
  `;
}

function createRecapTableMarkup(title, subtitle, headers, rows) {
  if (!rows.length) {
    return `
      <section class="recap-section">
        <div class="recap-section-head">
          <h3>${escapeHtml(title)}</h3>
          <p>${escapeHtml(subtitle)}</p>
        </div>
        <div class="empty-form">Aucune ligne a afficher pour cette vue.</div>
      </section>
    `;
  }

  return `
    <section class="recap-section">
      <div class="recap-section-head">
        <h3>${escapeHtml(title)}</h3>
        <p>${escapeHtml(subtitle)}</p>
      </div>
      <div class="recap-table-wrap">
        <table class="recap-table">
          <thead>
            <tr>
              ${headers.map((header, index) => `<th class="${index > 0 ? "numeric" : ""}">${escapeHtml(header)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${rows.map((row) => `
              <tr class="${row.total ? "total-row" : ""}">
                ${row.cells.map((cell) => `<td class="${cell.numeric ? "numeric" : ""}">${escapeHtml(cell.value || "-")}</td>`).join("")}
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderEditor() {
  if (state.mode !== "budget") {
    refs.formTitle.textContent = "Nouvelle fiche";
    refs.formSubtitle.textContent = "Chargez Budget_2025 Final.xlsx pour generer le formulaire Journalier.";
    refs.formFields.innerHTML = '<div class="empty-form">Le formulaire Date / Categories / Value apparaitra ici.</div>';
    return;
  }

  if (state.activeView === RECAP_SHEET_NAME) {
    refs.formTitle.textContent = "Vue recap";
    refs.formSubtitle.textContent = "Lecture seule dans l'app.";
    refs.formFields.innerHTML = `
      <div class="empty-form">
        Cette vue n'edite pas directement la feuille Recapitulatif d'Excel.
        Elle reconstruit une synthese lisible a partir de ${RECAP_SHEET_NAME}, ${TCD_SHEET_NAME}
        et de vos transactions ${JOURNAL_SHEET_NAME}. Pour modifier les donnees, revenez sur
        la vue Journalier.
      </div>
    `;
    return;
  }

  const editingRow =
    state.editorMode === "edit" &&
    state.editingIndex !== null &&
    state.budget.rows[state.editingIndex]
      ? state.budget.rows[state.editingIndex]
      : { Date: "", Categories: "", Value: "" };

  if (!state.budget.rows[state.editingIndex]) {
    state.editorMode = "create";
    state.editingIndex = null;
  }

  refs.formTitle.textContent = state.editorMode === "edit" ? "Modifier la transaction" : "Nouvelle transaction";
  refs.formSubtitle.textContent = "Saisie directe de Journalier!D:F avec categories predefinies depuis la colonne B.";
  refs.formFields.innerHTML = "";

  appendField(renderDateField(editingRow.Date));
  appendField(renderCategoryField(editingRow.Categories));
  appendField(renderValueField(editingRow.Value));
}

function renderDateField(value) {
  const wrapper = document.createElement("div");
  wrapper.className = "field-card";

  const label = document.createElement("label");
  label.setAttribute("for", "field-date");
  label.textContent = "Date";

  const input = document.createElement("input");
  input.id = "field-date";
  input.name = "Date";
  input.type = "date";
  input.value = toDateInputValue(value);

  wrapper.append(label, input);
  return wrapper;
}

function renderCategoryField(value) {
  const wrapper = document.createElement("div");
  wrapper.className = "field-card";

  const label = document.createElement("label");
  label.setAttribute("for", "field-categories");
  label.textContent = "Categories";

  const select = document.createElement("select");
  select.id = "field-categories";
  select.name = "Categories";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Choisir une categorie";
  select.appendChild(placeholder);

  state.budget.categories.forEach((category) => {
    const option = document.createElement("option");
    option.value = category;
    option.textContent = category;
    select.appendChild(option);
  });

  if (value && !state.budget.categories.includes(value)) {
    const currentOption = document.createElement("option");
    currentOption.value = value;
    currentOption.textContent = `${value} (hors liste)`;
    select.appendChild(currentOption);
  }

  select.value = value || "";

  const hint = document.createElement("p");
  hint.className = "field-hint";
  hint.textContent = `${state.budget.categories.length} categories disponibles depuis Journalier!B.`;

  wrapper.append(label, select, hint);
  return wrapper;
}

function renderValueField(value) {
  const wrapper = document.createElement("div");
  wrapper.className = "field-card";

  const label = document.createElement("label");
  label.setAttribute("for", "field-value");
  label.textContent = "Value";

  const input = document.createElement("input");
  input.id = "field-value";
  input.name = "Value";
  input.type = "text";
  input.inputMode = "decimal";
  input.placeholder = "-42.08";
  input.value = value || "";

  const hint = document.createElement("p");
  hint.className = "field-hint";
  hint.textContent = "Entrez un nombre negatif pour une depense, positif pour un revenu.";

  wrapper.append(label, input, hint);
  return wrapper;
}

function appendField(field) {
  refs.formFields.appendChild(field);
}

function createDetailMarkup(label, value) {
  return `
    <div class="detail-row">
      <span class="detail-label">${escapeHtml(label)}</span>
      <span class="detail-value">${escapeHtml(value || "-")}</span>
    </div>
  `;
}

function readCellText(cell) {
  if (!cell) {
    return "";
  }

  if (cell.w !== undefined && cell.w !== null && cell.w !== "") {
    return String(cell.w);
  }

  if (cell.v === undefined || cell.v === null) {
    return "";
  }

  return String(cell.v);
}

function readCellRawValue(cell) {
  return cell ? cell.v : "";
}

function normalizeHeaderName(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function normalizeDateValue(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return excelSerialToIso(value);
  }

  const text = String(value).trim();
  if (!text) {
    return "";
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }

  if (/^\d+(\.\d+)?$/.test(text)) {
    return excelSerialToIso(Number(text));
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }

  return "";
}

function excelSerialToIso(serial) {
  const epoch = Date.UTC(1899, 11, 30);
  const millis = Math.round(Number(serial) * 86400000);
  const date = new Date(epoch + millis);

  if (Number.isNaN(date.getTime())) {
    return "";
  }

  return date.toISOString().slice(0, 10);
}

function isoDateToExcelSerial(iso) {
  const date = new Date(`${iso}T00:00:00Z`);
  return (date.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
}

function normalizeAmountValue(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return trimTrailingZeros(value);
  }

  const text = String(value).trim();
  if (!text) {
    return "";
  }

  const parsed = parseAmount(text);
  return Number.isFinite(parsed) ? trimTrailingZeros(parsed) : text;
}

function trimTrailingZeros(value) {
  return String(Number(value));
}

function parseAmount(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const text = String(value ?? "").trim();
  if (!text) {
    return Number.NaN;
  }

  const normalized = text
    .replace(/\s/g, "")
    .replace(/\$/g, "")
    .replace(/\(/g, "-")
    .replace(/\)/g, "")
    .replace(/,/g, "");

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : Number.NaN;
}

function formatDateForDisplay(value) {
  const iso = normalizeDateValue(value);
  if (!iso) {
    return String(value ?? "").trim();
  }

  const parsed = new Date(`${iso}T00:00:00`);
  return new Intl.DateTimeFormat("fr-CA", { dateStyle: "medium" }).format(parsed);
}

function formatCurrency(value) {
  const amount = parseAmount(value);
  if (!Number.isFinite(amount)) {
    return String(value ?? "").trim();
  }

  return new Intl.NumberFormat("fr-CA", {
    style: "currency",
    currency: "CAD",
    maximumFractionDigits: 2,
  }).format(amount);
}

function formatSignedCurrency(value) {
  const amount = parseAmount(value);
  if (!Number.isFinite(amount)) {
    return String(value ?? "").trim();
  }

  return new Intl.NumberFormat("fr-CA", {
    style: "currency",
    currency: "CAD",
    maximumFractionDigits: 2,
    signDisplay: "auto",
  }).format(amount);
}

function toDateInputValue(value) {
  return normalizeDateValue(value);
}

function sortBudgetRowsInPlace(rows) {
  rows.sort((left, right) => compareBudgetRowsForStorage(left, right));
}

function compareBudgetRowsForStorage(left, right) {
  const leftDate = normalizeDateValue(left.Date);
  const rightDate = normalizeDateValue(right.Date);

  if (leftDate && rightDate && leftDate !== rightDate) {
    return leftDate.localeCompare(rightDate);
  }

  if (leftDate && !rightDate) {
    return -1;
  }

  if (!leftDate && rightDate) {
    return 1;
  }

  const categoryCompare = (left.Categories || "").localeCompare(right.Categories || "", "fr-CA", {
    sensitivity: "base",
  });

  if (categoryCompare !== 0) {
    return categoryCompare;
  }

  return (left.__id || "").localeCompare(right.__id || "");
}

function compareBudgetRowsForDisplay(leftRow, rightRow, leftIndex, rightIndex) {
  const storageCompare = compareBudgetRowsForStorage(leftRow, rightRow);
  if (storageCompare !== 0) {
    return storageCompare * -1;
  }

  return rightIndex - leftIndex;
}

function sanitizeBudgetRow(row) {
  return {
    __id: row?.__id || createId(),
    Date: normalizeDateValue(row?.Date),
    Categories: String(row?.Categories ?? "").trim(),
    Value: normalizeAmountValue(row?.Value),
  };
}

function createId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID();
  }

  return `record-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function setLastAction(message) {
  state.lastAction = message;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
