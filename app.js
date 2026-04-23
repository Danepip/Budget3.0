const STORAGE_KEY = "budget-2025-card-view-v2";

const JOURNAL_SHEET_NAME = "Journalier";
const RECAP_SHEET_NAME = "Recapitulatif";
const ANALYSIS_VIEW_NAME = "Comparaisons";
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
  sourceLink: null,
  sourceSafety: createEmptySourceSafety(),
  cloud: createEmptyCloudState(),
  draftSavedAt: "",
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
let sourceSaveQueue = Promise.resolve();
let budgetSourcePlugin = null;
let supabaseClient = null;
let supabaseAuthSubscription = null;
let supabaseRealtimeChannel = null;
let cloudRefreshTimer = null;
let cloudSyncQueue = Promise.resolve();

document.addEventListener("DOMContentLoaded", () => {
  cacheDom();
  bindEvents();
  restoreDraft();
  syncLibraryState();
  setupAppShell();
  void initSupabaseIntegration();
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

function createEmptyCloudState() {
  return {
    configured: false,
    ready: false,
    syncBusy: false,
    email: "",
    status: "Supabase non configure.",
    session: null,
    user: null,
    space: {
      id: "",
      name: "",
      joinCode: "",
    },
    lastPulledAt: "",
    lastPushedAt: "",
  };
}

function createEmptySourceSafety() {
  return {
    allowDirectWrite: false,
    reason: "Ecriture directe desactivee pour proteger l'integrite du classeur source.",
    issues: [],
  };
}

function normalizeActiveView(value) {
  if (value === RECAP_SHEET_NAME || value === ANALYSIS_VIEW_NAME) {
    return value;
  }

  return JOURNAL_SHEET_NAME;
}

function cacheDom() {
  refs.fileInput = document.getElementById("excel-file");
  refs.sheetSelect = document.getElementById("sheet-select");
  refs.recapYearField = document.getElementById("recap-year-field");
  refs.recapYearSelect = document.getElementById("recap-year-select");
  refs.recapMonthField = document.getElementById("recap-month-field");
  refs.recapMonthSelect = document.getElementById("recap-month-select");
  refs.searchInput = document.getElementById("search-input");
  refs.cloudStatus = document.getElementById("cloud-status");
  refs.cloudEmailInput = document.getElementById("cloud-email");
  refs.cloudCodeInput = document.getElementById("cloud-code");
  refs.cloudMagicLinkButton = document.getElementById("cloud-magic-link");
  refs.cloudSignOutButton = document.getElementById("cloud-sign-out");
  refs.cloudCreateSpaceButton = document.getElementById("cloud-create-space");
  refs.cloudJoinSpaceButton = document.getElementById("cloud-join-space");
  refs.cloudPushButton = document.getElementById("cloud-push");
  refs.cloudPullButton = document.getElementById("cloud-pull");
  refs.cloudSpaceHint = document.getElementById("cloud-space-hint");
  refs.openSourceButton = document.getElementById("open-source");
  refs.saveSourceButton = document.getElementById("save-source");
  refs.saveDraftButton = document.getElementById("save-draft");
  refs.restoreDraftButton = document.getElementById("restore-draft");
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
  refs.draftStatus = document.getElementById("draft-status");
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
  refs.cloudMagicLinkButton.addEventListener("click", () => {
    void onCloudMagicLinkRequested();
  });
  refs.cloudSignOutButton.addEventListener("click", () => {
    void onCloudSignOutRequested();
  });
  refs.cloudCreateSpaceButton.addEventListener("click", () => {
    void onCloudCreateSpaceRequested();
  });
  refs.cloudJoinSpaceButton.addEventListener("click", () => {
    void onCloudJoinSpaceRequested();
  });
  refs.cloudPushButton.addEventListener("click", () => {
    void onCloudPublishRequested();
  });
  refs.cloudPullButton.addEventListener("click", () => {
    void onCloudRefreshRequested();
  });
  refs.openSourceButton.addEventListener("click", onOpenSourceRequested);
  refs.saveSourceButton.addEventListener("click", () => {
    void saveSourceWorkbook();
  });
  refs.saveDraftButton.addEventListener("click", onSaveDraftRequested);
  refs.restoreDraftButton.addEventListener("click", onRestoreDraftRequested);
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

function getSupabaseRuntime() {
  return window.supabase || null;
}

function getSupabaseConfig() {
  const rawConfig = window.BUDGET_SUPABASE_CONFIG || {};

  return {
    url: String(rawConfig.url || "").trim(),
    anonKey: String(rawConfig.anonKey || "").trim(),
    defaultSpaceName: String(rawConfig.defaultSpaceName || "Budget partage 2025").trim() || "Budget partage 2025",
  };
}

function hasSupabaseSession() {
  return Boolean(state.cloud.user);
}

function hasCloudSpaceSelected() {
  return Boolean(state.cloud.space.id);
}

function canUseSupabaseCloud() {
  return Boolean(supabaseClient && hasSupabaseSession() && hasCloudSpaceSelected());
}

function hasLocalBudgetData() {
  return state.mode === "budget" && (
    state.budget.rows.length > 0 ||
    state.budget.categories.length > 0 ||
    state.recap.planTemplate.length > 0
  );
}

function buildSupabaseRedirectUrl() {
  return `${window.location.origin}${window.location.pathname}`;
}

function setCloudStatus(message) {
  state.cloud.status = message;
}

function setCloudBusy(nextBusy) {
  state.cloud.syncBusy = Boolean(nextBusy);
}

function clearCloudRefreshTimer() {
  if (!cloudRefreshTimer) {
    return;
  }

  window.clearTimeout(cloudRefreshTimer);
  cloudRefreshTimer = null;
}

function queueCloudRefresh() {
  clearCloudRefreshTimer();

  if (!canUseSupabaseCloud()) {
    return;
  }

  cloudRefreshTimer = window.setTimeout(() => {
    cloudRefreshTimer = null;
    void loadBudgetFromSupabase(state.cloud.space.id, {
      silent: true,
      preserveLastAction: true,
    }).catch(() => undefined);
  }, 700);
}

async function initSupabaseIntegration() {
  const runtime = getSupabaseRuntime();
  const config = getSupabaseConfig();

  if (!runtime || !config.url || !config.anonKey) {
    state.cloud.configured = false;
    state.cloud.ready = false;
    setCloudStatus("Supabase non configure. Renseignez supabase.config.js pour activer le partage.");
    renderAll();
    return;
  }

  try {
    supabaseClient = runtime.createClient(config.url, config.anonKey, {
      auth: {
        persistSession: true,
        autoRefreshToken: true,
        detectSessionInUrl: true,
      },
    });
    state.cloud.configured = true;
    state.cloud.ready = true;
    setCloudStatus("Supabase configure. Connectez-vous pour partager le budget.");

    if (supabaseAuthSubscription?.data?.subscription?.unsubscribe) {
      supabaseAuthSubscription.data.subscription.unsubscribe();
    }

    supabaseAuthSubscription = supabaseClient.auth.onAuthStateChange((_event, session) => {
      state.cloud.session = session || null;
      state.cloud.user = session?.user || null;
      if (session?.user?.email) {
        state.cloud.email = session.user.email;
      }

      if (!state.cloud.user) {
        stopSupabaseRealtime();
        setCloudStatus("Supabase configure. Connectez-vous pour partager le budget.");
        persistDraftIfPossible();
        renderAll();
        return;
      }

      setCloudStatus(hasCloudSpaceSelected()
        ? `Connecte a Supabase. Espace actif: ${state.cloud.space.name || state.cloud.space.joinCode || "budget partage"}.`
        : "Connecte a Supabase. Creez ou rejoignez un espace partage.");
      persistDraftIfPossible();
      renderAll();

      if (hasCloudSpaceSelected()) {
        void attachToCurrentCloudSpace({ silent: true, preserveLastAction: true });
      }
    });

    await syncSupabaseSession();
  } catch (error) {
    console.error(error);
    state.cloud.configured = false;
    state.cloud.ready = false;
    setCloudStatus("Supabase n'a pas pu etre initialise.");
    renderAll();
  }
}

async function syncSupabaseSession() {
  if (!supabaseClient) {
    return;
  }

  const { data, error } = await supabaseClient.auth.getSession();
  if (error) {
    console.error(error);
    setCloudStatus("Connexion Supabase indisponible pour le moment.");
    renderAll();
    return;
  }

  state.cloud.session = data.session || null;
  state.cloud.user = data.session?.user || null;
  if (data.session?.user?.email) {
    state.cloud.email = data.session.user.email;
  }

  if (state.cloud.user) {
    setCloudStatus(hasCloudSpaceSelected()
      ? `Connecte a Supabase. Espace actif: ${state.cloud.space.name || state.cloud.space.joinCode || "budget partage"}.`
      : "Connecte a Supabase. Creez ou rejoignez un espace partage.");
    if (hasCloudSpaceSelected()) {
      await attachToCurrentCloudSpace({
        silent: true,
        preserveLastAction: true,
      });
    }
  }

  renderAll();
}

function normalizeCloudSpaceRecord(record) {
  return {
    id: String(record?.space_id || record?.id || "").trim(),
    name: String(record?.space_name || record?.name || "").trim(),
    joinCode: String(record?.join_code || record?.joinCode || "").trim(),
  };
}

function applyCloudSpaceRecord(record) {
  const normalized = normalizeCloudSpaceRecord(record);
  state.cloud.space = normalized;

  if (refs.cloudCodeInput) {
    refs.cloudCodeInput.value = normalized.joinCode;
  }
}

function persistDraftIfPossible() {
  if (state.mode === "budget") {
    persistDraft();
  }
}

async function onCloudMagicLinkRequested() {
  if (!supabaseClient || !state.cloud.ready) {
    setLastAction("Supabase n'est pas encore configure.");
    renderAll();
    return;
  }

  const email = String(refs.cloudEmailInput.value || "").trim().toLowerCase();
  if (!email) {
    setLastAction("Saisissez votre email pour recevoir le lien magique.");
    renderAll();
    return;
  }

  try {
    state.cloud.email = email;
    setCloudBusy(true);
    setCloudStatus("Envoi du lien magique en cours...");
    renderAll();

    const { error } = await supabaseClient.auth.signInWithOtp({
      email,
      options: {
        emailRedirectTo: buildSupabaseRedirectUrl(),
      },
    });

    if (error) {
      throw error;
    }

    setCloudStatus(`Lien magique envoye a ${email}. Ouvrez votre email pour terminer la connexion.`);
    setLastAction(`Lien magique Supabase envoye a ${email}`);
  } catch (error) {
    console.error(error);
    setCloudStatus("Le lien magique n'a pas pu etre envoye.");
    setLastAction("Connexion Supabase impossible");
  } finally {
    setCloudBusy(false);
    renderAll();
  }
}

async function onCloudSignOutRequested() {
  if (!supabaseClient || !hasSupabaseSession()) {
    setLastAction("Aucune session Supabase active.");
    renderAll();
    return;
  }

  try {
    setCloudBusy(true);
    const { error } = await supabaseClient.auth.signOut();
    if (error) {
      throw error;
    }

    stopSupabaseRealtime();
    state.cloud.session = null;
    state.cloud.user = null;
    setCloudStatus("Session Supabase fermee.");
    setLastAction("Deconnexion Supabase terminee");
  } catch (error) {
    console.error(error);
    setCloudStatus("La deconnexion Supabase a echoue.");
    setLastAction("Deconnexion Supabase impossible");
  } finally {
    setCloudBusy(false);
    renderAll();
  }
}

async function onCloudCreateSpaceRequested() {
  if (!supabaseClient || !hasSupabaseSession()) {
    setLastAction("Connectez-vous d'abord a Supabase.");
    renderAll();
    return;
  }

  const suggestedName = state.workbookName
    ? state.workbookName.replace(/\.(xlsx|xls|csv)$/i, "")
    : getSupabaseConfig().defaultSpaceName;
  const desiredName = String(window.prompt("Nom du budget partage", suggestedName) || "").trim();

  if (!desiredName) {
    setLastAction("Creation d'espace annulee.");
    renderAll();
    return;
  }

  try {
    setCloudBusy(true);
    setCloudStatus("Creation de l'espace partage...");
    renderAll();

    const { data, error } = await supabaseClient.rpc("create_budget_space", {
      space_name: desiredName,
    });

    if (error) {
      throw error;
    }

    const createdSpace = Array.isArray(data) ? data[0] : data;
    applyCloudSpaceRecord(createdSpace);
    setCloudStatus(`Espace partage cree: ${state.cloud.space.name}.`);
    setLastAction(`Espace cloud cree: ${state.cloud.space.name}`);

    await attachToCurrentCloudSpace({
      silent: true,
      preserveLastAction: true,
    });

    if (hasLocalBudgetData()) {
      const publishNow = window.confirm("Publier vos donnees locales actuelles vers ce nouvel espace partage ?");
      if (publishNow) {
        await publishLocalBudgetToSupabase();
      }
    }
  } catch (error) {
    console.error(error);
    setCloudStatus("L'espace partage n'a pas pu etre cree.");
    setLastAction("Creation de l'espace cloud impossible");
  } finally {
    setCloudBusy(false);
    renderAll();
  }
}

async function onCloudJoinSpaceRequested() {
  if (!supabaseClient || !hasSupabaseSession()) {
    setLastAction("Connectez-vous d'abord a Supabase.");
    renderAll();
    return;
  }

  const joinCode = String(refs.cloudCodeInput.value || "").trim().toLowerCase();
  if (!joinCode) {
    setLastAction("Saisissez le code de l'espace partage.");
    renderAll();
    return;
  }

  try {
    setCloudBusy(true);
    setCloudStatus("Rejointure de l'espace partage...");
    renderAll();

    const { data, error } = await supabaseClient.rpc("join_budget_space", {
      space_join_code: joinCode,
    });

    if (error) {
      throw error;
    }

    const joinedSpace = Array.isArray(data) ? data[0] : data;
    applyCloudSpaceRecord(joinedSpace);
    setCloudStatus(`Espace partage rejoint: ${state.cloud.space.name}.`);
    setLastAction(`Espace cloud rejoint: ${state.cloud.space.name}`);

    const shouldLoadCloud = !hasLocalBudgetData() ||
      window.confirm("Charger les donnees de cet espace cloud et remplacer les donnees locales visibles ?");

    if (shouldLoadCloud) {
      await loadBudgetFromSupabase(state.cloud.space.id, {
        silent: true,
        preserveLastAction: true,
      });
    } else {
      await attachToCurrentCloudSpace({
        silent: true,
        preserveLastAction: true,
      });
    }
  } catch (error) {
    console.error(error);
    setCloudStatus("Impossible de rejoindre cet espace partage.");
    setLastAction("Rejointure cloud impossible");
  } finally {
    setCloudBusy(false);
    renderAll();
  }
}

async function onCloudPublishRequested() {
  if (!canUseSupabaseCloud()) {
    setLastAction("Connectez-vous et choisissez un espace cloud avant la publication.");
    renderAll();
    return;
  }

  if (!hasLocalBudgetData()) {
    setLastAction("Aucune donnee locale a publier vers Supabase.");
    renderAll();
    return;
  }

  await publishLocalBudgetToSupabase();
}

async function onCloudRefreshRequested() {
  if (!canUseSupabaseCloud()) {
    setLastAction("Connectez-vous et choisissez un espace cloud avant de recharger.");
    renderAll();
    return;
  }

  try {
    setCloudBusy(true);
    await loadBudgetFromSupabase(state.cloud.space.id);
    setLastAction("Budget recharge depuis Supabase.");
  } catch (error) {
    console.error(error);
    setLastAction("Le rechargement depuis Supabase a echoue.");
  } finally {
    setCloudBusy(false);
    renderAll();
  }
}

async function attachToCurrentCloudSpace(options = {}) {
  if (!canUseSupabaseCloud()) {
    return;
  }

  await loadCloudSpaceMetadata(state.cloud.space.id, options);
  startSupabaseRealtime(state.cloud.space.id);
}

async function loadCloudSpaceMetadata(spaceId, options = {}) {
  if (!supabaseClient || !spaceId) {
    return;
  }

  const { data, error } = await supabaseClient
    .from("budget_spaces")
    .select("id, name, join_code")
    .eq("id", spaceId)
    .maybeSingle();

  if (error) {
    console.error(error);
    if (!options.silent) {
      setCloudStatus("Impossible de lire les informations de l'espace Supabase.");
    }
    renderAll();
    return;
  }

  if (data) {
    applyCloudSpaceRecord(data);
    if (!options.silent) {
      setCloudStatus(`Espace partage actif: ${state.cloud.space.name}.`);
    }
    persistDraftIfPossible();
    renderAll();
  }
}

function startSupabaseRealtime(spaceId) {
  if (!supabaseClient || !spaceId) {
    return;
  }

  const currentTopic = `budget-space-${spaceId}`;
  if (supabaseRealtimeChannel?.topic === currentTopic) {
    return;
  }

  stopSupabaseRealtime();

  supabaseRealtimeChannel = supabaseClient
    .channel(currentTopic)
    .on("postgres_changes", {
      event: "*",
      schema: "public",
      table: "budget_transactions",
      filter: `space_id=eq.${spaceId}`,
    }, queueCloudRefresh)
    .on("postgres_changes", {
      event: "*",
      schema: "public",
      table: "budget_categories",
      filter: `space_id=eq.${spaceId}`,
    }, queueCloudRefresh)
    .on("postgres_changes", {
      event: "*",
      schema: "public",
      table: "budget_plan_rows",
      filter: `space_id=eq.${spaceId}`,
    }, queueCloudRefresh)
    .subscribe();
}

function stopSupabaseRealtime() {
  clearCloudRefreshTimer();

  if (!supabaseRealtimeChannel || !supabaseClient) {
    supabaseRealtimeChannel = null;
    return;
  }

  supabaseClient.removeChannel(supabaseRealtimeChannel);
  supabaseRealtimeChannel = null;
}

async function publishLocalBudgetToSupabase() {
  if (!canUseSupabaseCloud()) {
    return;
  }

  try {
    setCloudBusy(true);
    setCloudStatus(`Publication en cours vers ${state.cloud.space.name || "l'espace partage"}...`);
    renderAll();

    const spaceId = state.cloud.space.id;
    const categoriesPayload = buildSupabaseCategoryPayload(spaceId);
    const planPayload = buildSupabasePlanPayload(spaceId);
    const transactionsPayload = buildSupabaseTransactionPayload(spaceId);

    let query = supabaseClient.from("budget_transactions").delete();
    let { error } = await query.eq("space_id", spaceId);
    if (error) {
      throw error;
    }

    query = supabaseClient.from("budget_categories").delete();
    ({ error } = await query.eq("space_id", spaceId));
    if (error) {
      throw error;
    }

    query = supabaseClient.from("budget_plan_rows").delete();
    ({ error } = await query.eq("space_id", spaceId));
    if (error) {
      throw error;
    }

    if (categoriesPayload.length) {
      ({ error } = await supabaseClient.from("budget_categories").insert(categoriesPayload));
      if (error) {
        throw error;
      }
    }

    if (planPayload.length) {
      ({ error } = await supabaseClient.from("budget_plan_rows").insert(planPayload));
      if (error) {
        throw error;
      }
    }

    if (transactionsPayload.length) {
      ({ error } = await supabaseClient.from("budget_transactions").upsert(transactionsPayload));
      if (error) {
        throw error;
      }
    }

    state.cloud.lastPushedAt = new Date().toISOString();
    setCloudStatus(`Budget publie vers ${state.cloud.space.name}.`);
    setLastAction(`Donnees locales publiees vers ${state.cloud.space.name}`);
    persistDraftIfPossible();
  } catch (error) {
    console.error(error);
    setCloudStatus("La publication vers Supabase a echoue.");
    setLastAction("Publication Supabase impossible");
  } finally {
    setCloudBusy(false);
    renderAll();
  }
}

async function loadBudgetFromSupabase(spaceId, options = {}) {
  if (!supabaseClient || !spaceId) {
    return;
  }

  try {
    if (!options.silent) {
      setCloudStatus("Chargement des donnees cloud...");
      renderAll();
    }

    const [{ data: categories, error: categoriesError }, { data: planRows, error: planError }, { data: transactions, error: transactionsError }] = await Promise.all([
      supabaseClient
        .from("budget_categories")
        .select("name, position")
        .eq("space_id", spaceId)
        .order("position", { ascending: true })
        .order("name", { ascending: true }),
      supabaseClient
        .from("budget_plan_rows")
        .select("label, plan_amount, position")
        .eq("space_id", spaceId)
        .order("position", { ascending: true })
        .order("label", { ascending: true }),
      supabaseClient
        .from("budget_transactions")
        .select("id, entry_date, category, amount, sort_order")
        .eq("space_id", spaceId)
        .order("sort_order", { ascending: true })
        .order("entry_date", { ascending: true })
        .order("created_at", { ascending: true }),
    ]);

    if (categoriesError) {
      throw categoriesError;
    }

    if (planError) {
      throw planError;
    }

    if (transactionsError) {
      throw transactionsError;
    }

    state.mode = "budget";
    state.workbookName = state.workbookName || state.cloud.space.name || "Budget partage cloud";
    state.workbook = null;
    state.sourceLink = null;
    state.sourceSafety = createEmptySourceSafety();
    state.activeView = normalizeActiveView(state.activeView);
    state.search = "";
    state.editingIndex = null;
    state.editorMode = "create";
    state.budget = {
      headers: ["Date", "Categories", "Value"],
      categories: (categories || []).map((row) => String(row.name || "").trim()).filter(Boolean),
      rows: (transactions || []).map((row) => sanitizeBudgetRow({
        __id: row.id,
        Date: row.entry_date,
        Categories: row.category,
        Value: normalizeAmountValue(row.amount),
      })),
      clearEndRow: START_ROW + (transactions?.length || 0) + 8,
    };
    state.recap = {
      available: true,
      snapshotDate: `Supabase - ${formatDraftSavedAt(new Date().toISOString())}`,
      planTemplate: (planRows || []).map((row) => ({
        label: String(row.label || "").trim(),
        plan: normalizeAmountValue(row.plan_amount),
      })).filter((row) => row.label),
    };
    state.cloud.lastPulledAt = new Date().toISOString();

    if (!options.preserveLastAction) {
      setLastAction(`Budget charge depuis ${state.cloud.space.name || "Supabase"}`);
    }

    setCloudStatus(`Espace partage actif: ${state.cloud.space.name || "budget partage"}.`);
    persistDraft();
    startSupabaseRealtime(spaceId);
    renderAll();
  } catch (error) {
    console.error(error);
    if (!options.silent) {
      setCloudStatus("Le chargement des donnees cloud a echoue.");
      renderAll();
    }
    throw error;
  }
}

function buildSupabaseCategoryPayload(spaceId) {
  const categories = new Set();
  const ordered = [];

  state.budget.categories.forEach((category) => {
    const normalized = String(category || "").trim();
    if (!normalized || categories.has(normalized)) {
      return;
    }

    categories.add(normalized);
    ordered.push(normalized);
  });

  state.budget.rows.forEach((row) => {
    const normalized = String(row.Categories || "").trim();
    if (!normalized || categories.has(normalized)) {
      return;
    }

    categories.add(normalized);
    ordered.push(normalized);
  });

  return ordered.map((name, index) => ({
    space_id: spaceId,
    name,
    position: index,
  }));
}

function buildSupabasePlanPayload(spaceId) {
  return state.recap.planTemplate
    .map((row, index) => ({
      space_id: spaceId,
      label: String(row.label || "").trim(),
      plan_amount: Number.isFinite(parseAmount(row.plan)) ? parseAmount(row.plan) : null,
      position: index,
    }))
    .filter((row) => row.label);
}

function buildSupabaseTransactionPayload(spaceId) {
  return state.budget.rows
    .map((row, index) => ({
      id: String(row.__id || createId()),
      space_id: spaceId,
      entry_date: normalizeDateValue(row.Date) || null,
      category: String(row.Categories || "").trim(),
      amount: Number.isFinite(parseAmount(row.Value)) ? parseAmount(row.Value) : null,
      sort_order: index,
    }))
    .filter((row) => row.category || Number.isFinite(row.amount) || row.entry_date);
}

async function syncSingleTransactionToSupabase(record) {
  if (!canUseSupabaseCloud()) {
    return;
  }

  const categoryName = String(record.Categories || "").trim();

  if (categoryName) {
    const knownCategories = buildSupabaseCategoryPayload(state.cloud.space.id);
    const targetCategory = knownCategories.find((row) => row.name === categoryName) || {
      space_id: state.cloud.space.id,
      name: categoryName,
      position: knownCategories.length,
    };

    const { error: categoryError } = await supabaseClient
      .from("budget_categories")
      .upsert(targetCategory, {
        onConflict: "space_id,name",
      });

    if (categoryError) {
      throw categoryError;
    }
  }

  const payload = {
    id: String(record.__id || createId()),
    space_id: state.cloud.space.id,
    entry_date: normalizeDateValue(record.Date) || null,
    category: categoryName,
    amount: Number.isFinite(parseAmount(record.Value)) ? parseAmount(record.Value) : null,
    sort_order: Math.max(0, state.budget.rows.findIndex((row) => row.__id === record.__id)),
  };

  const { error } = await supabaseClient
    .from("budget_transactions")
    .upsert(payload);

  if (error) {
    throw error;
  }

  state.cloud.lastPushedAt = new Date().toISOString();
  setCloudStatus(`Derniere transaction synchronisee vers ${state.cloud.space.name || "Supabase"}.`);
  persistDraftIfPossible();
}

async function removeSingleTransactionFromSupabase(recordId) {
  if (!canUseSupabaseCloud()) {
    return;
  }

  const { error } = await supabaseClient
    .from("budget_transactions")
    .delete()
    .eq("space_id", state.cloud.space.id)
    .eq("id", String(recordId || ""));

  if (error) {
    throw error;
  }

  state.cloud.lastPushedAt = new Date().toISOString();
  setCloudStatus(`Suppression synchronisee vers ${state.cloud.space.name || "Supabase"}.`);
  persistDraftIfPossible();
}

function enqueueCloudSync(task) {
  const nextTask = cloudSyncQueue.then(() => task());
  cloudSyncQueue = nextTask.catch(() => undefined);
  return nextTask;
}

function isStandaloneMode() {
  return window.matchMedia?.("(display-mode: standalone)")?.matches || window.navigator.standalone === true;
}

function isAppleMobileDevice() {
  return /iphone|ipad|ipod/i.test(window.navigator.userAgent || "");
}

function readStoredDraft() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return null;
    }

    const draft = JSON.parse(raw);
    return draft && typeof draft === "object" ? draft : null;
  } catch (error) {
    console.error(error);
    return null;
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
    cloud: {
      email: state.cloud.email,
      space: state.cloud.space,
      lastPulledAt: state.cloud.lastPulledAt,
      lastPushedAt: state.cloud.lastPushedAt,
    },
    recapFilters: state.recapFilters,
    savedAt: new Date().toISOString(),
  };

  state.draftSavedAt = payload.savedAt;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function applyStoredDraft(draft) {
  state.mode = "budget";
  state.workbookName = draft.workbookName || "";
  state.workbook = null;
  state.sourceLink = null;
  state.activeView = normalizeActiveView(draft.activeView);
  state.search = "";
  state.editingIndex = null;
  state.editorMode = "create";
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
  state.cloud.email = String(draft.cloud?.email || state.cloud.email || "");
  state.cloud.space = {
    id: String(draft.cloud?.space?.id || state.cloud.space.id || ""),
    name: String(draft.cloud?.space?.name || state.cloud.space.name || ""),
    joinCode: String(draft.cloud?.space?.joinCode || draft.cloud?.space?.join_code || state.cloud.space.joinCode || ""),
  };
  state.cloud.lastPulledAt = String(draft.cloud?.lastPulledAt || state.cloud.lastPulledAt || "");
  state.cloud.lastPushedAt = String(draft.cloud?.lastPushedAt || state.cloud.lastPushedAt || "");
  state.sourceSafety = createEmptySourceSafety();
  state.draftSavedAt = String(draft.savedAt || "");

  if (refs.searchInput) {
    refs.searchInput.value = "";
  }

  if (refs.fileInput) {
    refs.fileInput.value = "";
  }
}

function restoreDraft(options = {}) {
  const manual = Boolean(options.manual);
  const draft = readStoredDraft();

  if (!draft || draft.mode !== "budget" || !Array.isArray(draft.rows)) {
    if (manual) {
      setLastAction("Aucun brouillon local a restaurer.");
    }
    return false;
  }

  applyStoredDraft(draft);
  state.lastAction = manual
    ? "Brouillon local restaure. Mode autonome actif avec vos donnees locales."
    : "Brouillon restaure. Mode autonome actif avec vos donnees locales.";
  return true;
}

function onSaveDraftRequested() {
  if (state.mode !== "budget") {
    setLastAction("Chargez ou restaurez un budget avant de sauvegarder en local.");
    renderAll();
    return;
  }

  persistDraft();
  setLastAction("Sauvegarde locale mise a jour.");
  renderAll();
}

function onRestoreDraftRequested() {
  restoreDraft({ manual: true });
  renderAll();
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
    await importWorkbookFile(file, {
      sourceLink: null,
      successMessage: `Classeur charge: ${file.name}`,
    });
  } catch (error) {
    console.error(error);
    state.workbook = null;
    state.sourceLink = null;
    setLastAction("Le fichier n'a pas pu etre lu");
    renderAll();
  }
}

async function onOpenSourceRequested() {
  if (!window.XLSX) {
    setLastAction("Import impossible: bibliotheque Excel absente");
    renderStats();
    return;
  }

  if (!canUseSourceLinkPicker()) {
    setLastAction(buildSourceLinkUnavailableMessage());
    renderStats();
    return;
  }

  try {
    if (canUseAndroidSourcePicker()) {
      const sourceResult = await getBudgetSourcePlugin().pickSource();
      if (!sourceResult?.data) {
        throw new Error("Aucune donnee recue depuis la source Android");
      }

      await importWorkbookBuffer(base64ToArrayBuffer(sourceResult.data), sourceResult.name || "Budget_2025 Final.xlsx", {
        sourceLink: {
          kind: "android-document",
          uri: String(sourceResult.uri || ""),
          name: String(sourceResult.name || ""),
        },
        successMessage: `Source liee: ${sourceResult.name || "Budget_2025 Final.xlsx"}`,
      });
      return;
    }

    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [
        {
          description: "Classeur Excel Budget",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
            "application/vnd.ms-excel": [".xls"],
          },
        },
      ],
    });

    if (!handle) {
      return;
    }

    const file = await handle.getFile();
    await importWorkbookFile(file, {
      sourceLink: {
        kind: "file-handle",
        handle,
      },
      successMessage: `Source liee: ${file.name}`,
    });
  } catch (error) {
    if (error?.name === "AbortError") {
      setLastAction("Liaison de la source annulee");
      renderStats();
      return;
    }

    console.error(error);
    setLastAction("La source n'a pas pu etre ouverte avec ecriture");
    renderStats();
  }
}

async function importWorkbookFile(file, options = {}) {
  const buffer = await file.arrayBuffer();
  await importWorkbookBuffer(buffer, file.name, options);
}

async function importWorkbookBuffer(buffer, fileName, options = {}) {
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    bookFiles: true,
    bookVBA: true,
  });

  if (!isBudgetWorkbook(workbook)) {
    resetBudgetStateForInvalidWorkbook(fileName);
    return;
  }

  state.workbookName = fileName;
  state.workbook = workbook;
  state.sourceLink = options.sourceLink || null;
  state.sourceSafety = analyzeWorkbookSourceSafety(workbook, fileName);
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
  setLastAction(
    state.sourceSafety.allowDirectWrite
      ? options.successMessage || `Classeur charge: ${fileName}`
      : `${options.successMessage || `Classeur charge: ${fileName}`} - source protegee, export copie uniquement`
  );
  renderAll();
}

function resetBudgetStateForInvalidWorkbook(fileName) {
  state.workbookName = fileName;
  state.workbook = null;
  state.sourceLink = null;
  state.sourceSafety = createEmptySourceSafety();
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

function analyzeWorkbookSourceSafety(workbook, fileName) {
  const issues = [];
  const rawKeys = Array.isArray(workbook?.keys)
    ? workbook.keys
    : Object.keys(workbook?.files || {});
  const normalizedKeys = rawKeys.map((value) => String(value || ""));
  const definedNames = Array.isArray(workbook?.Workbook?.Names) ? workbook.Workbook.Names : [];
  const normalizedSheetNames = new Set((workbook?.SheetNames || []).map((value) => String(value || "").trim().toLowerCase()));

  if (
    normalizedKeys.some((value) =>
      /^xl\/(pivotTables|pivotCache|drawings|charts|slicers|externalLinks|persons|threadedComments|ctrlProps|connections)/i.test(
        value
      )
    )
  ) {
    issues.push("objets Excel avances");
  }

  if (
    definedNames.some((item) => {
      const reference = String(item?.Ref || "");
      return reference.includes("#REF!") || /\[\d+\]/.test(reference);
    })
  ) {
    issues.push("noms definis ou liaisons complexes");
  }

  if (
    normalizedSheetNames.has("journalier") &&
    normalizedSheetNames.has("recapitulatif") &&
    normalizedSheetNames.has("tcd")
  ) {
    issues.push("modele Budget 2025 structure");
  }

  if (/budget_2025 final/i.test(String(fileName || ""))) {
    issues.push("classeur source sensible");
  }

  if (!issues.length) {
    return {
      allowDirectWrite: false,
      reason: "Ecriture directe desactivee pour proteger l'integrite du classeur source.",
      issues: ["ecriture source preservee par precaution"],
    };
  }

  return {
    allowDirectWrite: false,
    reason: `Ecriture directe desactivee: ${issues.join(", ")}. Utilisez Exporter Excel sur une copie pour proteger le fichier source.`,
    issues,
  };
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
  state.activeView = normalizeActiveView(event.target.value);
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
    void deleteRecord(index);
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

async function deleteRecord(index) {
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
  const actionLabel = `Transaction supprimee: ${title}`;
  setLastAction(actionLabel);
  renderAll();
  try {
    await enqueueCloudSync(() => removeSingleTransactionFromSupabase(target.__id));
  } catch (error) {
    console.error(error);
    setLastAction(`${actionLabel} - sync cloud en echec`);
    renderAll();
  }
  await enqueueSourceSave({
    automatic: true,
    baseAction: actionLabel,
  });
}

async function onSaveRecord(event) {
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

  let actionLabel = "Transaction enregistree";
  if (state.editorMode === "edit" && state.editingIndex !== null && state.budget.rows[state.editingIndex]) {
    nextRecord.__id = state.budget.rows[state.editingIndex].__id;
    state.budget.rows[state.editingIndex] = nextRecord;
    actionLabel = "Transaction mise a jour";
    setLastAction(actionLabel);
  } else {
    state.budget.rows.push(nextRecord);
    actionLabel = "Nouvelle transaction ajoutee";
    setLastAction(actionLabel);
  }

  sortBudgetRowsInPlace(state.budget.rows);
  state.editingIndex = state.budget.rows.findIndex((row) => row.__id === nextRecord.__id);
  state.editorMode = "edit";

  persistDraft();
  renderAll();
  try {
    await enqueueCloudSync(() => syncSingleTransactionToSupabase(nextRecord));
  } catch (error) {
    console.error(error);
    setLastAction(`${actionLabel} - sync cloud en echec`);
    renderAll();
  }
  await enqueueSourceSave({
    automatic: true,
    baseAction: actionLabel,
  });
}

async function exportWorkbook() {
  if (!window.XLSX) {
    setLastAction("Export impossible: bibliotheque Excel absente");
    renderStats();
    return;
  }

  if (state.mode !== "budget") {
    setLastAction("Chargez ou restaurez des donnees pour exporter une copie");
    renderStats();
    return;
  }

  try {
    const exportPayload = buildExportPayload();

    if (canUseNativeExcelExport()) {
      setLastAction(exportPayload.preparingMessage);
      renderStats();
      await exportWorkbookWithNativeShare(exportPayload.workbook, exportPayload.fileName);
      setLastAction(exportPayload.sharedMessage);
    } else {
      XLSX.writeFile(exportPayload.workbook, exportPayload.fileName);
      setLastAction(exportPayload.successMessage);
    }

    renderStats();
  } catch (error) {
    console.error(error);
    setLastAction(buildExportErrorMessage(error));
    renderStats();
  }
}

function buildExportPayload() {
  if (shouldUseSimplifiedSafeExport()) {
    return {
      workbook: buildSimplifiedExportWorkbook(),
      fileName: buildSimplifiedExportFileName(),
      preparingMessage: "Preparation d'une copie simplifiee du journal",
      sharedMessage: "Copie simplifiee du journal exportee et partagee depuis l'app mobile",
      successMessage: "Copie simplifiee du journal exportee sans toucher au classeur source",
    };
  }

  applyBudgetRowsToWorkbook(state.workbook, state.budget);

  return {
    workbook: state.workbook,
    fileName: buildExportFileName(),
    preparingMessage: "Preparation du fichier Excel pour l'app mobile",
    sharedMessage: "Classeur exporte et partage depuis l'app mobile",
    successMessage: "Classeur exporte avec Journalier mis a jour",
  };
}

function shouldUseSimplifiedSafeExport() {
  return Boolean(state.mode === "budget" && (!state.workbook || !state.sourceSafety.allowDirectWrite));
}

function enqueueSourceSave(options = {}) {
  if (!canSaveToSource()) {
    return Promise.resolve(false);
  }

  sourceSaveQueue = sourceSaveQueue
    .catch(() => false)
    .then(() => saveSourceWorkbook(options));

  return sourceSaveQueue;
}

async function saveSourceWorkbook(options = {}) {
  const automatic = Boolean(options.automatic);
  const baseAction = String(options.baseAction || "");

  if (!window.XLSX) {
    if (!automatic) {
      setLastAction("Enregistrement source impossible: bibliotheque Excel absente");
      renderStats();
    }
    return false;
  }

  if (!state.workbook || state.mode !== "budget") {
    if (!automatic) {
      setLastAction("Chargez d'abord Budget_2025 Final.xlsx");
      renderStats();
    }
    return false;
  }

  if (!canSaveToSource()) {
    if (!automatic) {
      setLastAction(buildSourceLinkUnavailableMessage());
      renderStats();
    }
    return false;
  }

  try {
    applyBudgetRowsToWorkbook(state.workbook, state.budget);
    setLastAction(
      automatic && baseAction
        ? `${baseAction} - synchronisation de la source`
        : "Ecriture en cours dans le fichier source"
    );
    renderStats();
    await saveWorkbookToLinkedSource(state.workbook, state.sourceLink);
    setLastAction(
      automatic && baseAction
        ? `${baseAction} - source mise a jour`
        : `Fichier source mis a jour: ${state.workbookName}`
    );
  } catch (error) {
    console.error(error);
    setLastAction(automatic && baseAction ? `${baseAction} - ${buildSourceSaveErrorMessage(error)}` : buildSourceSaveErrorMessage(error));
    renderStats();
    return false;
  }

  renderStats();
  return true;
}

function buildExportFileName() {
  const baseName = state.workbookName
    ? state.workbookName.replace(/\.(xlsx|xls)$/i, "")
    : "Budget_2025 Final";

  return `${sanitizeExportFileName(baseName)}-card-view.xlsx`;
}

function buildSimplifiedExportFileName() {
  const baseName = state.workbookName
    ? state.workbookName.replace(/\.(xlsx|xls)$/i, "")
    : "Budget_2025 Final";

  return `${sanitizeExportFileName(baseName)}-journalier-safe.xlsx`;
}

function buildSimplifiedExportWorkbook() {
  const workbook = XLSX.utils.book_new();
  const journalSheet = {};
  const lastRow = START_ROW + Math.max(state.budget.rows.length, state.budget.categories.length) + 2;

  journalSheet[`${CATEGORY_LIST_COL}${HEADER_ROW}`] = { t: "s", v: state.budget.headers[1] || "Categories" };
  journalSheet[`${DATE_COL}${HEADER_ROW}`] = { t: "s", v: state.budget.headers[0] || "Date" };
  journalSheet[`${CATEGORY_COL}${HEADER_ROW}`] = { t: "s", v: state.budget.headers[1] || "Categories" };
  journalSheet[`${VALUE_COL}${HEADER_ROW}`] = { t: "s", v: state.budget.headers[2] || "Value" };

  state.budget.categories.forEach((category, index) => {
    journalSheet[`${CATEGORY_LIST_COL}${START_ROW + index}`] = {
      t: "s",
      v: category,
    };
  });

  state.budget.rows.forEach((row, index) => {
    const sheetRow = START_ROW + index;
    const isoDate = normalizeDateValue(row.Date);
    const amountNumber = parseAmount(row.Value);

    if (isoDate) {
      journalSheet[`${DATE_COL}${sheetRow}`] = {
        t: "n",
        v: isoDateToExcelSerial(isoDate),
        z: "m/d/yyyy",
      };
    }

    if (row.Categories) {
      journalSheet[`${CATEGORY_COL}${sheetRow}`] = {
        t: "s",
        v: row.Categories,
      };
    }

    if (Number.isFinite(amountNumber)) {
      journalSheet[`${VALUE_COL}${sheetRow}`] = {
        t: "n",
        v: amountNumber,
      };
    } else if (row.Value) {
      journalSheet[`${VALUE_COL}${sheetRow}`] = {
        t: "s",
        v: row.Value,
      };
    }
  });

  journalSheet["!ref"] = `B2:F${lastRow}`;

  const infoSheet = XLSX.utils.aoa_to_sheet([
    ["Budget 2025 Card View"],
    ["Export simplifie du journal"],
    [
      "Cette copie preserve les transactions Journalier et la liste de categories, sans reecrire le modele Excel source complexe."
    ],
    ["Fichier source", state.workbookName || "Budget_2025 Final.xlsx"],
    ["Date export", new Date().toISOString()],
  ]);

  XLSX.utils.book_append_sheet(workbook, infoSheet, "Infos");
  XLSX.utils.book_append_sheet(workbook, journalSheet, JOURNAL_SHEET_NAME);

  return workbook;
}

function sanitizeExportFileName(value) {
  const normalized = String(value || "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-")
    .replace(/\s+/g, " ")
    .trim();

  return normalized || "Budget_2025 Final";
}

function canUseSourceLinkPicker() {
  return Boolean((canUseBrowserSourcePicker() || canUseAndroidSourcePicker()) && state.sourceSafety.allowDirectWrite);
}

function hasLinkedWritableSource() {
  return Boolean(
    (state.sourceLink?.kind === "file-handle" && state.sourceLink?.handle) ||
      (state.sourceLink?.kind === "android-document" && state.sourceLink?.uri)
  );
}

function canSaveToSource() {
  return Boolean(state.mode === "budget" && state.workbook && state.sourceSafety.allowDirectWrite && hasLinkedWritableSource());
}

async function saveWorkbookToLinkedSource(workbook, sourceLink) {
  if (!sourceLink) {
    throw new Error("Aucune source liee");
  }

  if (sourceLink.kind === "android-document") {
    await saveWorkbookToAndroidSource(workbook, sourceLink);
    return;
  }

  if (!sourceLink.handle) {
    throw new Error("Aucune source liee");
  }

  const hasPermission = await ensureSourceWritePermission(sourceLink.handle);
  if (!hasPermission) {
    throw new Error("Permission d'ecriture refusee");
  }

  const writable = await sourceLink.handle.createWritable();

  try {
    await writable.write(
      XLSX.write(workbook, {
        bookType: getWorkbookBookType(sourceLink.handle.name || state.workbookName),
        type: "array",
      })
    );
  } finally {
    await writable.close();
  }
}

async function saveWorkbookToAndroidSource(workbook, sourceLink) {
  const BudgetSource = getBudgetSourcePlugin();
  if (!BudgetSource?.saveSource) {
    throw new Error("Plugin source Android indisponible");
  }

  await BudgetSource.saveSource({
    uri: sourceLink.uri,
    data: XLSX.write(workbook, {
      bookType: getWorkbookBookType(sourceLink.name || state.workbookName),
      type: "base64",
    }),
    fileName: sourceLink.name || state.workbookName,
  });
}

async function ensureSourceWritePermission(fileHandle) {
  const permissionOptions = { mode: "readwrite" };

  if (typeof fileHandle.queryPermission === "function") {
    const currentPermission = await fileHandle.queryPermission(permissionOptions);
    if (currentPermission === "granted") {
      return true;
    }
  }

  if (typeof fileHandle.requestPermission === "function") {
    const requestedPermission = await fileHandle.requestPermission(permissionOptions);
    return requestedPermission === "granted";
  }

  return true;
}

function getWorkbookBookType(fileName) {
  const normalized = String(fileName || "").toLowerCase();

  if (normalized.endsWith(".xls")) {
    return "biff8";
  }

  return "xlsx";
}

function buildSourceLinkUnavailableMessage() {
  if (state.mode === "budget" && !state.sourceSafety.allowDirectWrite) {
    return state.sourceSafety.reason;
  }

  if (hasLinkedWritableSource()) {
    return "La source est deja liee";
  }

  if (canUseAndroidSourcePicker()) {
    return "Sur Android, utilisez Lier la source pour autoriser l'ecriture directe";
  }

  if (location.protocol === "file:") {
    return "Mode local: publiez l'app en HTTPS puis utilisez Lier la source pour ecrire dans le fichier d'origine";
  }

  if (!canUseSourceLinkPicker()) {
    return "Votre navigateur ne permet pas encore de lier directement le fichier source";
  }

  return "Liez d'abord la source avec le bouton dedie";
}

function buildSourceSaveErrorMessage(error) {
  const rawMessage = String(error?.message || "").toLowerCase();

  if (error?.name === "NotAllowedError" || rawMessage.includes("permission")) {
    return "Ecriture refusee sur le fichier source";
  }

  if (rawMessage.includes("stream") || rawMessage.includes("document")) {
    return "Le document source n'a pas accepte la reecriture directe";
  }

  return "L'enregistrement direct dans la source a echoue";
}

function getCapacitorRuntime() {
  return window.Capacitor || null;
}

function getCapacitorPlatformName() {
  return String(getCapacitorRuntime()?.getPlatform?.() || "");
}

function isAndroidNativeRuntime() {
  return isNativeAppRuntime() && getCapacitorPlatformName() === "android";
}

function getBudgetSourcePlugin() {
  if (budgetSourcePlugin) {
    return budgetSourcePlugin;
  }

  const Capacitor = getCapacitorRuntime();
  if (!Capacitor) {
    return null;
  }

  budgetSourcePlugin = typeof Capacitor.registerPlugin === "function"
    ? Capacitor.registerPlugin("BudgetSource")
    : Capacitor.Plugins?.BudgetSource || null;

  return budgetSourcePlugin;
}

function canUseBrowserSourcePicker() {
  return Boolean(window.isSecureContext && typeof window.showOpenFilePicker === "function");
}

function canUseAndroidSourcePicker() {
  return Boolean(isAndroidNativeRuntime() && getBudgetSourcePlugin());
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

function base64ToArrayBuffer(base64Value) {
  const binary = window.atob(String(base64Value || ""));
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes.buffer;
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
  renderCloudPanel();
  renderDraftStatus();
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

  [JOURNAL_SHEET_NAME, RECAP_SHEET_NAME, ANALYSIS_VIEW_NAME].forEach((viewName) => {
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

  if (state.activeView === ANALYSIS_VIEW_NAME) {
    refs.cardsKicker.textContent = "Comparaisons";
    refs.cardsTitle.textContent = "Income, expenses et savings en perspective";
    refs.cardsCaption.textContent =
      "Cette vue ajoute des comparaisons et des graphiques pertinents a partir de Journalier pour suivre l'equilibre entre income, expenses, savings et cash.";
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
  const analysisActive = hasBudget && state.activeView === ANALYSIS_VIEW_NAME;
  const draft = readStoredDraft();
  const hasStoredDraft = Boolean(draft && draft.mode === "budget" && Array.isArray(draft.rows));
  const availableYears = getAvailableRecapYears();
  const availableMonths = getAvailableRecapMonths(state.recapFilters.year);

  refs.sheetSelect.disabled = !hasBudget;
  refs.searchInput.disabled = !hasBudget;
  refs.searchInput.placeholder = recapActive
    ? "Chercher un poste ou une categorie du recap..."
    : analysisActive
      ? "Periode, indicateur, valeur..."
      : "Categorie, date, valeur...";
  refs.openSourceButton.disabled = !window.XLSX || !canUseSourceLinkPicker();
  refs.openSourceButton.textContent = state.mode === "budget" && !state.sourceSafety.allowDirectWrite ? "Source protegee" : hasLinkedWritableSource() ? "Relier la source" : "Lier la source";
  refs.openSourceButton.title = state.mode === "budget" && !state.sourceSafety.allowDirectWrite
    ? state.sourceSafety.reason
    : canUseSourceLinkPicker()
      ? "Ouvre le classeur avec autorisation d'ecriture directe"
      : buildSourceLinkUnavailableMessage();
  refs.saveSourceButton.disabled = !canSaveToSource();
  refs.saveSourceButton.textContent = state.mode === "budget" && !state.sourceSafety.allowDirectWrite ? "Source preservee" : "Enregistrer la source";
  refs.saveSourceButton.title = canSaveToSource()
    ? "Ecrit les changements dans le fichier d'origine sans passer par une copie exportee"
    : buildSourceLinkUnavailableMessage();
  refs.saveDraftButton.disabled = !hasBudget;
  refs.saveDraftButton.title = hasBudget
    ? "Memorise vos donnees actuelles dans le navigateur pour reprendre plus tard"
    : "Chargez ou restaurez un budget avant d'enregistrer un brouillon local";
  refs.restoreDraftButton.disabled = !hasStoredDraft;
  refs.restoreDraftButton.title = hasStoredDraft
    ? "Recharge le dernier brouillon local memorise dans l'app"
    : "Aucun brouillon local disponible pour le moment";
  refs.recapYearField.classList.toggle("hidden", !hasBudget);
  refs.recapMonthField.classList.toggle("hidden", !hasBudget);
  refs.recapYearSelect.disabled = !hasBudget || !availableYears.length;
  refs.recapMonthSelect.disabled = !hasBudget || !availableMonths.length;
  refs.addButton.disabled = !journalActive;
  refs.exportButton.disabled = !hasBudget || !window.XLSX;
  refs.saveButton.disabled = !journalActive;
  refs.cancelButton.disabled = !journalActive;
}

function renderCloudPanel() {
  const cloudReady = state.cloud.ready;
  const signedIn = hasSupabaseSession();
  const spaceSelected = hasCloudSpaceSelected();
  const busy = state.cloud.syncBusy;
  const canPublish = canUseSupabaseCloud() && hasLocalBudgetData();

  refs.cloudStatus.textContent = state.cloud.status;
  refs.cloudEmailInput.value = refs.cloudEmailInput.matches(":focus")
    ? refs.cloudEmailInput.value
    : state.cloud.email;
  refs.cloudCodeInput.value = refs.cloudCodeInput.matches(":focus")
    ? refs.cloudCodeInput.value
    : state.cloud.space.joinCode;

  refs.cloudEmailInput.disabled = !cloudReady || busy || signedIn;
  refs.cloudCodeInput.disabled = !cloudReady || busy || !signedIn;
  refs.cloudMagicLinkButton.disabled = !cloudReady || busy || signedIn;
  refs.cloudSignOutButton.disabled = !cloudReady || busy || !signedIn;
  refs.cloudCreateSpaceButton.disabled = !cloudReady || busy || !signedIn;
  refs.cloudJoinSpaceButton.disabled = !cloudReady || busy || !signedIn;
  refs.cloudPushButton.disabled = !canPublish || busy;
  refs.cloudPullButton.disabled = !canUseSupabaseCloud() || busy;

  refs.cloudMagicLinkButton.textContent = busy && !signedIn ? "Connexion..." : "Lien magique";
  refs.cloudSignOutButton.textContent = busy && signedIn ? "Patientez..." : "Deconnexion";

  const identityLabel = signedIn
    ? `Compte: ${state.cloud.user?.email || state.cloud.email || "connecte"}`
    : "Compte: non connecte";
  const spaceLabel = spaceSelected
    ? `Espace: ${state.cloud.space.name || "budget partage"}`
    : "Espace: aucun";
  const codeLabel = spaceSelected && state.cloud.space.joinCode
    ? `Code partage: ${state.cloud.space.joinCode}`
    : "Code partage: a creer ou rejoindre";
  const pullLabel = state.cloud.lastPulledAt
    ? `Dernier chargement: ${formatDraftSavedAt(state.cloud.lastPulledAt)}`
    : "Dernier chargement: aucun";
  const pushLabel = state.cloud.lastPushedAt
    ? `Derniere publication: ${formatDraftSavedAt(state.cloud.lastPushedAt)}`
    : "Derniere publication: aucune";

  refs.cloudSpaceHint.textContent = [identityLabel, spaceLabel, codeLabel, pullLabel, pushLabel].join(" | ");
}

function renderDraftStatus() {
  const draft = readStoredDraft();
  const hasStoredDraft = Boolean(draft && draft.mode === "budget" && Array.isArray(draft.rows));

  if (!hasStoredDraft) {
    refs.draftStatus.textContent = "Aucun brouillon local memorise.";
    return;
  }

  const savedAtLabel = formatDraftSavedAt(draft.savedAt);
  const suffix = savedAtLabel ? ` Derniere sauvegarde ${savedAtLabel}.` : "";

  if (canUseSupabaseCloud()) {
    refs.draftStatus.textContent = `Mode cloud partage actif.${suffix}`;
    return;
  }

  if (state.mode === "budget" && !state.workbook) {
    refs.draftStatus.textContent = `Mode autonome local actif.${suffix}`;
    return;
  }

  if (state.mode === "budget") {
    refs.draftStatus.textContent = `Brouillon local pret en secours.${suffix}`;
    return;
  }

  refs.draftStatus.textContent = `Brouillon local disponible.${suffix} Cliquez sur Restaurer pour reprendre vos donnees.`;
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
    refs.metricFile.textContent = buildWorkbookLabel();
    refs.metricSave.textContent = canUseSourceLinkPicker() ? "Chargez ou liez la source" : "Chargez le fichier";
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
    refs.metricFile.textContent = buildWorkbookLabel();
    refs.metricSave.textContent = getSaveCapabilityLabel();
    return;
  }

  if (state.activeView === ANALYSIS_VIEW_NAME) {
    const analysisView = buildLiveAnalysisView();
    refs.recordsLabel.textContent = "Periodes";
    refs.recordsCount.textContent = String(analysisView.seriesRows.length);
    refs.columnsLabel.textContent = "Graphiques";
    refs.columnsCount.textContent = String(analysisView.chartCount);
    refs.activeSheet.textContent = ANALYSIS_VIEW_NAME;
    refs.lastAction.textContent = state.lastAction;
    refs.metricMode.textContent = `Vue analyse - ${analysisView.periodLabel}`;
    refs.metricFile.textContent = buildWorkbookLabel();
    refs.metricSave.textContent = getSaveCapabilityLabel();
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
  refs.metricFile.textContent = buildWorkbookLabel();
  refs.metricSave.textContent = getSaveCapabilityLabel();
}

function getExportCapabilityLabel() {
  return canUseNativeExcelExport() ? "Partage natif Excel" : "Export vers le classeur";
}

function getSaveCapabilityLabel() {
  if (canUseSupabaseCloud()) {
    return "Supabase partage actif";
  }

  if (canSaveToSource()) {
    return "Source liee - auto-save actif";
  }

  if (state.mode === "budget" && !state.workbook) {
    return "Mode autonome - export copie locale";
  }

  if (state.mode === "budget" && !state.sourceSafety.allowDirectWrite) {
    return "Source preservee - export copie uniquement";
  }

  if (canUseBrowserSourcePicker() || canUseAndroidSourcePicker()) {
    return "Export ou liaison source";
  }

  return getExportCapabilityLabel();
}

function buildWorkbookLabel() {
  if (hasCloudSpaceSelected()) {
    const cloudName = state.cloud.space.name || state.cloud.space.joinCode || "Budget partage";
    if (state.workbookName) {
      return `${cloudName} - cloud partage`;
    }

    return `${cloudName} - cloud partage`;
  }

  if (!state.workbookName) {
    return state.mode === "budget" ? "Donnees locales" : "Aucun fichier";
  }

  if (state.mode === "budget" && !state.workbook) {
    return `${state.workbookName} - donnees locales`;
  }

  if (state.mode === "budget" && !state.sourceSafety.allowDirectWrite) {
    return `${state.workbookName} - source protegee`;
  }

  return hasLinkedWritableSource() ? `${state.workbookName} - source liee` : state.workbookName;
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

  if (state.activeView === ANALYSIS_VIEW_NAME) {
    renderAnalysisView();
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
      "<p>Importez une premiere fois votre budget ou restaurez vos donnees locales pour reconstruire la synthese.</p>",
    ].join("");
    return;
  }

  refs.cardsEmpty.classList.add("hidden");
  refs.recapView.innerHTML = createRecapMarkup(recapView);
}

function renderAnalysisView() {
  refs.cardsGrid.classList.add("hidden");
  refs.recapView.classList.remove("hidden");

  const analysisView = buildLiveAnalysisView();

  if (!analysisView.available) {
    refs.cardsEmpty.classList.remove("hidden");
    refs.recapView.classList.add("hidden");
    refs.cardsEmpty.innerHTML = [
      "<strong>La vue Comparaisons n'est pas disponible.</strong>",
      "<p>Importez une premiere fois votre budget ou restaurez vos donnees locales pour construire les graphiques et comparaisons.</p>",
    ].join("");
    return;
  }

  refs.cardsEmpty.classList.add("hidden");
  refs.recapView.innerHTML = createAnalysisMarkup(analysisView);
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

function buildLiveAnalysisView() {
  if (state.mode !== "budget") {
    return {
      available: false,
      periodLabel: "Toutes les donnees",
      chartCount: 0,
      metricCards: [],
      comparisonRows: [],
      seriesRows: [],
      snapshotDate: "",
      transactionCount: 0,
      trendTitle: "",
      trendSubtitle: "",
    };
  }

  const filteredRows = getFilteredRecapSourceRows();
  const snapshot = computeMetricSnapshot(buildActualAmountMap(filteredRows));
  const allSeriesRows = buildAnalysisSeriesRows();
  const seriesRows = filterAnalysisSeriesRows(allSeriesRows);

  return {
    available: true,
    periodLabel: buildRecapPeriodLabel(),
    chartCount: 2,
    metricCards: buildAnalysisMetricCards(snapshot),
    comparisonRows: buildAnalysisComparisonRows(snapshot),
    seriesRows,
    snapshotDate: state.recap.snapshotDate,
    transactionCount: filteredRows.length,
    trendTitle: buildAnalysisTrendTitle(),
    trendSubtitle: buildAnalysisTrendSubtitle(allSeriesRows.length, seriesRows.length),
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

function computeMetricSnapshot(actualMap) {
  const income = Math.abs(getActualAmount(actualMap, "Income"));
  const savings = Math.abs(getActualAmount(actualMap, "Savings"));
  const seasonalSavings = Math.abs(getActualAmount(actualMap, "Savings for seasonal exp."));
  const totalSavings = savings + seasonalSavings;
  const totalExpenses = computeTotalExpenses(actualMap);
  const cash = income - totalSavings - totalExpenses;

  return {
    income,
    expenses: totalExpenses,
    savings,
    seasonalSavings,
    totalSavings,
    cash,
  };
}

function buildRecapMetrics(actualMap) {
  const snapshot = computeMetricSnapshot(actualMap);

  return [
    { label: "Income", value: snapshot.income, tone: "positive" },
    { label: "Expenses", value: snapshot.expenses, tone: "negative" },
    { label: "Savings", value: snapshot.savings, tone: "neutral" },
    { label: "Cash", value: snapshot.cash, tone: snapshot.cash >= 0 ? "positive" : "negative" },
    { label: "Seasonal Savings", value: snapshot.seasonalSavings, tone: "neutral" },
  ];
}

function buildAnalysisMetricCards(snapshot) {
  return [
    { label: "Income", value: snapshot.income, tone: "positive" },
    { label: "Expenses", value: snapshot.expenses, tone: "negative" },
    { label: "Savings", value: snapshot.totalSavings, tone: "neutral" },
    { label: "Cash", value: snapshot.cash, tone: snapshot.cash >= 0 ? "positive" : "negative" },
  ];
}

function buildAnalysisComparisonRows(snapshot) {
  const rows = [
    {
      label: "Income",
      value: snapshot.income,
      displayValue: formatCurrency(snapshot.income),
      tone: "positive",
      caption: "Volume des revenus retenus pour la periode filtree.",
    },
    {
      label: "Expenses",
      value: snapshot.expenses,
      displayValue: formatCurrency(snapshot.expenses),
      tone: "negative",
      caption: "Somme des depenses hors postes de savings.",
    },
    {
      label: "Savings",
      value: snapshot.totalSavings,
      displayValue: formatCurrency(snapshot.totalSavings),
      tone: "neutral",
      caption: "Savings total incluant les seasonal savings.",
    },
    {
      label: "Cash",
      value: Math.abs(snapshot.cash),
      displayValue: formatSignedCurrency(snapshot.cash),
      tone: snapshot.cash >= 0 ? "positive" : "negative",
      caption: "Disponible net apres expenses et savings.",
    },
  ];

  const maxValue = Math.max(...rows.map((row) => row.value), 1);

  return rows.map((row) => ({
    ...row,
    percentage: buildChartScale(row.value, maxValue, 14),
  }));
}

function buildAnalysisSeriesRows() {
  const mode = getAnalysisSeriesMode();
  const buckets = new Map();

  state.budget.rows.forEach((row) => {
    const dateParts = getBudgetRowDateParts(row);
    if (!dateParts) {
      return;
    }

    const bucket = getAnalysisBucketForDate(dateParts, mode);
    if (!bucket) {
      return;
    }

    if (!buckets.has(bucket.key)) {
      buckets.set(bucket.key, {
        ...bucket,
        rows: [],
      });
    }

    buckets.get(bucket.key).rows.push(row);
  });

  let seriesRows = Array.from(buckets.values())
    .sort((left, right) => String(left.sortKey).localeCompare(String(right.sortKey)));

  if (mode === "recent_months") {
    seriesRows = seriesRows.slice(-8);
  }

  if (mode === "selected_period_window") {
    const selectedKey = buildYearMonthKey(state.recapFilters.year, state.recapFilters.month);
    const selectedIndex = seriesRows.findIndex((row) => row.key === selectedKey);
    seriesRows = selectedIndex >= 0
      ? seriesRows.slice(Math.max(0, selectedIndex - 5), selectedIndex + 1)
      : seriesRows.slice(-6);
  }

  return seriesRows.map((row) => {
    const snapshot = computeMetricSnapshot(buildActualAmountMap(row.rows));

    return {
      key: row.key,
      label: row.label,
      shortLabel: row.shortLabel,
      income: snapshot.income,
      expenses: snapshot.expenses,
      savings: snapshot.totalSavings,
      cash: snapshot.cash,
    };
  });
}

function getAnalysisSeriesMode() {
  if (state.recapFilters.year !== "all" && state.recapFilters.month === "all") {
    return "year_months";
  }

  if (state.recapFilters.year === "all" && state.recapFilters.month !== "all") {
    return "month_across_years";
  }

  if (state.recapFilters.year !== "all" && state.recapFilters.month !== "all") {
    return "selected_period_window";
  }

  return "recent_months";
}

function getAnalysisBucketForDate(dateParts, mode) {
  if (mode === "year_months") {
    if (dateParts.year !== state.recapFilters.year) {
      return null;
    }

    return {
      key: dateParts.month,
      sortKey: buildYearMonthKey(dateParts.year, dateParts.month),
      label: formatMonthLabel(dateParts.month),
      shortLabel: formatMonthShortLabel(dateParts.year, dateParts.month, false),
    };
  }

  if (mode === "month_across_years") {
    if (dateParts.month !== state.recapFilters.month) {
      return null;
    }

    return {
      key: dateParts.year,
      sortKey: dateParts.year,
      label: dateParts.year,
      shortLabel: dateParts.year,
    };
  }

  return {
    key: buildYearMonthKey(dateParts.year, dateParts.month),
    sortKey: buildYearMonthKey(dateParts.year, dateParts.month),
    label: formatMonthShortLabel(dateParts.year, dateParts.month, true),
    shortLabel: formatMonthShortLabel(dateParts.year, dateParts.month, false),
  };
}

function filterAnalysisSeriesRows(rows) {
  if (!state.search) {
    return rows;
  }

  return rows.filter((row) => matchesAnalysisSearch(row, state.search));
}

function matchesAnalysisSearch(row, query) {
  const haystack = [
    row.label,
    row.shortLabel,
    formatCurrency(row.income),
    formatCurrency(row.expenses),
    formatCurrency(row.savings),
    formatSignedCurrency(row.cash),
  ].join(" ").toLowerCase();

  return haystack.includes(query);
}

function buildAnalysisTrendTitle() {
  const mode = getAnalysisSeriesMode();

  if (mode === "year_months") {
    return `Lecture mensuelle de ${state.recapFilters.year}`;
  }

  if (mode === "month_across_years") {
    return `Comparaison annuelle pour ${formatMonthLabel(state.recapFilters.month)}`;
  }

  if (mode === "selected_period_window") {
    return `Fenetre autour de ${buildRecapPeriodLabel()}`;
  }

  return "Dernieres periodes disponibles";
}

function buildAnalysisTrendSubtitle(totalCount, visibleCount) {
  const mode = getAnalysisSeriesMode();
  let baseSubtitle = "Income, expenses et savings total sont compares periode par periode.";

  if (mode === "year_months") {
    baseSubtitle = "Chaque groupe compare les mois de l'annee filtree.";
  } else if (mode === "month_across_years") {
    baseSubtitle = "Le meme mois est compare d'une annee a l'autre.";
  } else if (mode === "selected_period_window") {
    baseSubtitle = "La selection montre la periode demandee et les periodes precedentes les plus proches.";
  }

  if (state.search && totalCount !== visibleCount) {
    return `${baseSubtitle} Recherche active: ${visibleCount} periode(s) visibles sur ${totalCount}.`;
  }

  return baseSubtitle;
}

function buildYearMonthKey(year, month) {
  return `${year}-${month}`;
}

function formatMonthShortLabel(year, month, includeYear) {
  const label = new Intl.DateTimeFormat("fr-CA", {
    month: "short",
    ...(includeYear ? { year: "numeric" } : {}),
    timeZone: "UTC",
  }).format(new Date(Date.UTC(Number(year), Number(month) - 1, 1, 12)));

  return label.charAt(0).toUpperCase() + label.slice(1);
}

function buildChartScale(value, maxValue, minimumPercent = 10) {
  if (!Number.isFinite(value) || value <= 0 || !Number.isFinite(maxValue) || maxValue <= 0) {
    return 0;
  }

  return Math.max(minimumPercent, (value / maxValue) * 100);
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

function createAnalysisMarkup(analysisView) {
  const comparisonMaxValue = Math.max(
    ...analysisView.comparisonRows.map((row) => row.value),
    1
  );
  const trendMaxValue = Math.max(
    ...analysisView.seriesRows.flatMap((row) => [row.income, row.expenses, row.savings]),
    1
  );
  const transactionLabel = `${analysisView.transactionCount} transaction${analysisView.transactionCount > 1 ? "s" : ""} retenue${analysisView.transactionCount > 1 ? "s" : ""}.`;

  return `
    <div class="analysis-shell">
      <div class="recap-metrics">
        ${analysisView.metricCards.map((metric) => createRecapMetricMarkup(metric)).join("")}
      </div>
      <div class="recap-note">
        <strong>Source:</strong> vue analytique construite depuis <code>${JOURNAL_SHEET_NAME}</code>.
        ${analysisView.snapshotDate ? ` Snapshot date: ${escapeHtml(analysisView.snapshotDate)}.` : ""}
        <br><strong>Periode analysee:</strong> ${escapeHtml(analysisView.periodLabel)}. ${escapeHtml(transactionLabel)}
        <br><strong>Lecture:</strong> les graphiques comparent Income, Expenses et Savings. Savings inclut aussi les seasonal savings.
      </div>

      <section class="recap-section">
        <div class="recap-section-head">
          <h3>Comparaison des indicateurs</h3>
          <p>Lecture rapide des masses budgetaires sur la periode filtree.</p>
        </div>
        <div class="analysis-bar-list">
          ${analysisView.comparisonRows.map((row) => createAnalysisMetricBarMarkup(row, comparisonMaxValue)).join("")}
        </div>
      </section>

      <section class="recap-section">
        <div class="recap-section-head">
          <h3>${escapeHtml(analysisView.trendTitle)}</h3>
          <p>${escapeHtml(analysisView.trendSubtitle)}</p>
        </div>
        ${analysisView.seriesRows.length ? `
          <div class="analysis-period-chart">
            ${analysisView.seriesRows.map((row) => createAnalysisPeriodGroupMarkup(row, trendMaxValue)).join("")}
          </div>
          <div class="analysis-legend">
            <span class="analysis-legend-item"><span class="analysis-legend-swatch income"></span>Income</span>
            <span class="analysis-legend-item"><span class="analysis-legend-swatch expenses"></span>Expenses</span>
            <span class="analysis-legend-item"><span class="analysis-legend-swatch savings"></span>Savings</span>
          </div>
        ` : `
          <div class="empty-form">
            Aucun groupe de comparaison ne correspond a la recherche ou au filtre actif.
          </div>
        `}
      </section>

      ${createRecapTableMarkup(
        "Tableau de comparaison",
        "Resume periode par periode pour comparer Income, Expenses, Savings et Cash.",
        ["Periode", "Income", "Expenses", "Savings", "Cash"],
        analysisView.seriesRows.map((row) => ({
          cells: [
            { value: row.label, numeric: false },
            { value: formatCurrency(row.income), numeric: true },
            { value: formatCurrency(row.expenses), numeric: true },
            { value: formatCurrency(row.savings), numeric: true },
            { value: formatSignedCurrency(row.cash), numeric: true },
          ],
          total: false,
        }))
      )}
    </div>
  `;
}

function createAnalysisMetricBarMarkup(row, maxValue) {
  return `
    <article class="analysis-bar-card analysis-bar-card-${row.tone}">
      <div class="analysis-bar-head">
        <span>${escapeHtml(row.label)}</span>
        <strong>${escapeHtml(row.displayValue)}</strong>
      </div>
      <div class="analysis-bar-track">
        <span
          class="analysis-bar-fill analysis-bar-fill-${row.tone}"
          style="width: ${buildChartScale(row.value, maxValue, 14)}%;"
        ></span>
      </div>
      <p class="analysis-bar-caption">${escapeHtml(row.caption)}</p>
    </article>
  `;
}

function createAnalysisPeriodGroupMarkup(row, maxValue) {
  return `
    <article class="analysis-period-group">
      <div class="analysis-period-bars">
        <span class="analysis-mini-bar income" style="height: ${buildChartScale(row.income, maxValue, 10)}%;"></span>
        <span class="analysis-mini-bar expenses" style="height: ${buildChartScale(row.expenses, maxValue, 10)}%;"></span>
        <span class="analysis-mini-bar savings" style="height: ${buildChartScale(row.savings, maxValue, 10)}%;"></span>
      </div>
      <strong class="analysis-period-label">${escapeHtml(row.shortLabel)}</strong>
      <span class="analysis-period-meta">${escapeHtml(formatSignedCurrency(row.cash))}</span>
    </article>
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

  if (state.activeView !== JOURNAL_SHEET_NAME) {
    const isRecapView = state.activeView === RECAP_SHEET_NAME;
    refs.formTitle.textContent = isRecapView ? "Vue recap" : "Vue comparaisons";
    refs.formSubtitle.textContent = "Lecture seule dans l'app.";
    refs.formFields.innerHTML = isRecapView
      ? `
        <div class="empty-form">
          Cette vue n'edite pas directement la feuille Recapitulatif d'Excel.
          Elle reconstruit une synthese lisible a partir de ${RECAP_SHEET_NAME}, ${TCD_SHEET_NAME}
          et de vos transactions ${JOURNAL_SHEET_NAME}. Pour modifier les donnees, revenez sur
          la vue Journalier.
        </div>
      `
      : `
        <div class="empty-form">
          Cette vue ajoute des comparaisons et des graphiques a partir des ecritures de ${JOURNAL_SHEET_NAME}.
          Elle est destinee a l'analyse. Pour modifier les donnees, revenez sur la vue Journalier.
        </div>
      `;
    return;
  }

  refs.formSubtitle.textContent = state.workbook
    ? "Ajoutez ou modifiez vos transactions. Les vues de recap et de comparaison se recalculent aussitot."
    : canUseSupabaseCloud()
      ? "Mode cloud partage. Ajoutez ou modifiez vos transactions, Supabase les republie pour les autres personnes."
      : "Mode autonome local. Ajoutez ou modifiez vos transactions, les graphiques se mettent a jour aussitot et l'export reste disponible.";

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
  refs.formSubtitle.textContent = state.workbook
    ? "Saisie directe de Journalier!D:F avec categories predefinies depuis la colonne B."
    : canUseSupabaseCloud()
      ? "Mode cloud partage: chaque enregistrement met a jour vos vues locales et synchronise Supabase."
      : "Mode autonome local: vos categories, recapitulatifs et graphiques se mettent a jour a chaque enregistrement.";
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

function formatDraftSavedAt(value) {
  if (!value) {
    return "";
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }

  return new Intl.DateTimeFormat("fr-CA", {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(parsed);
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
