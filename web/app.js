import {
  createViewerAccount,
  diagnoseCloud,
  cloudAccessLabel,
  initializeCloud,
  isCloudConnected,
  isCloudReadOnly,
  loadCloudState,
  mergeCloudState,
  queueCloudSave,
  grantViewer,
  revokeViewer,
  signIn,
  signOut,
  signUp
} from "./cloud-sync.js?v=18";
import { extractOIRecommendations, extractPostOiDocument, requestDailyReview } from "./openai-ai.js?v=6";
import {
  configureMicrosoft,
  connectMicrosoft,
  checkCalendarAvailability,
  createTeamsEvent,
  disconnectMicrosoft,
  initializeMicrosoft,
  isMicrosoftConnected,
  microsoftClientId
} from "./microsoft-graph.js?v=2";

let pdfjsLib;

async function getPdfLibrary() {
  if (pdfjsLib) return pdfjsLib;
  pdfjsLib = await import("./vendor/pdf.min.mjs");
  pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.mjs";
  return pdfjsLib;
}

const STORAGE_KEY = "sitplan-crm-v1";
const RECOVERY_KEY = "sitplan-crm-recovery-v1";
const SAFETY_BACKUP_KEY = "sitplan-crm-safety-backups-v1";
const SUPABASE_SESSION_KEY = "sitplan-supabase-session";
const LOCAL_SAVED_AT_KEY = "sitplan-crm-local-saved-at";
const SAFETY_BACKUP_LIMIT = 2;
const STORAGE_BACKUP_MAX_CHARS = 1200000;
const RECOVERY_BACKUP_LIMIT = 1;
const DAILY_REVIEW_PROMPT_KEY = "sitplan-daily-review-prompt-date";
const DAILY_REVIEW_SUGGEST_HOUR = 17;
const DAILY_REVIEW_COOLDOWN_HOURS = 23;
const DAILY_REVIEW_COOLDOWN_MS = DAILY_REVIEW_COOLDOWN_HOURS * 60 * 60 * 1000;
const TA_PORTFOLIO_DEFAULT_TARGET = 100;
const DEFAULT_WEEKLY_OI_GOAL = 25;
const PRIVATE_ACCESS_REQUIRED = window.location.protocol !== "file:";
const SITPLAN_APP_VERSION = "v139";
const OFFICIAL_DATA_SOURCE = {
  label: "SITPLAN Web CRM",
  version: SITPLAN_APP_VERSION,
  cloudProject: "bfkfuymbjhkyqyscyszu",
  canonicalWorkbook: "SITPLAN_FINAL (1).xlsx",
  canonicalWorkbookHash: "8902879ec090f2e5",
  decidedAt: "2026-08-09"
};
const STAGES = ["Base", "TA", "OI agendada", "OI realizada", "PC agendada", "PC realizada", "N", "C2"];
const LEGACY_STAGE_MAP = {
  "Novo contato": "Base",
  "Qualificação": "OI agendada",
  "Reunião": "OI realizada",
  "Proposta": "PC agendada",
  "Cliente": "N"
};
const BRAZIL_STATES = [
  ["AC", "Acre"], ["AL", "Alagoas"], ["AP", "Amapá"], ["AM", "Amazonas"],
  ["BA", "Bahia"], ["CE", "Ceará"], ["DF", "Distrito Federal"], ["ES", "Espírito Santo"],
  ["GO", "Goiás"], ["MA", "Maranhão"], ["MT", "Mato Grosso"], ["MS", "Mato Grosso do Sul"],
  ["MG", "Minas Gerais"], ["PA", "Pará"], ["PB", "Paraíba"], ["PR", "Paraná"],
  ["PE", "Pernambuco"], ["PI", "Piauí"], ["RJ", "Rio de Janeiro"], ["RN", "Rio Grande do Norte"],
  ["RS", "Rio Grande do Sul"], ["RO", "Rondônia"], ["RR", "Roraima"], ["SC", "Santa Catarina"],
  ["SP", "São Paulo"], ["SE", "Sergipe"], ["TO", "Tocantins"]
];
const PROFESSION_SUGGESTIONS = [
  "Advogado(a)", "Administrador(a)", "Arquiteto(a)", "Autônomo(a)", "Bancário(a)",
  "Cirurgião(ã)", "Comerciante", "Contador(a)", "Dentista", "Empresário(a)",
  "Engenheiro(a)", "Executivo(a)", "Farmacêutico(a)", "Fisioterapeuta", "Funcionário(a) público(a)",
  "Gestor(a)", "Médico(a)", "Militar", "Nutricionista", "Odontólogo(a)",
  "Piloto(a)", "Professor(a)", "Psicólogo(a)", "Servidor(a) público(a)", "Veterinário(a)"
];

let activeSettingsForNormalization = {};
let state = loadState();
let currentView = "dashboard";
let taskFilter = "pending";
let searchTerm = "";
let importedPdfContacts = [];
let pdfImportMode = "contacts";
let pdfReferrerId = "";
let pendingPostOiRecommendations = [];
let viewedWeekOffset = 0;
let agendaWeekOffset = 0;
let workFilter = "";
let assistantClientId = "";
let assistantMode = "next";
let assistantTone = "warm";
let assistantMessage = "";
let assistantIsLoading = false;
let assistantChatHistory = [];
let assistantChatClientId = "";
let dailyReviewLoading = false;
let activityInteractionMode = false;
let fridaySchedulingClientId = "";
let fridayPortfolioFilter = "current";
let fridayReferrerFilter = "all";
let fridaySosActive = false;
let fridaySosFocus = "";
let fridayPinnedClientId = "";
let fridayReferrerSearch = "";
let fridayQueueSearch = "";
let fridayTemperatureFilter = "all";
let fridaySessionActive = false;
let fridaySessionStartedAt = 0;
let fridaySessionClientIds = [];
let fridaySessionResults = [];
let fridayLockedQueueKey = "";
let fridayLockedClientIds = [];
let clientsPage = 1;
const CLIENTS_PAGE_SIZE = 25;
let crmTab = "list";
let resultsTab = "summary";
let c2RecoveryView = "now";
let fridayQueueVisibleLimit = 12;
const FRIDAY_QUEUE_PAGE_SIZE = 12;

const el = id => document.getElementById(id);
const currency = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
const shortDate = new Intl.DateTimeFormat("pt-BR", { day: "2-digit", month: "short" });

function offsetDate(days) {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

function defaultAreaCode(settings = activeSettingsForNormalization) {
  const digits = String(settings?.defaultAreaCode || "11").replace(/\D/g, "").slice(0, 2);
  return digits.length === 2 ? digits : "11";
}

function normalizeWhatsapp(value, areaCode = defaultAreaCode()) {
  let digits = String(value || "").replace(/\D/g, "").replace(/^0+/, "");
  if (digits.startsWith("55") && digits.length >= 12) digits = digits.slice(2);
  if (digits.length === 8 || digits.length === 9) digits = `${areaCode}${digits}`;
  return digits;
}

function contactDuplicate({ name = "", phone = "", email = "" }, ignoredId = "") {
  const normalizedPhone = normalizeWhatsapp(phone);
  const normalizedEmail = email.trim().toLocaleLowerCase("pt-BR");
  return state.clients.find(client => {
    if (client.id === ignoredId) return false;
    const clientPhone = normalizeWhatsapp(client.whatsapp || client.phone);
    const clientEmail = String(client.email || "").trim().toLocaleLowerCase("pt-BR");
    return (normalizedPhone && clientPhone === normalizedPhone)
      || (normalizedEmail && clientEmail === normalizedEmail);
  });
}

function normalizeClient(client) {
  const originalWhatsapp = client.whatsapp || client.phone || "";
  const normalizedWhatsapp = normalizeWhatsapp(originalWhatsapp);
  return {
    ...client,
    stage: LEGACY_STAGE_MAP[client.stage] || client.stage || "Base",
    phone: normalizedWhatsapp,
    whatsapp: normalizedWhatsapp,
    state: client.state || "",
    city: client.city || "",
    profession: client.profession || "",
    incomeRange: client.incomeRange || "",
    maritalStatus: client.maritalStatus || "",
    children: client.children || 0,
    referrer: client.referrer || "",
    referrerRelationship: client.referrerRelationship || client.relationship || "",
    referralContext: client.referralContext || "",
    temperature: client.temperature || "Morno",
    score: client.score || 0,
    lastInteraction: client.lastInteraction || "",
    nextAction: client.nextAction || (client.followUp ? "Realizar contato" : ""),
    taDate: client.taDate || "",
    oiScheduledDate: client.oiScheduledDate || "",
    oiTime: client.oiTime || "",
    oiMeetingLink: client.oiMeetingLink || "",
    oiMicrosoftEventId: client.oiMicrosoftEventId || "",
    oiDone: Boolean(client.oiDone),
    oiDate: client.oiDate || "",
    oiObjectives: client.oiObjectives || "",
    oiProtectionAreas: Array.isArray(client.oiProtectionAreas) ? client.oiProtectionAreas : [],
    oiBudget: client.oiBudget || "",
    oiDecision: client.oiDecision || "",
    oiObjection: client.oiObjection || "",
    oiUrgency: client.oiUrgency || "",
    oiQuickSummary: client.oiQuickSummary || "",
    oiDetailNotes: client.oiDetailNotes || "",
    oiSummaryNotes: client.oiSummaryNotes || "",
    oiRecsCaptured: Number(client.oiRecsCaptured || 0),
    postOiSuggestedPcDate: client.postOiSuggestedPcDate || "",
    postOiSuggestedPcTime: client.postOiSuggestedPcTime || "",
    postOiCompletedAt: client.postOiCompletedAt || "",
    planStatus: client.planStatus || "",
    planDueDate: client.planDueDate || "",
    planReadyDate: client.planReadyDate || "",
    planProfile: Boolean(client.planProfile),
    planGoals: Boolean(client.planGoals),
    planProtection: Boolean(client.planProtection),
    planBudget: Boolean(client.planBudget),
    planAlternatives: Boolean(client.planAlternatives),
    planNotes: client.planNotes || "",
    planMonthlyBudget: Number(client.planMonthlyBudget || 0),
    planFixedCost: Number(client.planFixedCost || 0),
    planLifestyle: client.planLifestyle || "",
    planRiskConcerns: client.planRiskConcerns || "",
    pcScheduledDate: client.pcScheduledDate || "",
    pcTime: client.pcTime || "",
    pcMeetingLink: client.pcMeetingLink || "",
    pcMicrosoftEventId: client.pcMicrosoftEventId || "",
    pcDone: Boolean(client.pcDone),
    pcDate: client.pcDate || "",
    outcome: client.outcome || "",
    outcomeDate: client.outcomeDate || "",
    c2NextAction: client.c2NextAction || "",
    c2Objection: client.c2Objection || "",
    c2EnteredAt: String(client.c2EnteredAt || (client.outcome === "C2" ? client.outcomeDate || "" : "")),
    c2LastOutcome: String(client.c2LastOutcome || ""),
    c2LastOutcomeDate: String(client.c2LastOutcomeDate || ""),
    c2TouchCount: Number(client.c2TouchCount || 0),
    c2RecoveredAt: String(client.c2RecoveredAt || ""),
    c2LegacyReview: String(client.c2LegacyReview || ""),
    cpfDone: Boolean(client.cpfDone),
    cpfDate: client.cpfDate || "",
    policyDelivered: Boolean(client.policyDelivered),
    deliveryDate: client.deliveryDate || "",
    policyStatus: client.policyStatus || "",
    nextRelationship: client.nextRelationship || "",
    relationshipStatus: client.relationshipStatus || "",
    recs: client.recs || 0,
    recsDate: client.recsDate || "",
    d7Done: Boolean(client.d7Done),
    d30Done: Boolean(client.d30Done),
    d60Done: Boolean(client.d60Done),
    d180Done: Boolean(client.d180Done),
    taPortfolioWeek: client.taPortfolioWeek || "",
    taPortfolioPinnedAt: client.taPortfolioPinnedAt || "",
    taNoInterestReason: client.taNoInterestReason || "",
    taNoInterestDate: client.taNoInterestDate || "",
    taPausedUntil: client.taPausedUntil || "",
    doNotContact: Boolean(client.doNotContact)
  };
}

function shouldStoreValue(value) {
  if (value === undefined || value === null || value === "") return false;
  if (value === false) return false;
  if (typeof value === "number" && value === 0) return false;
  if (Array.isArray(value) && !value.length) return false;
  if (typeof value === "object" && !Array.isArray(value) && !Object.keys(value).length) return false;
  return true;
}

function compactClientForStorage(client) {
  const normalized = normalizeClient(client);
  const compact = {};
  Object.entries(normalized).forEach(([key, value]) => {
    if (!shouldStoreValue(value)) return;
    if (key === "phone" && value === normalized.whatsapp) return;
    if (key === "stage" && value === "Base") return;
    if (key === "temperature" && value === "Morno") return;
    if (key === "priority" && value === "Média") return;
    compact[key] = value;
  });
  compact.id = normalized.id;
  compact.name = normalized.name;
  compact.whatsapp = normalized.whatsapp;
  compact.createdAt = normalized.createdAt;
  return compact;
}

function compactActivityForStorage(activity) {
  const compact = {};
  Object.entries(activity || {}).forEach(([key, value]) => {
    if (!shouldStoreValue(value)) return;
    compact[key] = value;
  });
  if (activity?.id) compact.id = activity.id;
  if (activity?.title) compact.title = activity.title;
  if (activity?.date) compact.date = activity.date;
  if (activity?.createdAt) compact.createdAt = activity.createdAt;
  return compact;
}

function compactStateForStorage(data = state) {
  return {
    clients: (data.clients || []).map(compactClientForStorage),
    activities: (data.activities || []).map(compactActivityForStorage),
    weeklyGoals: data.weeklyGoals || {},
    settings: data.settings || {}
  };
}

function serializeStateForStorage(data = state) {
  return JSON.stringify(compactStateForStorage(data));
}

function releaseLocalStoragePressure() {
  try {
    localStorage.removeItem(SAFETY_BACKUP_KEY);
    localStorage.removeItem(RECOVERY_KEY);
  } catch (_) {}
}

function setLocalStorageSafely(key, value, { critical = false } = {}) {
  try {
    localStorage.setItem(key, value);
    return true;
  } catch (error) {
    releaseLocalStoragePressure();
    try {
      localStorage.setItem(key, value);
      return true;
    } catch (retryError) {
      console.error("Falha ao salvar no armazenamento local:", retryError);
      if (critical) {
        showToast("Navegador sem espaço para salvar. A nuvem tentará proteger seus dados; exporte um backup assim que possível.", true);
      }
      return false;
    }
  }
}

function loadState() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (saved?.clients) {
      activeSettingsForNormalization = saved.settings || {};
      return {
        ...saved,
        clients: saved.clients.map(normalizeClient),
        activities: Array.isArray(saved.activities) ? saved.activities : [],
        weeklyGoals: saved.weeklyGoals || {},
        settings: saved.settings || {}
      };
    }
  } catch (_) {}
  activeSettingsForNormalization = {};
  return { clients: [], activities: [], weeklyGoals: {}, settings: { defaultAreaCode: "11" } };
}

function saveState() {
  if (isCloudReadOnly()) {
    showToast("Modo operacional: alterações bloqueadas para proteger a base.", true);
    refreshFromCloud().then(renderAll).catch(() => {});
    return;
  }
  state.settings ||= {};
  state.settings.defaultAreaCode = defaultAreaCode(state.settings);
  activeSettingsForNormalization = state.settings;
  const nextState = serializeStateForStorage(state);
  const previousState = localStorage.getItem(STORAGE_KEY);
  const localSaved = setLocalStorageSafely(STORAGE_KEY, nextState, { critical: true });
  if (localSaved && previousState !== nextState) setLocalStorageSafely(LOCAL_SAVED_AT_KEY, new Date().toISOString());
  if (localSaved && previousState && previousState !== nextState) storeRecoveryBackup(previousState);
  storeSafetyBackup(nextState, "auto");
  if (isCloudConnected()) setSyncStatus({ id: "connected" }, "Salvando...");
  queueCloudSave(state);
}

function activateStateSettings() {
  state.settings ||= {};
  state.settings.defaultAreaCode = defaultAreaCode(state.settings);
  activeSettingsForNormalization = state.settings;
}

function storeRecoveryBackup(serializedState) {
  if (!serializedState || serializedState.length > STORAGE_BACKUP_MAX_CHARS) return;
  try {
    const recovery = JSON.parse(localStorage.getItem(RECOVERY_KEY) || "[]");
    recovery.push({ savedAt: new Date().toISOString(), data: serializedState });
    setLocalStorageSafely(RECOVERY_KEY, JSON.stringify(recovery.slice(-RECOVERY_BACKUP_LIMIT)));
  } catch (_) {}
}

function storeSafetyBackup(serializedState, reason = "auto") {
  if (!serializedState || serializedState.length > STORAGE_BACKUP_MAX_CHARS) return;
  try {
    const today = offsetDate(0);
    const backups = JSON.parse(localStorage.getItem(SAFETY_BACKUP_KEY) || "[]");
    const cleanBackups = Array.isArray(backups) ? backups.filter(item => item?.data) : [];
    const entry = {
      date: today,
      savedAt: new Date().toISOString(),
      reason,
      clientsCount: state.clients.length,
      activitiesCount: state.activities.length,
      data: serializedState
    };
    const sameDayIndex = cleanBackups.findIndex(item => item.date === today && item.reason === reason);
    if (sameDayIndex >= 0) cleanBackups[sameDayIndex] = entry;
    else cleanBackups.push(entry);
    setLocalStorageSafely(SAFETY_BACKUP_KEY, JSON.stringify(cleanBackups.slice(-SAFETY_BACKUP_LIMIT)));
  } catch (_) {
    try {
      const backups = JSON.parse(localStorage.getItem(SAFETY_BACKUP_KEY) || "[]");
      setLocalStorageSafely(SAFETY_BACKUP_KEY, JSON.stringify((Array.isArray(backups) ? backups : []).slice(-1)));
    } catch (__) {}
  }
}

function undoLastChange() {
  if (!confirm("Desfazer a última alteração feita no CRM?")) return;
  try {
    const recovery = JSON.parse(localStorage.getItem(RECOVERY_KEY) || "[]");
    const snapshot = recovery.pop();
    if (!snapshot) {
      showToast("Nenhuma alteração anterior disponível", true);
      return;
    }
    const restored = JSON.parse(snapshot.data);
    activeSettingsForNormalization = restored.settings || {};
    state = {
      ...restored,
      clients: (restored.clients || []).map(normalizeClient),
      activities: Array.isArray(restored.activities) ? restored.activities : [],
      weeklyGoals: restored.weeklyGoals || {},
      settings: restored.settings || {}
    };
    activateStateSettings();
    setLocalStorageSafely(RECOVERY_KEY, JSON.stringify(recovery));
    setLocalStorageSafely(STORAGE_KEY, serializeStateForStorage(state), { critical: true });
    queueCloudSave(state);
    renderAll();
    showToast("Última alteração desfeita");
  } catch (_) {
    showToast("Não foi possível restaurar a alteração", true);
  }
}

function setSyncStatus(user = null, message = "") {
  if (user) {
    el("syncStatus").textContent = message || cloudAccessLabel();
    el("cloudAccessButton").textContent = "Sair da nuvem";
    el("cloudAccessButton").dataset.connected = "true";
  } else {
    el("syncStatus").textContent = message || "Modo local";
    el("cloudAccessButton").textContent = "Entrar e sincronizar";
    el("cloudAccessButton").dataset.connected = "false";
  }
  renderOfficialSourceNotice();
}

function officialSourceModeLabel() {
  if (isCloudReadOnly()) return "Modo sócio: leitura operacional";
  if (isCloudConnected()) return `Nuvem oficial ativa · ${OFFICIAL_DATA_SOURCE.cloudProject}`;
  return "Modo local: confira a origem antes de importar";
}

function renderOfficialSourceNotice() {
  const title = el("officialSourceTitle");
  const detail = el("officialSourceDetail");
  const card = el("officialSourceCard");
  if (!title || !detail || !card) return;
  title.textContent = `${OFFICIAL_DATA_SOURCE.label} ${OFFICIAL_DATA_SOURCE.version}`;
  detail.textContent = `${officialSourceModeLabel()}. Use somente CSV revisado do SITPLAN para carga de contatos. Planilha legada: ${OFFICIAL_DATA_SOURCE.canonicalWorkbook} (${OFFICIAL_DATA_SOURCE.canonicalWorkbookHash}).`;
  card.classList.toggle("connected", isCloudConnected());
  card.classList.toggle("local", !isCloudConnected());
}

function confirmOfficialImportSource(file) {
  const name = file?.name || "arquivo selecionado";
  const message = [
    "Conferência de origem do SITPLAN",
    "",
    `Fonte oficial dos dados vivos: ${OFFICIAL_DATA_SOURCE.label} ${OFFICIAL_DATA_SOURCE.version} + nuvem.`,
    `Planilha oficial legada para conferência: ${OFFICIAL_DATA_SOURCE.canonicalWorkbook}.`,
    "",
    `Arquivo selecionado: ${name}`,
    "",
    "Continue somente se este CSV veio de uma exportação revisada/gerada para o SITPLAN atual. Isso evita misturar base antiga ou errada no CRM."
  ].join("\n");
  return window.confirm(message);
}

const READ_ONLY_WRITE_SELECTORS = [
  "#newClientButton",
  "#importCsvButton",
  "#importPdfButton",
  "#exportData",
  "#importData",
  "#undoLastChange",
  "#viewerEmailInput",
  "#viewerPasswordInput",
  "#createViewerAccess",
  "#grantViewerAccess",
  "#revokeViewerAccess",
  "#runAutomationNow",
  "#runDailyReview",
  "#smartDailyReviewRun",
  "#saveWeeklyContract",
  "#newActivityButton",
  "#deleteActivity",
  "#deleteClient",
  "#confirmPdfImport",
  "#saveMeetingDetails",
  "#meetingCreateMicrosoft",
  "#meetingCompleteOi",
  "#meetingOiNoShow",
  "#meetingCloseN",
  "#meetingCloseC2",
  "#createRecommendedContact",
  "#importOiRecommendations",
  "#clientLogInteraction",
  "#quickInputForm button[type='submit']",
  "#clientForm button[type='submit']",
  "#activityForm button[type='submit']",
  "#c2UpdateForm button[type='submit']",
  "[data-c2-legacy-action]",
  "#postOiForm button[type='submit']",
  "[data-friday-result]",
  "[data-policy-status]",
  "[data-policy-delivered]",
  "[data-plan-ready]",
  "[data-c2-next]",
  "[data-input-client]",
  "[data-missed-call]",
  "[data-task-complete]",
  "[data-complete]",
  "[data-execution-plan-ready]"
].join(",");

const VIEWER_DEFAULT_VIEW = "friday";
const VIEWER_ALLOWED_VIEWS = new Set(["friday", "clients"]);

function applyReadOnlyMode() {
  const readOnly = isCloudReadOnly();
  document.body.classList.toggle("read-only-mode", readOnly);
  document.body.classList.toggle("ta-viewer-mode", readOnly);
  document.querySelectorAll(READ_ONLY_WRITE_SELECTORS).forEach(control => {
    control.disabled = readOnly;
    if (readOnly) control.title = "Modo operacional: alteração bloqueada";
    else if (control.title === "Modo operacional: alteração bloqueada") control.removeAttribute("title");
  });
  applyTaViewerScope(readOnly);
  if (readOnly && el("syncStatus")) el("syncStatus").textContent = "TA + CRM básico";
}

function applyTaViewerScope(readOnly = isCloudReadOnly()) {
  document.querySelectorAll(".nav-item").forEach(button => {
    const blocked = readOnly && !VIEWER_ALLOWED_VIEWS.has(button.dataset.view);
    button.hidden = blocked;
    button.disabled = blocked;
  });
  document.querySelectorAll(".mobile-nav-item").forEach(button => {
    const blocked = readOnly && !VIEWER_ALLOWED_VIEWS.has(button.dataset.mobileView);
    button.hidden = blocked;
    button.disabled = blocked;
  });
  const scoreSortOption = el("sortClients")?.querySelector("option[value='score']");
  if (scoreSortOption) scoreSortOption.hidden = readOnly;
  if (readOnly && el("sortClients")?.value === "score") el("sortClients").value = "name";
  if (readOnly && !VIEWER_ALLOWED_VIEWS.has(currentView)) switchView(VIEWER_DEFAULT_VIEW);
}

function applyPrivateAccessLock(locked = PRIVATE_ACCESS_REQUIRED && (!isCloudConnected() || !secureEntryConfirmed)) {
  document.documentElement.classList.toggle("secure-entry", Boolean(locked));
  if (!PRIVATE_ACCESS_REQUIRED) return;
  document.body.classList.toggle("app-locked", locked);
  if (locked) openAuthModal({ locked: true });
  else closeAuthModal(true);
}

function cloudErrorMessage(error) {
  const raw = error?.message || String(error || "");
  if (error?.code === "NETWORK" || /Failed to fetch|NetworkError|Falha de conexao/i.test(raw)) {
    return "Não consegui falar com o Supabase. Verifique internet, projeto ativo e tente novamente. Se o site abriu normal, pode ser bloqueio temporário da API.";
  }
  if (error?.status === 401 || error?.status === 403 || /JWT|permission|permissao|Unauthorized/i.test(raw)) {
    return "Conectou no login, mas a base negou permissão. Saia da nuvem, entre de novo e confira se o SQL de segurança foi executado.";
  }
  if (/sitplan_ta_viewer_contacts|PGRST202|function .* does not exist/i.test(raw)) {
    return "Acesso TA ainda não configurado no Supabase. Execute o arquivo supabase-viewer-access.sql atualizado no SQL Editor.";
  }
  if (error?.status === 404 || error?.code === "42P01" || /relation .* does not exist|contacts|activities|weekly_goals/i.test(raw)) {
    return "Banco sem as tabelas do SITPLAN. Rode o arquivo supabase-schema.sql no SQL Editor do Supabase.";
  }
  if (/Invalid login credentials/i.test(raw)) return "E-mail ou senha incorretos. Confira os dados ou clique em Criar minha conta.";
  if (/Email not confirmed/i.test(raw)) return "Seu e-mail ainda não foi confirmado. Abra o e-mail do Supabase e confirme o cadastro.";
  if (/User not found/i.test(raw)) return "Conta não encontrada. Clique em Criar minha conta para cadastrar este e-mail.";
  if (/Signup|signups/i.test(raw)) return "Cadastro por e-mail parece desativado no Supabase. Ative Authentication > Providers > Email.";
  return `Falha na nuvem: ${raw || "erro desconhecido"}.`;
}

function showAuthError(error) {
  const message = error?.userMessage || cloudErrorMessage(error);
  el("authMessage").textContent = message;
  el("authMessage").classList.add("error");
  if (el("liveRegion")) el("liveRegion").textContent = message;
}

function authConnectionError(error) {
  const authError = error instanceof Error ? error : new Error(String(error || "Falha de conexão"));
  authError.userMessage = "Não foi possível conectar. Verifique sua internet e tente novamente.";
  return authError;
}

async function runCloudDiagnostic() {
  openAuthModal();
  el("authMessage").classList.remove("error");
  el("authMessage").textContent = "Diagnosticando nuvem...";
  try {
    const result = await diagnoseCloud();
    const lines = result.checks.map(check => `${check.ok ? "OK" : "ERRO"} - ${check.name}: ${check.detail}`);
    el("authMessage").textContent = result.ok
      ? `Diagnóstico da nuvem OK.\n${lines.join("\n")}\nAgora tente entrar com e-mail e senha.`
      : `Diagnóstico encontrou problema:\n${lines.join("\n")}`;
    el("authMessage").classList.toggle("error", !result.ok);
  } catch (error) {
    showAuthError(error);
  }
}

function lastLocalSaveAt() {
  try {
    const recorded = localStorage.getItem(LOCAL_SAVED_AT_KEY);
    if (recorded) return recorded;
    const backups = JSON.parse(localStorage.getItem(SAFETY_BACKUP_KEY) || "[]");
    return (Array.isArray(backups) ? backups : [])
      .map(item => item?.savedAt || "")
      .reduce((latest, value) => (value > latest ? value : latest), "");
  } catch (_) {
    return "";
  }
}

// Combine cloud data with this device's data without discarding local edits
// that have not reached the cloud yet (see mergeCloudState in cloud-sync.js).
function adoptCloudState(remoteState) {
  const normalizedRemote = {
    ...remoteState,
    clients: remoteState.clients.map(normalizeClient)
  };
  const merged = isCloudReadOnly()
    ? { state: normalizedRemote, hasLocalChanges: false }
    : mergeCloudState(state, normalizedRemote, { localSavedAt: lastLocalSaveAt() });
  activeSettingsForNormalization = merged.state.settings || {};
  state = {
    clients: merged.state.clients.map(normalizeClient),
    activities: merged.state.activities,
    weeklyGoals: merged.state.weeklyGoals,
    settings: merged.state.settings || {}
  };
  activateStateSettings();
  setLocalStorageSafely(STORAGE_KEY, serializeStateForStorage(state), { critical: true });
  if (merged.hasLocalChanges) queueCloudSave(state);
  return merged.hasLocalChanges;
}

async function connectAuthenticatedUser(user) {
  applyPrivateAccessLock(false);
  setSyncStatus(user, "Sincronizando...");
  try {
    const remoteState = await loadCloudState();
    const remoteHasData = remoteState.clients.length || remoteState.activities.length || Object.keys(remoteState.weeklyGoals).length;
    if (remoteHasData || isCloudReadOnly()) {
      adoptCloudState(remoteState);
      runOperationalAutomation();
      renderAll();
    } else {
      queueCloudSave(state);
    }
    setSyncStatus(user);
    closeAuthModal();
    showToast("Nuvem conectada com sucesso");
  } catch (error) {
    console.error(error);
    setSyncStatus(null, "Nuvem não conectou");
    applyPrivateAccessLock(true);
    openAuthModal();
    showAuthError(error);
  }
}

async function refreshFromCloud() {
  if (!isCloudConnected()) return;
  try {
    const remoteState = await loadCloudState();
    adoptCloudState(remoteState);
    runOperationalAutomation();
    renderAll();
    setSyncStatus({ id: "connected" });
  } catch (error) {
    console.error(error);
  }
}

function openAuthModal(options = {}) {
  el("authMessage").textContent = "";
  el("authMessage").classList.remove("error");
  toggleStoredSessionActions(Boolean(options.storedSession));
  if (options.storedSession) {
    const email = options.user?.email || storedSessionUser()?.email || "";
    if (email && !el("authEmail").value) el("authEmail").value = email;
    el("authMessage").textContent = "Sessão salva neste aparelho. Por segurança, confirme antes de abrir o CRM ou saia para digitar a senha novamente.";
  } else if (options.locked) {
    el("authMessage").textContent = "Digite seu login para liberar o SITPLAN CRM.";
  }
  el("authModal").classList.remove("hidden");
  setTimeout(() => (options.storedSession ? el("continueStoredSessionButton") : el("authEmail"))?.focus(), 50);
}

function closeAuthModal(force = false) {
  if (!force && PRIVATE_ACCESS_REQUIRED && document.body.classList.contains("app-locked")) return;
  el("authModal").classList.add("hidden");
}

async function submitAuth(event) {
  event.preventDefault();
  el("authMessage").textContent = "Entrando...";
  el("authMessage").classList.remove("error");
  manualAuthInProgress = true;
  secureEntryConfirmed = true;
  try {
    const { error } = await signIn(el("authEmail").value.trim(), el("authPassword").value);
    if (error) {
      secureEntryConfirmed = false;
      showAuthError(error);
    }
  } catch (error) {
    secureEntryConfirmed = false;
    showAuthError(authConnectionError(error));
  } finally {
    manualAuthInProgress = false;
  }
}

async function createAccount() {
  if (!el("authForm").reportValidity()) return;
  el("authMessage").textContent = "Criando sua conta...";
  el("authMessage").classList.remove("error");
  manualAuthInProgress = true;
  secureEntryConfirmed = true;
  try {
    const { data, error } = await signUp(el("authEmail").value.trim(), el("authPassword").value);
    if (error) {
      secureEntryConfirmed = false;
      showAuthError(error);
      return;
    }
    el("authMessage").textContent = data.session
      ? "Conta criada e conectada."
      : "Conta criada. Confira seu e-mail para confirmar o acesso.";
    if (!data.session) secureEntryConfirmed = false;
  } catch (error) {
    secureEntryConfirmed = false;
    showAuthError(authConnectionError(error));
  } finally {
    manualAuthInProgress = false;
  }
}

function parseLocalDate(value) {
  if (!value) return null;
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day, 12);
}

function startOfWeek(offset = 0) {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  const day = date.getDay() || 7;
  date.setDate(date.getDate() - day + 1 + (offset * 7));
  return date;
}

function dateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function selectedWeek() {
  const start = startOfWeek(viewedWeekOffset);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  return { start, end, startKey: dateKey(start), endKey: dateKey(end) };
}

function weeklyGoal(week = selectedWeek()) {
  return Number(state.weeklyGoals[week.startKey] || DEFAULT_WEEKLY_OI_GOAL);
}

function inSelectedWeek(value, week = selectedWeek()) {
  return Boolean(value && value >= week.startKey && value <= week.endKey);
}

function selectedWeekDates(week = selectedWeek()) {
  return Array.from({ length: 7 }, (_, index) => {
    const date = new Date(week.start);
    date.setDate(week.start.getDate() + index);
    return {
      key: dateKey(date),
      date,
      label: ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"][index],
      short: date.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" })
    };
  });
}

function countRecsOnDate(clients, key) {
  return clients
    .filter(client => client.recsDate === key)
    .reduce((total, client) => total + Number(client.recs || 0), 0);
}

function weeklyBreakdownRows(week = selectedWeek()) {
  const clients = state.clients;
  const dates = selectedWeekDates(week);
  return dates.map(day => {
    const metrics = [
      { key: "ta", label: "TA", value: clients.filter(client => client.taDate === day.key).length },
      { key: "oiScheduled", label: "OI ag.", value: clients.filter(client => client.oiScheduledDate === day.key).length },
      { key: "oiDone", label: "OI real.", value: clients.filter(client => client.oiDone && client.oiDate === day.key).length },
      { key: "pcScheduled", label: "PC ag.", value: clients.filter(client => client.pcScheduledDate === day.key).length },
      { key: "pcDone", label: "PC real.", value: clients.filter(client => client.pcDone && client.pcDate === day.key).length },
      { key: "recs", label: "RECs", value: countRecsOnDate(clients, day.key) },
      { key: "sales", label: "N", value: clients.filter(client => client.outcome === "N" && client.outcomeDate === day.key).length }
    ];
    return {
      ...day,
      metrics,
      total: metrics.reduce((sum, metric) => sum + metric.value, 0)
    };
  });
}

function renderWeeklyBreakdown(week = selectedWeek()) {
  const grid = el("weeklyBreakdownGrid");
  if (!grid) return;

  const rows = weeklyBreakdownRows(week);
  const maxTotal = Math.max(...rows.map(row => row.total), 1);
  const todayKey = offsetDate(0);
  const todayRow = rows.find(row => row.key === todayKey);
  const bestRow = rows.reduce((best, row) => row.total > best.total ? row : best, rows[0]);
  const totalMovements = rows.reduce((sum, row) => sum + row.total, 0);
  const fridayRow = rows[4];
  const fridayTa = fridayRow?.metrics.find(metric => metric.key === "ta")?.value || 0;
  const fridayRecs = fridayRow?.metrics.find(metric => metric.key === "recs")?.value || 0;

  el("weeklyBreakdownHighlightLabel").textContent = todayRow ? "Movimentos hoje" : "Melhor dia";
  el("weeklyBreakdownToday").textContent = todayRow ? todayRow.total : bestRow.total;

  if (!totalMovements) {
    el("weeklyBreakdownSummary").textContent = "Ainda não há movimentos registrados nesta semana. O primeiro objetivo é criar tração em TA e OI agendadas.";
  } else if (fridayTa + fridayRecs > 0) {
    el("weeklyBreakdownSummary").textContent = `Melhor dia: ${bestRow.label} (${bestRow.total}). Sexta está com ${fridayTa} TA e ${fridayRecs} REC(s), o termômetro do próximo ciclo.`;
  } else {
    el("weeklyBreakdownSummary").textContent = `Melhor dia: ${bestRow.label} (${bestRow.total}). Use esta leitura para equilibrar TA, montagem de plano e fechamento.`;
  }

  grid.innerHTML = rows.map(row => {
    const intensity = Math.max(0.08, row.total / maxTotal);
    const metrics = row.metrics.map(metric => `
      <span class="weekly-breakdown-pill ${metric.value ? "active" : ""}" title="${metric.label}: ${metric.value}">
        <b>${metric.value}</b>${metric.label}
      </span>`).join("");
    return `
      <article class="weekly-breakdown-day ${row.key === todayKey ? "today" : ""}" style="--day-intensity:${intensity}">
        <div class="weekly-breakdown-daytop">
          <div><strong>${row.label}</strong><small>${row.short}</small></div>
          <em>${row.total}</em>
        </div>
        <div class="weekly-breakdown-bar"><i></i></div>
        <div class="weekly-breakdown-pills">${metrics}</div>
      </article>`;
  }).join("");
}

function dateStatus(value) {
  if (!value) return { text: "Sem data", className: "" };
  const date = parseLocalDate(value);
  const today = parseLocalDate(offsetDate(0));
  const diff = Math.round((date - today) / 86400000);
  if (diff < 0) return { text: "Atrasado", className: "overdue" };
  if (diff === 0) return { text: "Hoje", className: "today" };
  if (diff === 1) return { text: "Amanhã", className: "" };
  return { text: shortDate.format(date), className: "" };
}

function formatDate(value) {
  if (!value) return "Não agendado";
  return new Intl.DateTimeFormat("pt-BR", { day: "2-digit", month: "2-digit", year: "numeric" }).format(parseLocalDate(value));
}

function formatMeetingDate(date, time = "") {
  if (!date) return "Data não definida";
  return `${formatDate(date)}${time ? ` às ${time}` : ""}`;
}

function meetingDateTime(date, time, durationMinutes = 60) {
  const start = new Date(`${date}T${time || "09:00"}:00`);
  const end = new Date(start.getTime() + Number(durationMinutes || 60) * 60000);
  return { start, end };
}

function previousBusinessDate(value) {
  const date = parseLocalDate(value);
  date.setDate(date.getDate() - 1);
  if (date.getDay() === 0) date.setDate(date.getDate() - 2);
  if (date.getDay() === 6) date.setDate(date.getDate() - 1);
  return dateKey(date);
}

function businessDateAfter(value, days = 1) {
  const date = value ? parseLocalDate(value) : new Date();
  date.setHours(12, 0, 0, 0);
  let remaining = days;
  while (remaining > 0) {
    date.setDate(date.getDate() + 1);
    if (date.getDay() !== 0 && date.getDay() !== 6) remaining -= 1;
  }
  return dateKey(date);
}

function planDueDateFor(client) {
  const afterOi = businessDateAfter(client.oiDate || offsetDate(0), 2);
  const beforePc = client.pcScheduledDate ? previousBusinessDate(client.pcScheduledDate) : "";
  if (beforePc && beforePc < afterOi) return beforePc;
  return afterOi;
}

function ensurePlanTask(client) {
  if (!client.oiDone || client.planStatus === "Pronto") return;
  client.planStatus ||= "Em montagem";
  client.planDueDate ||= planDueDateFor(client);
  const existing = state.activities.find(activity =>
    activity.clientId === client.id && activity.type === "Plano" && !activity.completed
  );
  const task = {
    id: existing?.id || crypto.randomUUID(),
    title: `Montar plano - ${client.name}`,
    date: client.planDueDate,
    time: "09:00",
    type: "Plano",
    clientId: client.id,
    notes: [
      client.pcScheduledDate
        ? `Plano deve estar pronto antes do PC de ${formatDate(client.pcScheduledDate)}.`
        : client.postOiSuggestedPcDate
          ? `PC sugerido para ${formatDate(client.postOiSuggestedPcDate)}.`
          : "Montar plano com base nas necessidades levantadas na OI.",
      client.oiObjectives ? `Objetivo: ${client.oiObjectives}` : "",
      client.oiProtectionAreas?.length ? `Prioridades: ${client.oiProtectionAreas.join(", ")}.` : ""
    ].filter(Boolean).join(" "),
    completed: false,
    createdAt: existing?.createdAt || new Date().toISOString()
  };
  if (existing) Object.assign(existing, task);
  else state.activities.push(task);
}

function ensureOperationalActivity(client, prefix, options) {
  const existing = state.activities.find(activity =>
    activity.clientId === client.id
    && !activity.completed
    && activity.title.startsWith(prefix)
  );
  if (existing) return false;
  state.activities.push({
    id: crypto.randomUUID(),
    title: `${prefix}${client.name}`,
    date: options.date,
    time: options.time || "",
    type: options.type || "Follow-up",
    clientId: client.id,
    notes: options.notes || "",
    completed: false,
    createdAt: new Date().toISOString()
  });
  return true;
}

function nextFridayDate() {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  const days = (5 - date.getDay() + 7) % 7;
  date.setDate(date.getDate() + days);
  return dateKey(date);
}

function runOperationalAutomation({ notify = false } = {}) {
  const today = offsetDate(0);
  let created = 0;

  state.clients.forEach(client => {
    if (client.doNotContact) return;

    if (["Base", "TA"].includes(client.stage) && !client.oiScheduledDate && !client.followUp) {
      client.followUp = isRecommendation(client) ? nextBusinessDate() : nextFridayDate();
      client.nextAction ||= isRecommendation(client) ? "Fazer TA da recomendação" : "Fazer primeira abordagem";
      client.completed = false;
      created += 1;
    }

    if (client.oiScheduledDate && !client.oiDone) {
      const date = previousBusinessDate(client.oiScheduledDate);
      created += Number(ensureOperationalActivity(client, "Confirmar OI - ", {
        date: date < today ? today : date,
        time: "09:00",
        type: "Follow-up",
        notes: assistantWhatsappMessage(client, "short")
      }));
    }

    if (client.oiDone && client.planStatus !== "Pronto" && client.outcome !== "N") {
      const before = state.activities.length;
      ensurePlanTask(client);
      created += Number(state.activities.length > before);
    }

    if (client.planStatus === "Pronto" && !client.pcScheduledDate && !client.outcome) {
      created += Number(ensureOperationalActivity(client, "Agendar PC - ", {
        date: client.followUp || nextBusinessDate(),
        type: "Follow-up",
        notes: "Plano pronto. Oferecer duas opções de horário para a reunião PC."
      }));
      client.nextAction = "Agendar reunião PC";
      client.followUp ||= nextBusinessDate();
    }

    if (client.pcScheduledDate && !client.pcDone) {
      const date = previousBusinessDate(client.pcScheduledDate);
      created += Number(ensureOperationalActivity(client, "Confirmar PC - ", {
        date: date < today ? today : date,
        time: "09:00",
        type: "Follow-up",
        notes: assistantWhatsappMessage(client, "short")
      }));
    }

    if (client.pcDone && !client.outcome) {
      created += Number(ensureOperationalActivity(client, "Definir resultado do PC - ", {
        date: client.followUp || nextBusinessDate(),
        type: "Follow-up",
        notes: "Registrar N ou C2 e combinar um próximo passo com data."
      }));
    }

    if (client.outcome === "C2" && !client.followUp) {
      client.followUp = nextBusinessDate();
      client.nextAction = client.c2NextAction || "Realizar follow-up de fechamento";
      client.completed = false;
      created += 1;
    }

    if (client.outcome === "N" && !client.policyDelivered) {
      created += Number(ensureOperationalActivity(client, "Acompanhar emissão - ", {
        date: client.followUp || nextBusinessDate(),
        type: "Follow-up",
        notes: "Acompanhar emissão, pendências e preparação da entrega da apólice."
      }));
    }
  });

  state.settings.lastAutomationRun = new Date().toISOString();
  if (created) saveState();
  renderAutomationCenter(created);
  if (notify) showToast(created ? `${created} ação(ões) organizada(s) automaticamente` : "Sua operação já está organizada");
  return created;
}

function automationCenterData() {
  const today = offsetDate(0);
  const confirmations = state.activities.filter(activity =>
    !activity.completed && activity.title.startsWith("Confirmar ")
  );
  const plans = state.clients.filter(client => client.oiDone && client.planStatus !== "Pronto" && client.outcome !== "N");
  const pcs = state.clients.filter(client => client.planStatus === "Pronto" && !client.pcScheduledDate && !client.outcome);
  const c2s = state.clients.filter(client => client.outcome === "C2" && (!client.followUp || client.followUp <= today));
  const aftercare = state.clients.filter(client => client.outcome === "N" && !client.policyDelivered);
  const feed = [
    ...confirmations.map(activity => ({ label: activity.title, date: activity.date, clientId: activity.clientId })),
    ...plans.map(client => ({ label: `Montar plano - ${client.name}`, date: client.planDueDate, clientId: client.id })),
    ...pcs.map(client => ({ label: `Agendar PC - ${client.name}`, date: client.followUp, clientId: client.id })),
    ...c2s.map(client => ({ label: `Retomar C2 - ${client.name}`, date: client.followUp, clientId: client.id })),
    ...aftercare.map(client => ({ label: `Pós-venda - ${client.name}`, date: client.followUp, clientId: client.id }))
  ].sort((a, b) => (a.date || "0000").localeCompare(b.date || "0000")).slice(0, 5);
  return { confirmations, plans, pcs, c2s, aftercare, feed };
}

function dailyReviewData() {
  const today = offsetDate(0);
  const tomorrow = nextBusinessDate();
  const eventsToday = state.activities.filter(activity => activity.date === today);
  const priorities = [...state.clients]
    .filter(client => !client.doNotContact && client.outcome !== "N")
    .sort((a, b) => assistantPriorityScore(b) - assistantPriorityScore(a))
    .slice(0, 3)
    .map(client => ({
      id: client.id,
      primeiroNome: client.name.split(/\s+/)[0],
      nome: client.name,
      etapa: client.stage,
      score: assistantPriorityScore(client),
      proximaAcao: client.nextAction || assistantAction(client).action,
      data: client.followUp || tomorrow
    }));
  return {
    data: today,
    taRealizadas: state.clients.filter(client => client.taDate === today).length,
    oiAgendadas: state.clients.filter(client => client.oiScheduledDate === today).length,
    oiRealizadas: state.clients.filter(client => client.oiDone && client.oiDate === today).length,
    pcAgendadas: state.clients.filter(client => client.pcScheduledDate === today).length,
    pcRealizadas: state.clients.filter(client => client.pcDone && client.pcDate === today).length,
    vendasN: state.clients.filter(client => client.outcome === "N" && (client.outcomeDate === today || client.cpfDate === today)).length,
    novosC2: state.clients.filter(client => client.outcome === "C2" && client.outcomeDate === today).length,
    recomendacoes: state.clients.filter(client => client.recsDate === today).reduce((sum, client) => sum + Number(client.recs || 0), 0),
    tarefasConcluidas: eventsToday.filter(activity => activity.completed).length,
    tarefasPendentesHoje: eventsToday.filter(activity => !activity.completed).length,
    planosPendentes: state.clients.filter(client => client.oiDone && client.planStatus !== "Pronto" && client.outcome !== "N").length,
    pcsParaAgendar: state.clients.filter(client => client.planStatus === "Pronto" && !client.pcScheduledDate && !client.outcome).length,
    c2Vencidos: state.clients.filter(client => client.outcome === "C2" && (!client.followUp || client.followUp <= today)).length,
    prioridadesAmanha: priorities
  };
}

function renderDailyReview() {
  const review = state.settings.dailyReview;
  if (!review?.text) {
    el("dailyReview").classList.add("hidden");
    updateDailyReviewButton();
    return;
  }
  el("dailyReview").classList.remove("hidden");
  const today = offsetDate(0);
  const yesterday = offsetDate(-1);
  el("dailyReviewTitle").textContent = review.date === today
    ? "Fechamento de hoje"
    : review.date === yesterday
      ? "Briefing de hoje, gerado ontem"
      : `Fechamento de ${formatDate(review.date)}`;
  el("dailyReviewTime").textContent = review.createdAt
    ? new Date(review.createdAt).toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })
    : "";
  el("dailyReviewText").textContent = review.text;
  renderDailyReviewPriorities(review.priorities || dailyReviewData().prioridadesAmanha || []);
  updateDailyReviewButton();
}

function priorityColor(score = 0) {
  if (score >= 70) return "high";
  if (score >= 45) return "medium";
  return "low";
}

function renderDailyReviewPriorities(priorities = []) {
  const container = el("dailyReviewPriorities");
  if (!container) return;
  container.innerHTML = priorities.length ? `
    <small>Top 3 para começar o próximo dia</small>
    <div class="daily-review-priority-grid">
      ${priorities.slice(0, 3).map((item, index) => `
        <button class="daily-review-priority ${priorityColor(Number(item.score || 0))}" data-daily-review-client="${escapeHtml(item.id || "")}">
          <b>${index + 1}</b>
          <span>
            <strong>${escapeHtml(item.nome || item.primeiroNome || "Contato")}</strong>
            <em>${escapeHtml(item.etapa || "Base")} · ${escapeHtml(item.proximaAcao || "Definir próximo passo")}</em>
          </span>
          <i>${Math.round(Number(item.score || 0))}</i>
        </button>
      `).join("")}
    </div>
  ` : "";
  container.querySelectorAll("[data-daily-review-client]").forEach(button => {
    button.addEventListener("click", () => {
      const id = button.dataset.dailyReviewClient;
      if (id) openClientModal(id);
    });
  });
}

function registerPaidApiUsage(answer, purpose) {
  const usage = answer?.usage;
  if (!usage) return;
  state.settings.apiUsage = {
    month: offsetDate(0).slice(0, 7),
    requests: Number(usage.requests || 0),
    limit: Number(usage.limit || 0),
    lastPurpose: purpose,
    updatedAt: new Date().toISOString()
  };
}

function dailyReviewCooldownStatus() {
  const review = state.settings.dailyReview;
  const createdAt = review?.createdAt ? new Date(review.createdAt) : null;
  const timestamp = createdAt?.getTime();
  if (!timestamp || Number.isNaN(timestamp)) return { locked: false };
  const elapsed = Date.now() - timestamp;
  if (elapsed >= DAILY_REVIEW_COOLDOWN_MS) return { locked: false, createdAt };
  const remainingMs = DAILY_REVIEW_COOLDOWN_MS - elapsed;
  const remainingHours = Math.max(1, Math.ceil(remainingMs / 3600000));
  const dateLabel = review.date === offsetDate(0) ? "hoje" : `em ${formatDate(review.date)}`;
  const timeLabel = createdAt.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
  return {
    locked: true,
    createdAt,
    remainingHours,
    message: `Fechamento já gerado ${dateLabel} às ${timeLabel}. Próximo disponível em ${remainingHours}h.`
  };
}

async function runDailyReview() {
  if (dailyReviewLoading) return;
  const cooldown = dailyReviewCooldownStatus();
  if (cooldown.locked) {
    updateDailyReviewButton();
    showToast(cooldown.message, true);
    return;
  }
  if (!isCloudConnected()) {
    showToast("Entre e sincronize para gerar o fechamento com IA", true);
    return;
  }

  dailyReviewLoading = true;
  const button = el("runDailyReview");
  button.disabled = true;
  button.textContent = "Analisando o dia...";
  try {
    runOperationalAutomation();
    const summary = dailyReviewData();
    const answer = await requestDailyReview(summary);
    state.settings.dailyReview = {
      date: offsetDate(0),
      text: answer.text,
      model: answer.model,
      createdAt: new Date().toISOString(),
      priorities: summary.prioridadesAmanha
    };
    registerPaidApiUsage(answer, "Fechamento diário");
    saveState();
    renderDailyReview();
    renderSmartDailyReviewPrompt();
    showToast("Fechamento diário concluído com 1 chamada da API");
  } catch (error) {
    showToast(error.message || "Não foi possível analisar o dia", true);
  } finally {
    dailyReviewLoading = false;
    button.disabled = false;
    updateDailyReviewButton();
  }
}

function updateDailyReviewButton() {
  const buttons = [el("runDailyReview"), el("smartDailyReviewRun")].filter(Boolean);
  if (!buttons.length || dailyReviewLoading) return;
  const cooldown = dailyReviewCooldownStatus();
  buttons.forEach(button => {
    button.disabled = cooldown.locked;
    button.textContent = cooldown.locked
      ? `Disponível em ${cooldown.remainingHours}h`
      : "Fechar o dia com IA (1 chamada)";
    button.title = cooldown.locked ? cooldown.message : "Gerar fechamento diário com IA";
  });
}

function shouldShowSmartDailyReviewPrompt() {
  const today = offsetDate(0);
  if (dailyReviewCooldownStatus().locked) return false;
  if (!isCloudConnected()) return false;
  if (new Date().getHours() < DAILY_REVIEW_SUGGEST_HOUR) return false;
  return sessionStorage.getItem(DAILY_REVIEW_PROMPT_KEY) !== today;
}

function renderSmartDailyReviewPrompt() {
  const prompt = el("smartDailyReviewPrompt");
  if (!prompt) return;
  const show = shouldShowSmartDailyReviewPrompt();
  prompt.classList.toggle("hidden", !show);
  if (!show) return;
  const top3 = dailyReviewData().prioridadesAmanha || [];
  const list = el("smartDailyReviewTop3");
  if (!list) return;
  list.innerHTML = top3.length ? top3.map((item, index) => `
    <button class="smart-review-mini ${priorityColor(Number(item.score || 0))}" data-smart-review-client="${escapeHtml(item.id || "")}">
      <b>${index + 1}</b>
      <span>${escapeHtml(item.primeiroNome || item.nome || "Contato")}</span>
      <em>${escapeHtml(priorityLabel(item.score))}</em>
      <i>Score ${Math.round(Number(item.score || 0))}</i>
    </button>
  `).join("") : `<span class="smart-review-empty">Sem prioridade crítica agora.</span>`;
  list.querySelectorAll("[data-smart-review-client]").forEach(button => {
    button.addEventListener("click", () => {
      const id = button.dataset.smartReviewClient;
      if (id) openClientModal(id);
    });
  });
}

function renderAutomationCenter(created = 0) {
  if (!el("automationSummary")) return;
  const data = automationCenterData();
  const total = data.confirmations.length + data.plans.length + data.pcs.length + data.c2s.length + data.aftercare.length;
  el("automationConfirmations").textContent = data.confirmations.length;
  el("automationPlans").textContent = data.plans.length;
  el("automationPcs").textContent = data.pcs.length;
  el("automationC2").textContent = data.c2s.length;
  el("automationAftercare").textContent = data.aftercare.length;
  const apiUsage = state.settings.apiUsage;
  const currentMonth = apiUsage?.month === offsetDate(0).slice(0, 7);
  const requests = currentMonth ? Number(apiUsage.requests || 0) : 0;
  el("apiUsageMonth").textContent = `${requests} ${requests === 1 ? "chamada paga" : "chamadas pagas"} neste mês`;
  el("automationSummary").textContent = created
    ? `${created} nova(s) ação(ões) foram preparadas. ${total} item(ns) pedem sua atenção.`
    : total
      ? `${total} item(ns) estão organizados e aguardam sua execução.`
      : "Nenhuma lacuna operacional encontrada. Seu funil está organizado.";
  el("automationFeed").innerHTML = data.feed.length
    ? data.feed.map(item => `<div class="automation-feed-item">
        <span>${escapeHtml(item.label)}${item.date ? ` · ${formatDate(item.date)}` : ""}</span>
        <button data-automation-client="${item.clientId}">Abrir</button>
      </div>`).join("")
    : `<div class="automation-feed-item"><span>Nenhuma pendência operacional crítica.</span></div>`;
  document.querySelectorAll("[data-automation-client]").forEach(button => button.addEventListener("click", () => {
    openClientModal(button.dataset.automationClient);
    switchFormTab("management");
  }));
  renderDailyReview();
  renderSmartDailyReviewPrompt();
}

function renderLeakRadar() {
  if (!el("leakHealthScore")) return;
  const today = offsetDate(0);
  const active = state.clients.filter(client => client.outcome !== "N" && !client.doNotContact);
  const noNextStep = active.filter(client => !client.nextAction || !client.followUp);
  const oiNoPc = active.filter(client => client.oiDone && !client.pcScheduledDate && !client.pcDone);
  const planOverdue = active.filter(client =>
    client.oiDone
    && client.planStatus !== "Pronto"
    && client.planDueDate
    && client.planDueDate < today
  );
  const pcNoOutcome = active.filter(client => client.pcDone && !client.outcome);
  const c2Stale = active.filter(client =>
    client.outcome === "C2"
    && (!client.followUp || client.followUp <= today || (daysSince(client.lastInteraction) ?? 0) >= 6)
  );
  const score = Math.max(0, 100
    - (noNextStep.length * 4)
    - (oiNoPc.length * 8)
    - (planOverdue.length * 10)
    - (pcNoOutcome.length * 12)
    - (c2Stale.length * 10));
  const totalLeaks = noNextStep.length + oiNoPc.length + planOverdue.length + pcNoOutcome.length + c2Stale.length;
  const label = score >= 85 ? "Blindado" : score >= 65 ? "Atencao" : "Risco";
  const headline = !totalLeaks
    ? "Funil limpo: nenhuma fuga crítica agora."
    : `${totalLeaks} vazamento(s) pedem decisão comercial.`;
  const worst = [
    { label: "contatos sem próxima ação", count: noNextStep.length },
    { label: "OIs sem PC", count: oiNoPc.length },
    { label: "planos vencidos", count: planOverdue.length },
    { label: "PCs sem decisão", count: pcNoOutcome.length },
    { label: "C2s esfriando", count: c2Stale.length }
  ].sort((a, b) => b.count - a.count)[0];

  el("leakHealthScore").textContent = score;
  el("leakHealthLabel").textContent = label;
  el("leakRadarHeadline").textContent = headline;
  el("leakRadarSummary").textContent = worst?.count
    ? `Maior vazamento agora: ${worst.count} ${worst.label}. Clique no cartao e resolva em lote.`
    : "Use este painel como auditoria semanal: ele mostra onde o CRM esta protegendo sua carteira.";
  el("leakNoNextStep").textContent = noNextStep.length;
  el("leakOiNoPc").textContent = oiNoPc.length;
  el("leakPlanOverdue").textContent = planOverdue.length;
  el("leakPcNoOutcome").textContent = pcNoOutcome.length;
  el("leakC2Stale").textContent = c2Stale.length;
}

function markPlanReady(client) {
  client.planStatus = "Pronto";
  client.planReadyDate = offsetDate(0);
  client.planProfile = true;
  client.planGoals = true;
  client.planProtection = true;
  client.planBudget = true;
  client.planAlternatives = true;
  state.activities
    .filter(activity => activity.clientId === client.id && activity.type === "Plano")
    .forEach(activity => { activity.completed = true; });
  client.nextAction = client.pcScheduledDate ? "Confirmar PC e revisar apresentação" : "Agendar reunião PC";
  client.followUp = client.pcScheduledDate || nextBusinessDate();
  logClientEvent(client.id, "Plano pronto", "Checklist de preparação concluído.", "Historico");
}

function compactUtc(date) {
  return date.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
}

function graphLocalDateTime(date) {
  const pad = value => String(value).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}:00`;
}

function meetingDetailsFromModal() {
  const client = state.clients.find(item => item.id === el("meetingClientId").value);
  const kind = el("meetingKind").value;
  const date = el("meetingDate").value;
  const time = el("meetingTime").value;
  const duration = Number(el("meetingDuration").value || 60);
  const link = el("meetingLink").value.trim();
  const email = el("meetingEmail").value.trim();
  const title = `${kind} - ${client?.name || "Contato"}`;
  const description = [
    kind === "OI" ? "Reunião inicial de planejamento." : "Reunião de fechamento e próximos passos.",
    link ? `Link da reunião: ${link}` : "",
    "Evento preparado pelo SITPLAN CRM."
  ].filter(Boolean).join("\n\n");
  const { start, end } = meetingDateTime(date, time, duration);
  return { client, kind, date, time, duration, link, email, title, description, start, end };
}

function meetingConfirmationMessage(details) {
  const firstName = details.client?.name?.trim().split(/\s+/)[0] || "";
  const purpose = details.kind === "OI"
    ? "nossa conversa inicial de planejamento"
    : "nossa reunião de fechamento";
  const linkText = details.link ? `\n\nLink para entrar: ${details.link}` : "";
  return `Olá, ${firstName}! Passando para confirmar ${purpose} no dia ${formatMeetingDate(details.date, details.time)}. Reservei ${details.duration} minutos para conversarmos com tranquilidade.${linkText}\n\nSe precisar ajustar o horário, me avise por aqui. Combinado?`;
}

function meetingDayCommitmentMessage(details) {
  const firstName = details.client?.name?.trim().split(/\s+/)[0] || "";
  const purpose = details.kind === "OI"
    ? "nossa conversa de planejamento"
    : "nossa reunião de fechamento";
  const linkText = details.link ? `\n\nSegue o link do Teams para entrar no horário combinado:\n${details.link}` : "";
  return `Olá, ${firstName}! Tudo bem?\n\nPassando para reforçar nosso compromisso de hoje, ${formatMeetingDate(details.date, details.time)}, para ${purpose}. Já deixei esse horário reservado para você e vou conduzir de forma objetiva, com calma e foco no seu momento.${linkText}\n\nSe puder, entre alguns minutinhos antes para evitar qualquer problema de acesso. Combinado?`;
}

function updateMeetingPreview() {
  if (!el("meetingDate").value || !el("meetingTime").value) return;
  const details = meetingDetailsFromModal();
  el("meetingDateSummary").textContent = `${details.kind} · ${formatMeetingDate(details.date, details.time)} · ${details.duration} min`;
  el("meetingMessage").value = meetingConfirmationMessage(details);
}

function updateMeetingOutcomePanel(kind) {
  const isOi = kind === "OI";
  el("meetingOutcomeTitle").textContent = isOi ? "Registrar resultado da OI" : "Registrar resultado do PC";
  el("meetingOutcomeHint").textContent = isOi
    ? "Registre o diagnóstico, as recomendações e transforme a OI em um plano de ação."
    : "Registre a decisão para manter sua conversão e seus follow-ups corretos.";
  el("meetingOiOutcomeActions").classList.toggle("hidden", !isOi);
  el("meetingPcOutcomeActions").classList.toggle("hidden", isOi);
}

function openMeetingModal(kind, explicitClientId = "") {
  const clientId = explicitClientId || el("clientId").value;
  const client = state.clients.find(item => item.id === clientId);
  if (!client) {
    showToast("Salve o contato antes de preparar o convite", true);
    return;
  }
  const isOi = kind === "OI";
  el("meetingClientId").value = client.id;
  el("meetingKind").value = kind;
  el("meetingModalTitle").textContent = `Preparar convite da ${kind}`;
  el("meetingContactName").textContent = client.name;
  el("meetingDate").value = (isOi ? el("clientOiScheduledDate").value : el("clientPcScheduledDate").value)
    || (isOi ? client.oiScheduledDate : client.pcScheduledDate);
  el("meetingTime").value = (isOi ? el("clientOiTime").value : el("clientPcTime").value)
    || (isOi ? client.oiTime : client.pcTime)
    || "19:00";
  el("meetingLink").value = (isOi ? el("clientOiMeetingLink").value : el("clientPcMeetingLink").value)
    || (isOi ? client.oiMeetingLink : client.pcMeetingLink);
  el("meetingEmail").value = client.email || "";
  el("meetingDuration").value = "60";
  updateMeetingOutcomePanel(kind);
  updateMicrosoftStatus();
  el("meetingModal").classList.remove("hidden");
  updateMeetingPreview();
  if (kind === "PC" && client.planStatus !== "Pronto") {
    showToast("Atenção: este PC está sem plano marcado como pronto", true);
  }
}

function updateMicrosoftStatus(message = "") {
  const connected = isMicrosoftConnected();
  el("microsoftStatus").textContent = message || (connected ? "Microsoft 365 conectado" : "Agenda Microsoft não conectada");
  el("microsoftAccessButton").textContent = connected ? "Microsoft 365 conectado" : "Conectar Microsoft 365";
  el("meetingCheckAvailability").disabled = !connected;
  el("meetingCreateMicrosoft").disabled = !connected;
  el("meetingMicrosoftHint").textContent = connected
    ? "Modo manual assistido ativo. Se quiser, use as opções avançadas para Microsoft automático."
    : "Abra o Teams ou Outlook, gere o link, cole acima e salve no CRM.";
}

function openMicrosoftModal() {
  el("microsoftClientId").value = microsoftClientId();
  el("microsoftRedirectUri").textContent = `${window.location.origin}${window.location.pathname}`;
  el("microsoftMessage").textContent = isMicrosoftConnected()
    ? "Sua conta Microsoft 365 está conectada."
    : "Informe o ID público criado no portal Microsoft Entra.";
  el("microsoftModal").classList.remove("hidden");
}

function closeMicrosoftModal() {
  el("microsoftModal").classList.add("hidden");
}

async function submitMicrosoftConnection() {
  const clientId = el("microsoftClientId").value.trim();
  if (!/^[0-9a-f-]{36}$/i.test(clientId)) {
    el("microsoftMessage").textContent = "O ID deve ter o formato 00000000-0000-0000-0000-000000000000.";
    el("microsoftMessage").classList.add("error");
    return;
  }
  configureMicrosoft(clientId);
  el("microsoftMessage").classList.remove("error");
  el("microsoftMessage").textContent = "Abrindo a autorização segura da Microsoft...";
  await connectMicrosoft();
}

async function createMeetingInMicrosoft() {
  if (!isMicrosoftConnected()) {
    openMicrosoftModal();
    return;
  }
  const details = meetingDetailsFromModal();
  if (!details.client || !details.date || !details.time) {
    showToast("Informe data e horário", true);
    return;
  }
  const button = el("meetingCreateMicrosoft");
  button.disabled = true;
  button.textContent = "Criando evento e Teams...";
  try {
    const result = await createTeamsEvent({
      subject: details.title,
      body: details.description,
      startDateTime: graphLocalDateTime(details.start),
      endDateTime: graphLocalDateTime(details.end),
      attendeeEmail: details.email,
      attendeeName: details.client.name,
      transactionId: `${details.client.id}-${details.kind}-${details.date}-${details.time}`
    });
    el("meetingLink").value = result.joinUrl;
    updateMeetingPreview();
    if (details.kind === "OI") details.client.oiMicrosoftEventId = result.id;
    else details.client.pcMicrosoftEventId = result.id;
    saveMeetingDetails({ close: false });
    showToast(result.joinUrl ? "Evento e Teams criados" : "Evento criado; confira o Teams no Outlook");
  } catch (error) {
    console.error(error);
    showToast(error.message || "Falha ao criar evento Microsoft", true);
    if (/conect|expir/i.test(error.message || "")) updateMicrosoftStatus("Reconecte sua conta Microsoft");
  } finally {
    button.disabled = !isMicrosoftConnected();
    button.textContent = "Criar evento + Teams automaticamente";
  }
}

async function checkMeetingAvailability() {
  if (!isMicrosoftConnected()) {
    openMicrosoftModal();
    return;
  }
  const details = meetingDetailsFromModal();
  if (!details.client || !details.date || !details.time) {
    showToast("Informe data e horário antes de consultar a agenda", true);
    return;
  }
  const button = el("meetingCheckAvailability");
  const previousText = button.textContent;
  button.disabled = true;
  button.textContent = "Consultando agenda...";
  try {
    const excludeEventId = details.kind === "OI"
      ? details.client.oiMicrosoftEventId
      : details.client.pcMicrosoftEventId;
    const result = await checkCalendarAvailability({
      startDateTime: graphLocalDateTime(details.start),
      endDateTime: graphLocalDateTime(details.end),
      excludeEventId
    });
    if (result.available) {
      showToast("Agenda livre neste horário");
      el("meetingMicrosoftHint").textContent = "Agenda livre. Pode criar o evento + Teams com segurança.";
      return;
    }
    const conflicts = result.conflicts.slice(0, 2).map(item => item.subject).join(", ");
    showToast(`Conflito encontrado: ${conflicts}`, true);
    el("meetingMicrosoftHint").textContent = `Atenção: conflito com ${conflicts}. Escolha outro horário antes de criar o Teams.`;
  } catch (error) {
    console.error(error);
    showToast(error.message || "Não foi possível consultar a agenda Microsoft", true);
    if (/conect|expir/i.test(error.message || "")) updateMicrosoftStatus("Reconecte sua conta Microsoft");
  } finally {
    button.disabled = !isMicrosoftConnected();
    button.textContent = previousText;
  }
}

function closeMeetingModal() {
  fridaySchedulingClientId = "";
  el("meetingModal").classList.add("hidden");
}

function openPostOiModal(clientId) {
  const client = state.clients.find(item => item.id === clientId);
  if (!client) {
    showToast("Salve o contato antes de registrar o Pós-OI", true);
    return;
  }
  resetPostOiAiReview();
  if (!el("clientModal").classList.contains("hidden")) closeClientModal();
  el("postOiForm").reset();
  el("postOiClientId").value = client.id;
  el("postOiClientName").textContent = client.name;
  el("postOiQuickSummary").value = client.oiQuickSummary || client.oiSummaryNotes || "";
  el("postOiObjectives").value = client.oiObjectives || "";
  el("postOiBudget").value = client.oiBudget || "";
  el("postOiDecision").value = client.oiDecision || "";
  el("postOiObjection").value = client.oiObjection || "";
  el("postOiUrgency").value = client.oiUrgency || "";
  el("postOiNotes").value = client.oiDetailNotes || "";
  el("postOiRecs").value = Number(client.oiRecsCaptured || 0);
  el("postOiPcDate").value = client.postOiSuggestedPcDate || client.pcScheduledDate || businessDateAfter(client.oiDate || offsetDate(0), 5);
  el("postOiPcTime").value = client.postOiSuggestedPcTime || client.pcTime || "19:00";
  const selectedProtections = new Set(client.oiProtectionAreas || []);
  document.querySelectorAll("[data-post-oi-protection]").forEach(input => {
    input.checked = selectedProtections.has(input.value);
  });
  el("postOiModal").classList.remove("hidden");
  setTimeout(() => el("postOiObjectives").focus(), 50);
}

function closePostOiModal() {
  el("postOiModal").classList.add("hidden");
}

function resetPostOiAiReview() {
  pendingPostOiRecommendations = [];
  el("postOiAiReview")?.classList.add("hidden");
  if (el("postOiAiReview")) el("postOiAiReview").innerHTML = "";
  if (el("postOiDocumentFile")) el("postOiDocumentFile").value = "";
}

function setSelectByText(selectId, value) {
  const select = el(selectId);
  const normalized = normalizedPersonName(value);
  if (!select || !normalized) return;
  const option = [...select.options].find(item => {
    const candidate = normalizedPersonName(item.value || item.textContent);
    return candidate === normalized || candidate.includes(normalized) || normalized.includes(candidate);
  });
  if (option) select.value = option.value;
}

function applyPostOiDiagnosis(result) {
  const diagnosis = result?.diagnosis || {};
  if (diagnosis.objectives) el("postOiObjectives").value = String(diagnosis.objectives).slice(0, 600);
  if (diagnosis.notes) {
    el("postOiQuickSummary").value = String(diagnosis.notes).slice(0, 900);
    el("postOiNotes").value = String(diagnosis.notes).slice(0, 600);
  }
  setSelectByText("postOiBudget", diagnosis.budget);
  setSelectByText("postOiDecision", diagnosis.decision);
  setSelectByText("postOiObjection", diagnosis.objection);
  setSelectByText("postOiUrgency", diagnosis.urgency);

  const protections = Array.isArray(diagnosis.protections) ? diagnosis.protections.map(normalizedPersonName) : [];
  document.querySelectorAll("[data-post-oi-protection]").forEach(input => {
    const key = normalizedPersonName(input.value);
    input.checked = protections.some(item => item === key || item.includes(key) || key.includes(item));
  });

  if (/^\d{4}-\d{2}-\d{2}$/.test(String(diagnosis.pcDate || ""))) el("postOiPcDate").value = diagnosis.pcDate;
  if (/^\d{2}:\d{2}$/.test(String(diagnosis.pcTime || ""))) el("postOiPcTime").value = diagnosis.pcTime;

  pendingPostOiRecommendations = (Array.isArray(result.recommendations) ? result.recommendations : [])
    .map(item => ({
      name: String(item.name || "").trim().replace(/\s+/g, " "),
      phone: normalizeWhatsapp(item.phone || ""),
      email: "",
      referrer: "",
      source: "Indicação",
      relationship: String(item.relationship || "").trim(),
      notes: String(item.notes || "").trim(),
      confidence: String(item.confidence || "medium"),
      selected: true,
      duplicate: false,
      invalid: false
    }))
    .filter(item => item.name);

  const recsCount = Math.max(Number(diagnosis.recsCount || 0), pendingPostOiRecommendations.length);
  if (recsCount) el("postOiRecs").value = recsCount;

  pendingPostOiRecommendations = pendingPostOiRecommendations.map(item => {
    const duplicate = recommendationDuplicate(item.name, item.phone);
    const validPhone = item.phone.length >= 10 && item.phone.length <= 11;
    return { ...item, selected: Boolean(validPhone && !duplicate), duplicate: Boolean(duplicate), invalid: !validPhone };
  });

  const review = el("postOiAiReview");
  if (!review) return;
  const recNames = pendingPostOiRecommendations.slice(0, 8).map(item => `<li>${escapeHtml(item.name)}${item.phone ? ` · ${escapeHtml(item.phone)}` : " · sem WhatsApp"}</li>`).join("");
  review.innerHTML = `
    <strong>Leitura automática pronta para revisão</strong>
    <div>Objetivo: ${escapeHtml(el("postOiObjectives").value || "não identificado")}</div>
    <div>PC sugerido: ${escapeHtml(el("postOiPcDate").value || "sem data")} ${escapeHtml(el("postOiPcTime").value || "")}</div>
    <div>RECs encontradas: ${pendingPostOiRecommendations.length || recsCount || 0}</div>
    ${recNames ? `<ul>${recNames}</ul>` : ""}
    <small>Revise os campos antes de salvar. Nada foi gravado ainda.</small>
  `;
  review.classList.remove("hidden");
}

function openPendingPostOiRecommendations(referrerId) {
  if (!pendingPostOiRecommendations.length) {
    openPdfModal("recommendations", referrerId);
    return;
  }
  pdfImportMode = "recommendations";
  pdfReferrerId = referrerId;
  importedPdfContacts = pendingPostOiRecommendations.map(item => ({ ...item }));
  closePostOiModal();
  el("pdfImportMode").value = "recommendations";
  el("pdfReferrer").innerHTML = state.clients.length
    ? state.clients.map(client => `<option value="${client.id}">${escapeHtml(client.name)}</option>`).join("")
    : `<option value="">Nenhum contato cadastrado</option>`;
  el("pdfReferrer").value = referrerId;
  updatePdfImportMode();
  el("pdfUploadStep").classList.add("hidden");
  el("pdfLoadingStep").classList.add("hidden");
  renderPdfReview("Documento da OI");
  el("pdfReviewStep").classList.remove("hidden");
  el("pdfModal").classList.remove("hidden");
}

async function processPostOiDocumentFile(file) {
  if (!file || (!file.name.toLowerCase().endsWith(".pdf") && file.type !== "application/pdf")) {
    showToast("Selecione um PDF da OI", true);
    return;
  }
  const button = el("postOiReadDocument");
  button.disabled = true;
  button.textContent = "Lendo documento...";
  try {
    const { lines } = await extractPdfLines(file);
    const result = await extractPostOiDocument({
      fileName: file.name,
      fileData: await fileAsDataUrl(file),
      extractedText: lines.join("\n")
    });
    applyPostOiDiagnosis(result);
    if (result.usage) registerPaidApiUsage(result, "Diagnóstico Pós-OI por PDF");
    showToast("Documento da OI lido. Revise antes de salvar.");
  } catch (error) {
    console.error(error);
    showToast(error.message || "Não foi possível ler o documento da OI", true);
  } finally {
    button.disabled = false;
    button.textContent = "Escolher PDF da OI";
    el("postOiDocumentFile").value = "";
  }
}

function postOiBriefing(data) {
  return [
    data.quick ? `Resumo livre da OI: ${data.quick}` : "",
    data.objectives ? `Objetivo principal: ${data.objectives}` : "",
    `Prioridades: ${data.protections.length ? data.protections.join(", ") : "a validar"}`,
    `Faixa de conforto: ${data.budget || "não definida"}`,
    `Decisão: ${data.decision || "processo decisório não informado"}`,
    `Ponto de atenção: ${data.objection || "nenhum identificado"}`,
    `Urgência: ${data.urgency || "não definida"}`,
    data.notes ? `Contexto para o PC: ${data.notes}` : ""
  ].filter(Boolean).join("\n");
}

function budgetValueFromRange(value = "") {
  const text = String(value || "");
  if (/Acima/i.test(text)) return 2500;
  if (/300/.test(text)) return 300;
  if (/600/.test(text)) return 600;
  if (/1\.000|1000/.test(text)) return 1000;
  if (/2\.000|2000/.test(text)) return 2000;
  return 0;
}

function plannerClientFromForm() {
  const existing = state.clients.find(client => client.id === el("clientId")?.value) || {};
  return {
    ...existing,
    name: el("clientName")?.value.trim() || existing.name || "",
    incomeRange: el("clientIncomeRange")?.value || existing.incomeRange || "",
    value: Number(el("clientValue")?.value || existing.value || 0),
    planMonthlyBudget: Number(el("clientPlanMonthlyBudget")?.value || existing.planMonthlyBudget || 0),
    planFixedCost: Number(el("clientPlanFixedCost")?.value || existing.planFixedCost || 0),
    planLifestyle: el("clientPlanLifestyle")?.value.trim() || existing.planLifestyle || "",
    planRiskConcerns: el("clientPlanRiskConcerns")?.value.trim() || existing.planRiskConcerns || ""
  };
}

function pcPlannerStrategy(client = {}) {
  const firstName = String(client.name || "cliente").trim().split(/\s+/)[0] || "cliente";
  const budget = Number(client.planMonthlyBudget || client.value || budgetValueFromRange(client.oiBudget) || 0);
  const fixedCost = Number(client.planFixedCost || 0);
  const riskText = compactText(
    client.planRiskConcerns
      || (Array.isArray(client.oiProtectionAreas) ? client.oiProtectionAreas.join(", ") : "")
      || client.oiObjectives
      || client.oiDetailNotes
      || client.oiSummaryNotes
      || client.notes,
    "Riscos ainda não registrados."
  );
  const lifestyle = compactText(client.planLifestyle, "Padrão de vida desejado ainda não registrado.");
  const missing = [];
  if (!budget) missing.push("capacidade mensal");
  if (!fixedCost) missing.push("custo fixo");
  if (!client.planLifestyle) missing.push("padrão de vida desejado");
  if (!client.planRiskConcerns && !client.oiProtectionAreas?.length && !client.oiObjectives) missing.push("riscos prioritários");

  let angle = "Validar prioridade e orçamento antes de apresentar alternativa.";
  if (budget > 0 && budget <= 300) angle = "Começar por uma proteção essencial, simples e defensável.";
  if (budget > 300 && budget <= 600) angle = "Priorizar renda familiar e risco mais doloroso, sem abrir muitas alternativas.";
  if (budget > 600 && budget <= 1000) angle = "Montar um núcleo forte de proteção e deixar evolução futura clara.";
  if (budget > 1000 && budget <= 2000) angle = "Trabalhar camadas: proteção principal, dependentes e continuidade de renda.";
  if (budget > 2000) angle = "Construir estratégia mais completa, conectando renda, sucessão e longo prazo.";

  const nextQuestion = missing.length
    ? `Antes do PC, confirmar ${missing.slice(0, 2).join(" e ")}.`
    : "Você já tem insumos suficientes para entrar no PC com direção.";
  const closingLine = `No PC com ${firstName}, comece retomando o risco principal, mostre que respeitou a capacidade mensal e conduza para decisão com duas alternativas claras.`;

  return { budget, fixedCost, riskText, lifestyle, missing, angle, nextQuestion, closingLine };
}

function pcReadinessData(client = {}, plan = pcPlannerStrategy(client)) {
  const checks = [
    {
      label: "Dor ou objetivo claro",
      ok: Boolean(client.oiObjectives || client.oiQuickSummary || client.oiSummaryNotes || client.planNotes),
      points: 18
    },
    {
      label: "Risco dominante identificado",
      ok: Boolean(client.planRiskConcerns || client.oiProtectionAreas?.length || client.oiObjectives || client.oiDetailNotes || client.oiSummaryNotes),
      points: 18
    },
    {
      label: "Capacidade mensal",
      ok: Boolean(plan.budget),
      points: 14
    },
    {
      label: "Decisor mapeado",
      ok: Boolean(client.oiDecision),
      points: 12
    },
    {
      label: "Urgência percebida",
      ok: Boolean(client.oiUrgency),
      points: 10
    },
    {
      label: "Padrão de vida",
      ok: Boolean(client.planLifestyle),
      points: 10
    },
    {
      label: "PC com data sugerida",
      ok: Boolean(client.pcScheduledDate || client.postOiSuggestedPcDate),
      points: 10
    },
    {
      label: "Plano encaminhado",
      ok: Boolean(client.planStatus === "Pronto" || client.planNotes),
      points: 8
    }
  ];
  const score = checks.reduce((sum, item) => sum + (item.ok ? item.points : 0), 0);
  const missing = checks.filter(item => !item.ok).map(item => item.label);
  const label = score >= 80 ? "PC forte" : score >= 60 ? "PC promissor" : score >= 40 ? "PC frágil" : "OI incompleta";
  const tone = score >= 80 ? "good" : score >= 60 ? "warn" : "risk";
  const action = score >= 80
    ? "Entre no PC com duas alternativas claras e peça decisão."
    : missing.length
      ? `Antes do PC, confirme: ${missing.slice(0, 3).join(", ")}.`
      : "Revise o diagnóstico antes de apresentar o plano.";
  return { score, checks, missing, label, tone, action };
}

function renderPcPlanner(client = null) {
  if (!el("clientPcPlanner")) return;
  const draft = client || plannerClientFromForm();
  const plan = pcPlannerStrategy(draft);
  const readiness = pcReadinessData(draft, plan);
  el("clientPcPlanner").innerHTML = `
    <div>
      <p class="eyebrow">PLANEJADOR DE PC</p>
      <h4>Estratégia comercial para chegar no fechamento com clareza</h4>
      <p>Use como preparação. A proposta técnica e valores finais continuam no sistema oficial da Prudential.</p>
    </div>
    <div class="pc-readiness ${readiness.tone}">
      <div>
        <small>Prontidão para PC</small>
        <strong>${readiness.label}</strong>
        <span>${escapeHtml(readiness.action)}</span>
      </div>
      <div class="pc-readiness-score">
        <b>${readiness.score}</b>
        <i><em style="width:${readiness.score}%"></em></i>
      </div>
    </div>
    <div class="pc-planner-grid">
      <article>
        <small>Capacidade mensal</small>
        <strong>${plan.budget ? currency.format(plan.budget) : "A validar"}</strong>
        <span>${plan.fixedCost ? `Custo fixo informado: ${currency.format(plan.fixedCost)}.` : "Registre custo fixo se souber."}</span>
      </article>
      <article>
        <small>Risco dominante</small>
        <strong>${escapeHtml(plan.riskText.slice(0, 74))}${plan.riskText.length > 74 ? "..." : ""}</strong>
        <span>${escapeHtml(plan.lifestyle.slice(0, 92))}${plan.lifestyle.length > 92 ? "..." : ""}</span>
      </article>
      <article>
        <small>Direção do plano</small>
        <strong>${escapeHtml(plan.angle)}</strong>
        <span>${escapeHtml(plan.nextQuestion)}</span>
      </article>
    </div>
    <div class="pc-planner-script">${escapeHtml(plan.closingLine)}</div>
  `;
}

function savePostOi(action = "save") {
  const form = el("postOiForm");
  if (!form.reportValidity()) return false;
  const client = state.clients.find(item => item.id === el("postOiClientId").value);
  if (!client) return false;
  const protections = [...document.querySelectorAll("[data-post-oi-protection]:checked")].map(input => input.value);
  const previousOiRecs = Number(client.oiRecsCaptured || 0);
  const oiRecs = Math.max(0, Number(el("postOiRecs").value || 0));
  const suggestedPcDate = el("postOiPcDate").value;
  const suggestedPcTime = el("postOiPcTime").value || "19:00";
  const quick = el("postOiQuickSummary").value.trim();
  const data = {
    quick,
    objectives: el("postOiObjectives").value.trim(),
    protections,
    budget: el("postOiBudget").value,
    decision: el("postOiDecision").value,
    objection: el("postOiObjection").value,
    urgency: el("postOiUrgency").value,
    notes: el("postOiNotes").value.trim()
  };
  const wasCompleted = Boolean(client.postOiCompletedAt);
  if (!quick && !data.objectives && !data.notes && !protections.length) {
    showToast("Registre ao menos um resumo rápido da OI ou uma informação comercial importante.", true);
    return false;
  }

  client.oiObjectives = data.objectives;
  client.oiProtectionAreas = data.protections;
  client.oiBudget = data.budget;
  client.oiDecision = data.decision;
  client.oiObjection = data.objection;
  client.oiUrgency = data.urgency;
  client.oiQuickSummary = data.quick;
  client.oiDetailNotes = data.notes;
  client.oiSummaryNotes = [data.quick, data.notes].filter(Boolean).join("\n\n");
  client.oiRecsCaptured = oiRecs;
  client.postOiSuggestedPcDate = suggestedPcDate;
  client.postOiSuggestedPcTime = suggestedPcTime;
  client.postOiCompletedAt = new Date().toISOString();
  client.recs = Math.max(0, Number(client.recs || 0) - previousOiRecs + oiRecs);
  if (oiRecs) client.recsDate = offsetDate(0);
  client.oiDone = true;
  client.oiDate ||= offsetDate(0);
  client.lastInteraction = offsetDate(0);
  client.planStatus = client.planStatus === "Pronto" ? "Pronto" : "Em montagem";
  client.planProfile = Boolean(data.quick || data.objectives || data.decision || data.notes);
  client.planGoals = Boolean(data.quick || data.objectives);
  client.planProtection = data.protections.length > 0;
  client.planBudget = Boolean(data.budget);
  client.planNotes = postOiBriefing(data);
  if (!client.planMonthlyBudget) client.planMonthlyBudget = budgetValueFromRange(data.budget);
  if (!client.value && client.planMonthlyBudget) client.value = client.planMonthlyBudget;
  if (!client.planRiskConcerns) client.planRiskConcerns = data.protections.length ? data.protections.join(", ") : data.objectives;
  client.planDueDate = client.planStatus === "Pronto"
    ? client.planDueDate
    : planDueDateFor({ ...client, pcScheduledDate: suggestedPcDate });
  completeMeetingActivity(client.id, "OI");
  applyProcessStage(client, "OI realizada");
  ensurePlanTask(client);
  client.followUp = client.planDueDate;
  client.nextAction = "Montar plano com base no diagnóstico Pós-OI";
  client.completed = false;

  logClientEvent(
    client.id,
    wasCompleted ? "Diagnóstico Pós-OI revisado" : "OI realizada e diagnosticada",
    `${data.protections.length ? data.protections.join(", ") : "Prioridades a validar"} · ${oiRecs} REC(s) registrada(s).`,
    "OI"
  );
  saveState();
  renderAll();
  closePostOiModal();

  if (action === "recs") {
    openPendingPostOiRecommendations(client.id);
    showToast(pendingPostOiRecommendations.length ? "Diagnóstico salvo. Revise as RECs encontradas." : "Diagnóstico salvo. Agora envie o PDF das recomendações");
    return true;
  }
  if (action === "pc") {
    openMeetingModal("PC", client.id);
    el("meetingDate").value = suggestedPcDate;
    el("meetingTime").value = suggestedPcTime;
    updateMeetingPreview();
    showToast("Diagnóstico salvo. Revise e confirme o PC");
    return true;
  }
  showToast("Pós-OI salvo e plano incluído na agenda");
  return true;
}

function saveMeetingDetails({ close = true } = {}) {
  const details = meetingDetailsFromModal();
  if (!details.client || !details.date || !details.time) {
    showToast("Informe a data e o horário da reunião", true);
    return false;
  }
  const isOi = details.kind === "OI";
  if (details.email) details.client.email = details.email;
  if (isOi) {
    details.client.oiScheduledDate = details.date;
    details.client.oiTime = details.time;
    details.client.oiMeetingLink = details.link;
    el("clientOiScheduledDate").value = details.date;
    el("clientOiTime").value = details.time;
    el("clientOiMeetingLink").value = details.link;
    applyProcessStage(details.client, "OI agendada");
  } else {
    details.client.pcScheduledDate = details.date;
    details.client.pcTime = details.time;
    details.client.pcMeetingLink = details.link;
    el("clientPcScheduledDate").value = details.date;
    el("clientPcTime").value = details.time;
    el("clientPcMeetingLink").value = details.link;
    applyProcessStage(details.client, "PC agendada");
  }
  details.client.followUp = details.date;
  details.client.completed = false;
  const existingActivity = state.activities.find(activity =>
    activity.clientId === details.client.id
    && activity.type === details.kind
    && activity.title.startsWith(`${details.kind} -`)
  );
  const meetingChanged = !existingActivity
    || existingActivity.date !== details.date
    || existingActivity.time !== details.time
    || !existingActivity.notes.includes(details.link || "Link da reunião ainda não informado.");
  const activity = {
    id: existingActivity?.id || crypto.randomUUID(),
    title: details.title,
    date: details.date,
    time: details.time,
    type: details.kind,
    clientId: details.client.id,
    notes: details.link ? `Link: ${details.link}` : "Link da reunião ainda não informado.",
    completed: false,
    createdAt: existingActivity?.createdAt || new Date().toISOString()
  };
  if (existingActivity) Object.assign(existingActivity, activity);
  else state.activities.push(activity);
  const reminderTitle = `Confirmar ${details.kind} - ${details.client.name}`;
  const suggestedReminderDate = previousBusinessDate(details.date);
  const reminderDate = suggestedReminderDate < offsetDate(0) ? offsetDate(0) : suggestedReminderDate;
  const existingReminder = state.activities.find(item =>
    item.clientId === details.client.id
    && item.type === "Follow-up"
    && item.title.startsWith(`Confirmar ${details.kind} -`)
  );
  const reminder = {
    id: existingReminder?.id || crypto.randomUUID(),
    title: reminderTitle,
    date: reminderDate,
    time: "09:00",
    type: "Follow-up",
    clientId: details.client.id,
    notes: meetingConfirmationMessage(details),
    completed: false,
    createdAt: existingReminder?.createdAt || new Date().toISOString()
  };
  if (existingReminder) Object.assign(existingReminder, reminder);
  else state.activities.push(reminder);
  if (meetingChanged) {
    logClientEvent(details.client.id, `${details.kind} preparada`, `${formatMeetingDate(details.date, details.time)}${details.link ? " · link registrado" : ""}`, "Historico");
  }
  if (isOi && fridaySchedulingClientId === details.client.id) {
    const alreadyLogged = state.activities.some(activity =>
      activity.clientId === details.client.id
      && activity.date === offsetDate(0)
      && activity.title === "Sexta - OI marcada"
    );
    if (!alreadyLogged) logClientEvent(details.client.id, "Sexta - OI marcada", "OI agendada pelo Modo Sexta.", "Ligacao");
    fridaySchedulingClientId = "";
  }
  saveState();
  renderAll();
  if (close) closeMeetingModal();
  showToast(`${details.kind} salva no CRM`);
  return true;
}

function completeMeetingActivity(clientId, kind) {
  state.activities
    .filter(activity => activity.clientId === clientId && activity.type === kind)
    .forEach(activity => { activity.completed = true; });
}

function registerMeetingOutcome(outcome) {
  const details = meetingDetailsFromModal();
  const client = details.client;
  if (!client) return;
  const today = offsetDate(0);

  if (outcome === "OI_DONE") {
    closeMeetingModal();
    openPostOiModal(client.id);
    return;
  }

  if (outcome === "OI_NO_SHOW") {
    completeMeetingActivity(client.id, "OI");
    client.stage = "OI agendada";
    client.followUp = nextBusinessDate();
    client.nextAction = "Remarcar OI após não comparecimento";
    client.completed = false;
    logClientEvent(client.id, "Não compareceu à OI", "Nova tentativa de agendamento criada.", "Historico");
    saveState();
    renderAll();
    closeMeetingModal();
    showToast("Ausência registrada e retorno agendado");
    return;
  }

  if (client.planStatus !== "Pronto") {
    showToast("Antes de registrar resultado do PC, marque o plano como Pronto.", true);
    closeMeetingModal();
    openClientModal(client.id);
    switchFormTab("management");
    el("commercialProcessSection").open = true;
    setTimeout(() => el("clientPlanStatus").focus(), 80);
    return;
  }

  client.pcDone = true;
  client.pcDate = today;
  completeMeetingActivity(client.id, "PC");
  if (outcome === "N") {
    applyProcessStage(client, "N");
    client.followUp = nextBusinessDate();
    client.nextAction = "Acompanhar emissão e entrega da apólice";
    logClientEvent(client.id, "Venda fechada - N", "Iniciar acompanhamento da emissão.", "Historico");
    showToast("Parabens: venda registrada como N");
  } else {
    applyProcessStage(client, "PC realizada");
    client.followUp = nextBusinessDate();
    client.nextAction = "Registrar motivo do C2 e próximo follow-up";
    logClientEvent(client.id, "PC realizada sem fechamento", "Informe o motivo do C2 para concluir a etapa.", "Historico");
    saveState();
    renderAll();
    closeMeetingModal();
    openClientModal(client.id);
    switchFormTab("management");
    el("commercialProcessSection").open = true;
    el("clientOutcome").value = "C2";
    el("clientOutcomeDate").value = today;
    el("clientFollowUp").value = client.followUp;
    setTimeout(() => el("clientC2Objection").focus(), 80);
    showToast("Para virar C2, informe motivo, próxima ação e follow-up.", true);
    return;
  }
  client.completed = false;
  saveState();
  renderAll();
  closeMeetingModal();
}

function openGoogleCalendar() {
  const details = meetingDetailsFromModal();
  if (!saveMeetingDetails({ close: false })) return;
  const url = new URL("https://calendar.google.com/calendar/render");
  url.searchParams.set("action", "TEMPLATE");
  url.searchParams.set("text", details.title);
  url.searchParams.set("dates", `${compactUtc(details.start)}/${compactUtc(details.end)}`);
  url.searchParams.set("details", details.description);
  if (details.email) url.searchParams.set("add", details.email);
  window.open(url.toString(), "_blank", "noopener");
}

function openOutlookCalendar() {
  const details = meetingDetailsFromModal();
  if (!saveMeetingDetails({ close: false })) return;
  const url = new URL("https://outlook.office.com/calendar/0/deeplink/compose");
  url.searchParams.set("subject", details.title);
  url.searchParams.set("startdt", details.start.toISOString());
  url.searchParams.set("enddt", details.end.toISOString());
  url.searchParams.set("body", details.description);
  if (details.email) url.searchParams.set("to", details.email);
  window.open(url.toString(), "_blank", "noopener");
}

function downloadMeetingIcs() {
  const details = meetingDetailsFromModal();
  if (!saveMeetingDetails({ close: false })) return;
  const escapeIcs = value => String(value).replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/,/g, "\\,").replace(/;/g, "\\;");
  const ics = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//SITPLAN//CRM Comercial//PT-BR",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    "BEGIN:VEVENT",
    `UID:${crypto.randomUUID()}@sitplan.pages.dev`,
    `DTSTAMP:${compactUtc(new Date())}`,
    `DTSTART:${compactUtc(details.start)}`,
    `DTEND:${compactUtc(details.end)}`,
    `SUMMARY:${escapeIcs(details.title)}`,
    `DESCRIPTION:${escapeIcs(details.description)}`,
    details.link ? `URL:${escapeIcs(details.link)}` : "",
    "END:VEVENT",
    "END:VCALENDAR"
  ].filter(Boolean).join("\r\n");
  const blob = new Blob([ics], { type: "text/calendar;charset=utf-8" });
  const anchor = document.createElement("a");
  anchor.href = URL.createObjectURL(blob);
  anchor.download = `${details.kind.toLowerCase()}-${details.client.name.toLowerCase().replace(/[^a-z0-9]+/gi, "-")}.ics`;
  anchor.click();
  URL.revokeObjectURL(anchor.href);
}

function openMeetingWhatsapp() {
  const details = meetingDetailsFromModal();
  if (!saveMeetingDetails({ close: false })) return;
  const digits = normalizeWhatsapp(details.client.whatsapp || details.client.phone);
  if (digits.length < 10) {
    showToast("WhatsApp inválido para este contato", true);
    return;
  }
  const number = digits.startsWith("55") ? digits : `55${digits}`;
  const message = el("meetingMessage").value.trim() || meetingConfirmationMessage(details);
  window.open(`https://wa.me/${number}?text=${encodeURIComponent(message)}`, "_blank", "noopener");
  logClientEvent(details.client.id, "Confirmação de reunião preparada", `${details.kind} · mensagem aberta no WhatsApp.`, "WhatsApp");
  saveState();
}

function openMeetingDayWhatsapp() {
  const details = meetingDetailsFromModal();
  if (!saveMeetingDetails({ close: false })) return;
  const digits = normalizeWhatsapp(details.client.whatsapp || details.client.phone);
  if (digits.length < 10) {
    showToast("WhatsApp inválido para este contato", true);
    return;
  }
  if (!details.link) {
    showToast("Cole ou gere o link do Teams antes de enviar a mensagem do dia", true);
    return;
  }
  const number = digits.startsWith("55") ? digits : `55${digits}`;
  const message = meetingDayCommitmentMessage(meetingDetailsFromModal());
  window.open(`https://wa.me/${number}?text=${encodeURIComponent(message)}`, "_blank", "noopener");
  logClientEvent(details.client.id, `Mensagem do dia enviada - ${details.kind}`, "Link e compromisso reforçados no WhatsApp.", "WhatsApp");
  saveState();
}

function openTeamsScheduler() {
  const details = meetingDetailsFromModal();
  const url = new URL("https://teams.microsoft.com/l/meeting/new");
  url.searchParams.set("subject", details.title);
  if (details.email) url.searchParams.set("attendees", details.email);
  window.open(url.toString(), "_blank", "noopener");
}

function daysSince(value) {
  if (!value) return null;
  const diff = Math.floor((parseLocalDate(offsetDate(0)) - parseLocalDate(value)) / 86400000);
  return Math.max(diff, 0);
}

function contactAlert(client) {
  const status = dateStatus(client.followUp);
  if (status.className === "overdue") return { text: "Atrasado", className: "overdue" };
  if (status.className === "today") return { text: "Hoje", className: "today" };
  if (!client.followUp) return { text: "Sem data", className: "today" };
  return { text: "Em dia", className: "ok" };
}

async function loadCitiesForState(stateCode, selectedCity = "") {
  el("clientCity").innerHTML = `<option value="">${stateCode ? "Carregando municípios..." : "Escolha primeiro o estado"}</option>`;
  el("clientCity").disabled = true;
  if (!stateCode) return;

  try {
    const response = await fetch(`https://servicodados.ibge.gov.br/api/v1/localidades/estados/${stateCode}/municipios?orderBy=nome`);
    if (!response.ok) throw new Error("IBGE indisponível");
    const cities = await response.json();
    const selectedExists = cities.some(city => city.nome === selectedCity);
    el("clientCity").innerHTML = `
      <option value="">Selecione o município</option>
      ${cities.map(city => `<option value="${escapeHtml(city.nome)}">${escapeHtml(city.nome)}</option>`).join("")}
      ${selectedCity && !selectedExists ? `<option value="${escapeHtml(selectedCity)}">${escapeHtml(selectedCity)}</option>` : ""}
    `;
    el("clientCity").disabled = false;
    el("clientCity").value = selectedCity || "";
  } catch (_) {
    el("clientCity").innerHTML = selectedCity
      ? `<option value="${escapeHtml(selectedCity)}">${escapeHtml(selectedCity)}</option>`
      : `<option value="">Municípios indisponíveis</option>`;
    el("clientCity").disabled = !selectedCity;
    showToast("Base de municípios indisponível no momento", true);
  }
}

function applyProcessStage(client, stage) {
  const today = offsetDate(0);
  const previousStage = client.stage;
  const previousOutcome = client.outcome;
  client.stage = stage;
  if (stage === "TA" && !client.taDate) client.taDate = today;
  if (stage === "OI agendada") {
    if (!client.taDate) client.taDate = today;
    if (!client.oiScheduledDate) client.oiScheduledDate = client.followUp || today;
    client.nextAction = "Confirmar presença e enviar os dados da OI";
  }
  if (stage === "OI realizada") {
    client.oiDone = true;
    if (!client.oiDate) client.oiDate = today;
    client.planStatus ||= "Em montagem";
    client.planDueDate ||= planDueDateFor(client);
    ensurePlanTask(client);
    client.nextAction = "Montar plano do cliente";
    client.followUp = client.planDueDate;
  }
  if (stage === "PC agendada") {
    if (!client.pcScheduledDate) client.pcScheduledDate = client.followUp || today;
    if (client.planStatus !== "Pronto") {
      client.planDueDate = planDueDateFor(client);
      ensurePlanTask(client);
      client.nextAction = "Concluir plano antes do PC";
      client.followUp = client.planDueDate;
    } else {
      client.nextAction = "Confirmar o PC";
    }
  }
  if (stage === "PC realizada") {
    client.pcDone = true;
    if (!client.pcDate) client.pcDate = today;
    client.nextAction = "Registrar decisão ou combinar follow-up";
    client.followUp ||= nextBusinessDate();
  }
  if (stage === "N" || stage === "C2") {
    client.outcome = stage;
    if (!client.outcomeDate) client.outcomeDate = today;
    client.nextAction = stage === "N" ? "Acompanhar emissão e entrega da apólice" : (client.c2NextAction || "Realizar follow-up de fechamento");
    if (stage === "C2") {
      client.c2EnteredAt = previousOutcome === "C2" ? (client.c2EnteredAt || client.outcomeDate) : today;
      if (previousOutcome !== "C2") client.c2RecoveredAt = "";
      client.followUp = client.followUp && client.followUp > today ? client.followUp : nextBusinessDate();
    }
    if (stage === "N" && previousOutcome === "C2") client.c2RecoveredAt ||= today;
  }
  if (previousStage && previousStage !== stage) {
    logClientEvent(client.id, "Mudança de etapa", `${previousStage} → ${stage}`, "Historico");
  }
}

function focusProcessField(fieldId, message) {
  switchFormTab("management");
  const section = el("commercialProcessSection");
  if (section) section.open = true;
  showToast(message, true);
  setTimeout(() => el(fieldId)?.focus(), 80);
}

function hasPostOiDiagnosis(client, data = {}) {
  return Boolean(
    data.planNotes
    || client?.postOiCompletedAt
    || client?.oiSummaryNotes
    || client?.oiObjectives
  );
}

function validateClientProcessGate(data, existing = null) {
  if (data.oiScheduledDate && !data.oiTime) {
    return { field: "clientOiTime", message: "OI agendada precisa de horário." };
  }
  if (data.oiDone && !data.oiDate) {
    return { field: "clientOiDate", message: "OI realizada precisa da data da reunião." };
  }
  if (data.oiDone && !hasPostOiDiagnosis(existing, data)) {
    return { field: "reviewPostOi", message: "Antes de avançar, registre o diagnóstico Pós-OI." };
  }
  if (data.planStatus === "Pronto") {
    const planReady = data.planProfile && data.planGoals && data.planProtection && data.planBudget && data.planAlternatives;
    if (!planReady) return { field: "clientPlanProfile", message: "Plano pronto exige checklist completo." };
    if (!data.planNotes) return { field: "clientPlanNotes", message: "Plano pronto precisa de observações para o PC." };
  }
  if (data.pcScheduledDate && !data.pcTime) {
    return { field: "clientPcTime", message: "PC agendada precisa de horário." };
  }
  if (data.pcDone && data.planStatus !== "Pronto") {
    return { field: "clientPlanStatus", message: "Antes de realizar PC, marque o plano como Pronto." };
  }
  if (data.pcDone && !data.pcDate) {
    return { field: "clientPcDate", message: "PC realizada precisa da data da reunião." };
  }
  if (data.pcDone && !data.outcome) {
    return { field: "clientOutcome", message: "PC realizada precisa virar N ou C2." };
  }
  if (data.outcome && !data.pcDone) {
    return { field: "clientPcDone", message: "Antes de registrar N ou C2, marque a PC como realizada." };
  }
  if (data.outcome === "N" && !data.outcomeDate) {
    return { field: "clientOutcomeDate", message: "N precisa da data do fechamento." };
  }
  if (data.outcome === "C2") {
    if (!data.outcomeDate) return { field: "clientOutcomeDate", message: "C2 precisa da data do PC/decisão." };
    if (!data.c2Objection) return { field: "clientC2Objection", message: "C2 precisa de motivo principal." };
    if (!data.c2NextAction) return { field: "clientC2NextAction", message: "C2 precisa de próxima ação objetiva." };
    if (!data.followUp) return { field: "clientFollowUp", message: "C2 precisa de data de follow-up." };
  }
  return null;
}

function escapeHtml(value = "") {
  const div = document.createElement("div");
  div.textContent = value;
  return div.innerHTML;
}

function displayActivityType(type = "") {
  const labels = {
    Ligacao: "Ligação",
    Historico: "Histórico",
    Recomendacao: "Recomendação"
  };
  return labels[type] || type;
}

function displaySystemText(value = "") {
  return String(value)
    .replace(/\bNao\b/g, "Não")
    .replace(/\bnao\b/g, "não")
    .replace(/\bVoce\b/g, "Você")
    .replace(/\bvoce\b/g, "você")
    .replace(/\bproxima\b/g, "próxima")
    .replace(/\bproximo\b/g, "próximo")
    .replace(/\bdecisao\b/g, "decisão")
    .replace(/\bconversao\b/g, "conversão")
    .replace(/\bemissao\b/g, "emissão")
    .replace(/\bapolice\b/g, "apólice")
    .replace(/\bligacao\b/g, "ligação")
    .replace(/\bLigacao\b/g, "Ligação")
    .replace(/\bHistorico\b/g, "Histórico")
    .replace(/\bRecomendacao\b/g, "Recomendação");
}

function filteredClients() {
  return state.clients.filter(client => {
    const haystack = `${client.name} ${client.email} ${client.phone} ${client.whatsapp} ${client.city} ${client.source} ${client.referrer} ${client.referrerRelationship} ${client.referralContext} ${client.nextAction}`.toLowerCase();
    return haystack.includes(searchTerm.toLowerCase());
  });
}

function whatsappLink(client, afterMissedCall = false) {
  const digits = (client.whatsapp || client.phone || "").replace(/\D/g, "");
  if (digits.length < 10) return "";
  const number = digits.startsWith("55") ? digits : `55${digits}`;
  const firstName = client.name.trim().split(/\s+/)[0];
  const context = client.referrer
    ? `O(a) ${client.referrer} me passou seu contato.`
    : client.source === "Indicação" ? "Recebi seu contato por uma recomendação." : "";
  const message = afterMissedCall
    ? fridayMissedCallMessage(client)
    : `Oi, ${firstName}! Tudo bem? ${context} Posso te ligar rapidinho para explicar o motivo do contato?`;
  return `https://wa.me/${number}?text=${encodeURIComponent(message.replace(/\s+/g, " ").trim())}`;
}

function fridayMissedCallMessage(client) {
  const firstName = client.name.trim().split(/\s+/)[0] || client.name || "";
  const referrer = compactText(client.referrer, "");
  const greeting = firstName ? `Oi, ${firstName}!` : "Oi!";
  const context = referrer ? `O(a) ${referrer} me passou seu contato.` : "Queria falar rapidinho com você.";
  return `${greeting} Tudo bem? Tentei te ligar agora. ${context} Posso te ligar mais tarde ou amanhã?`
    .replace(/\s+/g, " ")
    .trim();
}

function fridayReferrerWarmupMessage(group) {
  const names = group.clients.slice(0, 4).map(client => client.name.trim().split(/\s+/)[0]).filter(Boolean);
  if (group.name === "Sem recomendante") {
    return "Vou ligar para alguns contatos agora. Vou manter a abordagem simples e tentar falar por telefone.";
  }
  const list = names.length ? `: ${names.join(", ")}` : "";
  return `Oi, ${group.name}! Tudo bem? Vou tentar falar hoje com alguns nomes que você me passou${list}. Tem algum contexto rápido que me ajude antes de ligar?`;
}

function assistantDays(client) {
  return daysSince(client.lastInteraction) ?? daysSince(client.createdAt?.slice(0, 10)) ?? 0;
}

function assistantPriorityScore(client) {
  const days = assistantDays(client);
  let score = Math.min(days, 20);
  if (client.followUp && client.followUp <= offsetDate(0) && !client.completed) score += 30;
  if (client.oiDone && !client.pcScheduledDate && !client.pcDone && client.outcome !== "N") score += 28;
  if (client.pcDone && client.outcome !== "N") score += 32;
  if (client.outcome === "C2") score += 24;
  if (client.referrer || client.source === "Indicação") score += 12;
  if (client.temperature === "Quente") score += 15;
  if (!client.lastInteraction) score += 10;
  if (client.outcome === "N") score -= 60;
  return score;
}

function sitplanOpportunityScore(client) {
  const reasons = [];
  let score = 0;
  const days = assistantDays(client);
  if (client.followUp && client.followUp <= offsetDate(0) && !client.completed) {
    score += 24;
    reasons.push("follow-up vence agora");
  }
  if (client.referrer || String(client.source || "").includes("Indica")) {
    score += 18;
    reasons.push("veio por recomendação");
  }
  if (client.oiDone && !client.pcScheduledDate && !client.pcDone && client.outcome !== "N") {
    score += 28;
    reasons.push("OI sem PC");
  }
  if (client.pcDone && client.outcome !== "N") {
    score += 30;
    reasons.push("PC sem decisão");
  }
  if (client.outcome === "C2") {
    score += days >= 7 ? 24 : 14;
    reasons.push("C2 para fechar");
  }
  if (client.planStatus !== "Pronto" && client.oiDone && client.outcome !== "N") {
    score += 18;
    reasons.push("plano pendente");
  }
  if (client.temperature === "Quente") {
    score += 12;
    reasons.push("quente");
  }
  if (client.priority === "Alta") {
    score += 10;
    reasons.push("prioridade alta");
  }
  if (!client.lastInteraction) {
    score += 10;
    reasons.push("sem contato registrado");
  } else if (days >= 10) {
    score += 10;
    reasons.push(`${days} dias sem contato`);
  }
  if (client.outcome === "N") score -= 80;
  return { score: Math.max(0, Math.min(100, Math.round(score))), reasons: reasons.slice(0, 3) };
}

function nextBestMovement(client) {
  const today = offsetDate(0);
  const days = assistantDays(client);
  if (client.outcome === "N") {
    if (!client.policyDelivered) return {
      action: "Acompanhar apólice",
      reason: "N fechado sem entrega registrada",
      urgency: "media",
      view: "tasks",
      score: 55
    };
    if (client.nextRelationship && client.nextRelationship <= today) return {
      action: "Cuidar do cliente",
      reason: "relacionamento de pos-venda venceu",
      urgency: "media",
      view: "tasks",
      score: 50
    };
    return { action: "Cliente protegido", reason: "sem pendência crítica", urgency: "baixa", view: "clients", score: 5 };
  }
  if (client.pcDone && !client.outcome) return {
    action: "Registrar N ou C2",
    reason: "PC realizada sem decisão",
    urgency: "crítica",
    view: "clients",
    score: 100
  };
  if (client.outcome === "C2") return {
    action: "Fechar follow-up C2",
    reason: client.c2Objection ? `motivo: ${client.c2Objection}` : `${days} dias desde ultimo contato`,
    urgency: (!client.followUp || client.followUp <= today || days >= 7) ? "crítica" : "alta",
    view: "clients",
    score: (!client.followUp || client.followUp <= today || days >= 7) ? 94 : 76
  };
  if (client.pcScheduledDate && !client.pcDone) {
    if (client.planStatus !== "Pronto") return {
      action: "Montar plano",
      reason: "PC marcada com plano incompleto",
      urgency: "crítica",
      view: "tasks",
      score: 92
    };
    return {
      action: "Confirmar PC",
      reason: `PC marcada para ${formatDate(client.pcScheduledDate)}`,
      urgency: client.pcScheduledDate <= today ? "alta" : "media",
      view: "tasks",
      score: client.pcScheduledDate <= today ? 78 : 60
    };
  }
  if (client.oiDone && !client.pcScheduledDate && !client.pcDone) return {
    action: "Marcar PC",
    reason: "OI realizada sem próximo compromisso",
    urgency: daysSince(client.oiDate) >= 2 ? "crítica" : "alta",
    view: "clients",
    score: daysSince(client.oiDate) >= 2 ? 90 : 74
  };
  if (client.oiDone && client.planStatus !== "Pronto") return {
    action: "Montar plano",
    reason: "diagnóstico Pós-OI precisa virar plano",
    urgency: client.planDueDate && client.planDueDate < today ? "crítica" : "alta",
    view: "tasks",
    score: client.planDueDate && client.planDueDate < today ? 88 : 70
  };
  if (client.oiScheduledDate && !client.oiDone) return {
    action: "Confirmar OI",
    reason: `OI marcada para ${formatDate(client.oiScheduledDate)}`,
    urgency: client.oiScheduledDate <= today ? "alta" : "media",
    view: "tasks",
    score: client.oiScheduledDate <= today ? 72 : 52
  };
  const isReferral = client.referrer || String(client.source || "").includes("Indica");
  if (["Base", "TA"].includes(client.stage) && !client.oiScheduledDate) return {
    action: "Ligar para OI",
    reason: isReferral ? "indicação sem OI marcada" : (client.lastInteraction ? "TA sem OI marcada" : "contato sem abordagem"),
    urgency: isReferral || !client.lastInteraction ? "alta" : "media",
    view: "friday",
    score: isReferral ? 68 : 48
  };
  if (!client.nextAction || !client.followUp) return {
    action: "Definir destino",
    reason: "contato sem próxima ação clara",
    urgency: "media",
    view: "input",
    score: 42
  };
  return {
    action: client.nextAction || "Manter cadência",
    reason: client.followUp ? `follow-up em ${formatDate(client.followUp)}` : "sem urgência crítica",
    urgency: client.followUp <= today ? "alta" : "baixa",
    view: "tasks",
    score: client.followUp <= today ? 58 : 25
  };
}

function rankedNextMovements(limit = 5) {
  return state.clients
    .filter(client => !client.doNotContact)
    .map(client => ({ client, movement: nextBestMovement(client), opportunity: sitplanOpportunityScore(client) }))
    .filter(item => item.client.outcome !== "N" || item.movement.score >= 50)
    .sort((a, b) => (b.movement.score + b.opportunity.score) - (a.movement.score + a.opportunity.score))
    .slice(0, limit);
}

function priorityLabel(score = 0) {
  const value = Number(score || 0);
  if (value >= 75) return "Prioridade máxima";
  if (value >= 55) return "Alta prioridade";
  if (value >= 35) return "Atenção";
  return "Baixa urgência";
}

function c2PatternData() {
  const c2s = state.clients.filter(client => client.outcome === "C2" || client.pcDone && client.outcome !== "N");
  const groups = new Map();
  c2s.forEach(client => {
    const key = client.c2Objection || client.oiObjection || "Não informado";
    groups.set(key, (groups.get(key) || 0) + 1);
  });
  const ranking = [...groups.entries()].sort((a, b) => b[1] - a[1]);
  return { total: c2s.length, top: ranking[0], ranking };
}

function renderGodMode() {
  if (!el("godModeTop5")) return;
  const week = selectedWeek();
  const goal = weeklyGoal(week);
  const oiScheduled = state.clients.filter(client => inSelectedWeek(client.oiScheduledDate, week)).length;
  const missing = Math.max(goal - oiScheduled, 0);
  const backupAge = state.settings.lastBackupAt ? daysSince(state.settings.lastBackupAt) : null;
  const ranked = state.clients
    .filter(client => client.outcome !== "N")
    .map(client => ({ client, ...sitplanOpportunityScore(client) }))
    .sort((a, b) => b.score - a.score)
    .slice(0, 5);
  const referral = referralRankingData()[0];
  const c2Pattern = c2PatternData();

  const backupWarning = backupAge === null || backupAge >= 7
    ? " Backup pendente: exporte uma copia hoje."
    : "";
  el("godModeBriefing").textContent = (missing
    ? `Faltam ${missing} OI(s) para a meta. Hoje eu priorizaria recomendações, OIs sem PC e C2 com decisão pendente.`
    : "Meta semanal de OIs protegida. Agora o jogo é presença, PC bem preparado e fechamento.") + backupWarning;
  el("godModeTop5").innerHTML = ranked.length ? ranked.map((item, index) => `
    <button class="god-mode-contact" data-god-client="${item.client.id}">
      <b>${index + 1}</b>
      <span><strong>${escapeHtml(item.client.name)}</strong><small>${escapeHtml(priorityLabel(item.score))} · ${escapeHtml(item.reasons.join(", ") || "manter cadência")}</small></span>
    </button>
  `).join("") : `<div class="mini-empty">Sem contatos abertos para priorizar.</div>`;
  el("godModeReferralHub").textContent = referral?.name || "Sem dados";
  el("godModeReferralHint").textContent = referral
    ? `${referral.n} N(s), ${referral.leads} indicado(s), ${referral.declared} REC(s) declarada(s).`
    : "Cadastre o campo Indicador nos contatos recomendados.";
  el("godModeC2Pattern").textContent = c2Pattern.top ? c2Pattern.top[0] : "Sem leitura";
  el("godModeC2Hint").textContent = c2Pattern.top
    ? `${c2Pattern.top[1]} caso(s) em ${c2Pattern.total}. Use isso para ajustar seu próximo PC.`
    : "Quando registrar C2, informe o motivo principal.";
  document.querySelectorAll("[data-god-client]").forEach(button => button.addEventListener("click", () => {
    assistantClientId = button.dataset.godClient;
    assistantMode = "next";
    switchView("assistant");
  }));
}

function renderNextBestMovement() {
  if (!el("nextMoveList")) return;
  const items = rankedNextMovements(4);
  if (!items.length) {
    el("nextMoveSummary").textContent = "Nenhum movimento crítico agora. Use o tempo para alimentar novas recomendações.";
    el("nextMoveList").innerHTML = `<div class="mini-empty">Base sem urgência crítica neste momento.</div>`;
    return;
  }
  const first = items[0];
  el("nextMoveSummary").textContent = `${first.movement.action}: ${first.client.name}. ${first.movement.reason}.`;
  el("nextMoveList").innerHTML = items.map((item, index) => {
    const { client, movement, opportunity } = item;
    const whatsapp = whatsappLink(client);
    const urgency = movement.urgency || "media";
    const urgencyClass = urgency.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    return `
      <article class="next-move-item">
        <b class="next-move-rank">${index + 1}</b>
        <div class="next-move-copy">
          <strong>${escapeHtml(movement.action)} - ${escapeHtml(client.name)}</strong>
          <span>${escapeHtml(priorityLabel(movement.score))} · ${escapeHtml(movement.reason)} · ${escapeHtml(opportunity.reasons.join(", ") || client.stage)}</span>
          <i class="next-move-pill ${urgencyClass}">${escapeHtml(urgency)}</i>
        </div>
        <div class="next-move-actions">
          ${whatsapp ? `<a href="${whatsapp}" target="_blank" rel="noopener" data-next-whatsapp="${client.id}">WhatsApp</a>` : ""}
          <button data-next-open="${client.id}">Abrir ficha</button>
          <button data-next-view="${movement.view}" data-next-client="${client.id}">Registrar ação</button>
        </div>
      </article>`;
  }).join("");
  document.querySelectorAll("[data-next-open]").forEach(button => button.addEventListener("click", () => openClientModal(button.dataset.nextOpen)));
  document.querySelectorAll("[data-next-view]").forEach(button => button.addEventListener("click", () => {
    const view = button.dataset.nextView || "clients";
    const clientId = button.dataset.nextClient;
    if (view === "clients") {
      openClientModal(clientId);
      return;
    }
    if (view === "friday") workFilter = "friday-all";
    switchView(view);
    if (view === "assistant") {
      assistantClientId = clientId;
      assistantMode = "next";
      renderAssistant();
    }
  }));
  document.querySelectorAll("[data-next-whatsapp]").forEach(link => link.addEventListener("click", () => {
    const client = state.clients.find(item => item.id === link.dataset.nextWhatsapp);
    if (!client) return;
    client.lastInteraction = offsetDate(0);
    if (!client.taDate && ["Base", "TA"].includes(client.stage)) client.taDate = offsetDate(0);
    client.nextAction = "Aguardar retorno no WhatsApp";
    client.followUp = nextBusinessDate();
    client.completed = false;
    logClientEvent(client.id, "WhatsApp iniciado pelo Proximo Melhor Movimento", "Contato acionado pela prioridade do SITPLAN.", "WhatsApp");
    saveState();
    renderAll();
  }));
}

function weeklyContract(week = selectedWeek()) {
  const campaign = campaignData();
  const nTarget = Math.max(campaign.settings.threeWWeeklyTarget, Math.ceil(campaign.neededPerWeek));
  const defaults = {
    oi: Math.max(weeklyGoal(week), nTarget * 4, DEFAULT_WEEKLY_OI_GOAL),
    pc: Math.max(5, nTarget * 2),
    c2: Math.max(5, nTarget),
    plans: Math.max(5, nTarget * 2),
    hub: ""
  };
  return { ...defaults, ...(state.settings.weeklyContracts?.[week.startKey] || {}) };
}

function weeklyBoardData(week = selectedWeek()) {
  const contract = weeklyContract(week);
  const clients = state.clients;
  const oi = clients.filter(client => inSelectedWeek(client.oiScheduledDate, week)).length;
  const pc = clients.filter(client => client.pcDone && inSelectedWeek(client.pcDate, week)).length;
  const c2Touched = clients.filter(client => client.outcome === "C2" && inSelectedWeek(client.lastInteraction, week)).length;
  const plans = clients.filter(client => client.planStatus === "Pronto" && inSelectedWeek(client.planReadyDate, week)).length;
  const ta = clients.filter(client => inSelectedWeek(client.taDate, week)).length;
  const sales = clients.filter(client => client.outcome === "N" && inSelectedWeek(client.outcomeDate, week)).length;
  const oiWithoutPc = clients.filter(client => client.oiDone && !client.pcScheduledDate && !client.pcDone && client.outcome !== "N");
  const pcWithoutDecision = clients.filter(client => client.pcDone && client.outcome !== "N");
  const staleC2 = clients.filter(client => client.outcome === "C2" && (daysSince(client.lastInteraction) ?? 0) >= 7);
  const topRef = referralRankingData()[0];
  return { contract, oi, pc, c2Touched, plans, ta, sales, oiWithoutPc, pcWithoutDecision, staleC2, topRef };
}

function weekFromOffset(offset = 0) {
  const start = startOfWeek(offset);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  return { start, end, startKey: dateKey(start), endKey: dateKey(end) };
}

function currentWeek() {
  return weekFromOffset(0);
}

function activityText(activity) {
  return `${activity.title || ""} ${activity.type || ""} ${activity.notes || ""}`.toLocaleLowerCase("pt-BR");
}

function dailyRitualData() {
  const today = offsetDate(0);
  const week = currentWeek();
  const activitiesToday = state.activities.filter(activity => activity.date === today);
  const board = weeklyBoardData(week);
  const recsToday = state.clients
    .filter(client => client.recsDate === today)
    .reduce((sum, client) => sum + Number(client.recs || 0), 0);
  const taActivities = activitiesToday.filter(activity => {
    const text = activityText(activity);
    return text.includes("ta") || text.includes("ligacao") || text.includes("ligação") || text.includes("sexta");
  }).length;
  const taToday = Math.max(
    state.clients.filter(client => client.taDate === today).length,
    taActivities
  );
  const oiDoneToday = state.clients.filter(client => client.oiDone && client.oiDate === today).length;
  const plansToday = state.clients.filter(client => client.planStatus === "Pronto" && client.planReadyDate === today).length;
  const c2Due = state.clients.filter(client => client.outcome === "C2" && (!client.followUp || client.followUp <= today));
  const c2TouchedToday = state.clients.filter(client => client.outcome === "C2" && client.lastInteraction === today).length;
  const targets = { ta: 10, oi: 1, recs: 3, plans: 1, c2: Math.max(1, Math.min(c2Due.length, 3)) };
  const checks = [
    { key: "ta", title: "TA / abordagens", actual: taToday, target: targets.ta, detail: "ligações ou abordagens registradas hoje", action: "Abrir TA", view: "friday" },
    { key: "oi", title: "OIs realizadas", actual: oiDoneToday, target: targets.oi, detail: "reuniões de OI concluídas hoje", action: "Lançar OI", view: "input" },
    { key: "recs", title: "RECs recebidas", actual: recsToday, target: targets.recs, detail: "recomendacoes capturadas hoje", action: "Lancar REC", view: "input" },
    { key: "plans", title: "Planos montados", actual: plansToday, target: targets.plans, detail: "planos prontos para PC", action: "Ver agenda", view: "tasks" },
    { key: "c2", title: "Follow-up de C2", actual: c2TouchedToday, target: targets.c2, detail: `${c2Due.length} C2(s) pedem acompanhamento`, action: "Ver C2", view: "clients", filter: "c2-stale" }
  ];
  const pending = [
    ...state.clients
      .filter(client => client.oiDone && !client.pcScheduledDate && !client.pcDone && client.outcome !== "N")
      .map(client => ({ client, tag: "OI sem PC", text: "Marcar PC ou definir motivo real para esperar." })),
    ...state.clients
      .filter(client => client.oiDone && client.outcome !== "N" && client.planStatus !== "Pronto")
      .map(client => ({ client, tag: "Plano", text: "Montar plano antes do PC esfriar." })),
    ...c2Due.map(client => ({ client, tag: "C2", text: client.nextAction || "Fazer follow-up e registrar objeção atual." }))
  ].filter((item, index, list) => list.findIndex(entry => entry.client.id === item.client.id && entry.tag === item.tag) === index)
    .slice(0, 6);
  const weekRows = selectedWeekDates(week).map(day => {
    const dayActivities = state.activities.filter(activity => activity.date === day.key);
    const dayTaActivities = dayActivities.filter(activity => {
      const text = activityText(activity);
      return text.includes("ta") || text.includes("ligacao") || text.includes("ligação") || text.includes("sexta");
    }).length;
    const metrics = [
      { key: "ta", label: "TA", value: Math.max(state.clients.filter(client => client.taDate === day.key).length, dayTaActivities) },
      { key: "oi", label: "OI", value: state.clients.filter(client => client.oiScheduledDate === day.key || (client.oiDone && client.oiDate === day.key)).length },
      { key: "pc", label: "PC", value: state.clients.filter(client => client.pcScheduledDate === day.key || (client.pcDone && client.pcDate === day.key)).length },
      { key: "recs", label: "RECs", value: countRecsOnDate(state.clients, day.key) },
      { key: "plans", label: "Planos", value: state.clients.filter(client => client.planStatus === "Pronto" && client.planReadyDate === day.key).length },
      { key: "c2", label: "C2", value: state.clients.filter(client => client.outcome === "C2" && client.lastInteraction === day.key).length },
      { key: "sales", label: "N", value: state.clients.filter(client => client.outcome === "N" && (client.outcomeDate === day.key || client.cpfDate === day.key)).length }
    ];
    const total = metrics.reduce((sum, metric) => sum + metric.value, 0);
    const scheduled = state.clients.filter(client => client.followUp === day.key && !client.completed).length;
    return { ...day, metrics, total, scheduled };
  });
  return { today, week, board, checks, pending, weekRows };
}

function ritualProgress(done, target) {
  const pct = Math.min(Math.round((done / Math.max(target, 1)) * 100), 100);
  return pct;
}

function goalCpfDate(client) {
  if (!(client.cpfDone || client.outcome === "N")) return "";
  return client.cpfDate || client.outcomeDate || client.pcDate || "";
}

function cpfsInWeek(week) {
  return state.clients.filter(client => inSelectedWeek(goalCpfDate(client), week)).length;
}

function ritualEvolutionData() {
  const goal = annualGoalData();
  const campaign = campaignData();
  const target3w = campaign.settings.threeWWeeklyTarget;
  const weekly = [];
  for (let offset = -7; offset <= 0; offset += 1) {
    const week = weekRangeFromOffset(offset);
    const n = salesInWeek(week);
    weekly.push({
      ...week,
      offset,
      n,
      cpfs: cpfsInWeek(week),
      ok: n >= target3w,
      label: offset === 0 ? "Atual" : week.start.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" })
    });
  }
  let streak = 0;
  for (let offset = 0; offset > -52; offset -= 1) {
    const week = weekRangeFromOffset(offset);
    if (salesInWeek(week) >= target3w) streak += 1;
    else break;
  }
  const deadline = parseLocalDate(goal.deadlineValue) || new Date(2026, 9, 31, 12);
  const months = [];
  for (let month = 0; month <= deadline.getMonth(); month += 1) {
    const key = `${deadline.getFullYear()}-${String(month + 1).padStart(2, "0")}`;
    const date = new Date(deadline.getFullYear(), month, 1, 12);
    const cpfs = state.clients.filter(client => String(goalCpfDate(client)).startsWith(key)).length;
    const n = monthlyStats(key).n;
    months.push({
      key,
      label: date.toLocaleDateString("pt-BR", { month: "short" }).replace(".", ""),
      cpfs,
      n
    });
  }
  const maxMonth = Math.max(...months.map(month => month.cpfs), 1);
  return { goal, campaign, target3w, weekly, streak, months, maxMonth };
}

function renderDailyRitual() {
  if (!el("ritualChecklist")) return;
  const data = dailyRitualData();
  const doneChecks = data.checks.filter(item => item.actual >= item.target).length;
  el("ritualDateLabel").textContent = shortDate.format(parseLocalDate(data.today)).toUpperCase();
  el("ritualScore").textContent = `${doneChecks}/${data.checks.length}`;
  el("ritualChecklist").innerHTML = data.checks.map(item => `
    <article class="ritual-check-row ${item.actual >= item.target ? "done" : ""}">
      <input type="checkbox" ${item.actual >= item.target ? "checked" : ""} disabled aria-label="${escapeHtml(item.title)}">
      <div>
        <strong>${escapeHtml(item.title)}</strong>
        <small>${escapeHtml(item.detail)}</small>
      </div>
      <div class="ritual-check-count">${item.actual}/${item.target}</div>
    </article>
  `).join("");

  const weekMetrics = [
    { label: "OIs agendadas", value: data.board.oi, target: data.board.contract.oi },
    { label: "PCs realizadas", value: data.board.pc, target: data.board.contract.pc },
    { label: "Planos prontos", value: data.board.plans, target: data.board.contract.plans },
    { label: "Ns fechados", value: data.board.sales, target: Math.max(3, Math.ceil((data.board.contract.oi || 20) * .15)) },
    { label: "C2 acompanhados", value: data.board.c2Touched, target: data.board.contract.c2 }
  ];
  el("ritualWeekGrid").innerHTML = weekMetrics.map(item => `
    <article class="ritual-week-metric">
      <small>${escapeHtml(item.label)}</small>
      <strong>${item.value}/${item.target}</strong>
      <span><i style="width:${ritualProgress(item.value, item.target)}%"></i></span>
    </article>
  `).join("");

  const maxDayTotal = Math.max(...data.weekRows.map(row => row.total + row.scheduled), 1);
  el("ritualWeekDays").innerHTML = data.weekRows.map(row => {
    const intensity = Math.max(.08, (row.total + row.scheduled) / maxDayTotal);
    const hasMovement = row.total > 0;
    const action = row.key === data.today
      ? (hasMovement ? "Hoje ja tem movimento. Complete os pilares que faltam." : "Hoje ainda esta zerado. Comece com TA ou lance um movimento.")
      : row.scheduled
        ? `${row.scheduled} pendência(s) ou follow-up neste dia.`
        : hasMovement
          ? "Dia com tracao registrada."
          : "Sem movimento registrado.";
    return `
      <article class="ritual-day-card ${row.key === data.today ? "today" : ""}" style="--ritual-day-intensity:${intensity}">
        <div class="ritual-day-top">
          <div><strong>${row.label}</strong><small>${row.short}</small></div>
          <em>${row.total}</em>
        </div>
        <div class="ritual-day-bar"><i></i></div>
        <div class="ritual-day-pills">
          ${row.metrics.map(metric => `
            <span class="ritual-day-pill ${metric.value ? "active" : ""}">
              ${escapeHtml(metric.label)} <b>${metric.value}</b>
            </span>
          `).join("")}
        </div>
        <div class="ritual-day-action">${escapeHtml(action)}</div>
      </article>
    `;
  }).join("");

  const evolution = ritualEvolutionData();
  const goalPct = Math.min((evolution.goal.done / evolution.goal.target) * 100, 100);
  const weeklyNeed = Math.max(evolution.target3w, Math.ceil(evolution.goal.perWeek || 0));
  el("ritualEvolutionSummary").textContent = evolution.goal.projection >= evolution.goal.target
    ? `Projeção atual: ${evolution.goal.projection} CPFs até outubro. O jogo agora é proteger consistência e qualidade.`
    : `Projecao atual: ${evolution.goal.projection} CPFs. Para outubro, precisa de ${evolution.goal.perWeek.toFixed(1).replace(".", ",")} CPF(s)/semana.`;
  el("ritualEvolutionCards").innerHTML = `
    <article class="ritual-evolution-card ${evolution.goal.projection >= evolution.goal.target ? "good" : "warn"}">
      <small>Meta até outubro</small>
      <strong>${evolution.goal.done}/${evolution.goal.target}</strong>
      <span>${evolution.goal.remaining} faltam · ${Math.round(goalPct)}% concluído</span>
    </article>
    <article class="ritual-evolution-card ${evolution.campaign.currentWeekN >= evolution.target3w ? "good" : "warn"}">
      <small>3W semana atual</small>
      <strong>${evolution.campaign.currentWeekN}/${evolution.target3w} N</strong>
      <span>${evolution.campaign.current3wNeed ? `faltam ${evolution.campaign.current3wNeed} N para proteger` : "semana protegida"}</span>
    </article>
    <article class="ritual-evolution-card ${evolution.streak ? "good" : "warn"}">
      <small>Sequencia 3W</small>
      <strong>${evolution.streak}</strong>
      <span>semana(s) consecutivas com ${evolution.target3w}+ N · alvo final 50</span>
    </article>
  `;
  el("ritualEvolutionWeeks").innerHTML = evolution.weekly.map(week => `
    <article class="ritual-evolution-week ${week.offset === 0 ? "current" : ""} ${week.ok ? "ok" : "risk"}">
      <small>${week.label}</small>
      <strong>${week.n} N</strong>
      <span>${week.ok ? "3W ok" : `faltou ${evolution.target3w - week.n}`}</span>
    </article>
  `).join("");
  el("ritualEvolutionMonths").innerHTML = evolution.months.map(month => `
    <article class="ritual-evolution-month">
      <small>${escapeHtml(month.label)}</small>
      <span><i style="width:${Math.min((month.cpfs / evolution.maxMonth) * 100, 100)}%"></i></span>
      <strong>${month.cpfs}</strong>
    </article>
  `).join("");

  el("ritualPendingList").innerHTML = data.pending.length ? data.pending.map(item => `
    <article class="ritual-pending-item">
      <div>
        <strong><span class="ritual-tag">${escapeHtml(item.tag)}</span>${escapeHtml(item.client.name)}</strong>
        <small>${escapeHtml(item.text)} Etapa: ${escapeHtml(item.client.stage || "Base")}.</small>
      </div>
      <button class="secondary-button" data-ritual-client="${item.client.id}">Abrir contato</button>
    </article>
  `).join("") : `<div class="ritual-empty">Nada crítico pendente agora. Bom sinal: mantenha o ritmo e gere nomes novos.</div>`;

  document.querySelectorAll("[data-ritual-client]").forEach(button => button.addEventListener("click", () => {
    openClientModal(button.dataset.ritualClient);
  }));
}

function truthStatus(done, target) {
  if (!target) return "neutro";
  const pct = done / target;
  if (pct >= 1) return "ok";
  if (pct >= .65) return "warn";
  return "danger";
}

function renderBoardCommercial() {
  if (!el("boardTruths")) return;
  const week = selectedWeek();
  const data = weeklyBoardData(week);
  const { contract } = data;
  el("contractOi").value = contract.oi;
  el("contractPc").value = contract.pc;
  el("contractC2").value = contract.c2;
  el("contractPlans").value = contract.plans;
  el("contractHub").value = contract.hub || "";
  const execution = [data.oi / Math.max(contract.oi, 1), data.pc / Math.max(contract.pc, 1), data.c2Touched / Math.max(contract.c2, 1), data.plans / Math.max(contract.plans, 1)];
  const avg = Math.round((execution.reduce((sum, value) => sum + Math.min(value, 1), 0) / execution.length) * 100);
  el("boardSummary").textContent = `Execução do contrato semanal: ${avg}%. ${avg < 65 ? "O sistema está pedindo foco, não mais planejamento." : "Ritmo em construção; proteja os próximos movimentos."}`;
  const truths = [
    { title: "Verdade de Volume", value: `${data.oi}/${contract.oi} OIs`, text: data.oi >= contract.oi ? "Volume protegido." : `Faltam ${contract.oi - data.oi} OI(s). Priorize TA e recomendacoes.`, status: truthStatus(data.oi, contract.oi) },
    { title: "Verdade de Conversão", value: `${data.sales} N`, text: data.pcWithoutDecision.length ? `${data.pcWithoutDecision.length} PC(s) sem decisão. O gargalo está no fechamento.` : "Sem PC pendente de decisão crítica.", status: data.pcWithoutDecision.length ? "danger" : "ok" },
    { title: "Verdade de Timing", value: `${data.oiWithoutPc.length + data.staleC2.length}`, text: `${data.oiWithoutPc.length} OI(s) sem PC e ${data.staleC2.length} C2(s) frios.`, status: data.oiWithoutPc.length + data.staleC2.length ? "warn" : "ok" },
    { title: "Verdade de Origem", value: data.topRef?.name || "Sem hub", text: data.topRef ? `${data.topRef.n} N(s) e ${data.topRef.leads} indicado(s). Cuide desse ativo.` : "Ainda falta mapear recomendantes.", status: data.topRef ? "ok" : "warn" },
    { title: "Verdade de Execucao", value: `${avg}%`, text: `PCs ${data.pc}/${contract.pc}, C2s ${data.c2Touched}/${contract.c2}, planos ${data.plans}/${contract.plans}.`, status: avg >= 80 ? "ok" : avg >= 55 ? "warn" : "danger" }
  ];
  el("boardTruths").innerHTML = truths.map(item => `
    <article class="board-truth ${item.status}">
      <small>${item.title}</small>
      <strong>${escapeHtml(item.value)}</strong>
      <p>${escapeHtml(item.text)}</p>
    </article>
  `).join("");
}

function referrerLastTouch(name) {
  const client = state.clients.find(item => normalizedPersonName(item.name) === normalizedPersonName(name));
  return client?.lastInteraction || client?.followUp || "";
}

function renderVipReferrers() {
  if (!el("vipGold")) return;
  const ranking = referralRankingData();
  const clients = state.clients;
  const gold = ranking.filter(item => item.n > 0).slice(0, 5);
  const volume = ranking.filter(item => item.leads >= 3 && item.n === 0).slice(0, 5);
  const dormant = ranking.filter(item => item.leads > 0 && (daysSince(referrerLastTouch(item.name)) ?? 999) >= 14).slice(0, 5);
  const referredNames = new Set(ranking.map(item => normalizedPersonName(item.name)));
  const potential = clients
    .filter(client => client.outcome === "N" && !Number(client.recs || 0) && !referredNames.has(normalizedPersonName(client.name)))
    .slice(0, 5)
    .map(client => ({ name: client.name, leads: 0, n: 0, declared: 0, hint: "Cliente N sem REC registrada" }));
  const card = (item, fallback = "") => `
    <button class="vip-referrer-item" data-vip-referrer="${escapeHtml(item.name)}">
      <strong>${escapeHtml(item.name)}</strong>
      <small>${escapeHtml(item.hint || `${item.leads} indicados - ${item.n} Ns - ${item.declared} RECs`)}</small>
    </button>`;
  el("vipReferrerSummary").textContent = `${ranking.length} hubs mapeados`;
  el("vipGold").innerHTML = gold.length ? gold.map(item => card(item)).join("") : `<div class="mini-empty">Ainda sem hub Ouro.</div>`;
  el("vipVolume").innerHTML = volume.length ? volume.map(item => card(item)).join("") : `<div class="mini-empty">Sem hub de volume parado.</div>`;
  el("vipDormant").innerHTML = dormant.length ? dormant.map(item => card(item)).join("") : `<div class="mini-empty">Nenhum hub dormindo.</div>`;
  el("vipPotential").innerHTML = potential.length ? potential.map(item => card(item)).join("") : `<div class="mini-empty">Sem potencial pendente.</div>`;
  document.querySelectorAll("[data-vip-referrer]").forEach(button => button.addEventListener("click", () => {
    switchView("clients");
    renderClients();
    if (el("referrerFilter")) {
      el("referrerFilter").value = button.dataset.vipReferrer;
      renderClients();
    }
  }));
}

function weeksBetween(startDate, endDate) {
  return Math.max(Math.ceil((endDate - startDate) / (7 * 86400000)), 0);
}

function campaignSettings() {
  return {
    octoberTarget: Number(state.settings.campaignOctoberTarget || 100),
    octoberDeadline: state.settings.campaignOctoberDeadline || "2026-10-31",
    threeWStartWeek: Number(state.settings.campaign3wStartWeek || 1),
    threeWCurrentWeek: Number(state.settings.campaign3wCurrentWeek || 6),
    threeWTotalWeeks: Number(state.settings.campaign3wTotalWeeks || 50),
    threeWWeeklyTarget: Number(state.settings.campaign3wWeeklyTarget || 3)
  };
}

function weekRangeFromOffset(offset) {
  const start = startOfWeek(offset);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  return { start, end, startKey: dateKey(start), endKey: dateKey(end) };
}

function salesInWeek(week) {
  return state.clients.filter(client => client.outcome === "N" && inSelectedWeek(client.outcomeDate, week)).length;
}

function campaignTimelineData() {
  const settings = campaignSettings();
  const weeks = [];
  for (let offset = -5; offset <= 4; offset += 1) {
    const week = weekRangeFromOffset(offset);
    const n = salesInWeek(week);
    weeks.push({
      ...week,
      offset,
      n,
      threeWOk: n >= settings.threeWWeeklyTarget,
      label: offset === 0 ? "Atual" : `${week.start.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" })}`
    });
  }
  return weeks;
}

function campaignData() {
  const settings = campaignSettings();
  const deadline = parseLocalDate(settings.octoberDeadline);
  const today = new Date();
  today.setHours(12, 0, 0, 0);
  const totalN = state.clients.filter(client => client.outcome === "N").length;
  const remainingN = Math.max(settings.octoberTarget - totalN, 0);
  const weeksLeft = Math.max(weeksBetween(today, deadline), 1);
  const neededPerWeek = remainingN / weeksLeft;
  const currentWeek = selectedWeek();
  const currentWeekN = salesInWeek(currentWeek);
  const timeline = campaignTimelineData();
  const completed3w = timeline.filter(week => week.offset <= 0 && week.threeWOk).length;
  const current3wNeed = Math.max(settings.threeWWeeklyTarget - currentWeekN, 0);
  const projected = Math.round(totalN + (currentWeekN || 0) * weeksLeft);
  const deadlineLabel = deadline.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });
  return { settings, totalN, remainingN, weeksLeft, neededPerWeek, currentWeekN, current3wNeed, completed3w, projected, timeline, deadlineLabel };
}

function renderCampaignCenter() {
  if (!el("campaign100Progress")) return;
  const data = campaignData();
  const pct100 = Math.min((data.totalN / data.settings.octoberTarget) * 100, 100);
  const threeWPct = Math.min((data.settings.threeWCurrentWeek / data.settings.threeWTotalWeeks) * 100, 100);
  const weeklyTarget = Math.max(data.settings.threeWWeeklyTarget, Math.ceil(data.neededPerWeek));
  el("campaign100Progress").textContent = `${data.totalN}/${data.settings.octoberTarget} N`;
  el("campaign100Detail").textContent = `${data.remainingN} N faltam até ${data.deadlineLabel}. ${data.weeksLeft} semanas: precisa de ${data.neededPerWeek.toFixed(1).replace(".", ",")} N/semana.`;
  el("campaign100Bar").style.width = `${pct100}%`;
  el("campaign3wProgress").textContent = `Semana ${data.settings.threeWCurrentWeek}/${data.settings.threeWTotalWeeks}`;
  el("campaign3wDetail").textContent = data.current3wNeed
    ? `Faltam ${data.current3wNeed} N para salvar a semana 3W.`
    : "3W protegido nesta semana. Agora jogue para 100N.";
  el("campaign3wBar").style.width = `${threeWPct}%`;
  el("campaignWeeklyNeed").textContent = `${weeklyTarget} N`;
  el("campaignWeeklyHint").textContent = data.currentWeekN >= weeklyTarget
    ? "Semana acima do alvo combinado."
    : `Semana atual tem ${data.currentWeekN} N. Faltam ${weeklyTarget - data.currentWeekN} para o alvo de campanha.`;
  el("campaignStatus").textContent = data.projected >= data.settings.octoberTarget ? "No ritmo" : "Atenção";
  el("campaignSummary").textContent = data.projected >= data.settings.octoberTarget
    ? `Projetando ${data.projected} N até ${data.deadlineLabel}. Continue protegendo 3W e acelerando PCs.`
    : `Projetando ${data.projected} N até ${data.deadlineLabel}. Ajuste necessário: elevar PC -> N e recuperar C2.`;
  el("campaignTimeline").innerHTML = data.timeline.map(week => `
    <article class="campaign-week ${week.offset === 0 ? "current" : ""} ${week.threeWOk ? "ok" : "risk"}">
      <small>${week.label}</small>
      <strong>${week.n} N</strong>
      <span>${week.threeWOk ? "3W ok" : `faltou ${data.settings.threeWWeeklyTarget - week.n}`}</span>
    </article>
  `).join("");
}

function assistantRisk(client) {
  const days = assistantDays(client);
  const overdue = client.followUp && client.followUp < offsetDate(0) && !client.completed;
  if ((client.pcDone && client.outcome !== "N") || days >= 14 || (client.outcome === "C2" && days >= 7)) {
    return { level: "Alto", reason: client.pcDone ? "PC realizado sem decisão registrada" : `${days} dias sem contato` };
  }
  if (overdue || (client.oiDone && !client.pcScheduledDate) || days >= 7) {
    return { level: "Médio", reason: overdue ? "follow-up atrasado" : client.oiDone ? "OI realizada sem PC marcada" : `${days} dias sem contato` };
  }
  return { level: "Baixo", reason: "cadência comercial em dia" };
}

function assistantAction(client) {
  const firstName = client.name.trim().split(/\s+/)[0];
  if (client.outcome === "N") {
    return {
      title: "Cuidar do relacionamento",
      action: `Confirme se a entrega e os próximos passos com ${firstName} estão em dia.`,
      talk: "Faça um contato de pós-venda, reforce disponibilidade e prepare o momento certo para solicitar recomendações.",
      channel: "WhatsApp curto ou ligação de relacionamento."
    };
  }
  if (client.pcDone && client.outcome !== "N") {
    return {
      title: "Retomar a decisão do PC",
      action: `Entre em contato com ${firstName} hoje e identifique a objeção que ainda impede a decisão.`,
      talk: "Não reapresente tudo. Pergunte qual ponto precisa ficar mais claro para que a pessoa se sinta segura para decidir.",
      channel: "Ligação direta; WhatsApp apenas para combinar o horário."
    };
  }
  if (client.outcome === "C2") {
    return {
      title: "Follow-up de fechamento",
      action: `Retome o C2 com uma pergunta objetiva e uma razão concreta para conversar agora.`,
      talk: "Resgate a prioridade mencionada no PC e pergunte o que mudou desde a última conversa.",
      channel: "WhatsApp personalizado seguido de ligação."
    };
  }
  if (client.oiDone && !client.pcScheduledDate && !client.pcDone) {
    return {
      title: "Converter OI em PC",
      action: `Agende o PC de ${firstName}; esta é a principal pendência do funil.`,
      talk: "Conecte o próximo encontro ao que a pessoa valorizou na OI e ofereça duas opções de horário.",
      channel: "WhatsApp hoje e ligação se não houver resposta."
    };
  }
  if (client.pcScheduledDate && !client.pcDone) {
    return {
      title: "Preparar e confirmar o PC",
      action: `Confirme data e compromisso do PC com ${firstName}.`,
      talk: "Relembre brevemente o objetivo do encontro e organize as informações necessárias para a decisão.",
      channel: "WhatsApp de confirmação no dia anterior e no dia."
    };
  }
  if (client.oiScheduledDate && !client.oiDone) {
    return {
      title: "Firmar o compromisso da OI",
      action: `Confirme a OI com ${firstName} e facilite a entrada na reunião.`,
      talk: "Envie horário, duração prevista e link, reforçando que será uma conversa de entendimento.",
      channel: "WhatsApp no dia anterior e algumas horas antes."
    };
  }
  return {
    title: client.referrer || client.source === "Indicação" ? "Abordar a recomendação" : "Realizar o primeiro contato",
    action: `Faça a abordagem de ${firstName} e busque sair da conversa com uma OI marcada.`,
    talk: client.referrer
      ? `Use a confiança da indicação de ${client.referrer}, sem pressionar, e ofereça duas opções de horário.`
      : "Apresente o motivo do contato com naturalidade, faça uma pergunta curta e ofereça duas opções de horário.",
    channel: "Ligação primeiro; se não atender, WhatsApp imediatamente."
  };
}

function assistantWhatsappMessage(client, tone = "warm") {
  const firstName = client.name.trim().split(/\s+/)[0];
  const referrer = client.referrer ? ` O(a) ${client.referrer} compartilhou seu contato comigo.` : "";
  const shortReferrer = client.referrer ? ` O(a) ${client.referrer} me passou seu contato.` : "";
  let message;
  if (client.pcDone && client.outcome !== "N") {
    message = `Olá, ${firstName}! Fiquei pensando na nossa conversa e queria entender se ficou algum ponto que eu possa esclarecer para você tomar sua decisão com tranquilidade. Podemos falar rapidamente hoje?`;
  } else if (client.outcome === "C2") {
    message = `Olá, ${firstName}! Tudo bem? Queria retomar o ponto que conversamos sobre seu planejamento e entender o que mudou desde então. Posso te ligar por alguns minutos hoje ou amanhã?`;
  } else if (client.oiDone && !client.pcScheduledDate) {
    message = `Olá, ${firstName}! Gostei muito da nossa conversa. O próximo passo é organizarmos as alternativas com base no que você me contou. Você prefere fazermos isso em [opção 1] ou [opção 2]?`;
  } else if (client.pcScheduledDate && !client.pcDone) {
    message = `Olá, ${firstName}! Passando para confirmar nosso encontro de fechamento em ${formatMeetingDate(client.pcScheduledDate, client.pcTime)}.${client.pcMeetingLink ? ` Link: ${client.pcMeetingLink}.` : ""} Vou deixar tudo organizado com base no que conversamos. Combinado?`;
  } else if (client.oiScheduledDate && !client.oiDone) {
    message = `Olá, ${firstName}! Passando para confirmar nossa conversa em ${formatMeetingDate(client.oiScheduledDate, client.oiTime)}.${client.oiMeetingLink ? ` Link: ${client.oiMeetingLink}.` : ""} Será um encontro objetivo para entender suas prioridades e seu momento. Combinado?`;
  } else {
    message = `Olá, ${firstName}! Tudo bem?${referrer} Trabalho ajudando pessoas a organizarem proteção e planejamento financeiro. Gostaria de combinar uma conversa breve para conhecer seu momento. Você prefere [opção 1] ou [opção 2]?`;
  }
  if (tone === "professional") message = message.replace("Olá,", "Olá,").replace("Tudo bem?", "Espero que esteja bem.");
  if (tone === "short") {
    if (client.oiDone && !client.pcScheduledDate) return `Olá, ${firstName}! Vamos marcar nosso próximo encontro para organizar as alternativas? Você prefere [opção 1] ou [opção 2]?`;
    if (client.pcDone && client.outcome !== "N") return `Olá, ${firstName}! Ficou algum ponto do nosso encontro que eu possa esclarecer? Podemos falar rapidamente hoje?`;
    return `Oi, ${firstName}! Tudo bem?${shortReferrer} Posso te ligar rapidinho para explicar o motivo do contato?`;
  }
  return message.replace(/\s+/g, " ").trim();
}

function contactMemoryContext(client) {
  return [
    client.referrer ? `Indicado por ${client.referrer}.` : "",
    client.referrerRelationship ? `Relação com o recomendante: ${client.referrerRelationship}.` : "",
    client.referralContext ? `Contexto da indicação: ${compactText(client.referralContext).slice(0, 220)}.` : "",
    client.profession ? `Profissão: ${client.profession}.` : "",
    client.notes ? `Observação: ${compactText(client.notes).slice(0, 220)}.` : "",
    client.oiSummaryNotes ? `Resumo da OI: ${compactText(client.oiSummaryNotes).slice(0, 180)}.` : "",
    client.planNotes ? `Plano: ${compactText(client.planNotes).slice(0, 160)}.` : ""
  ].filter(Boolean);
}

function approachMemory(client) {
  const firstName = client.name.trim().split(/\s+/)[0] || client.name;
  const recommendation = assistantAction(client);
  const context = contactMemoryContext(client);
  const referrer = client.referrer ? `Veio por indicação de ${client.referrer}.` : "Sem recomendante registrado.";
  const opener = client.referrer
    ? `Comece ancorando a confiança da indicação de ${client.referrer}${client.referrerRelationship ? `, ${client.referrerRelationship}` : ""}, sem parecer venda fria.`
    : "Comece validando o melhor momento da pessoa e o motivo da conversa.";
  const reminder = context.length
    ? context.join(" ")
    : "Ainda há pouco contexto salvo. Pergunte como a pessoa conhece o indicador, momento de vida e principal prioridade.";
  return {
    title: recommendation.title,
    referrer,
    opener,
    reminder,
    talk: recommendation.talk,
    whatsapp: assistantWhatsappMessage(client, "short"),
    objective: `Objetivo com ${firstName}: sair com uma OI marcada ou com próximo horário claro.`
  };
}

function fridayFamilyProfile(client) {
  const marital = compactText(client.maritalStatus, "");
  const children = Number(client.children || 0);
  if (children > 0) return marital ? `${marital} · ${children} filho(s)` : `${children} filho(s)`;
  return marital || "Família não informada";
}

function fridayGeneratedHook(client) {
  const profession = compactText(client.profession, "");
  const marital = compactText(client.maritalStatus, "");
  const children = Number(client.children || 0);
  const lowerProfession = profession.toLocaleLowerCase("pt-BR");
  if (children > 0) {
    return "Puxe por família, renda e continuidade do padrão de vida dos filhos.";
  }
  if (["médico", "medico", "advogado", "dentista", "empresário", "empresario", "autônomo", "autonomo"].some(term => lowerProfession.includes(term))) {
    return `Puxe pelo momento profissional como ${profession} e pela proteção da renda que sustenta o planejamento.`;
  }
  if (marital && !["não informado", "nao informado"].includes(marital.toLocaleLowerCase("pt-BR"))) {
    return "Puxe por planejamento do casal/família e proteção de renda no longo prazo.";
  }
  return "Comece simples: motivo da indicação, momento de vida e permissão para explicar a ligação.";
}

function fridayOpeningTemplate(client, referrer, firstName) {
  const children = Number(client.children || 0);
  const profession = compactText(client.profession, "");
  if (referrer) {
    return `Oi, ${firstName}, tudo bem? Aqui é o Rafa. O(a) ${referrer} me passou seu contato e achei melhor te ligar rapidinho. Posso te explicar em 20 segundos o motivo?`;
  }
  if (children > 0) {
    return `Oi, ${firstName}, tudo bem? Aqui é o Rafa. Queria te ligar rapidinho sobre proteção e planejamento familiar. Posso te explicar em 20 segundos?`;
  }
  if (profession) {
    return `Oi, ${firstName}, tudo bem? Aqui é o Rafa. Queria te ligar rapidinho pelo seu momento profissional como ${profession}. Posso te explicar em 20 segundos?`;
  }
  return `Oi, ${firstName}, tudo bem? Aqui é o Rafa. Queria te ligar rapidinho para entender se faz sentido uma conversa objetiva. Posso te explicar em 20 segundos?`;
}

function fridayTaDossier(client) {
  const firstName = client.name.trim().split(/\s+/)[0] || client.name || "contato";
  const referrer = compactText(client.referrer, "");
  const profession = compactText(client.profession, "");
  const city = compactText(client.city, "");
  const notes = compactText(client.notes || client.oiQuickSummary || client.oiSummaryNotes || "", "");
  const source = compactText(client.source || client.origin || "", "");
  const facts = [
    referrer ? `Indicação de ${referrer}` : "",
    profession,
    fridayFamilyProfile(client),
    city,
    source && source !== "Indicação" ? source : "",
    client.temperature ? `Temperatura ${client.temperature}` : "",
    Number(client.children || 0) ? `${client.children} filho(s)` : ""
  ].filter(Boolean).slice(0, 4);
  const referrerLabel = referrer || "Sem recomendante";
  const referrerHint = referrer
    ? `Use ${referrer} como ponte de confiança. Não venda de primeira; peça permissão para explicar o motivo da ligação.`
    : "Comece criando contexto rápido. Descubra origem, momento de vida e se faz sentido marcar uma conversa.";
  const subjectHook = fridayGeneratedHook(client);
  const opening = fridayOpeningTemplate(client, referrer, firstName);
  const nextQuestion = "Se fizer sentido, eu queria te propor uma conversa objetiva de diagnóstico. Você prefere começo ou fim do dia?";
  const objective = "Marcar OI. A ligação não é para vender produto; é para conquistar permissão, confiança e dois horários possíveis.";
  const memory = notes
    ? compactText(notes).slice(0, 180)
    : "Sem observação salva. Na ligação, descubra relação com o recomendante, momento familiar/profissional e principal preocupação.";
  const summary = facts.length
    ? facts.join(" · ")
    : subjectHook;
  return {
    referrerLabel,
    familyLabel: fridayFamilyProfile(client),
    professionLabel: profession || "Profissão não informada",
    phoneLabel: client.whatsapp || client.phone || "Sem WhatsApp",
    referrerHint,
    subjectHook,
    opening,
    nextQuestion,
    objective,
    memory,
    summary,
    whatsapp: assistantWhatsappMessage(client, "short")
  };
}

function fridayCardMemory(client) {
  const dossier = fridayTaDossier(client);
  return client.notes
    ? compactText(client.notes).slice(0, 150)
    : dossier.summary;
}

function fridaySessionFactsHtml(dossier) {
  return `<div class="friday-session-facts">
    <span>${escapeHtml(dossier.phoneLabel)}</span>
    <span>${escapeHtml(dossier.professionLabel)}</span>
    <span>${escapeHtml(dossier.familyLabel)}</span>
  </div>`;
}

function fridayLocalScript(client) {
  const memory = approachMemory(client);
  const dossier = fridayTaDossier(client);
  return [
    "FICHA TA - ANTES DE LIGAR",
    "",
    `Quem indicou: ${dossier.referrerLabel}`,
    `Puxar assunto: ${dossier.subjectHook}`,
    `Memória salva: ${dossier.memory}`,
    "",
    "ABERTURA SUGERIDA",
    "",
    dossier.opening,
    "",
    `Pergunta de avanço: ${dossier.nextQuestion}`,
    "",
    `Objetivo: ${dossier.objective}`,
    "",
    `Regra de ouro: ${memory.talk}`,
    "",
    `WhatsApp se não atender: ${dossier.whatsapp}`
  ].join("\n");
}

function fridayTaDossierHtml(client, dossier, index, compact = false) {
  return `<details class="friday-ta-dossier" ${index === 0 && !compact ? "open" : ""}>
    <summary>
      <span>${compact ? "Ver ficha" : "Ficha TA"}</span>
      <strong>${escapeHtml(dossier.summary)}</strong>
    </summary>
    <div class="friday-ta-dossier-grid">
      <article>
        <small>Quem indicou</small>
        <strong>${escapeHtml(dossier.referrerLabel)}</strong>
        <p>${escapeHtml(dossier.referrerHint)}</p>
      </article>
      <article>
        <small>Puxar assunto</small>
        <strong>Gancho de entrada</strong>
        <p>${escapeHtml(dossier.subjectHook)}</p>
      </article>
      <article>
        <small>Memória do contato</small>
        <strong>O que lembrar</strong>
        <p>${escapeHtml(dossier.memory)}</p>
      </article>
    </div>
    <div class="friday-ta-opening">
      <small>Primeira frase da ligação</small>
      <p ${compact ? `contenteditable="true" spellcheck="true" aria-label="Editar primeira frase da ligação"` : ""}>${escapeHtml(dossier.opening)}</p>
    </div>
    <div class="friday-ta-next">
      <span>${escapeHtml(dossier.objective)}</span>
      <strong>${escapeHtml(dossier.nextQuestion)}</strong>
    </div>
    <div class="friday-ta-copy-actions">
      <button type="button" data-copy-ta-opener="${escapeHtml(client.id)}">Copiar abertura</button>
      <button type="button" data-copy-ta-whatsapp="${escapeHtml(client.id)}">Copiar WhatsApp</button>
    </div>
  </details>`;
}

function assistantNextRegister(client) {
  if (client.outcome === "N") return "Registrar pós-venda, entrega da apólice e novo pedido de recomendações.";
  if (client.pcDone && client.outcome !== "N") return "Registrar objeção real. Se não fechou, classificar como C2 com data de follow-up.";
  if (client.outcome === "C2") return "Atualizar motivo do C2, próxima data e probabilidade de retomada.";
  if (client.oiDone && !client.pcScheduledDate && !client.pcDone) return "Registrar data do PC ou motivo pelo qual o PC ainda não foi marcado.";
  if (client.pcScheduledDate && !client.pcDone) return "Confirmar presença, link e status do plano antes do PC.";
  if (client.oiScheduledDate && !client.oiDone) return "Registrar confirmação da OI ou remarcação se a pessoa não puder comparecer.";
  return "Registrar TA: atendeu, não atendeu, WhatsApp enviado ou OI marcada.";
}

function assistantConversationObjective(client) {
  if (client.outcome === "N") return "Manter relacionamento e abrir caminho para novas recomendações.";
  if (client.pcDone && client.outcome !== "N") return "Descobrir a objeção que travou a decisão e buscar um próximo sim possível.";
  if (client.outcome === "C2") return "Reabrir a decisão sem parecer cobrança, usando o motivo que ficou pendente.";
  if (client.oiDone && !client.pcScheduledDate && !client.pcDone) return "Transformar interesse da OI em PC marcado.";
  if (client.pcScheduledDate && !client.pcDone) return "Chegar no PC com compromisso confirmado e plano pronto.";
  if (client.oiScheduledDate && !client.oiDone) return "Reduzir no-show e firmar o compromisso da OI.";
  return "Conseguir resposta e sair com uma OI agendada.";
}

function assistantPlaybook(client) {
  const recommendation = assistantAction(client);
  return [
    { label: "Objetivo", value: assistantConversationObjective(client) },
    { label: "Fala-chave", value: recommendation.talk },
    { label: "Depois registre", value: assistantNextRegister(client) }
  ];
}

function compactText(value = "", fallback = "") {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  return text || fallback;
}

function assistantKnownContext(client) {
  const history = clientHistory(client.id).slice(0, 3);
  const facts = [];
  if (client.referrer) facts.push(`Indicado por ${client.referrer}.`);
  if (client.source) facts.push(`Origem: ${client.source}.`);
  if (client.profession) facts.push(`Profissão: ${client.profession}.`);
  if (client.city) facts.push(`Cidade: ${client.city}.`);
  if (client.maritalStatus) facts.push(`Estado civil: ${client.maritalStatus}.`);
  if (client.children !== "" && client.children !== undefined && client.children !== null) facts.push(`Filhos: ${client.children}.`);
  if (client.incomeRange) facts.push(`Renda: ${client.incomeRange}.`);
  if (client.notes) facts.push(`Observação: ${compactText(client.notes).slice(0, 180)}.`);
  if (client.oiSummaryNotes) facts.push(`Resumo da OI: ${compactText(client.oiSummaryNotes).slice(0, 180)}.`);
  if (client.planNotes) facts.push(`Plano: ${compactText(client.planNotes).slice(0, 180)}.`);
  if (client.c2Objection) facts.push(`Objeção C2: ${client.c2Objection}.`);
  history.forEach(activity => {
    facts.push(`${formatDate(activity.date)} - ${activity.title}${activity.notes ? `: ${compactText(activity.notes).slice(0, 120)}` : ""}.`);
  });
  return facts.slice(0, 7);
}

function assistantInfoGaps(client) {
  const gaps = [];
  if (!client.profession) gaps.push("profissão");
  if (!client.incomeRange) gaps.push("faixa de renda");
  if (!client.maritalStatus) gaps.push("estado civil/família");
  if (!client.notes && !client.oiSummaryNotes && !client.planNotes) gaps.push("motivo real da conversa");
  if (!client.followUp && client.outcome !== "N") gaps.push("próxima data");
  if ((client.source === "Indicação" || client.referrer) && !client.referrer) gaps.push("quem indicou");
  return gaps.slice(0, 5);
}

function assistantDiscoveryQuestions(client) {
  if (client.oiDone || client.pcDone || client.outcome === "C2") {
    return [
      "O que ficou mais importante para você desde nossa última conversa?",
      "Existe algum ponto que ainda precisa ficar mais claro para você decidir?",
      "O que faria essa proteção fazer sentido agora, e não depois?"
    ];
  }
  if (client.referrer || String(client.source || "").includes("Indica")) {
    return [
      `Como você conhece ${client.referrer || "a pessoa que indicou"}?`,
      "Hoje, o que você mais quer proteger ou organizar na sua vida financeira?",
      "Faz sentido marcarmos uma conversa breve para eu entender seu momento?"
    ];
  }
  return [
    "Hoje, qual é sua principal prioridade financeira ou familiar?",
    "Você já tem algo estruturado para proteger renda, família ou patrimônio?",
    "Faz sentido marcarmos uma conversa breve para eu entender seu momento?"
  ];
}

function logClientEvent(clientId, title, notes = "", type = "Historico", date = offsetDate(0)) {
  if (!clientId || !title) return;
  state.activities.push({
    id: crypto.randomUUID(),
    title,
    date,
    time: new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" }),
    type,
    clientId,
    notes,
    completed: true,
    createdAt: new Date().toISOString()
  });
}

function clientHistory(clientId) {
  return state.activities
    .filter(activity => activity.clientId === clientId)
    .sort((a, b) => `${b.date}${b.time || ""}${b.createdAt || ""}`.localeCompare(`${a.date}${a.time || ""}${a.createdAt || ""}`));
}

function renderClientHistory(clientId) {
  if (!el("clientTimeline")) return;
  if (!clientId) {
    el("clientTimeline").innerHTML = `<div class="mini-empty">Salve o contato para começar o histórico.</div>`;
    el("clientLogInteraction").disabled = true;
    return;
  }
  el("clientLogInteraction").disabled = false;
  const history = clientHistory(clientId);
  el("clientTimeline").innerHTML = history.length
    ? history.slice(0, 30).map(activity => `
      <article class="timeline-item">
        <i class="timeline-dot"></i>
        <div><strong>${escapeHtml(displaySystemText(activity.title))}</strong><small>${escapeHtml(displayActivityType(activity.type || "Historico"))}</small><p>${escapeHtml(displaySystemText(activity.notes || "Sem observações"))}</p></div>
        <time>${formatDate(activity.date)}${activity.time ? `<br>${escapeHtml(activity.time)}` : ""}</time>
      </article>`).join("")
    : `<div class="mini-empty">Nenhuma interação registrada ainda.</div>`;
}

function client360Steps(client) {
  const outcomeDone = client.outcome === "N" || client.outcome === "C2";
  return [
    { label: "TA", done: Boolean(client.taDate), active: ["Base", "TA"].includes(client.stage), detail: client.taDate ? formatDate(client.taDate) : "abordar" },
    { label: "OI", done: Boolean(client.oiDone), active: Boolean(client.oiScheduledDate && !client.oiDone), detail: client.oiDone ? formatDate(client.oiDate) : client.oiScheduledDate ? formatDate(client.oiScheduledDate) : "marcar" },
    { label: "Plano", done: client.planStatus === "Pronto", active: client.oiDone && client.planStatus !== "Pronto" && client.outcome !== "N", detail: client.planStatus || "pendente" },
    { label: "PC", done: Boolean(client.pcDone), active: Boolean(client.pcScheduledDate && !client.pcDone), detail: client.pcDone ? formatDate(client.pcDate) : client.pcScheduledDate ? formatDate(client.pcScheduledDate) : "marcar" },
    { label: "Decisão", done: outcomeDone, active: client.pcDone && !outcomeDone, detail: client.outcome || "N/C2" },
    { label: "Apólice", done: Boolean(client.policyDelivered), active: client.outcome === "N" && !client.policyDelivered, detail: client.policyStatus || (client.outcome === "N" ? "acompanhar" : "futuro") }
  ];
}

function renderClient360(client = null) {
  if (!el("client360Panel")) return;
  if (!client) {
    el("client360Panel").classList.add("hidden");
    return;
  }
  const movement = nextBestMovement(client);
  const risk = assistantRisk(client);
  const care = clientCareStatus(client, "management");
  const opportunity = sitplanOpportunityScore(client);
  const score = Math.max(Number(client.score || 0), opportunity.score, movement.score || 0);
  const memory = approachMemory(client);
  const tags = [
    client.stage || "Base",
    client.source || "",
    client.referrer ? `Indicado por ${client.referrer}` : "",
    client.followUp ? `Follow-up ${formatDate(client.followUp)}` : "Sem follow-up",
    client.lastInteraction ? `${assistantDays(client)} dias sem contato` : "Sem contato registrado"
  ].filter(Boolean);

  el("client360Action").textContent = movement.action;
  el("client360Reason").textContent = `${movement.reason}. ${client.nextAction ? `Próxima ação registrada: ${client.nextAction}.` : "Defina o próximo passo para proteger este contato."}`;
  el("client360Score").textContent = Math.round(score);
  el("client360Risk").textContent = `${care.label}: ${care.reason || risk.reason}`;
  el("client360Tags").innerHTML = tags.map(tag => `<span>${escapeHtml(tag)}</span>`).join("");
  if (el("clientApproachMemory")) {
    el("clientApproachMemory").innerHTML = `
      <article>
        <small>Memória de abordagem</small>
        <strong>${escapeHtml(memory.referrer)}</strong>
        <p>${escapeHtml(memory.reminder)}</p>
      </article>
      <article>
        <small>O que falar agora</small>
        <strong>${escapeHtml(memory.title)}</strong>
        <p>${escapeHtml(memory.opener)}</p>
        <div class="talk-script">${escapeHtml(memory.talk)}</div>
      </article>
    `;
  }
  el("client360Timeline").innerHTML = client360Steps(client).map(step => `
    <article class="client-360-step ${step.done ? "done" : ""} ${step.active ? "active" : ""}">
      <strong>${escapeHtml(step.label)}</strong>
      <small>${escapeHtml(step.detail || "")}</small>
    </article>
  `).join("");
  el("client360Panel").classList.remove("hidden");
}

function clientCareStatus(client, tab) {
  const today = offsetDate(0);
  const tomorrow = offsetDate(1);
  if (!client) return { tone: "neutral", label: "Novo", reason: "preencha nome e WhatsApp" };
  if (tab === "profile") {
    if (!client.name || !(client.whatsapp || client.phone)) return { tone: "warn", label: "Atenção", reason: "dados mínimos incompletos" };
    const usefulMissing = [client.profession, client.source, client.referrer, client.incomeRange].filter(Boolean).length < 2;
    return usefulMissing
      ? { tone: "warn", label: "Atenção", reason: "cadastro útil pode melhorar" }
      : { tone: "good", label: "Em dia", reason: "dados principais salvos" };
  }
  if (tab === "management") {
    const days = assistantDays(client);
    if (client.followUp && client.followUp < today) return { tone: "danger", label: "Atrasado", reason: `follow-up venceu em ${formatDate(client.followUp)}` };
    if (client.outcome === "C2" && (!client.followUp || days >= 7)) return { tone: "danger", label: "Atrasado", reason: "C2 esfriando" };
    if (client.pcDone && !client.outcome) return { tone: "danger", label: "Atrasado", reason: "PC sem decisão registrada" };
    if (client.oiDone && !client.pcScheduledDate && !client.pcDone && client.outcome !== "N") return { tone: "warn", label: "Atenção", reason: "OI sem PC marcado" };
    if (client.oiDone && client.planStatus !== "Pronto" && client.outcome !== "N") return { tone: "warn", label: "Atenção", reason: "plano pendente" };
    if (client.followUp && client.followUp <= tomorrow) return { tone: "warn", label: "Atenção", reason: "follow-up próximo" };
    if (!client.nextAction || !client.followUp) return { tone: "warn", label: "Atenção", reason: "sem próximo passo claro" };
    return { tone: "good", label: "Em dia", reason: "próximo passo protegido" };
  }
  if (tab === "aftercare") {
    if (client.outcome !== "N" && !client.cpfDone) return { tone: "neutral", label: "Futuro", reason: "pós-venda ainda não começou" };
    const saleDays = daysSince(policySaleDate(client)) ?? 0;
    if (client.policyDelivered || client.policyStatus === "Entregue") return { tone: "good", label: "Em dia", reason: "apólice entregue" };
    if (!["Emitida", "Entregue"].includes(client.policyStatus) && saleDays > 10) return { tone: "danger", label: "Atrasado", reason: "emissão demorando" };
    if (client.nextRelationship && client.nextRelationship < today) return { tone: "danger", label: "Atrasado", reason: "pós-venda vencido" };
    return { tone: "warn", label: "Atenção", reason: "emissão ou entrega pendente" };
  }
  return { tone: "neutral", label: "Sem leitura", reason: "" };
}

function renderClientCareStatus(client = null) {
  document.querySelectorAll("[data-form-tab]").forEach(button => {
    const status = clientCareStatus(client, button.dataset.formTab);
    button.classList.remove("status-good", "status-warn", "status-danger", "status-neutral");
    button.classList.add(`status-${status.tone}`);
    button.title = `${status.label}: ${status.reason}`;
    const dot = button.querySelector(".form-tab-status");
    if (dot) dot.setAttribute("aria-label", `${status.label}: ${status.reason}`);
  });
}

function conversionRate(numerator, denominator, fallback = 0) {
  return denominator ? Math.min(numerator / denominator, 1) : fallback;
}

function renderPerformanceIntelligence() {
  const clients = state.clients;
  const totalTa = clients.filter(client => client.taDate).length;
  const totalOiScheduled = clients.filter(client => client.oiScheduledDate).length;
  const totalOiDone = clients.filter(client => client.oiDone || client.oiDate).length;
  const totalPcScheduled = clients.filter(client => client.pcScheduledDate).length;
  const totalPcDone = clients.filter(client => client.pcDone || client.pcDate).length;
  const totalSales = clients.filter(client => client.outcome === "N").length;
  const salesFromPc = clients.filter(client => (client.pcDone || client.pcDate) && client.outcome === "N").length;
  const salesFromTa = clients.filter(client => client.taDate && client.outcome === "N").length;

  const taOi = conversionRate(totalOiScheduled, totalTa, .3);
  const oiShow = conversionRate(totalOiDone, totalOiScheduled, .75);
  const oiPc = conversionRate(totalPcScheduled, totalOiDone, .65);
  const pcN = conversionRate(salesFromPc, totalPcDone, .3);
  const taN = conversionRate(salesFromTa, totalTa, 0);
  const rates = [
    ["conversionTaOi", "conversionTaOiDetail", taOi, totalTa, `${totalOiScheduled} de ${totalTa}`],
    ["conversionOiShow", "conversionOiShowDetail", oiShow, totalOiScheduled, `${totalOiDone} de ${totalOiScheduled}`],
    ["conversionOiPc", "conversionOiPcDetail", oiPc, totalOiDone, `${totalPcScheduled} de ${totalOiDone}`],
    ["conversionPcN", "conversionPcNDetail", pcN, totalPcDone, `${salesFromPc} de ${totalPcDone}`],
    ["conversionTaN", "conversionTaNDetail", taN, totalTa, `${salesFromTa} N(s) em ${totalTa} TA(s)`]
  ];
  rates.forEach(([valueId, detailId, rate, denominator, detail]) => {
    el(valueId).textContent = `${Math.round(rate * 100)}%`;
    el(detailId).textContent = denominator ? detail : "referência inicial";
  });

  const week = selectedWeek();
  const goal = weeklyGoal(week);
  const scheduledThisWeek = clients.filter(client => inSelectedWeek(client.oiScheduledDate, week)).length;
  const remainingOi = Math.max(goal - scheduledThisWeek, 0);
  const requiredTa = remainingOi ? Math.ceil(remainingOi / Math.max(taOi, .1)) : 0;
  const expectedSales = scheduledThisWeek * oiShow * oiPc * pcN;
  const evidence = totalTa + totalOiScheduled + totalOiDone + totalPcDone;
  el("forecastConfidence").textContent = evidence >= 40 ? "Boa base histórica" : evidence >= 15 ? "Base em aprendizado" : "Base inicial";
  el("forecastRequiredTa").textContent = remainingOi
    ? `${requiredTa} ${requiredTa === 1 ? "nova abordagem" : "novas abordagens"}`
    : "Meta semanal alcançada";
  el("forecastExplanation").textContent = remainingOi
    ? `Faltam ${remainingOi} OI(s). A projeção usa sua conversão TA → OI de ${Math.round(taOi * 100)}%.`
    : "Proteja o comparecimento das OIs e prepare a passagem para PC.";
  el("forecastSales").textContent = expectedSales < 1 && expectedSales > 0 ? expectedSales.toFixed(1).replace(".", ",") : Math.round(expectedSales);
  el("forecastSalesDetail").textContent = `com ${scheduledThisWeek} OI(s) agendada(s) nesta semana`;
}

function normalizedPersonName(value = "") {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("pt-BR")
    .replace(/\s+/g, " ")
    .trim();
}

function priorityClass(value = "") {
  const key = normalizedPersonName(value);
  if (key === "alta") return "prio-alta";
  if (key === "baixa") return "prio-baixa";
  return "prio-media";
}

function temperatureClass(value = "") {
  const key = normalizedPersonName(value);
  if (key === "quente") return "temp-quente";
  if (key === "frio") return "temp-frio";
  return "temp-morno";
}

function referralRankingData() {
  const groups = new Map();
  const clientsByName = new Map(state.clients.map(client => [normalizedPersonName(client.name), client]));
  state.clients.filter(client => client.referrer?.trim()).forEach(client => {
    const key = normalizedPersonName(client.referrer);
    if (!key) return;
    const existing = groups.get(key) || {
      key,
      name: client.referrer.trim(),
      leads: 0,
      oi: 0,
      pc: 0,
      n: 0,
      c2: 0,
      declared: Number(clientsByName.get(key)?.recs || 0)
    };
    existing.leads += 1;
    if (client.oiDone || client.oiScheduledDate) existing.oi += 1;
    if (client.pcDone || client.pcScheduledDate) existing.pc += 1;
    if (client.outcome === "N") existing.n += 1;
    if (client.outcome === "C2") existing.c2 += 1;
    groups.set(key, existing);
  });

  state.clients.filter(client => Number(client.recs || 0) > 0).forEach(client => {
    const key = normalizedPersonName(client.name);
    if (groups.has(key)) return;
    groups.set(key, {
      key,
      name: client.name,
      leads: 0,
      oi: 0,
      pc: 0,
      n: 0,
      c2: 0,
      declared: Number(client.recs || 0)
    });
  });

  return [...groups.values()].sort((a, b) =>
    b.n - a.n
    || b.leads - a.leads
    || (b.leads ? b.n / b.leads : 0) - (a.leads ? a.n / a.leads : 0)
    || b.declared - a.declared
    || a.name.localeCompare(b.name, "pt-BR")
  );
}

function referralConversion(item) {
  return item.leads ? item.n / item.leads : 0;
}

function referralTagContext(ranking) {
  return {
    maxLeads: Math.max(0, ...ranking.map(item => item.leads)),
    maxDeclared: Math.max(0, ...ranking.map(item => item.declared))
  };
}

function referralTags(item, context = referralTagContext([item])) {
  const tags = [];
  const conversion = referralConversion(item);
  const lastTouchDays = daysSince(referrerLastTouch(item.name));

  if (item.leads > 0 && item.leads === context.maxLeads) {
    tags.push({ label: "Top REC", type: "top", text: "Mais indicados na base." });
  }
  if (item.leads >= 2 && conversion >= 0.5) {
    tags.push({ label: "Alta conversão", type: "high", text: "Indica pouco ou muito, mas fecha bem." });
  }
  if (item.leads >= 3 && conversion < 0.25) {
    tags.push({ label: "Vazamento", type: "leak", text: "Gera volume, mas pouca venda. Ajuste abordagem." });
  }
  if (item.leads > 0 && (lastTouchDays ?? 999) >= 14) {
    tags.push({ label: "Adormecido", type: "sleep", text: "Ja indicou e precisa de reativacao." });
  }
  if (!tags.length && item.n > 0) {
    tags.push({ label: "Gerou N", type: "win", text: "Ja provou valor comercial." });
  }
  if (!tags.length && item.declared > item.leads) {
    tags.push({ label: "REC pendente", type: "pending", text: "Tem REC declarada ainda não vinculada." });
  }
  if (!tags.length) {
    tags.push({ label: "Mapear", type: "neutral", text: "Ainda precisa de histórico." });
  }
  return tags;
}

function primaryReferralTag(item, context) {
  return referralTags(item, context)[0];
}

function renderReferralIntelligence() {
  if (!el("referralRanking")) return;
  const ranking = referralRankingData();
  const tagContext = referralTagContext(ranking);
  const linked = ranking.reduce((sum, item) => sum + item.leads, 0);
  const sales = ranking.reduce((sum, item) => sum + item.n, 0);
  const rate = linked ? sales / linked : 0;
  const top = ranking[0];
  el("referredContactsTotal").textContent = linked;
  el("referredContactsNs").textContent = `${sales} ${sales === 1 ? "virou N" : "viraram N"}`;
  el("referralToNRate").textContent = `${Math.round(rate * 100)}%`;
  el("referralToNDetail").textContent = linked ? `${sales} N(s) em ${linked} indicado(s)` : "Cadastre quem indicou cada contato";
  el("referralConversionSummary").textContent = linked ? `${Math.round(rate * 100)}% indicado → N` : "Aguardando dados";
  el("topReferrerName").textContent = top?.name || "Sem dados";
  el("topReferrerDetail").textContent = top
    ? top.leads
      ? `${top.n} N(s) em ${top.leads} indicado(s) · ${Math.round((top.n / top.leads) * 100)}%`
      : `${top.declared} REC(s) declarada(s), ainda sem indicados vinculados`
    : "Cadastre o indicador nos contatos recomendados.";

  const topRec = ranking.filter(item => item.leads > 0).sort((a, b) => b.leads - a.leads || b.n - a.n)[0];
  const highConversion = ranking.filter(item => item.leads >= 2 && referralConversion(item) >= 0.5).sort((a, b) => referralConversion(b) - referralConversion(a) || b.n - a.n)[0];
  const leakage = ranking.filter(item => item.leads >= 3 && referralConversion(item) < 0.25).sort((a, b) => b.leads - a.leads)[0];
  const dormant = ranking.filter(item => item.leads > 0 && (daysSince(referrerLastTouch(item.name)) ?? 999) >= 14).sort((a, b) => b.n - a.n || b.leads - a.leads)[0];
  const tagCards = [
    { label: "Top REC", item: topRec, text: topRec ? `${topRec.leads} indicados. Pedir novas RECs com contexto.` : "Ainda sem volume mapeado.", type: "top" },
    { label: "Alta conversão", item: highConversion, text: highConversion ? `${Math.round(referralConversion(highConversion) * 100)}% de conversão. Cuidar como ativo VIP.` : "Buscando quem converte acima de 50%.", type: "high" },
    { label: "Vazamento", item: leakage, text: leakage ? `${leakage.leads} indicados e ${leakage.n} N. Revisar abordagem.` : "Sem vazamento relevante agora.", type: "leak" },
    { label: "Adormecido", item: dormant, text: dormant ? `${daysSince(referrerLastTouch(dormant.name)) ?? 0} dias sem contato. Reativar.` : "Nenhum hub critico dormindo.", type: "sleep" }
  ];
  if (el("referralTagStrip")) {
    el("referralTagStrip").innerHTML = tagCards.map(card => `
      <button class="referral-tag-card ${card.type}" ${card.item ? `data-referrer-filter="${escapeHtml(card.item.name)}"` : "disabled"}>
        <small>${card.label}</small>
        <strong>${escapeHtml(card.item?.name || "Sem nome")}</strong>
        <span>${escapeHtml(card.text)}</span>
      </button>
    `).join("");
  }

  el("referralRanking").innerHTML = ranking.length ? ranking.slice(0, 8).map((item, index) => {
    const conversion = item.leads ? Math.round((item.n / item.leads) * 100) : 0;
    const volume = Math.max(item.leads, item.declared);
    const tag = primaryReferralTag(item, tagContext);
    return `<article class="referral-row">
      <span class="referral-rank">${index + 1}</span>
      <div class="referral-person">
        <strong>${escapeHtml(item.name)}</strong>
        <small>${item.leads} indicado(s) cadastrado(s)${item.declared > item.leads ? ` · ${item.declared} REC(s) declarada(s)` : ""}</small>
      </div>
      <div><small>OIs</small><strong>${item.oi}</strong></div>
      <div><small>PCs</small><strong>${item.pc}</strong></div>
      <div class="referral-n"><small>Ns</small><strong>${item.n}</strong></div>
      <div><small>Conversão</small><strong>${item.leads ? `${conversion}%` : "—"}</strong></div>
      <div><small>Selo</small><strong class="referral-tag ${tag.type}" title="${escapeHtml(tag.text)}">${escapeHtml(tag.label)}</strong></div>
      <div class="referral-volume" title="${volume} recomendacoes registradas"><i style="width:${Math.min(volume * 12, 100)}%"></i></div>
      <button class="text-button" data-referrer-filter="${escapeHtml(item.name)}">Ver indicados</button>
    </article>`;
  }).join("") : `<div class="mini-empty">Cadastre o campo Indicador nos contatos recomendados para formar o ranking.</div>`;
  document.querySelectorAll("[data-referrer-filter]").forEach(button => button.addEventListener("click", () => {
    workFilter = "";
    switchView("clients");
    renderClients();
    el("referrerFilter").value = button.dataset.referrerFilter;
    renderClients();
  }));
}

function annualGoalData() {
  const target = Math.max(1, Number(state.settings.annualGoalTarget || 134));
  const deadlineValue = state.settings.annualGoalDeadline || "2026-10-31";
  const deadline = parseLocalDate(deadlineValue) || new Date(2026, 9, 31, 12);
  const today = new Date();
  today.setHours(12, 0, 0, 0);
  const yearStart = new Date(deadline.getFullYear(), 0, 1, 12);
  const done = state.clients.filter(client => client.cpfDone || client.outcome === "N").length;
  const remaining = Math.max(target - done, 0);
  const weeksLeft = Math.max(Math.ceil((deadline - today) / (7 * 86400000)), 0);
  const elapsedWeeks = Math.max((today - yearStart) / (7 * 86400000), 1);
  const totalWeeks = Math.max((deadline - yearStart) / (7 * 86400000), 1);
  const projection = Math.round((done / elapsedWeeks) * totalWeeks);
  const perWeek = weeksLeft ? remaining / weeksLeft : remaining;
  return { target, deadlineValue, done, remaining, weeksLeft, projection, perWeek };
}

function renderAnnualGoal() {
  if (!el("annualGoalDone")) return;
  const goal = annualGoalData();
  const pct = Math.min((goal.done / goal.target) * 100, 100);
  el("annualGoalDone").textContent = goal.done;
  el("annualGoalTarget").value = goal.target;
  el("annualGoalDeadline").value = goal.deadlineValue;
  el("annualGoalBar").style.width = `${pct}%`;
  el("annualGoalRemaining").textContent = goal.remaining;
  el("annualGoalWeeks").textContent = goal.weeksLeft;
  el("annualGoalPerWeek").textContent = goal.perWeek < 10 ? goal.perWeek.toFixed(1).replace(".", ",") : Math.ceil(goal.perWeek);
  el("annualGoalProjection").textContent = goal.projection;
  el("annualGoalMessage").textContent = goal.remaining === 0
    ? "Meta alcançada. Agora proteja a qualidade e o pós-venda."
    : goal.projection >= goal.target
      ? `Seu ritmo atual projeta ${goal.projection}. Mantenha a consistência.`
      : `No ritmo atual, a projeção é ${goal.projection}. O plano semanal precisa acelerar.`;
}

function monthlyStats(monthValue) {
  const prefix = monthValue || offsetDate(0).slice(0, 7);
  const inMonth = value => Boolean(value && value.startsWith(prefix));
  return {
    ta: state.clients.filter(client => inMonth(client.taDate)).length,
    oi: state.clients.filter(client => client.oiDone && inMonth(client.oiDate)).length,
    pc: state.clients.filter(client => client.pcDone && inMonth(client.pcDate)).length,
    n: state.clients.filter(client => client.outcome === "N" && inMonth(client.outcomeDate || client.cpfDate)).length,
    c2: state.clients.filter(client => client.outcome === "C2" && inMonth(client.outcomeDate)).length,
    recs: state.clients.filter(client => inMonth(client.recsDate)).reduce((sum, client) => sum + Number(client.recs || 0), 0)
  };
}

function diagnosisCopy(stats) {
  if (!stats.ta && !stats.oi && !stats.pc && !stats.n) {
    return {
      headline: "Ainda não há produção suficiente neste mês.",
      text: "Registre suas abordagens, reuniões e resultados para o SITPLAN localizar o gargalo real.",
      action: "Ação: comece pelo Modo Sexta e registre cada resultado de ligação."
    };
  }
  const taOi = stats.ta ? stats.oi / stats.ta : 0;
  const oiPc = stats.oi ? stats.pc / stats.oi : 0;
  const pcN = stats.pc ? stats.n / stats.pc : 0;
  if (stats.ta < 20) {
    return {
      headline: "O principal gargalo está no volume de abordagens.",
      text: `Foram ${stats.ta} TAs no mês. Sem alimentar o início do funil, OIs, PCs e Ns ficam limitados.`,
      action: "Ação: aumente a fila de recomendações e defina blocos fixos de ligação."
    };
  }
  if (taOi < .3) {
    return {
      headline: "Há atividade, mas poucas abordagens viram OI.",
      text: `Sua passagem aproximada de TA para OI realizada está em ${Math.round(taOi * 100)}%. Trabalhe abertura, confiança da indicação e oferta de dois horários.`,
      action: "Ação: revise o roteiro das próximas 10 ligações e compare os resultados."
    };
  }
  if (oiPc < .6) {
    return {
      headline: "A maior oportunidade está na passagem da OI para o PC.",
      text: `${stats.oi} OIs geraram ${stats.pc} PCs realizados. O próximo encontro precisa ser combinado ainda durante a OI.`,
      action: "Ação: não encerre uma OI positiva sem propor duas datas para o PC."
    };
  }
  if (pcN < .3) {
    return {
      headline: "O funil chega ao PC, mas o fechamento pede atenção.",
      text: `${stats.pc} PCs resultaram em ${stats.n} Ns e ${stats.c2} C2s. Registre a objeção principal e dê prazo ao follow-up.`,
      action: "Ação: ataque primeiro os C2s mais recentes com uma pergunta objetiva."
    };
  }
  if (stats.recs < stats.oi) {
    return {
      headline: "Seu funil está saudável; agora amplie as recomendações.",
      text: `Você realizou ${stats.oi} OIs e recebeu ${stats.recs} RECs. Cada boa conversa pode alimentar as próximas sextas.`,
      action: "Ação: inclua um pedido natural de recomendações nas conversas bem-sucedidas."
    };
  }
  return {
    headline: "O mês mostra um funil equilibrado.",
    text: `Você registrou ${stats.ta} TAs, ${stats.oi} OIs, ${stats.pc} PCs e ${stats.n} Ns. Preserve o ritmo e cuide dos ${stats.c2} C2s.`,
    action: "Ação: repita o que funcionou e concentre o próximo bloco nos follow-ups de maior potencial."
  };
}

function renderMonthlyDiagnosis() {
  const stats = monthlyStats(el("diagnosisMonth").value);
  const copy = diagnosisCopy(stats);
  el("diagnosisTa").textContent = stats.ta;
  el("diagnosisOi").textContent = stats.oi;
  el("diagnosisPc").textContent = stats.pc;
  el("diagnosisN").textContent = stats.n;
  el("diagnosisC2").textContent = stats.c2;
  el("diagnosisRecs").textContent = stats.recs;
  el("diagnosisHeadline").textContent = copy.headline;
  el("diagnosisText").textContent = copy.text;
  el("diagnosisAction").textContent = copy.action;
}

function openMonthlyDiagnosis() {
  el("diagnosisMonth").value ||= offsetDate(0).slice(0, 7);
  renderMonthlyDiagnosis();
  el("monthlyDiagnosisModal").classList.remove("hidden");
}

function closeMonthlyDiagnosis() {
  el("monthlyDiagnosisModal").classList.add("hidden");
}

function dailyCommandClients() {
  const eligible = [...state.clients]
    .filter(client => !client.doNotContact && (client.outcome !== "N" || client.nextRelationship))
  if (new Date().getDay() === 5) {
    return eligible.sort((a, b) => {
      const fridayScore = client => {
        if (!["Base", "TA"].includes(client.stage) || client.oiScheduledDate) return 1000 - assistantPriorityScore(client);
        const recommendation = client.source === "Indicação" || Boolean(client.referrer) || Number(client.recs || 0) > 0;
        if (recommendation && !client.lastInteraction) return 0;
        if (recommendation) return 1;
        if (!client.lastInteraction && !client.taDate) return 2;
        return 3;
      };
      return fridayScore(a) - fridayScore(b);
    }).slice(0, 3);
  }
  return eligible.sort((a, b) => assistantPriorityScore(b) - assistantPriorityScore(a)).slice(0, 3);
}

function isRecommendation(client) {
  return client.source === "Indicação" || Boolean(client.referrer) || Number(client.recs || 0) > 0;
}

function taPortfolioTarget() {
  state.settings ||= {};
  const target = Number(state.settings.taPortfolioWeeklyTarget || TA_PORTFOLIO_DEFAULT_TARGET);
  return Math.max(10, Math.min(500, target || TA_PORTFOLIO_DEFAULT_TARGET));
}

function taPortfolioWeekKey(offset = 0) {
  return dateKey(startOfWeek(offset));
}

function taPortfolioWeekLabel(weekKey) {
  const date = parseLocalDate(weekKey);
  return date ? `Semana de ${shortDate.format(date)}` : "Semana sem data";
}

function activeTaPortfolioWeeks() {
  return new Set([taPortfolioWeekKey(0), taPortfolioWeekKey(1)]);
}

function isTaPortfolioEligible(client) {
  return ["Base", "TA"].includes(client.stage)
    && !client.oiScheduledDate
    && client.outcome !== "N"
    && !client.doNotContact;
}

function isTaPortfolioCarryover(client) {
  return Boolean(client.taPortfolioWeek && !activeTaPortfolioWeeks().has(client.taPortfolioWeek));
}

function taPortfolioScore(client) {
  let score = fridayTitanScore(client);
  if (isRecommendation(client) && !client.lastInteraction) score += 18;
  if (isTaPortfolioCarryover(client)) score += 16;
  if (client.temperature === "Quente") score += 10;
  if (client.priority === "Alta") score += 8;
  if (client.referrer) score += 6;
  if (!client.followUp) score += 4;
  return score;
}

function cleanTaPortfolioAssignments() {
  let changed = false;
  state.clients.forEach(client => {
    if (!client.taPortfolioWeek) return;
    if (!isTaPortfolioEligible(client)) {
      client.taPortfolioWeek = "";
      client.taPortfolioPinnedAt = "";
      changed = true;
    }
  });
  return changed;
}

function assignTaPortfolio(weekKey, target = taPortfolioTarget()) {
  cleanTaPortfolioAssignments();
  const alreadyAssigned = state.clients
    .filter(client => isTaPortfolioEligible(client) && client.taPortfolioWeek === weekKey)
    .sort((a, b) => taPortfolioScore(b) - taPortfolioScore(a));
  const missing = Math.max(target - alreadyAssigned.length, 0);
  if (!missing) return 0;
  const candidates = state.clients
    .filter(client => isTaPortfolioEligible(client) && (!client.taPortfolioWeek || isTaPortfolioCarryover(client)))
    .sort((a, b) => {
      const carryoverPriority = Number(isTaPortfolioCarryover(b)) - Number(isTaPortfolioCarryover(a));
      if (carryoverPriority) return carryoverPriority;
      const score = taPortfolioScore(b) - taPortfolioScore(a);
      return score || String(a.name || "").localeCompare(String(b.name || ""), "pt-BR");
    })
    .slice(0, missing);
  candidates.forEach(client => {
    client.taPortfolioWeek = weekKey;
    client.taPortfolioPinnedAt = new Date().toISOString();
  });
  return candidates.length;
}

function taWorkedInWeek(client, weekKey) {
  const weekStart = parseLocalDate(weekKey);
  if (!weekStart) return false;
  const week = {
    start: weekStart,
    end: new Date(weekStart),
    startKey: weekKey,
    endKey: ""
  };
  week.end.setDate(week.start.getDate() + 6);
  week.endKey = dateKey(week.end);
  return inSelectedWeek(client.taDate, week)
    || inSelectedWeek(client.lastInteraction, week)
    || clientHistory(client.id).some(activity => inSelectedWeek(activity.date, week) && /Sexta|TA|contato|lig/i.test(activity.title || ""));
}

function taPortfolioData() {
  cleanTaPortfolioAssignments();
  const currentKey = taPortfolioWeekKey(0);
  const nextKey = taPortfolioWeekKey(1);
  const eligible = state.clients.filter(isTaPortfolioEligible);
  const current = eligible.filter(client => client.taPortfolioWeek === currentKey);
  const next = eligible.filter(client => client.taPortfolioWeek === nextKey);
  const carryover = eligible.filter(isTaPortfolioCarryover);
  const bank = eligible.filter(client => !client.taPortfolioWeek);
  const currentWorked = current.filter(client => taWorkedInWeek(client, currentKey));
  const currentUntouched = current.filter(client => !client.lastInteraction && !client.taDate);
  return {
    target: taPortfolioTarget(),
    currentKey,
    nextKey,
    eligible,
    current,
    next,
    carryover,
    bank,
    currentWorked,
    currentUntouched
  };
}

function taPortfolioLabelForClient(client) {
  if (!client.taPortfolioWeek) return "Banco de nomes";
  if (client.taPortfolioWeek === taPortfolioWeekKey(0)) return "Carteira desta semana";
  if (client.taPortfolioWeek === taPortfolioWeekKey(1)) return "Carteira da próxima semana";
  return `Sobra de ${taPortfolioWeekLabel(client.taPortfolioWeek)}`;
}

function fridayReferrerName(client) {
  return compactText(client.referrer, "Sem recomendante");
}

function fridaySearchText(value) {
  return String(value || "")
    .toLocaleLowerCase("pt-BR")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function fridayTemperatureValue(client) {
  const value = compactText(client.temperature, "Morno");
  if (/quente/i.test(value)) return "Quente";
  if (/frio/i.test(value)) return "Frio";
  return "Morno";
}

function fridayTemperatureRank(client) {
  const order = { Quente: 0, Morno: 1, Frio: 2 };
  return order[fridayTemperatureValue(client)] ?? 1;
}

function fridayReceivedKey(client) {
  return client.recsDate
    || client.taPortfolioPinnedAt?.slice(0, 10)
    || client.createdAt?.slice(0, 10)
    || client.taDate
    || "9999-12-31";
}

function fridayNextCallDue(client) {
  return Boolean(client.followUp && client.followUp <= offsetDate(0) && !client.oiScheduledDate && !client.doNotContact);
}

function fridayQueueSearchMatches(client) {
  const query = fridaySearchText(fridayQueueSearch).trim();
  if (!query) return true;
  return [
    client.name,
    client.referrer,
    client.phone,
    client.whatsapp,
    client.notes,
    client.nextAction,
    fridayGeneratedHook(client)
  ].some(value => fridaySearchText(value).includes(query));
}

function clearFridayQueueFineFilters() {
  fridayQueueSearch = "";
  fridayTemperatureFilter = "all";
}

function fridayApplyQueueControls(queue, options = {}) {
  const ignoreControls = Boolean(options.ignoreControls);
  const searched = ignoreControls ? queue : queue.filter(fridayQueueSearchMatches);
  const counts = { all: searched.length, Quente: 0, Morno: 0, Frio: 0 };
  searched.forEach(client => {
    counts[fridayTemperatureValue(client)] += 1;
  });
  const filtered = ignoreControls || fridayTemperatureFilter === "all"
    ? searched
    : searched.filter(client => fridayTemperatureValue(client) === fridayTemperatureFilter);
  const pinned = fridayPinnedClientId ? queue.find(client => client.id === fridayPinnedClientId) : null;
  const stableFiltered = pinned && !filtered.some(client => client.id === pinned.id)
    ? [pinned, ...filtered]
    : filtered;
  return { searched, filtered: stableFiltered, counts };
}

function fridayReferrerSequenceGroups(clients) {
  const groups = new Map();
  clients.forEach(client => {
    const name = fridayReferrerName(client);
    if (!groups.has(name)) {
      groups.set(name, {
        name,
        clients: [],
        hot: 0,
        uncalled: 0,
        worked: 0,
        score: 0
      });
    }
    const group = groups.get(name);
    group.clients.push(client);
    if (isRecommendation(client) && !client.lastInteraction) group.hot += 1;
    if (!client.lastInteraction && !client.taDate) group.uncalled += 1;
    if (client.lastInteraction || client.taDate) group.worked += 1;
    group.score += fridayTitanScore(client);
  });
  return [...groups.values()]
    .sort((a, b) => {
      const named = (a.name === "Sem recomendante" ? 1 : 0) - (b.name === "Sem recomendante" ? 1 : 0);
      return named || b.clients.length - a.clients.length || b.hot - a.hot || b.score - a.score;
    });
}

function fridayReferrerActiveGroup(groups) {
  if (!groups.length) return null;
  if (fridayReferrerFilter !== "all") {
    return groups.find(group => group.name === fridayReferrerFilter) || null;
  }
  return groups.find(group => group.name !== "Sem recomendante" && group.clients.length >= 2)
    || groups.find(group => group.name !== "Sem recomendante")
    || groups[0];
}

function fridayReferrerBlockDetail(group, remaining = 0) {
  if (!group) return "Sem bloco ativo.";
  const suggestedOi = Math.max(1, Math.min(3, remaining || Math.ceil(group.clients.length / 3) || 1));
  if (group.name === "Sem recomendante") {
    return `${group.clients.length} contato(s) sem recomendante. Use abordagem curta, descubra contexto e busque ${suggestedOi} OI(s).`;
  }
  return `${group.clients.length} indicado(s), ${group.hot} REC nova(s) e ${group.uncalled} sem tentativa. Use o mesmo gancho e busque ${suggestedOi} OI(s).`;
}

function fridayReferrerBlockObject(group, remaining = 0) {
  if (!group) return null;
  return {
    key: "referrer",
    number: "1",
    title: group.name === "Sem recomendante" ? "Bloco sem recomendante" : `Bloco ${group.name}`,
    count: group.clients.length,
    detail: fridayReferrerBlockDetail(group, remaining),
    action: "Executar bloco",
    referrer: group.name,
    group
  };
}

function fridayReferrerPlaybookHtml(group, selected = false) {
  if (!group) {
    return `<div class="mini-empty">Cadastre o campo Indicador nos contatos para ativar um discurso por recomendante.</div>`;
  }
  const firstNames = group.clients.slice(0, 5).map(client => client.name.trim().split(/\s+/)[0]).filter(Boolean);
  const names = firstNames.length ? firstNames.join(", ") : "nomes da fila";
  const opener = group.name === "Sem recomendante"
    ? "Oi, tudo bem? Aqui é o Rafa Villela. Trabalho com proteção e planejamento financeiro e queria te explicar em 20 segundos o motivo da minha ligação."
    : `Oi, tudo bem? Aqui é o Rafa Villela. O(a) ${group.name} me passou seu contato e achei melhor te ligar rapidinho. Posso te explicar em 20 segundos o motivo?`;
  const hook = group.name === "Sem recomendante"
    ? "Descubra origem, momento familiar/profissional e se existe abertura para uma conversa objetiva."
    : `Comece perguntando como a pessoa conhece ${group.name}. O objetivo é transferir confiança antes de falar de produto.`;
  return `<article>
    <div>
      <small>${selected ? "DISCURSO ATIVO" : "SUGESTÃO DE BLOCO"}</small>
      <strong>${escapeHtml(group.name === "Sem recomendante" ? "Contatos sem recomendante" : group.name)}</strong>
      <span>${group.clients.length} contato(s) · ${escapeHtml(names)}</span>
    </div>
    <div>
      <small>Primeira frase</small>
      <p>${escapeHtml(opener)}</p>
    </div>
    <div>
      <small>Gancho do bloco</small>
      <p>${escapeHtml(hook)}</p>
    </div>
    <div>
      <small>Fechamento</small>
      <p>Se fizer sentido, eu queria te propor uma conversa objetiva. Você prefere começo ou fim do dia?</p>
    </div>
  </article>`;
}

function fridayReferrerSequenceScript(group) {
  const firstNames = group.clients.slice(0, 4).map(client => client.name.trim().split(/\s+/)[0]).filter(Boolean).join(", ");
  if (group.name === "Sem recomendante") {
    return [
      "SEQUÊNCIA TA - SEM RECOMENDANTE",
      "",
      `Bloco com ${group.clients.length} contato(s): ${firstNames || "nomes da fila"}.`,
      "Abertura: Oi, tudo bem? Aqui é o Rafa Villela. Trabalho com proteção e planejamento financeiro e te liguei para entender se faz sentido uma conversa rápida. Posso te explicar em 20 segundos?",
      "Gancho: descubra o momento de vida/profissional antes de falar de produto.",
      "Objetivo: conseguir permissão e oferecer dois horários para a OI.",
      "Regra: se não atender, WhatsApp curto e próxima tentativa registrada."
    ].join("\n");
  }
  return [
    `SEQUÊNCIA TA - INDICAÇÕES DE ${group.name}`,
    "",
    `Bloco com ${group.clients.length} indicado(s): ${firstNames || "nomes da fila"}.`,
    `Abertura: Oi, tudo bem? Aqui é o Rafa Villela. O(a) ${group.name} me passou seu contato com bastante cuidado e falou que faria sentido eu te conhecer. Posso te explicar em 20 segundos o motivo da minha ligação?`,
    `Gancho: comece perguntando como a pessoa conhece ${group.name}. Use a confiança da indicação antes de falar do produto.`,
    "Objetivo: marcar OI, não vender na ligação.",
    "Pergunta de avanço: Se fizer sentido, você prefere uma conversa rápida no começo ou no fim do dia?"
  ].join("\n");
}

function fridaySosPriorityClients(clients) {
  return [...clients].sort((a, b) => {
    const aScore = fridayTitanScore(a) + (isRecommendation(a) ? 18 : 0) + (!a.lastInteraction && !a.taDate ? 10 : 0);
    const bScore = fridayTitanScore(b) + (isRecommendation(b) ? 18 : 0) + (!b.lastInteraction && !b.taDate ? 10 : 0);
    return bScore - aScore;
  });
}

function fridaySosRetakeClients(clients) {
  const today = offsetDate(0);
  return fridaySosPriorityClients(clients.filter(client =>
    client.followUp === today
    || client.lastInteraction
    || client.taDate
    || client.nextAction
  ));
}

function fridaySosData(clients, remaining = 0) {
  const groups = fridayReferrerSequenceGroups(clients);
  const bestGroup = fridayReferrerActiveGroup(groups);
  const top10 = fridaySosPriorityClients(clients).slice(0, 10);
  const retake = fridaySosRetakeClients(clients).slice(0, 10);
  const hot = clients.filter(client => isRecommendation(client) && !client.lastInteraction).length;
  const blocks = [
    fridayReferrerBlockObject(bestGroup, remaining) || {
      key: "referrer",
      number: "1",
      title: "Bloco de confiança",
      count: 0,
      detail: "Sem recomendante forte nesta carteira.",
      action: "Executar bloco",
      referrer: ""
    },
    {
      key: "top10",
      number: "2",
      title: "Top 10 menos difíceis",
      count: top10.length,
      detail: top10.length
        ? "Nomes com mais contexto, indicação ou prioridade. Use quando a sexta travar."
        : "Sem nomes elegíveis para montar Top 10.",
      action: "Ver Top 10",
      clients: top10
    },
    {
      key: "retake",
      number: "3",
      title: "Retentativa limpa",
      count: retake.length,
      detail: retake.length
        ? "Pessoas já tocadas ou com próximo passo. Menos frio que começar do zero."
        : "Sem retentativas claras agora.",
      action: "Ver retentativas",
      clients: retake
    }
  ];
  const summary = clients.length
    ? `Respira. Você tem ${clients.length} nome(s) nesta visão, ${hot} REC(s) nova(s) e ${remaining} OI(s) faltando. Ataque um bloco por vez.`
    : "Sem nomes nesta visão. Monte a carteira da semana ou mude o filtro para Todos.";
  return { groups, bestGroup, top10, retake, blocks, summary };
}

function fridaySosQueue(clients) {
  const data = fridaySosData(clients);
  if (fridaySosFocus === "top10") return data.top10;
  if (fridaySosFocus === "retake") return data.retake;
  return null;
}

function fridayQueueLockKey() {
  return [
    fridayPortfolioFilter || "all",
    fridayReferrerFilter || "all",
    fridaySosActive ? "sos" : "normal",
    fridaySosFocus || "referrer",
    fridaySessionActive ? "session" : "free"
  ].join("|");
}

function resetFridayStableQueue() {
  fridayLockedQueueKey = "";
  fridayLockedClientIds = [];
}

function removeFromFridayStableQueue(clientId) {
  if (!clientId) return;
  fridayLockedClientIds = fridayLockedClientIds.filter(id => id !== clientId);
  fridaySessionClientIds = fridaySessionClientIds.filter(id => id !== clientId);
}

function fridayStableQueue(queue) {
  const key = fridayQueueLockKey();
  if (!queue.length) {
    fridayLockedQueueKey = key;
    fridayLockedClientIds = [];
    return fridayFocusedQueue(queue);
  }
  if (fridaySessionActive && fridaySessionClientIds.length) {
    const byId = new Map(queue.map(client => [client.id, client]));
    const ordered = fridaySessionClientIds.map(id => byId.get(id)).filter(Boolean);
    const additions = queue.filter(client => !fridaySessionClientIds.includes(client.id));
    return fridayFocusedQueue([...ordered, ...additions]);
  }
  if (fridayLockedQueueKey !== key || !fridayLockedClientIds.length) {
    fridayLockedQueueKey = key;
    fridayLockedClientIds = queue.map(client => client.id);
    return fridayFocusedQueue(queue);
  }
  const byId = new Map(queue.map(client => [client.id, client]));
  const ordered = fridayLockedClientIds.map(id => byId.get(id)).filter(Boolean);
  const additions = queue.filter(client => !fridayLockedClientIds.includes(client.id));
  fridayLockedClientIds = [...ordered.map(client => client.id), ...additions.map(client => client.id)];
  return fridayFocusedQueue([...ordered, ...additions]);
}

function fridaySelectedBlock(data) {
  if (fridayReferrerFilter !== "all") {
    const selectedGroup = data.groups.find(group => group.name === fridayReferrerFilter);
    return fridayReferrerBlockObject(selectedGroup) || data.blocks[0];
  }
  if (fridaySosFocus) {
    return data.blocks.find(block => block.key === fridaySosFocus) || data.blocks[0];
  }
  return data.blocks.find(block => block.referrer && block.count) || data.blocks.find(block => block.count) || data.blocks[0];
}

function fridaySelectedBlockLabel() {
  if (fridayReferrerFilter !== "all") return `Bloco: ${fridayReferrerFilter}`;
  if (fridaySosFocus === "top10") return "Bloco: Top 10";
  if (fridaySosFocus === "retake") return "Bloco: retentativas";
  return "Bloco: recomendante";
}

function fridaySelectedBlockPlan(data) {
  const block = fridaySelectedBlock(data);
  if (block?.referrer) {
    if (block.referrer === "Sem recomendante") {
      return "Use abordagem curta e humana. Pergunte se pode explicar rapidamente o motivo da ligação, descubra o contexto e ofereça dois horários para uma OI.";
    }
    return `Abra citando ${block.referrer}. Confirme rapidamente como a pessoa conhece o recomendante, explique que a ligação é curta e feche sempre com duas opções de horário para a OI.`;
  }
  if (block?.key === "top10") {
    return "Use velocidade. São os nomes com menor atrito: ligue, explique o motivo em poucos segundos e tente sair com uma OI marcada. Se não atender, WhatsApp curto e próximo.";
  }
  if (block?.key === "retake") {
    return "Reabra sem cobrança. Diga que tentou falar antes, pergunte se pode explicar rapidamente o motivo da ligação e ofereça uma conversa objetiva em dois horários.";
  }
  return "Comece pelo bloco com maior confiança. Uma ligação por vez, sem trocar de estratégia no meio.";
}

function selectFridaySosBlock(block) {
  if (!block || !block.count) return;
  if (block.referrer) {
    fridayReferrerFilter = block.referrer;
    fridaySosFocus = "";
  } else {
    fridayReferrerFilter = "all";
    fridaySosFocus = block.key || "top10";
  }
  fridaySosActive = false;
  fridaySessionActive = false;
  clearFridayQueueFineFilters();
  resetFridayStableQueue();
  renderFridayMode();
}

function currentFridayExecutionQueue(portfolioQueue) {
  const sosQueue = (fridaySosActive || fridaySessionActive || fridaySosFocus) && fridayReferrerFilter === "all"
    ? fridaySosQueue(portfolioQueue)
    : null;
  return fridayStableQueue(sosQueue || fridayQueueClients(fridayPortfolioFilter, fridayReferrerFilter));
}

function startFridaySession(data) {
  const selected = fridaySelectedBlock(data);
  if (!selected?.count) {
    showToast("Escolha um bloco com contatos antes de começar", true);
    return;
  }
  if (selected.referrer) {
    fridayReferrerFilter = selected.referrer;
    fridaySosFocus = "";
  } else {
    fridayReferrerFilter = "all";
    fridaySosFocus = selected.key || "top10";
  }
  fridaySosActive = true;
  fridaySessionActive = true;
  fridaySessionStartedAt = Date.now();
  fridayPinnedClientId = "";
  clearFridayQueueFineFilters();
  resetFridayStableQueue();
  const portfolioQueue = fridayQueueClients(fridayPortfolioFilter, "all");
  fridaySessionClientIds = currentFridayExecutionQueue(portfolioQueue).map(client => client.id);
  fridayLockedClientIds = [...fridaySessionClientIds];
  fridayLockedQueueKey = fridayQueueLockKey();
  fridaySessionResults = [];
  renderFridayMode();
  focusFridayQueue(`Sessão iniciada: ${fridaySelectedBlockLabel()}`);
}

function endFridaySession(message = "Bloco encerrado. Veja a recuperação do TA.") {
  fridaySessionActive = false;
  fridaySosActive = false;
  fridaySosFocus = "";
  fridayReferrerFilter = "all";
  fridayPinnedClientId = "";
  clearFridayQueueFineFilters();
  resetFridayStableQueue();
  renderFridayMode();
  focusFridayQueue(message);
}

function recordFridaySessionResult(clientId, result) {
  if (!fridaySessionActive) return;
  fridaySessionResults.push({ clientId, result, at: Date.now() });
}

function fridaySessionMetrics() {
  const attempted = fridaySessionResults.length;
  const effective = fridaySessionResults.filter(item => ["talked", "scheduled", "declined"].includes(item.result)).length;
  const scheduled = fridaySessionResults.filter(item => item.result === "scheduled").length;
  const intervals = fridaySessionResults.slice(1).map((item, index) => item.at - fridaySessionResults[index].at);
  const avgIntervalMs = intervals.length ? Math.round(intervals.reduce((sum, value) => sum + value, 0) / intervals.length) : 0;
  let longestStreak = attempted ? 1 : 0;
  let currentStreak = attempted ? 1 : 0;
  intervals.forEach(interval => {
    if (interval <= 120000) currentStreak += 1;
    else currentStreak = 1;
    longestStreak = Math.max(longestStreak, currentStreak);
  });
  const contactRate = attempted ? Math.round((effective / attempted) * 100) : 0;
  const rhythm = fridayRhythmFromInterval(avgIntervalMs, attempted);
  return { attempted, effective, scheduled, avgIntervalMs, longestStreak, contactRate, rhythm };
}

function fridayRhythmFromInterval(avgIntervalMs, attempted = 0) {
  if (attempted <= 1 || !avgIntervalMs) return "Ritmo bom";
  if (avgIntervalMs <= 120000) return "Ritmo bom";
  if (avgIntervalMs <= 240000) return "Ritmo atenção";
  return "Ritmo caindo";
}

function fridayCurrentRhythmLabel() {
  if (!fridaySessionActive) return "Ritmo bom";
  const metrics = fridaySessionMetrics();
  if (metrics.attempted <= 1) return "Ritmo bom";
  const last = fridaySessionResults.at(-1)?.at || fridaySessionStartedAt || Date.now();
  const sinceLast = Date.now() - last;
  if (sinceLast > 240000) return "Ritmo caindo";
  if (sinceLast > 120000) return "Ritmo atenção";
  return metrics.rhythm;
}

function fridayOiProgressMeta() {
  const start = startOfWeek(0);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  const week = { start, end, startKey: dateKey(start), endKey: dateKey(end) };
  const scheduled = state.clients.filter(client => inSelectedWeek(client.oiScheduledDate, week)).length;
  const target = Math.max(weeklyGoal(week), 25);
  return {
    scheduled,
    target,
    percent: Math.min(Math.round((scheduled / target) * 100), 100),
    label: `${scheduled} OIs · meta 20-25`
  };
}

function fridaySosPlanText(data) {
  const names = list => list.slice(0, 10).map((client, index) => `${index + 1}. ${client.name}${client.referrer ? ` - indicação de ${client.referrer}` : ""}`).join("\n");
  return [
    "SEXTA SOS - PLANO DE 90 MINUTOS",
    "",
    data.summary,
    "",
    "BLOCO 1 - CONFIANÇA",
    data.bestGroup ? fridayReferrerSequenceScript(data.bestGroup) : "Sem recomendante forte. Comece pelo Top 10.",
    "",
    "BLOCO 2 - TOP 10 MENOS DIFÍCEIS",
    names(data.top10) || "Sem nomes elegíveis.",
    "",
    "BLOCO 3 - RETENTATIVA LIMPA",
    names(data.retake) || "Sem retentativas claras.",
    "",
    "REGRA DO SOS",
    "Não tente resolver a sexta inteira. Execute um bloco, registre resultado e avance para o próximo."
  ].join("\n");
}

function focusFridayQueue(message = "") {
  if (message) showToast(message);
  setTimeout(() => {
    el("fridayQueuePanel")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }, 80);
}

function fridayWorkedToday(client) {
  const today = offsetDate(0);
  return state.activities.some(activity =>
    activity.clientId === client.id
    && activity.date === today
    && [
      "Sexta - nao atendeu",
      "Sexta - não atendeu",
      "Sexta - atendeu sem agenda",
      "Sexta - OI marcada",
      "Sexta - sem interesse",
      "Sexta - sem interesse com retomada",
      "Sexta - sem interesse virou C2 leve",
      "Ligação sem atendimento"
    ].includes(activity.title)
  );
}

function fridayActiveBlockMeta(queue) {
  const count = queue.length;
  if (fridaySessionActive) {
    const current = queue[0];
    const progress = fridayOiProgressMeta();
    return {
      kicker: `${fridaySelectedBlockLabel()} · ${fridayCurrentRhythmLabel()}`,
      title: current ? `Agora: ${current.name}` : "Bloco concluído",
      detail: current
        ? `${progress.label}. Registre o resultado para avançar automaticamente.`
        : "Todos os contatos deste bloco foram trabalhados. Veja a recuperação e escolha o próximo bloco.",
      clearable: true,
      clearText: "Encerrar bloco",
      progress
    };
  }
  if (fridayReferrerFilter !== "all") {
    return {
      kicker: "Sequência por recomendante",
      title: `Bloco ativo: ${fridayReferrerFilter}`,
      detail: `${count} contato(s) filtrado(s). Use o mesmo gancho de confiança e registre um resultado antes de trocar de bloco.`,
      clearable: true
    };
  }
  if (fridaySosActive && fridaySosFocus === "top10") {
    return {
      kicker: "Sexta SOS",
      title: "Bloco ativo: Top 10 menos difíceis",
      detail: `${count} contato(s) com melhor combinação de contexto, indicação e prioridade. Ligue um por vez.`,
      clearable: true
    };
  }
  if (fridaySosActive && fridaySosFocus === "retake") {
    return {
      kicker: "Sexta SOS",
      title: "Bloco ativo: Retentativas limpas",
      detail: `${count} contato(s) já tocado(s) ou com próximo passo. Foque em reabrir conversa e oferecer ligação.`,
      clearable: true
    };
  }
  const current = queue[0];
  if (!current) return null;
  const referrer = current.referrer ? `Indicação de ${current.referrer}` : "sem recomendante";
  return {
    kicker: "Modo piloto TA",
    title: `Agora: ${current.name}`,
    detail: `${fridayCallLabel(current)} · ${referrer}. Ligue, tente marcar OI e registre o resultado para o SITPLAN avançar.`,
    clearable: false
  };
}

function clearFridayActiveBlock() {
  fridaySessionActive = false;
  fridaySosActive = false;
  fridaySosFocus = "";
  fridayReferrerFilter = "all";
  fridayPinnedClientId = "";
  clearFridayQueueFineFilters();
  resetFridayStableQueue();
  renderFridayMode();
  focusFridayQueue("Fila completa liberada");
}

function renderFridayActiveStrip(queue) {
  const strip = el("fridayActiveStrip");
  if (!strip) return;
  const meta = fridayActiveBlockMeta(queue);
  strip.classList.toggle("hidden", !meta);
  if (!meta) return;
  el("fridayActiveKicker").textContent = meta.kicker;
  el("fridayActiveTitle").textContent = meta.title;
  el("fridayActiveDetail").textContent = meta.detail;
  const progress = el("fridaySessionProgress");
  if (progress) {
    progress.classList.toggle("hidden", !meta.progress);
    const bar = progress.querySelector("span");
    if (bar) bar.style.width = `${meta.progress?.percent || 0}%`;
  }
  const clearButton = el("fridayClearActiveBlock");
  if (clearButton) {
    clearButton.classList.toggle("hidden", !meta.clearable);
    clearButton.textContent = meta.clearText || "Voltar para fila completa";
    clearButton.onclick = clearFridayActiveBlock;
  }
}

function fridayRecoveryData() {
  const today = offsetDate(0);
  const noAnswerIds = new Set(state.activities
    .filter(activity => activity.date === today && ["Sexta - nao atendeu", "Sexta - não atendeu", "Ligação sem atendimento"].includes(activity.title))
    .map(activity => activity.clientId)
    .filter(Boolean));
  const byId = id => state.clients.find(client => client.id === id);
  const noAnswer = [...noAnswerIds]
    .map(byId)
    .filter(Boolean)
    .filter(client => !client.oiScheduledDate && !client.doNotContact)
    .slice(0, 8);
  const declinedToday = state.clients
    .filter(client => client.taNoInterestDate === today)
    .sort((a, b) => (b.updatedAt || b.createdAt || "").localeCompare(a.updatedAt || a.createdAt || ""))
    .slice(0, 8);
  const paused = state.clients
    .filter(client => client.taPausedUntil && client.taPausedUntil > today && client.taNoInterestDate !== today && !client.doNotContact)
    .sort((a, b) => String(a.taPausedUntil).localeCompare(String(b.taPausedUntil)))
    .slice(0, 8);
  return { noAnswer, declinedToday, paused };
}

function fridayRecoveryItem(client, kind) {
  const referrer = client.referrer ? `Indicação de ${client.referrer}` : "Sem indicador";
  const reason = kind === "no-answer"
    ? `Próxima tentativa: ${client.followUp ? formatDate(client.followUp) : "definir"}`
    : client.taNoInterestReason
      ? `Motivo: ${client.taNoInterestReason}`
      : client.taPausedUntil
        ? `Pausado até ${formatDate(client.taPausedUntil)}`
        : "Revisar contato";
  const whatsapp = kind === "no-answer" ? whatsappLink(client, true) : "";
  return `<article class="friday-recovery-item">
    <div>
      <strong>${escapeHtml(client.name)}</strong>
      <small>${escapeHtml(referrer)}</small>
      <p>${escapeHtml(reason)}</p>
    </div>
    <div>
      ${whatsapp ? `<a href="${whatsapp}" target="_blank" rel="noopener">WhatsApp</a>` : ""}
      <button type="button" data-friday-recovery-focus="${client.id}">Focar</button>
      <button type="button" data-friday-open="${client.id}">Abrir</button>
    </div>
  </article>`;
}

function renderFridayRecoveryPanel() {
  const panel = el("fridayRecoveryPanel");
  if (!panel) return;
  const { noAnswer, declinedToday, paused } = fridayRecoveryData();
  const hasItems = noAnswer.length || declinedToday.length || paused.length;
  const totalRecovery = noAnswer.length + declinedToday.length + paused.length;
  if (fridaySessionActive) {
    panel.innerHTML = `<button type="button" class="friday-recovery-badge" data-friday-end-session>
      <span>Recuperação do TA</span>
      <strong>${totalRecovery}</strong>
    </button>`;
    document.querySelector("[data-friday-end-session]")?.addEventListener("click", () => endFridaySession());
    return;
  }
  if (!hasItems) {
    panel.innerHTML = "";
    return;
  }
  const metrics = fridaySessionMetrics();
  const metricHtml = metrics.attempted ? `<div class="friday-session-summary">
    <article><small>Tentativas no bloco</small><strong>${metrics.attempted}</strong></article>
    <article><small>Contato efetivo</small><strong>${metrics.contactRate}%</strong></article>
    <article><small>OIs marcadas</small><strong>${metrics.scheduled}</strong></article>
    <article><small>Ritmo do bloco</small><strong>${metrics.rhythm}</strong></article>
    <article><small>Maior sequência</small><strong>${metrics.longestStreak}</strong></article>
  </div>` : "";
  panel.innerHTML = `<div class="friday-recovery-header">
    <div>
      <p class="eyebrow">RECUPERAÇÃO DO TA</p>
      <h3>Não perca o rastro do que saiu da ligação principal</h3>
      <p>Controle rápido dos não atendidos, recusas e contatos pausados sem misturar com a fila principal.</p>
    </div>
  </div>
  ${metricHtml}
  <div class="friday-recovery-grid">
    <section>
      <h4>Não atenderam hoje <span>${noAnswer.length}</span></h4>
      ${noAnswer.length ? noAnswer.map(client => fridayRecoveryItem(client, "no-answer")).join("") : `<div class="mini-empty">Nenhum não atendimento registrado hoje.</div>`}
    </section>
    <section>
      <h4>Sem interesse / motivo <span>${declinedToday.length}</span></h4>
      ${declinedToday.length ? declinedToday.map(client => fridayRecoveryItem(client, "declined")).join("") : `<div class="mini-empty">Nenhum sem interesse registrado hoje.</div>`}
    </section>
    <section>
      <h4>Pausados para retomar <span>${paused.length}</span></h4>
      ${paused.length ? paused.map(client => fridayRecoveryItem(client, "paused")).join("") : `<div class="mini-empty">Nenhum contato pausado para retomar.</div>`}
    </section>
  </div>`;
  document.querySelectorAll("[data-friday-recovery-focus]").forEach(button => button.addEventListener("click", () => {
    fridayPinnedClientId = button.dataset.fridayRecoveryFocus;
    fridaySosActive = false;
    fridaySosFocus = "";
    fridayReferrerFilter = "all";
    fridayPortfolioFilter = "all";
    renderFridayMode();
    focusFridayQueue("Contato fixado na fila");
  }));
}

function fridayDeclinePanelHtml(client) {
  if (fridaySessionActive) {
    const reasons = ["Não quer contato", "Sem tempo agora", "Já tem seguro", "Não vê prioridade", "Não quis conversar", "Pediu retorno futuro"];
    return `<div class="friday-decline-panel friday-decline-quick hidden" id="friday-decline-${client.id}">
      <strong>Escolha o motivo</strong>
      <div class="friday-decline-reason-grid">
        ${reasons.map(reason => `<button type="button" data-friday-decline-quick="${client.id}" data-friday-decline-quick-reason="${escapeHtml(reason)}">${escapeHtml(reason)}</button>`).join("")}
      </div>
    </div>`;
  }
  return `<div class="friday-decline-panel hidden" id="friday-decline-${client.id}">
    <div>
      <label>Motivo principal
        <select data-friday-decline-reason="${client.id}">
          <option value="">Selecione</option>
          <option>Não quer contato</option>
          <option>Sem tempo agora</option>
          <option>Já tem seguro</option>
          <option>Não vê prioridade</option>
          <option>Não quis conversar</option>
          <option>Não é perfil</option>
          <option>Pediu retorno futuro</option>
          <option>Outro</option>
        </select>
      </label>
      <label>Destino
        <select data-friday-decline-action="${client.id}">
          <option value="remove">Retirar da fila</option>
          <option value="future">Acompanhar no futuro</option>
        </select>
      </label>
    </div>
    <label>Observação curta
      <input data-friday-decline-note="${client.id}" type="text" maxlength="140" placeholder="Ex.: pediu para não ligar mais, já tem cobertura, retomar no próximo trimestre">
    </label>
    <div class="friday-decline-actions">
      <button type="button" class="success" data-friday-decline-save="${client.id}">Salvar motivo</button>
      <button type="button" class="muted" data-friday-decline-cancel="${client.id}">Cancelar</button>
    </div>
  </div>`;
}

function renderFridaySosPanel(clients, remaining = 0) {
  if (!el("fridaySosBlocks")) return;
  const data = fridaySosData(clients, remaining);
  const selected = fridaySelectedBlock(data);
  el("fridaySosPanel")?.classList.toggle("active", fridaySosActive);
  el("fridaySosSummary").textContent = fridaySosActive
    ? `${data.summary} SOS ativo: execute a fila mostrada abaixo, sem trocar de estratégia a cada ligação.`
    : "Escolha um bloco, leia o plano de ataque e comece. Durante a execução, o SITPLAN mostra só um contato por vez.";
  const startButton = el("fridaySosStart");
  if (startButton) {
    startButton.disabled = !selected?.count || fridaySessionActive;
    startButton.textContent = fridaySessionActive ? "Bloco em execução" : "Começar bloco";
    startButton.onclick = () => startFridaySession(data);
  }
  if (el("fridaySosPlan")) {
    el("fridaySosPlan").innerHTML = `<small>Plano de ataque · ${escapeHtml(selected?.title || "bloco")}</small>
      <p>${escapeHtml(fridaySelectedBlockPlan(data))}</p>`;
  }
  el("fridaySosBlocks").innerHTML = data.blocks.map(block => {
    const isActive = selected && ((block.key && block.key === selected.key) || (block.referrer && block.referrer === selected.referrer));
    const recommended = block === data.blocks[0];
    const disabled = !block.count;
    const action = block.referrer
      ? `data-friday-sos-referrer="${escapeHtml(block.referrer)}"`
      : `data-friday-sos-focus="${escapeHtml(block.key)}"`;
    return `<article class="friday-sos-card ${isActive ? "active" : ""}">
      <span>${block.number}</span>
      <div>
        <small>${block.count} contato(s)${recommended ? " · recomendado" : ""}</small>
        <strong>${escapeHtml(block.title)}</strong>
        <p>${escapeHtml(block.detail)}</p>
        <button type="button" ${action} ${disabled ? "disabled" : ""}>Selecionar bloco</button>
      </div>
    </article>`;
  }).join("");
  document.querySelectorAll("[data-friday-sos-referrer]").forEach(button => button.addEventListener("click", () => {
    const block = data.blocks.find(item => item.referrer === button.dataset.fridaySosReferrer);
    selectFridaySosBlock(block);
  }));
  document.querySelectorAll("[data-friday-sos-focus]").forEach(button => button.addEventListener("click", () => {
    const block = data.blocks.find(item => item.key === button.dataset.fridaySosFocus);
    selectFridaySosBlock(block);
  }));
}

function renderFridayReferrerSequence(clients, allClients = clients) {
  if (!el("fridayReferrerGroups")) return;
  const viewGroups = fridayReferrerSequenceGroups(clients);
  const allGroups = fridayReferrerSequenceGroups(allClients);
  const selected = fridayReferrerFilter !== "all" ? allGroups.find(group => group.name === fridayReferrerFilter) : null;
  const namedGroups = viewGroups.filter(group => group.name !== "Sem recomendante");
  const allNamedGroups = allGroups.filter(group => group.name !== "Sem recomendante");
  const activeGroup = selected || fridayReferrerActiveGroup(viewGroups) || fridayReferrerActiveGroup(allGroups);
  const viewGroupNames = new Set(viewGroups.map(group => group.name));
  el("fridayReferrerSummary").textContent = selected
    ? `Filtro ativo: ${selected.name}. ${selected.clients.length} contato(s) na base TA. Use o mesmo discurso antes de trocar de bloco.`
    : allNamedGroups.length
      ? `${namedGroups.length} recomendante(s) nesta visão · ${allNamedGroups.length} na base TA completa. Busque pelo nome para abrir qualquer bloco.`
      : "Ainda não há recomendantes suficientes nesta visão. Cadastre o campo Indicador para ativar blocos melhores.";
  const selector = el("fridayReferrerSelect");
  if (selector) {
    const currentValue = allGroups.some(group => group.name === fridayReferrerFilter) ? fridayReferrerFilter : "all";
    selector.innerHTML = `<option value="all">Todos os recomendantes da base TA</option>` + allGroups.map(group => {
      const label = group.name === "Sem recomendante"
        ? `Sem recomendante (${group.clients.length})`
        : `${group.name} (${group.clients.length})${viewGroupNames.has(group.name) ? "" : " · fora da visão atual"}`;
      return `<option value="${escapeHtml(group.name)}">${escapeHtml(label)}</option>`;
    }).join("");
    selector.value = currentValue;
    selector.onchange = event => {
      fridayReferrerFilter = event.target.value || "all";
      if (fridayReferrerFilter !== "all" && !viewGroupNames.has(fridayReferrerFilter)) {
        fridayPortfolioFilter = "all";
      }
      fridaySosFocus = "";
      fridaySosActive = false;
      fridaySessionActive = false;
      fridayPinnedClientId = "";
      clearFridayQueueFineFilters();
      resetFridayStableQueue();
      renderFridayMode();
      focusFridayQueue(fridayReferrerFilter === "all" ? "Filtro por recomendante limpo" : `Filtro aberto: ${fridayReferrerFilter}`);
    };
  }
  const searchInput = el("fridayReferrerSearch");
  if (searchInput) {
    if (searchInput.value !== fridayReferrerSearch) searchInput.value = fridayReferrerSearch;
    searchInput.oninput = event => {
      fridayReferrerSearch = event.target.value || "";
      renderFridayReferrerSequence(clients, allClients);
    };
  }
  const playbook = el("fridayReferrerPlaybook");
  if (playbook) {
    playbook.innerHTML = fridayReferrerPlaybookHtml(activeGroup, Boolean(selected));
  }
  const copyActive = el("fridayReferrerCopyActive");
  if (copyActive) {
    copyActive.disabled = !activeGroup;
    copyActive.onclick = () => activeGroup && copyFridayReferrerSequence(activeGroup.name);
  }
  const warmupActive = el("fridayReferrerWarmupActive");
  if (warmupActive) {
    warmupActive.disabled = !activeGroup || activeGroup.name === "Sem recomendante";
    warmupActive.onclick = () => activeGroup && copyFridayReferrerWarmup(activeGroup.name);
  }
  const clearButton = el("fridayReferrerClear");
  if (clearButton) {
    clearButton.disabled = fridayReferrerFilter === "all";
    clearButton.textContent = fridayReferrerFilter === "all" ? "Ver todos" : "Limpar sequência";
    clearButton.onclick = () => {
      fridayReferrerFilter = "all";
      fridaySosActive = false;
      fridaySosFocus = "";
      fridaySessionActive = false;
      fridayPinnedClientId = "";
      fridayReferrerSearch = "";
      clearFridayQueueFineFilters();
      resetFridayStableQueue();
      renderFridayMode();
      focusFridayQueue("Sequência limpa");
    };
  }
  const searchTerm = fridaySearchText(fridayReferrerSearch).trim();
  const matchingGroups = allGroups.filter(group => !searchTerm || fridaySearchText(group.name).includes(searchTerm));
  const visibleGroups = (searchTerm || selected) ? matchingGroups : matchingGroups.slice(0, 12);
  el("fridayReferrerGroups").innerHTML = visibleGroups.length ? visibleGroups.map(group => {
    const active = group.name === fridayReferrerFilter;
    const title = group.name === "Sem recomendante" ? "Sem recomendante" : group.name;
    const sample = group.clients.slice(0, 3).map(client => client.name.trim().split(/\s+/)[0]).filter(Boolean).join(", ");
    const outsideView = !viewGroupNames.has(group.name);
    return `<article class="friday-referrer-card ${active ? "active" : ""}">
      <button type="button" data-friday-referrer-sequence="${escapeHtml(group.name)}">
        <span>${group.clients.length}</span>
        <div>
          <strong>${escapeHtml(title)}</strong>
          <small>${group.hot} REC nova(s) · ${group.uncalled} sem tentativa · ${group.worked} já tocado(s)${outsideView ? " · fora da visão atual" : ""}</small>
          <em>${escapeHtml(sample ? `Sequência: ${sample}` : "Sem nomes para amostra")}</em>
        </div>
      </button>
      <button type="button" class="friday-referrer-copy warmup" data-copy-friday-referrer-warmup="${escapeHtml(group.name)}">Copiar aquecimento</button>
      <button type="button" class="friday-referrer-copy" data-copy-friday-referrer-script="${escapeHtml(group.name)}">Copiar fala do bloco</button>
    </article>`;
  }).join("") : `<div class="mini-empty">Nenhum recomendante encontrado nessa busca.</div>`;
  document.querySelectorAll("[data-friday-referrer-sequence]").forEach(button => button.addEventListener("click", () => {
    fridayReferrerFilter = button.dataset.fridayReferrerSequence || "all";
    if (fridayReferrerFilter !== "all" && !viewGroupNames.has(fridayReferrerFilter)) {
      fridayPortfolioFilter = "all";
    }
    fridaySosFocus = "";
    fridaySosActive = false;
    fridaySessionActive = false;
    fridayPinnedClientId = "";
    clearFridayQueueFineFilters();
    resetFridayStableQueue();
    renderFridayMode();
    focusFridayQueue(`Sequência ${fridayReferrerFilter} aberta na fila`);
  }));
  document.querySelectorAll("[data-copy-friday-referrer-script]").forEach(button => button.addEventListener("click", event => {
    event.preventDefault();
    event.stopPropagation();
    copyFridayReferrerSequence(button.dataset.copyFridayReferrerScript);
  }));
  document.querySelectorAll("[data-copy-friday-referrer-warmup]").forEach(button => button.addEventListener("click", event => {
    event.preventDefault();
    event.stopPropagation();
    copyFridayReferrerWarmup(button.dataset.copyFridayReferrerWarmup);
  }));
}

function renderFridayQueueControls(baseQueue, displayQueue, counts, callsToday) {
  const searchInput = el("fridayQueueSearch");
  if (searchInput) {
    if (searchInput.value !== fridayQueueSearch) searchInput.value = fridayQueueSearch;
    searchInput.oninput = event => {
      fridayQueueSearch = event.target.value || "";
      renderFridayMode();
    };
  }
  const setCount = (id, value) => {
    if (el(id)) el(id).textContent = value;
  };
  setCount("fridayTempAll", counts.all);
  setCount("fridayTempHot", counts.Quente);
  setCount("fridayTempWarm", counts.Morno);
  setCount("fridayTempCold", counts.Frio);
  document.querySelectorAll("[data-friday-temperature]").forEach(button => {
    const value = button.dataset.fridayTemperature || "all";
    button.classList.toggle("active", fridayTemperatureFilter === value);
    button.onclick = () => {
      fridayTemperatureFilter = fridayTemperatureFilter === value ? "all" : value;
      renderFridayMode();
    };
  });
  const hasActiveFilter = Boolean(fridayQueueSearch.trim()) || fridayTemperatureFilter !== "all";
  const total = Math.max(hasActiveFilter ? displayQueue.length : baseQueue.length, callsToday);
  const pct = total ? Math.min(Math.round((callsToday / total) * 100), 100) : 0;
  if (el("fridayCallsProgressText")) el("fridayCallsProgressText").textContent = `${callsToday} de ${total} ligados hoje`;
  if (el("fridayCallsProgressBar")) el("fridayCallsProgressBar").style.width = `${pct}%`;
}

async function copyFridayReferrerSequence(referrerName) {
  const group = fridayReferrerSequenceGroups(fridayQueueClients("all", "all"))
    .find(item => item.name === referrerName);
  if (!group) {
    showToast("Sequência não encontrada nesta fila", true);
    return;
  }
  try {
    await copyTextToClipboard(fridayReferrerSequenceScript(group));
    showToast("Fala do bloco copiada");
  } catch (_) {
    showToast("Não foi possível copiar a fala do bloco", true);
  }
}

async function copyFridayReferrerWarmup(referrerName) {
  const group = fridayReferrerSequenceGroups(fridayQueueClients("all", "all"))
    .find(item => item.name === referrerName);
  if (!group) {
    showToast("Recomendante não encontrado nesta fila", true);
    return;
  }
  try {
    await copyTextToClipboard(fridayReferrerWarmupMessage(group));
    showToast("Mensagem para aquecer recomendante copiada");
  } catch (_) {
    showToast("Não foi possível copiar a mensagem", true);
  }
}

function fridayQueueClients(portfolioFilter = "all", referrerFilter = "all", options = {}) {
  const currentKey = taPortfolioWeekKey(0);
  const nextKey = taPortfolioWeekKey(1);
  const today = offsetDate(0);
  const includeWorkedToday = Boolean(options.includeWorkedToday);
  return state.clients
    .filter(client => ["Base", "TA"].includes(client.stage) && !client.oiScheduledDate && !client.doNotContact)
    .filter(client => !client.taPausedUntil || client.taPausedUntil <= today)
    .filter(client => includeWorkedToday || !fridayWorkedToday(client))
    .filter(client => {
      if (portfolioFilter === "current") return client.taPortfolioWeek === currentKey;
      if (portfolioFilter === "next") return client.taPortfolioWeek === nextKey;
      if (portfolioFilter === "carryover") return isTaPortfolioCarryover(client);
      if (portfolioFilter === "bank") return !client.taPortfolioWeek;
      return true;
    })
    .filter(client => referrerFilter === "all" || fridayReferrerName(client) === referrerFilter)
    .sort((a, b) => {
      const score = client => {
        if (isRecommendation(client) && !client.lastInteraction) return 0;
        if (isRecommendation(client)) return 1;
        if (!client.lastInteraction && !client.taDate) return 2;
        return 3;
      };
      const due = Number(fridayNextCallDue(b)) - Number(fridayNextCallDue(a));
      if (due) return due;
      const temperature = fridayTemperatureRank(a) - fridayTemperatureRank(b);
      if (temperature) return temperature;
      const received = String(fridayReceivedKey(a)).localeCompare(String(fridayReceivedKey(b)));
      if (received) return received;
      const priority = score(a) - score(b);
      return priority || assistantPriorityScore(b) - assistantPriorityScore(a);
    });
}

function fridayFocusedQueue(queue) {
  if (!fridayPinnedClientId) return queue;
  const pinned = queue.find(client => client.id === fridayPinnedClientId)
    || fridayQueueClients("all", "all", { includeWorkedToday: true }).find(client => client.id === fridayPinnedClientId);
  if (!pinned) {
    fridayPinnedClientId = "";
    return queue;
  }
  return [pinned, ...queue.filter(client => client.id !== fridayPinnedClientId)];
}

function fridayCallLabel(client) {
  if (isRecommendation(client) && !client.lastInteraction) return "Recomendação nova";
  if (isRecommendation(client)) return "Recomendação para retomar";
  if (!client.lastInteraction && !client.taDate) return "Nunca abordado";
  return "TA sem OI marcada";
}

function fridayTitanScore(client) {
  let score = assistantPriorityScore(client);
  if (isRecommendation(client)) score += client.lastInteraction ? 30 : 45;
  if (!client.lastInteraction && !client.taDate) score += 25;
  if (client.stage === "TA") score += 10;
  if ((daysSince(client.lastInteraction) ?? 0) >= 7) score += 12;
  return Math.min(Math.round(score), 100);
}

function fridayTitanReason(client) {
  if (isRecommendation(client) && !client.lastInteraction) return "REC nova: atacar primeiro";
  if (isRecommendation(client)) return "Indicação: alta chance de OI";
  if (!client.lastInteraction && !client.taDate) return "Contato limpo: primeira abordagem";
  if (client.stage === "TA") return "TA aberto: falta converter em OI";
  return "Retomar e buscar dois horários";
}

function fridayActivityCount(titles) {
  const today = offsetDate(0);
  return state.activities.filter(activity => activity.date === today && titles.includes(activity.title)).length;
}

function nameEngineData(week, goal, oiScheduled) {
  const queue = fridayQueueClients();
  const hot = queue.filter(client => isRecommendation(client));
  const uncalled = queue.filter(client => !client.lastInteraction && !client.taDate);
  const taOi = conversionRate(
    state.clients.filter(client => client.oiScheduledDate).length,
    state.clients.filter(client => client.taDate).length,
    .3
  );
  const oiRemaining = Math.max(goal - oiScheduled, 0);
  const requiredTa = oiRemaining ? Math.ceil(oiRemaining / Math.max(taOi, .12)) : 0;
  const healthyStock = Math.max(30, requiredTa + 15);
  const missingNames = Math.max(healthyStock - queue.length, 0);
  const ranking = referralRankingData();
  const hubs = ranking
    .filter(item => item.leads || item.declared || item.n)
    .sort((a, b) => {
      const score = item => (item.n * 4) + (item.oi * 2) + item.leads + Math.max(item.declared - item.leads, 0);
      return score(b) - score(a);
    })
    .slice(0, 3);
  const summary = missingNames
    ? `Seu estoque atual tem ${queue.length} nome(s). Para sustentar TA sem sofrer, busque pelo menos ${missingNames} novo(s) nome(s) nesta semana.`
    : `Estoque saudável: ${queue.length} nome(s) prontos. Agora execute por ordem de qualidade.`;
  return { queue, hot, uncalled, requiredTa, healthyStock, missingNames, hubs, summary };
}

function renderNameEngine(week, goal, oiScheduled) {
  if (!el("nameEngineSummary")) return;
  const data = nameEngineData(week, goal, oiScheduled);
  el("nameEngineSummary").textContent = data.summary;
  el("nameEngineReady").textContent = data.queue.length;
  el("nameEngineReadyDetail").textContent = `${data.hot.length} indicação(ões), ${data.uncalled.length} nunca abordado(s).`;
  el("nameEngineHot").textContent = data.hot.length;
  el("nameEngineNeed").textContent = data.missingNames;
  el("nameEngineNeedDetail").textContent = data.requiredTa
    ? `Para buscar ${Math.max(goal - oiScheduled, 0)} OI(s), o funil estima ${data.requiredTa} TA(s).`
    : "Meta de OIs da semana já protegida.";
  el("nameEngineHubs").innerHTML = data.hubs.length ? data.hubs.map((item, index) => {
    const pending = Math.max(Number(item.declared || 0) - Number(item.leads || 0), 0);
    const action = pending
      ? `${pending} REC(s) declarada(s) ainda sem virar contato.`
      : item.n
        ? "Gerou N: peça nomes parecidos com esse perfil."
        : "Reativar e pedir 2 bons nomes para TA.";
    return `<button class="name-hub-card" data-referrer-filter="${escapeHtml(item.name)}">
      <span>${index + 1}</span>
      <div>
        <strong>${escapeHtml(item.name)}</strong>
        <small>${item.leads} indicado(s) · ${item.oi} OI(s) · ${item.n} N(s)</small>
        <em>${escapeHtml(action)}</em>
      </div>
    </button>`;
  }).join("") : `<div class="mini-empty">Cadastre indicadores nos contatos para o SITPLAN sugerir quem pedir recomendações.</div>`;
  document.querySelectorAll(".name-hub-card").forEach(button => button.addEventListener("click", () => {
    workFilter = "";
    switchView("clients");
    renderClients();
    if (el("referrerFilter")) {
      el("referrerFilter").value = button.dataset.referrerFilter;
      renderClients();
    }
  }));
}

function renderTaPortfolioPanel() {
  if (!el("taPortfolioCurrentCount")) return;
  const data = taPortfolioData();
  const currentMissing = Math.max(data.target - data.current.length, 0);
  const nextMissing = Math.max(data.target - data.next.length, 0);
  el("taPortfolioTarget").value = data.target;
  el("taPortfolioCurrentCount").textContent = `${data.current.length}/${data.target}`;
  el("taPortfolioCurrentDetail").textContent = currentMissing
    ? `Faltam ${currentMissing} nome(s) para fechar a carteira.`
    : `${data.currentWorked.length} já trabalhado(s), ${data.currentUntouched.length} sem primeira tentativa.`;
  el("taPortfolioWorked").textContent = data.currentWorked.length;
  el("taPortfolioUntouched").textContent = data.currentUntouched.length;
  el("taPortfolioNextCount").textContent = `${data.next.length}/${data.target}`;
  el("taPortfolioNextDetail").textContent = nextMissing
    ? `Ainda cabe(m) ${nextMissing} nome(s) para antecipar a próxima sexta.`
    : "Próxima semana já protegida.";
  el("taPortfolioCarryoverCount").textContent = data.carryover.length;
  el("taPortfolioCarryoverDetail").textContent = data.carryover.length
    ? "Nomes que sobraram e devem ser reciclados antes de esfriar."
    : "Nenhuma sobra crítica da semana anterior.";
  el("taPortfolioBankCount").textContent = data.bank.length;
  el("taPortfolioBankDetail").textContent = data.bank.length
    ? "Banco disponível para completar novas carteiras."
    : "Sem nomes livres fora das carteiras atuais.";
  document.querySelectorAll("[data-ta-portfolio-filter]").forEach(button => {
    button.classList.toggle("active", button.dataset.taPortfolioFilter === fridayPortfolioFilter);
  });
}

function buildTaPortfolio(offset = 0) {
  if (isCloudReadOnly()) {
    showToast("Modo operacional: montagem de carteira bloqueada.", true);
    return;
  }
  const weekKey = taPortfolioWeekKey(offset);
  const added = assignTaPortfolio(weekKey);
  saveState();
  fridayPortfolioFilter = offset === 0 ? "current" : "next";
  renderFridayMode();
  showToast(added
    ? `${added} nome(s) adicionados à ${offset === 0 ? "carteira desta semana" : "próxima carteira"}.`
    : `${offset === 0 ? "Carteira desta semana" : "Próxima carteira"} já está no alvo.`);
}

function updateTaPortfolioTarget(value) {
  if (isCloudReadOnly()) {
    showToast("Modo operacional: meta de carteira bloqueada.", true);
    return;
  }
  const target = Math.max(10, Math.min(500, Number(value) || TA_PORTFOLIO_DEFAULT_TARGET));
  state.settings ||= {};
  state.settings.taPortfolioWeeklyTarget = target;
  saveState();
  renderFridayMode();
  showToast(`Meta da Carteira TA ajustada para ${target} nomes por semana.`);
}

function renderFridayMode() {
  if (!el("fridayCallList")) return;
  el("fridayView")?.classList.toggle("friday-session-running", fridaySessionActive);
  const taViewer = isCloudReadOnly();
  const start = startOfWeek(0);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  const week = { start, end, startKey: dateKey(start), endKey: dateKey(end) };
  const goal = weeklyGoal(week);
  const oiScheduled = state.clients.filter(client => inSelectedWeek(client.oiScheduledDate, week)).length;
  const remaining = Math.max(goal - oiScheduled, 0);
  const noInterestTitles = ["Sexta - sem interesse", "Sexta - sem interesse com retomada", "Sexta - sem interesse virou C2 leve"];
  const calls = fridayActivityCount(["Sexta - nao atendeu", "Sexta - não atendeu", "Sexta - atendeu sem agenda", "Sexta - OI marcada", ...noInterestTitles]);
  const talks = fridayActivityCount(["Sexta - atendeu sem agenda", "Sexta - OI marcada", ...noInterestTitles]);
  const totalTa = state.clients.filter(client => client.taDate).length;
  const totalOi = state.clients.filter(client => client.oiScheduledDate).length;
  const fullQueue = fridayQueueClients("all", "all");
  const fullReferrerQueue = fridayQueueClients("all", "all", { includeWorkedToday: true });
  if (fridayPortfolioFilter === "current" && !fridayQueueClients("current", "all").length) fridayPortfolioFilter = "all";
  let portfolioQueue = fridayQueueClients(fridayPortfolioFilter, "all");
  if (fridayReferrerFilter !== "all" && !portfolioQueue.some(client => fridayReferrerName(client) === fridayReferrerFilter)) {
    if (fullQueue.some(client => fridayReferrerName(client) === fridayReferrerFilter)) {
      fridayPortfolioFilter = "all";
      portfolioQueue = fridayQueueClients(fridayPortfolioFilter, "all");
    }
  }
  const baseQueue = currentFridayExecutionQueue(portfolioQueue);
  const queueControls = fridayApplyQueueControls(baseQueue, { ignoreControls: fridaySessionActive });
  const queue = queueControls.filtered;
  if (fridaySessionActive && !queue.length) {
    fridaySessionActive = false;
    fridaySosActive = false;
  }
  const displayQueue = fridaySessionActive ? queue.slice(0, 1) : queue.slice(0, fridayQueueVisibleLimit);
  const ritualRecs = fullQueue.filter(client => isRecommendation(client) && !client.lastInteraction).length;
  const ritualUncalled = fullQueue.filter(client => !isRecommendation(client) && !client.lastInteraction && !client.taDate).length;
  const ritualTa = fullQueue.filter(client => client.stage === "TA" || client.taDate).length;
  const ritualRetry = fullQueue.filter(client => client.followUp === offsetDate(0) || (client.lastInteraction && (daysSince(client.lastInteraction) ?? 0) >= 1)).length;

  el("fridayOiScheduled").textContent = oiScheduled;
  el("fridayOiGoal").textContent = goal;
  el("fridayOiRemaining").textContent = remaining ? `Faltam ${remaining}` : "Meta atingida";
  el("fridayCallsToday").textContent = calls;
  el("fridayTalksToday").textContent = talks;
  el("fridayOisWeek").textContent = oiScheduled;
  el("fridayConversion").textContent = `${totalTa ? Math.round((totalOi / totalTa) * 100) : 0}%`;
  el("fridayRitualRecs").textContent = ritualRecs;
  el("fridayRitualUncalled").textContent = ritualUncalled;
  el("fridayRitualTa").textContent = ritualTa;
  el("fridayRitualRetry").textContent = ritualRetry;
  el("fridayHeroMessage").textContent = taViewer
    ? "Acesso seguro: execute apenas a fila de TA liberada pelo dono da base."
    : remaining
      ? `Faltam ${remaining} OI(s) para sua meta. Comece pelas recomendações.`
      : "Meta atingida. Continue construindo a próxima semana.";
  let queueSummary = `${queue.length} contato(s) nesta visão. Recomendações e pessoas nunca abordadas aparecem primeiro.`;
  if (fridaySessionActive) queueSummary = `${queue.length} contato(s) restantes neste bloco. A fila fica oculta para manter foco.`;
  else if (taViewer) queueSummary = `${queue.length} contato(s) liberado(s) para abordagem.`;
  else if (fridaySosActive && fridaySosFocus === "top10") queueSummary = `${queue.length} contato(s) no SOS Top 10. Ligue sem negociar com a própria cabeça: um por vez.`;
  else if (fridaySosActive && fridaySosFocus === "retake") queueSummary = `${queue.length} contato(s) no SOS Retentativa. Foque em reabrir conversa e oferecer dois horários.`;
  else if (fridayReferrerFilter !== "all") queueSummary = `${queue.length} contato(s) na sequência de ${fridayReferrerFilter}. Use o mesmo gancho e registre cada resultado.`;
  if (fridayQueueSearch.trim()) queueSummary = `Busca ativa: ${queue.length} contato(s) encontrados por nome, recomendante ou memória do TA.`;
  if (fridayTemperatureFilter !== "all") queueSummary += ` Temperatura filtrada: ${fridayTemperatureFilter}.`;
  if (!fridaySessionActive && queue.length) {
    queueSummary += ` Exibindo ${displayQueue.length} de ${queue.length}. Ordem travada até você registrar um resultado ou atualizar a fila.`;
  }
  el("fridayQueueSummary").textContent = queueSummary;
  renderFridayQueueControls(baseQueue, queue, queueControls.counts, calls);
  renderFridayActiveStrip(queue);
  if (!taViewer) renderNameEngine(week, goal, oiScheduled);
  renderTaPortfolioPanel();
  renderFridaySosPanel(portfolioQueue, remaining);
  renderFridayReferrerSequence(portfolioQueue, fullReferrerQueue);

  el("fridayCallList").innerHTML = displayQueue.length ? displayQueue.map((client, index) => {
    const digits = normalizeWhatsapp(client.whatsapp || client.phone);
    const number = digits.startsWith("55") ? digits : `55${digits}`;
    const whatsapp = digits.length >= 10 ? `https://wa.me/${number}?text=${encodeURIComponent(assistantWhatsappMessage(client, "short"))}` : "";
    const missedWhatsapp = digits.length >= 10 ? `https://wa.me/${number}?text=${encodeURIComponent(fridayMissedCallMessage(client))}` : "";
    const titanScore = fridayTitanScore(client);
    const referrerLine = client.referrer
      ? `<p class="friday-referrer-line">${escapeHtml(client.referrer)}</p>`
      : `<p class="friday-referrer-line">Sem indicador registrado</p>`;
    const memoryLine = fridayCardMemory(client);
    const taDossier = fridayTaDossier(client);
    const pinned = client.id === fridayPinnedClientId;
    return `<article class="friday-call-card">
      <span class="friday-position">${index + 1}</span>
      <div class="friday-contact">
        <span class="friday-reason">${escapeHtml(fridayCallLabel(client))}</span>
        <h3>${escapeHtml(client.name)}</h3>
        ${referrerLine}
        ${fridaySessionActive ? fridaySessionFactsHtml(taDossier) : ""}
        <p class="friday-phone-line">${escapeHtml(client.whatsapp || client.phone || "Sem WhatsApp")}</p>
        <p class="friday-memory-line">${escapeHtml(memoryLine)}</p>
        ${fridayTaDossierHtml(client, taDossier, index, fridaySessionActive)}
        ${taViewer
          ? `<div class="friday-score-line"><small>Modo TA: ligar ou chamar no WhatsApp para tentar agendar OI.</small></div>`
          : `<div class="friday-score-line">
              <span>Score ${titanScore}</span>
              <span>${escapeHtml(taPortfolioLabelForClient(client))}</span>
              ${pinned ? `<span class="friday-focus-chip">Em foco</span>` : ""}
              <small>${escapeHtml(fridayTitanReason(client))}</small>
            </div>`}
      </div>
      <div class="friday-primary-actions">
        ${whatsapp ? `<a class="friday-whatsapp-button" href="${whatsapp}" target="_blank" rel="noopener">WhatsApp</a>` : ""}
        ${taViewer ? "" : `<button class="friday-open-button" data-friday-local="${client.id}">Ver roteiro</button>
        <button class="friday-open-button" data-copy-ta-missed="${client.id}">Copiar não atendeu</button>
        <button class="friday-open-button" data-friday-pin="${client.id}">${pinned ? "Soltar foco" : "Manter na tela"}</button>
        <button class="friday-open-button" data-friday-open="${client.id}">Abrir contato</button>`}
      </div>
      <div class="friday-ai-script hidden" id="friday-ai-${client.id}"></div>
      ${taViewer ? "" : `<div class="friday-result-actions">
        <small>Resultado da ligação</small>
        <button data-friday-result="no-answer" data-client-id="${client.id}">Não atendeu</button>
        <button data-friday-result="talked" data-client-id="${client.id}">Atendeu, sem agenda</button>
        <button class="success" data-friday-result="scheduled" data-client-id="${client.id}">OI marcada</button>
        ${missedWhatsapp ? `<a class="friday-result-whatsapp" href="${missedWhatsapp}" target="_blank" rel="noopener">WhatsApp não atendeu</a>` : ""}
        <button class="muted" data-friday-decline-toggle="${client.id}">Sem interesse / motivo</button>
      </div>
      ${fridayDeclinePanelHtml(client)}`}
    </article>`;
  }).join("") : `<div class="empty-state friday-empty"><h3>Fila concluída</h3><p>Todos os contatos elegíveis foram trabalhados ou já possuem OI.</p></div>`;
  const loadMoreButton = el("fridayLoadMore");
  if (loadMoreButton) {
    const hasMore = !fridaySessionActive && displayQueue.length < queue.length;
    loadMoreButton.classList.toggle("hidden", !hasMore);
    loadMoreButton.textContent = hasMore ? `Carregar mais (${queue.length - displayQueue.length} restantes)` : "Carregar mais contatos";
  }

  renderFridayRecoveryPanel();

  document.querySelectorAll("[data-friday-open]").forEach(button => button.addEventListener("click", () => openClientModal(button.dataset.fridayOpen)));
  document.querySelectorAll("[data-friday-local]").forEach(button => button.addEventListener("click", () => {
    fridayPinnedClientId = button.dataset.fridayLocal;
    showFridayLocalScript(button.dataset.fridayLocal, button);
  }));
  document.querySelectorAll("[data-friday-pin]").forEach(button => button.addEventListener("click", () => {
    fridayPinnedClientId = fridayPinnedClientId === button.dataset.fridayPin ? "" : button.dataset.fridayPin;
    renderFridayMode();
  }));
  document.querySelectorAll("[data-friday-decline-toggle]").forEach(button => button.addEventListener("click", () => {
    fridayPinnedClientId = button.dataset.fridayDeclineToggle;
    toggleFridayDeclinePanel(button.dataset.fridayDeclineToggle, true);
  }));
  document.querySelectorAll("[data-friday-decline-cancel]").forEach(button => button.addEventListener("click", () => {
    toggleFridayDeclinePanel(button.dataset.fridayDeclineCancel, false);
  }));
  document.querySelectorAll("[data-friday-decline-save]").forEach(button => button.addEventListener("click", () => {
    registerFridayDeclineFromPanel(button.dataset.fridayDeclineSave);
  }));
  document.querySelectorAll("[data-friday-decline-quick]").forEach(button => button.addEventListener("click", () => {
    registerFridayQuickDecline(button.dataset.fridayDeclineQuick, button.dataset.fridayDeclineQuickReason);
  }));
  document.querySelectorAll("[data-copy-ta-opener]").forEach(button => button.addEventListener("click", event => {
    event.preventDefault();
    event.stopPropagation();
    copyFridayTaText(button.dataset.copyTaOpener, "opener");
  }));
  document.querySelectorAll("[data-copy-ta-whatsapp]").forEach(button => button.addEventListener("click", event => {
    event.preventDefault();
    event.stopPropagation();
    copyFridayTaText(button.dataset.copyTaWhatsapp, "whatsapp");
  }));
  document.querySelectorAll("[data-copy-ta-missed]").forEach(button => button.addEventListener("click", event => {
    event.preventDefault();
    event.stopPropagation();
    copyFridayTaText(button.dataset.copyTaMissed, "missed");
  }));
  document.querySelectorAll("[data-friday-result]").forEach(button => button.addEventListener("click", () => {
    registerFridayResult(button.dataset.clientId, button.dataset.fridayResult);
  }));
}

async function copyFridayTaText(clientId, kind) {
  const client = state.clients.find(item => item.id === clientId);
  if (!client) return;
  const dossier = fridayTaDossier(client);
  const text = kind === "missed" ? fridayMissedCallMessage(client) : kind === "whatsapp" ? dossier.whatsapp : dossier.opening;
  try {
    await copyTextToClipboard(text);
    showToast(kind === "missed" ? "Mensagem de não atendimento copiada" : kind === "whatsapp" ? "WhatsApp da abordagem copiado" : "Abertura da ligação copiada");
  } catch (_) {
    showToast("Não foi possível copiar o texto", true);
  }
}

function showFridayLocalScript(clientId, button) {
  if (isCloudReadOnly()) {
    showToast("Modo operacional: roteiro interno protegido.", true);
    return;
  }
  const client = state.clients.find(item => item.id === clientId);
  const output = el(`friday-ai-${clientId}`);
  if (!client || !output) return;
  output.textContent = fridayLocalScript(client);
  output.classList.remove("hidden");
  button.textContent = "Roteiro aberto";
}

function toggleFridayDeclinePanel(clientId, open) {
  const panel = el(`friday-decline-${clientId}`);
  if (!panel) return;
  panel.classList.toggle("hidden", !open);
  if (open) panel.querySelector("select, input")?.focus();
}

function registerFridayDeclineFromPanel(clientId) {
  if (isCloudReadOnly()) {
    showToast("Modo operacional: registro bloqueado.", true);
    return;
  }
  const client = state.clients.find(item => item.id === clientId);
  if (!client) return;
  const reason = document.querySelector(`[data-friday-decline-reason="${clientId}"]`)?.value || "";
  const action = document.querySelector(`[data-friday-decline-action="${clientId}"]`)?.value || "remove";
  const note = document.querySelector(`[data-friday-decline-note="${clientId}"]`)?.value || "";
  if (!reason) {
    showToast("Escolha o motivo principal", true);
    return;
  }
  registerFridayNoInterest(client, offsetDate(0), { reason, action, note });
  saveState();
  renderAll();
}

function finishFridaySessionStep(message = "Registrado. Avançando para o próximo contato...") {
  if (!fridaySessionActive) {
    renderAll();
    return;
  }
  showToast(message);
  window.setTimeout(() => {
    renderAll();
    focusFridayQueue();
  }, 1500);
}

function registerFridayQuickDecline(clientId, reason = "") {
  if (isCloudReadOnly()) {
    showToast("Modo operacional: registro bloqueado.", true);
    return;
  }
  const client = state.clients.find(item => item.id === clientId);
  if (!client) return;
  const cleanReason = compactText(reason, "Sem motivo informado");
  const action = cleanReason === "Pediu retorno futuro" ? "future" : "remove";
  recordFridaySessionResult(client.id, "declined");
  registerFridayNoInterest(client, offsetDate(0), { reason: cleanReason, action });
  saveState();
  finishFridaySessionStep("Motivo salvo. Próximo contato em instantes...");
}

function registerFridayNoInterest(client, today, details = {}) {
  const cleanReason = compactText(details.reason, "Sem motivo informado");
  const note = compactText(details.note, "");
  const requestedAction = details.action === "c2" ? "future" : (details.action || "remove");
  const action = cleanReason === "Não quer contato" ? "remove" : requestedAction;
  const future = action === "future";
  client.taNoInterestReason = cleanReason;
  client.taNoInterestDate = today;
  client.priority = "Baixa";
  client.lastInteraction = today;
  client.taDate ||= today;
  if (client.stage === "Base") client.stage = "TA";
  const decisionText = future
    ? "Acompanhar no futuro, sem insistir agora."
    : "Retirado da fila de TA.";
  client.notes = [
    compactText(client.notes),
    `[TA sem interesse em ${formatDate(today)}] Motivo: ${cleanReason}.${note ? ` Obs.: ${note}.` : ""} ${decisionText}`
  ].filter(Boolean).join("\n");
  if (future) {
    client.doNotContact = false;
    client.taPausedUntil = offsetDate(60);
    client.followUp = client.taPausedUntil;
    client.nextAction = `Retomar com cuidado: ${cleanReason}`;
    client.completed = false;
    logClientEvent(client.id, "Sexta - sem interesse com retomada", `Motivo: ${cleanReason}.${note ? ` Obs.: ${note}.` : ""} Retomar em ${formatDate(client.taPausedUntil)}.`, "Ligacao");
    showToast("Motivo salvo para acompanhamento futuro");
  } else {
    client.doNotContact = true;
    client.taPausedUntil = "";
    client.followUp = "";
    client.nextAction = `Sem interesse: ${cleanReason}`;
    client.completed = true;
    logClientEvent(client.id, "Sexta - sem interesse", `Motivo: ${cleanReason}.${note ? ` Obs.: ${note}.` : ""} Contato retirado da fila de ligações.`, "Ligacao");
    showToast("Motivo salvo e contato retirado da fila");
  }
  if (fridayPinnedClientId === client.id) fridayPinnedClientId = "";
  removeFromFridayStableQueue(client.id);
}

function registerFridayResult(clientId, result) {
  if (isCloudReadOnly()) {
    showToast("Modo operacional: registro bloqueado.", true);
    return;
  }
  const client = state.clients.find(item => item.id === clientId);
  if (!client) return;
  const today = offsetDate(0);
  client.taDate ||= today;
  client.lastInteraction = today;
  if (client.stage === "Base") client.stage = "TA";
  client.completed = false;
  fridayPinnedClientId = client.id;
  recordFridaySessionResult(client.id, result);

  if (result === "scheduled") {
    fridaySchedulingClientId = client.id;
    if (fridayPinnedClientId === client.id) fridayPinnedClientId = "";
    removeFromFridayStableQueue(client.id);
    saveState();
    openMeetingModal("OI", client.id);
    showToast("Informe data e horario da OI");
    return;
  }
  if (result === "no-answer") {
    client.followUp = nextBusinessDate();
    client.nextAction = "Tentar nova ligação e enviar WhatsApp de não atendimento";
    logClientEvent(client.id, "Sexta - não atendeu", "Nova tentativa criada para o próximo dia útil.", "Ligacao");
    const link = whatsappLink(client, true);
    if (link) window.open(link, "_blank", "noopener");
    fridayPinnedClientId = "";
    removeFromFridayStableQueue(client.id);
    showToast("Tentativa registrada e WhatsApp preparado");
  }
  if (result === "talked") {
    client.followUp = nextBusinessDate();
    client.nextAction = "Retomar conversa e oferecer dois horários para OI";
    logClientEvent(client.id, "Sexta - atendeu sem agenda", "Contato realizado, mas OI ainda não agendada.", "Ligacao");
    fridayPinnedClientId = "";
    removeFromFridayStableQueue(client.id);
    showToast("Conversa registrada com novo follow-up");
  }
  saveState();
  if (fridaySessionActive && ["no-answer", "talked"].includes(result)) {
    finishFridaySessionStep();
    return;
  }
  renderAll();
}

function renderDailyCommand() {
  if (!el("weeklyPaceScore")) return;
  const today = new Date();
  const week = selectedWeek();
  const goal = weeklyGoal(week);
  const scheduled = state.clients.filter(client => inSelectedWeek(client.oiScheduledDate, week)).length;
  const pace = Math.min(Math.round((scheduled / goal) * 100), 100);
  const commands = dailyCommandClients();
  const isFriday = today.getDay() === 5;
  const total = executionCallClients().length + executionPlanClients().length + executionC2Clients().length;
  if (el("commandWeeklyOiScheduled")) el("commandWeeklyOiScheduled").textContent = scheduled;
  if (el("weeklyOiGoalScore")) el("weeklyOiGoalScore").textContent = goal;
  el("weeklyPaceScore").textContent = `${pace}%`;
  el("dailyCommandTitle").textContent = isFriday ? "Sexta de conquista: transforme contatos em OIs" : "Seu plano comercial de hoje";
  el("dailyCommandSummary").textContent = isFriday
    ? `Priorize recomendações e contatos ainda sem abordagem. Faltam ${Math.max(goal - scheduled, 0)} OI(s) para a meta.`
    : total ? `${total} movimento(s) comerciais para proteger hoje. Comece pela primeira fila e registre o avanço no Lançar.` : "Nenhuma frente crítica agora. Use o tempo para pedir RECs, fazer TA e alimentar a próxima semana.";
  if (!el("dailyCommandList")) return;
  el("dailyCommandList").innerHTML = commands.length ? commands.map((client, index) => {
    const recommendation = assistantAction(client);
    const message = assistantWhatsappMessage(client, "short");
    const digits = normalizeWhatsapp(client.whatsapp || client.phone);
    const number = digits.startsWith("55") ? digits : `55${digits}`;
    const whatsapp = digits.length >= 10 ? `https://wa.me/${number}?text=${encodeURIComponent(message)}` : "";
    return `<article class="command-item">
      <span class="command-number">${index + 1}</span>
      <div>
        <strong>${escapeHtml(client.name)}</strong>
        <p>${escapeHtml(recommendation.title)} · ${escapeHtml(recommendation.action)}</p>
        <div class="command-actions">
          <button data-command-client="${client.id}">Abrir contato</button>
          <button data-command-assistant="${client.id}">Preparar conversa</button>
          ${whatsapp ? `<a href="${whatsapp}" target="_blank" rel="noopener" data-command-whatsapp="${client.id}">WhatsApp</a>` : ""}
        </div>
      </div>
    </article>`;
  }).join("") : `<div class="mini-empty">Nenhuma prioridade disponível.</div>`;

  document.querySelectorAll("[data-command-client]").forEach(button => button.addEventListener("click", () => openClientModal(button.dataset.commandClient)));
  document.querySelectorAll("[data-command-assistant]").forEach(button => button.addEventListener("click", () => {
    assistantClientId = button.dataset.commandAssistant;
    assistantMode = "next";
    switchView("assistant");
  }));
  document.querySelectorAll("[data-command-whatsapp]").forEach(link => link.addEventListener("click", () => {
    const client = state.clients.find(item => item.id === link.dataset.commandWhatsapp);
    if (!client) return;
    client.lastInteraction = offsetDate(0);
    client.taDate ||= offsetDate(0);
    if (client.stage === "Base") client.stage = "TA";
    client.followUp = nextBusinessDate();
    client.nextAction = "Aguardar retorno no WhatsApp";
    client.completed = false;
    logClientEvent(client.id, "Mensagem preparada no WhatsApp", "Contato iniciado pelo Comando do Dia.", "WhatsApp");
    saveState();
  }));
}

function renderAssistantResult() {
  const client = state.clients.find(item => item.id === assistantClientId);
  const result = el("assistantResult");
  const whatsapp = el("assistantWhatsapp");
  const brief = el("assistantClientBrief");
  const memory = el("assistantMemory");
  const playbook = el("assistantPlaybook");
  if (assistantChatClientId && assistantChatClientId !== assistantClientId) resetAssistantChat();
  if (!client) {
    result.textContent = "Cadastre ou selecione um contato para iniciar a análise.";
    if (brief) brief.innerHTML = "";
    if (memory) memory.innerHTML = "";
    if (playbook) playbook.innerHTML = "";
    whatsapp.classList.add("hidden");
    return;
  }

  const risk = assistantRisk(client);
  const opportunity = sitplanOpportunityScore(client);
  const days = assistantDays(client);
  if (brief) {
    brief.innerHTML = `
      <article>
        <small>Etapa</small>
        <strong>${escapeHtml(client.stage || "Base")}</strong>
      </article>
      <article>
        <small>Score</small>
        <strong>${opportunity.score}/100</strong>
      </article>
      <article>
        <small>Risco</small>
        <strong>${escapeHtml(risk.level)}</strong>
      </article>
      <article>
        <small>Sem contato</small>
        <strong>${days}d</strong>
      </article>
      <article>
        <small>Origem</small>
        <strong>${escapeHtml(client.referrer || client.source || "Base")}</strong>
      </article>`;
  }

  if (memory) {
    const context = assistantKnownContext(client);
    const gaps = assistantInfoGaps(client);
    const questions = assistantDiscoveryQuestions(client);
    memory.innerHTML = `
      <div class="assistant-memory-main">
        <small>LEIA ANTES DE LIGAR</small>
        <strong>${escapeHtml(client.referrer ? `Recomendação de ${client.referrer}` : client.source || "Contato da base")}</strong>
        <p>${escapeHtml(context[0] || "Ainda há pouco contexto registrado. Comece validando quem indicou, momento de vida e objetivo da conversa.")}</p>
      </div>
      <div class="assistant-memory-grid">
        <article>
          <small>O que já sei</small>
          <ul>${context.length ? context.map(item => `<li>${escapeHtml(item)}</li>`).join("") : "<li>Nenhum detalhe registrado ainda.</li>"}</ul>
        </article>
        <article>
          <small>O que falta descobrir</small>
          <ul>${gaps.length ? gaps.map(item => `<li>${escapeHtml(item)}</li>`).join("") : "<li>Contexto minimo suficiente para agir.</li>"}</ul>
        </article>
        <article>
          <small>Perguntas boas</small>
          <ul>${questions.map(item => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
        </article>
      </div>`;
  }

  if (playbook) {
    playbook.innerHTML = assistantPlaybook(client).map(item => `
      <article>
        <small>${escapeHtml(item.label)}</small>
        <p>${escapeHtml(item.value)}</p>
      </article>`).join("");
  }

  el("assistantTone").classList.toggle("hidden", assistantMode !== "message");
  if (assistantMode === "message") {
    assistantMessage = assistantWhatsappMessage(client, assistantTone);
    result.textContent = assistantMessage;
    const digits = normalizeWhatsapp(client.whatsapp || client.phone);
    if (digits.length >= 10) {
      const number = digits.startsWith("55") ? digits : `55${digits}`;
      whatsapp.href = `https://wa.me/${number}?text=${encodeURIComponent(assistantMessage)}`;
      whatsapp.classList.remove("hidden");
    } else {
      whatsapp.classList.add("hidden");
    }
    return;
  }

  whatsapp.classList.add("hidden");
  if (assistantMode === "risk") {
    result.textContent = `RISCO ${risk.level.toUpperCase()}\n\nMotivo principal: ${risk.reason}.\n\nSinal observado: ${assistantDays(client)} dia(s) desde a última interação registrada.\n\nAção nas próximas 24 horas: ${assistantAction(client).action}`;
    return;
  }

  const recommendation = assistantAction(client);
  result.textContent = `SCORE SITPLAN ${opportunity.score}/100\nSinais: ${opportunity.reasons.join(", ") || "cadência normal"}\n\n${recommendation.title.toUpperCase()}\n\nAção agora: ${recommendation.action}\n\nO que falar: ${recommendation.talk}\n\nMelhor canal: ${recommendation.channel}`;
}

function setAssistantProvider(text, connected = false) {
  if (!el("assistantProvider")) return;
  el("assistantProvider").textContent = text;
  el("assistantProvider").classList.toggle("connected", connected);
}

function assistantProviderLabel(answer) {
  const usage = answer?.usage;
  return usage?.limit
    ? `${answer.model} · ${usage.requests}/${usage.limit} análises no mês`
    : `${answer.model} conectado`;
}

function resetAssistantChat() {
  assistantChatHistory = [];
  assistantChatClientId = "";
  if (el("assistantChatMessages")) el("assistantChatMessages").innerHTML = "";
  el("assistantChat")?.classList.add("hidden");
}

function appendAssistantChat(text, role) {
  const messages = el("assistantChatMessages");
  if (!messages) return null;
  const bubble = document.createElement("div");
  bubble.className = `assistant-chat-message ${role === "user" ? "user" : "ai"}`;
  bubble.textContent = text;
  messages.appendChild(bubble);
  messages.scrollTop = messages.scrollHeight;
  return bubble;
}

function updateAssistantWhatsapp(client, message) {
  const whatsapp = el("assistantWhatsapp");
  const digits = normalizeWhatsapp(client.whatsapp || client.phone);
  if (assistantMode !== "message" || digits.length < 10) {
    whatsapp.classList.add("hidden");
    return;
  }
  const number = digits.startsWith("55") ? digits : `55${digits}`;
  whatsapp.href = `https://wa.me/${number}?text=${encodeURIComponent(message)}`;
  whatsapp.classList.remove("hidden");
}

function refreshLocalAssistant() {
  renderAssistantResult();
  el("assistantChat")?.classList.remove("hidden");
  setAssistantProvider("Modo econômico · sem custo");
  showToast("Análise local atualizada sem usar a API");
}

function assistantClientBriefing(client, extraQuestion = "") {
  const week = selectedWeek();
  const board = weeklyBoardData(week);
  const review = dailyReviewData();
  const recommendation = assistantAction(client);
  const risk = assistantRisk(client);
  const opportunity = sitplanOpportunityScore(client);
  const knownContext = assistantKnownContext(client).map(item => `- ${item}`).join("\n") || "- Pouco contexto registrado.";
  const gaps = assistantInfoGaps(client).map(item => `- ${item}`).join("\n") || "- Sem lacunas críticas para a próxima ação.";
  const questions = assistantDiscoveryQuestions(client).map(item => `- ${item}`).join("\n");
  const history = state.activities
    .filter(activity => activity.clientId === client.id)
    .sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")))
    .slice(0, 6)
    .map(activity => `- ${formatDate(activity.date)} | ${activity.title}: ${activity.notes || "sem observação"}`)
    .join("\n") || "- Sem histórico registrado.";

  return [
    "Você é meu agente comercial especialista em Prudential/LifePlanner.",
    "Use as informações abaixo para me dar uma orientação prática, curta e acionável.",
    "",
    "CONTEXTO DO CLIENTE",
    `Nome: ${client.name}`,
    `Etapa: ${client.stage}`,
    `Temperatura: ${client.temperature}`,
    `Prioridade: ${client.priority}`,
    `Origem: ${client.source || "não informada"}`,
    `Indicador/recomendante: ${client.referrer || "não informado"}`,
    `Profissão: ${client.profession || "não informada"}`,
    `Cidade: ${client.city || "não informada"}`,
    `Última interação: ${client.lastInteraction ? formatDate(client.lastInteraction) : "não registrada"}`,
    `Próximo follow-up: ${client.followUp ? formatDate(client.followUp) : "não definido"}`,
    `OI: ${client.oiDone ? "realizada" : client.oiScheduledDate ? `agendada para ${formatMeetingDate(client.oiScheduledDate, client.oiTime)}` : "não agendada"}`,
    `PC: ${client.pcDone ? "realizada" : client.pcScheduledDate ? `agendada para ${formatMeetingDate(client.pcScheduledDate, client.pcTime)}` : "não agendada"}`,
    `Resultado: ${client.outcome || "em aberto"}`,
    `Status do plano: ${client.planStatus || "não informado"}`,
    `Objeção registrada: ${client.c2Objection || client.objection || "não registrada"}`,
    `Observações: ${client.notes || "sem observações"}`,
    "",
    "MEMÓRIA DA RECOMENDAÇÃO",
    knownContext,
    "",
    "LACUNAS A DESCOBRIR",
    gaps,
    "",
    "PERGUNTAS BOAS PARA A CONVERSA",
    questions,
    "",
    "LEITURA DO SITPLAN",
    `Score de oportunidade: ${opportunity.score}/100`,
    `Sinais: ${opportunity.reasons.join(", ") || "cadência normal"}`,
    `Risco: ${risk.level} (${risk.reason})`,
    `Ação recomendada agora: ${recommendation.action}`,
    `O que falar: ${recommendation.talk}`,
    `Melhor canal: ${recommendation.channel}`,
    "",
    "HISTÓRICO RECENTE",
    history,
    "",
    "RITMO DA SEMANA",
    `Meta de OIs: ${board.contract.oi}`,
    `OIs agendadas: ${board.oiScheduled}`,
    `OIs realizadas: ${board.oiDone}`,
    `PCs realizadas: ${board.pcDone}`,
    `N/vendas: ${board.sales}`,
    `C2: ${board.c2}`,
    `RECs recebidas hoje: ${review.recomendacoes}`,
    `Planos pendentes: ${review.planosPendentes}`,
    `C2 vencidos: ${review.c2Vencidos}`,
    "",
    "O QUE EU QUERO DE VOCÊ",
    extraQuestion || "Me dê a melhor próxima ação para este contato, uma mensagem curta de WhatsApp e uma estratégia de fechamento/follow-up.",
    "",
    "Responda em português do Brasil, direto ao ponto, com roteiro prático e sem prometer resultado."
  ].join("\n");
}

async function copyTextToClipboard(text) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }
  const area = document.createElement("textarea");
  area.value = text;
  area.setAttribute("readonly", "");
  area.style.position = "fixed";
  area.style.left = "-9999px";
  document.body.appendChild(area);
  area.select();
  document.execCommand("copy");
  area.remove();
}

async function copyAssistantBriefing(extraQuestion = "") {
  const client = state.clients.find(item => item.id === assistantClientId);
  if (!client) {
    showToast("Selecione um contato para copiar o briefing", true);
    return;
  }
  const briefing = assistantClientBriefing(client, extraQuestion);
  await copyTextToClipboard(briefing);
  el("assistantChat")?.classList.remove("hidden");
  appendAssistantChat("Briefing copiado. Agora cole no seu agente do ChatGPT.", "assistant");
  setAssistantProvider("Briefing ChatGPT · sem API");
  showToast("Briefing copiado para o ChatGPT");
}

function renderAssistant() {
  if (!el("assistantClient")) return;
  const previous = assistantClientId || el("assistantClient").value;
  const sorted = [...state.clients].sort((a, b) => assistantPriorityScore(b) - assistantPriorityScore(a));
  el("assistantAttentionCount").textContent = sorted.filter(client => assistantPriorityScore(client) >= 25).length;
  el("assistantClient").innerHTML = sorted.length
    ? sorted.map(client => `<option value="${client.id}">${escapeHtml(client.name)} · ${escapeHtml(client.stage)}</option>`).join("")
    : `<option value="">Nenhum contato cadastrado</option>`;
  assistantClientId = sorted.some(client => client.id === previous) ? previous : (sorted[0]?.id || "");
  el("assistantClient").value = assistantClientId;
  el("assistantEditContact").disabled = !assistantClientId;

  el("assistantPriorityList").innerHTML = sorted.length
    ? sorted.slice(0, 6).map((client, index) => {
      const risk = assistantRisk(client);
      return `<button class="assistant-priority" data-assistant-client="${client.id}">
        <span class="assistant-priority-index">${index + 1}</span>
        <span><strong>${escapeHtml(client.name)}</strong><small>${escapeHtml(assistantAction(client).title)} · ${assistantDays(client)}d sem contato</small></span>
        <em class="assistant-risk">${risk.level}</em>
      </button>`;
    }).join("")
    : `<div class="mini-empty">Cadastre contatos para receber prioridades.</div>`;
  document.querySelectorAll("[data-assistant-client]").forEach(button => button.addEventListener("click", () => {
    assistantClientId = button.dataset.assistantClient;
    el("assistantClient").value = assistantClientId;
    renderAssistantResult();
  }));
  renderAssistantResult();
}

function nextBusinessDate() {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  date.setDate(date.getDate() + 1);
  if (date.getDay() === 6) date.setDate(date.getDate() + 2);
  if (date.getDay() === 0) date.setDate(date.getDate() + 1);
  return dateKey(date);
}

function businessDateAfterToday(days = 1) {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  let remaining = Math.max(1, Number(days) || 1);
  while (remaining > 0) {
    date.setDate(date.getDate() + 1);
    if (date.getDay() !== 0 && date.getDay() !== 6) remaining -= 1;
  }
  return dateKey(date);
}

function quickInputQueueClients() {
  const today = offsetDate(0);
  return [...state.clients]
    .filter(client => client.outcome !== "N")
    .map(client => {
      let reason = "";
      let score = 0;
      if (!client.nextAction || !client.followUp) {
        reason = "Sem próximo passo definido";
        score += 80;
      } else if (client.followUp < today && !client.completed) {
        reason = "Follow-up atrasado";
        score += 70;
      } else if (client.oiDone && client.planStatus !== "Pronto") {
        reason = "Plano precisa ser montado";
        score += 65;
      } else if (client.oiDone && !client.pcScheduledDate && !client.pcDone) {
        reason = "OI realizada sem PC marcada";
        score += 60;
      } else if (client.pcDone && client.outcome !== "N") {
        reason = "PC realizada sem decisão final";
        score += 55;
      } else if (["Base", "TA"].includes(client.stage) && !client.oiScheduledDate) {
        reason = "Pode virar OI";
        score += String(client.source || "").includes("Indica") || client.referrer ? 50 : 35;
      }
      score += client.priority === "Alta" ? 20 : client.priority === "Média" || client.priority === "Media" ? 10 : 0;
      return { client, reason, score };
    })
    .filter(item => item.reason)
    .sort((a, b) => b.score - a.score)
    .slice(0, 8);
}

function setQuickInputClient(clientId) {
  if (!el("quickInputClient")) return;
  el("quickInputClient").value = clientId || "";
  updateQuickInputPreview();
}

function updateQuickInputPreview() {
  const client = state.clients.find(item => item.id === el("quickInputClient")?.value);
  if (!el("quickInputPreview")) return;
  if (!client) {
    el("quickInputPreview").textContent = "Escolha um contato para ver o resumo.";
    return;
  }
  const next = client.followUp ? formatDate(client.followUp) : "sem data";
  const origin = client.referrer ? `Indicado por ${client.referrer}` : (client.source || "Origem não informada");
  const relationship = client.referrerRelationship || "Relação ainda não registrada";
  const context = client.referralContext || compactText(client.notes, "Ainda não há memória comercial salva.").slice(0, 240);
  el("quickInputPreview").innerHTML = `
    <div class="record-context-heading">
      <div><small>CONTATO SELECIONADO</small><strong>${escapeHtml(client.name)}</strong></div>
      <span>${escapeHtml(client.stage)}</span>
    </div>
    <div class="record-context-grid">
      <article class="record-origin-memory"><small>ORIGEM DE CONFIANÇA</small><strong>${escapeHtml(origin)}</strong><span>${escapeHtml(relationship)}</span></article>
      <article><small>MEMÓRIA</small><p>${escapeHtml(context)}</p></article>
      <article><small>PRÓXIMO PASSO</small><strong>${escapeHtml(client.nextAction || "Definir próxima ação")}</strong><span>${escapeHtml(next)}</span></article>
    </div>
    <button class="text-button quick-referral-button" type="button" data-create-recommendation="${client.id}">Cadastrar uma indicação deste contato</button>
  `;
  el("quickInputPreview").querySelector("[data-create-recommendation]")?.addEventListener("click", event => {
    openRecommendedContactFrom(event.currentTarget.dataset.createRecommendation);
  });
}

function updateQuickInputFields() {
  const action = el("quickInputAction")?.value || "ta_call";
  const visibleByField = {
    "date": true,
    "meeting-time": ["oi_schedule", "pc_schedule"].includes(action),
    "meeting-link": ["oi_schedule", "pc_schedule"].includes(action),
    "recs": ["recs", "oi_done"].includes(action),
    "followup": ["ta_call", "c2", "note"].includes(action),
    "objection": action === "c2"
  };
  document.querySelectorAll("[data-quick-field]").forEach(field => {
    field.classList.toggle("quick-field-hidden", !visibleByField[field.dataset.quickField]);
  });
  const notesLabels = {
    ta_call: "Como foi a abordagem e qual é o próximo passo?",
    oi_schedule: "Contexto importante para preparar a OI",
    oi_done: "O que você descobriu sobre esta pessoa?",
    plan_ready: "Ponto central do plano preparado",
    pc_schedule: "Contexto importante para preparar o PC",
    n: "Resumo do fechamento",
    c2: "Motivo real e estratégia de retomada",
    recs: "Nomes, contexto ou observação sobre as recomendações",
    note: "Memória comercial e próximo passo"
  };
  if (el("quickInputNotesLabel")) el("quickInputNotesLabel").textContent = notesLabels[action] || "Observação / próximo passo";
}

function resetQuickInputForm(keepClient = true) {
  const clientId = keepClient ? el("quickInputClient")?.value : "";
  el("quickInputForm")?.reset();
  el("quickInputDate").value = offsetDate(0);
  el("quickInputTime").value = "19:00";
  el("quickInputRecs").value = "0";
  if (clientId) el("quickInputClient").value = clientId;
  updateQuickInputPreview();
  updateQuickInputFields();
}

function renderQuickInput() {
  if (!el("quickInputClient")) return;
  const selectedId = el("quickInputClient").value;
  const sorted = [...state.clients].sort((a, b) => a.name.localeCompare(b.name, "pt-BR"));
  el("quickInputClient").innerHTML = sorted.length
    ? sorted.map(client => `<option value="${client.id}">${escapeHtml(client.name)} - ${escapeHtml(client.stage)}</option>`).join("")
    : `<option value="">Nenhum contato cadastrado</option>`;
  if (selectedId && sorted.some(client => client.id === selectedId)) el("quickInputClient").value = selectedId;
  if (!el("quickInputDate").value) el("quickInputDate").value = offsetDate(0);
  updateQuickInputPreview();
  updateQuickInputFields();

  const queue = quickInputQueueClients();
  el("quickInputQueue").innerHTML = queue.length ? queue.map(({ client, reason, score }) => `
    <article class="quick-input-queue-item">
      <div>
        <strong>${escapeHtml(client.name)}</strong>
        <small>${escapeHtml(reason)} - ${escapeHtml(client.stage)}</small>
      </div>
      <button class="text-button" data-quick-pick="${client.id}">Lancar</button>
    </article>
  `).join("") : `<div class="mini-empty">Nenhum lancamento urgente agora.</div>`;
  document.querySelectorAll("[data-quick-pick]").forEach(button => button.addEventListener("click", () => setQuickInputClient(button.dataset.quickPick)));
}

function submitQuickInput(event) {
  event.preventDefault();
  const client = state.clients.find(item => item.id === el("quickInputClient").value);
  if (!client) {
    showToast("Escolha um contato para lancar", true);
    return;
  }
  const action = el("quickInputAction").value;
  const date = el("quickInputDate").value || offsetDate(0);
  const time = el("quickInputTime").value;
  const recs = Number(el("quickInputRecs").value || 0);
  const followUp = el("quickInputFollowUp").value;
  const objection = el("quickInputObjection").value;
  const link = el("quickInputMeetingLink").value.trim();
  const notes = el("quickInputNotes").value.trim();
  const noteText = notes || "Lancamento feito pela Central de Input.";

  const blockQuickInput = message => {
    showToast(message, true);
    return true;
  };
  if ((action === "oi_schedule" || action === "pc_schedule") && !time) {
    if (blockQuickInput("Reunião agendada precisa de horário.")) return;
  }
  if (action === "oi_done" && !notes && !hasPostOiDiagnosis(client)) {
    if (blockQuickInput("OI realizada precisa de diagnóstico ou observações.")) return;
  }
  if (action === "pc_done") {
    if (blockQuickInput("Depois do PC, registre direto como N ou C2.")) return;
  }
  if ((action === "n" || action === "c2") && client.planStatus !== "Pronto") {
    if (blockQuickInput("Antes de registrar N ou C2, marque o plano como Pronto.")) return;
  }
  if (action === "n" && !notes) {
    if (blockQuickInput("N precisa de uma observacao curta do fechamento.")) return;
  }
  if (action === "c2" && (!objection || !notes || !followUp)) {
    if (blockQuickInput("C2 precisa de motivo, próxima ação e data de follow-up.")) return;
  }
  if (action === "recs" && recs <= 0) {
    if (blockQuickInput("Informe quantas RECs foram recebidas.")) return;
  }

  if (action === "ta_call") {
    client.taDate = date;
    client.lastInteraction = date;
    client.stage = "TA";
    client.followUp = followUp || nextBusinessDate();
    client.nextAction = notes || "Retomar TA e tentar evoluir para OI";
    client.completed = false;
    logClientEvent(client.id, "TA registrada", noteText, "Ligacao", date);
  }
  if (action === "oi_schedule") {
    client.taDate ||= date;
    client.oiScheduledDate = date;
    client.oiTime = time;
    if (link) client.oiMeetingLink = link;
    applyProcessStage(client, "OI agendada");
    client.followUp = date;
    client.nextAction = notes || "Confirmar presença na OI";
    logClientEvent(client.id, "OI agendada", noteText, "Reunião", date);
  }
  if (action === "oi_done") {
    client.oiDone = true;
    client.oiDate = date;
    applyProcessStage(client, "OI realizada");
    if (notes) client.planNotes = [client.planNotes, notes].filter(Boolean).join("\n");
    if (notes) client.oiSummaryNotes = notes;
    logClientEvent(client.id, "OI realizada", noteText, "Reunião", date);
  }
  if (action === "plan_ready") {
    markPlanReady(client);
    if (followUp) client.followUp = followUp;
    if (notes) client.planNotes = [client.planNotes, notes].filter(Boolean).join("\n");
  }
  if (action === "pc_schedule") {
    client.pcScheduledDate = date;
    client.pcTime = time;
    if (link) client.pcMeetingLink = link;
    applyProcessStage(client, "PC agendada");
    client.followUp = date;
    client.nextAction = notes || "Confirmar PC e revisar plano";
    logClientEvent(client.id, "PC agendada", noteText, "Reunião", date);
  }
  if (action === "pc_done") {
    client.pcDone = true;
    client.pcDate = date;
    applyProcessStage(client, "PC realizada");
    client.followUp = followUp || nextBusinessDate();
    client.nextAction = notes || "Registrar N ou conduzir follow-up de fechamento";
    logClientEvent(client.id, "PC realizada", noteText, "Reunião", date);
  }
  if (action === "n") {
    const recoveredFromC2 = client.outcome === "C2";
    client.pcDone = true;
    client.pcDate ||= date;
    client.outcome = "N";
    client.outcomeDate = date;
    if (recoveredFromC2) client.c2RecoveredAt = date;
    applyProcessStage(client, "N");
    client.followUp = followUp || client.nextRelationship || "";
    client.nextAction = notes || "Acompanhar emissão e entrega da apólice";
    logClientEvent(client.id, recoveredFromC2 ? "C2 recuperado em N" : "Venda fechada", noteText, "Historico", date);
  }
  if (action === "c2") {
    const enteringC2 = client.outcome !== "C2";
    client.pcDone = true;
    client.pcDate ||= date;
    client.outcome = "C2";
    client.outcomeDate = date;
    client.c2EnteredAt = enteringC2 ? date : (client.c2EnteredAt || date);
    if (enteringC2) client.c2RecoveredAt = "";
    client.c2Objection = objection || client.c2Objection || "";
    client.c2NextAction = notes || client.c2NextAction || "Realizar follow-up de fechamento";
    applyProcessStage(client, "C2");
    client.followUp = followUp || client.followUp || nextBusinessDate();
    client.nextAction = client.c2NextAction;
    client.completed = false;
    logClientEvent(client.id, "C2 registrado", noteText, "Historico", date);
  }
  if (action === "recs") {
    client.recs = Number(client.recs || 0) + recs;
    client.recsDate = date;
    client.lastInteraction = date;
    if (followUp) client.followUp = followUp;
    if (notes) client.nextAction = notes;
    logClientEvent(client.id, "RECs recebidas", `${recs} recomendação(ões). ${noteText}`, "Recomendacao", date);
  }
  if (action === "note") {
    client.lastInteraction = date;
    if (followUp) client.followUp = followUp;
    if (notes) client.nextAction = notes;
    client.completed = false;
    logClientEvent(client.id, "Nota comercial", noteText, "Historico", date);
  }

  saveState();
  renderAll();
  setQuickInputClient(client.id);
  resetQuickInputForm(true);
  if (action === "recs" && recs > 0) {
    setTimeout(() => {
      if (confirm(`${client.name} indicou ${recs} pessoa(s). Quer cadastrar a primeira indicação agora?`)) {
        openRecommendedContactFrom(client.id);
      } else {
        showToast("RECs registradas. Cadastre os nomes quando estiver pronto.");
      }
    }, 120);
  } else {
    showToast("Lançamento salvo com segurança");
  }
}

function renderAll() {
  renderOfficialSourceNotice();
  renderGeneralSettings();
  organizeIndicatorsView();
  organizeResultsWorkspace();
  renderDashboard();
  renderOperationWorkspace();
  renderPolicies();
  renderAnnualGoal();
  renderReferralIntelligence();
  if (currentView === "friday") renderFridayMode();
  if (currentView === "clients") {
    renderClients();
    if (crmTab === "funnel") renderPipeline();
  }
  if (currentView === "tasks") renderTasks();
  if (currentView === "assistant") renderAssistant();
  renderDailyRitual();
  renderDailyCommand();
  renderAutomationCenter();
  renderLeakRadar();
  renderPerformanceIntelligence();
  renderTruthMode();
  renderDecisionMode();
  renderQuickInput();
  renderGodMode();
  renderNextBestMovement();
  renderBoardCommercial();
  renderVipReferrers();
  renderCampaignCenter();
  renderFunnelRay();
  renderCeoCommandPanel();
  renderExecutiveDigest();
  applyReadOnlyMode();
}

function renderOperationWorkspace() {
  if (!el("operationAgendaList")) return;
  const today = offsetDate(0);
  const tasks = agendaTasksData()
    .filter(task => !task.completed && task.date <= today)
    .sort((a, b) => `${a.date || ""}${a.time || ""}`.localeCompare(`${b.date || ""}${b.time || ""}`));
  const overdue = tasks.filter(task => task.date < today).length;
  const ritual = dailyRitualData();
  const completedChecks = ritual.checks.filter(check => check.actual >= check.target).length;
  const ritualScore = Math.round((completedChecks / Math.max(ritual.checks.length, 1)) * 100);

  el("operationAgendaCount").textContent = `${tasks.length} ${tasks.length === 1 ? "pendência" : "pendências"}`;
  el("operationRitualScore").textContent = `${ritualScore}% concluído`;
  el("operationWorkspaceSummary").textContent = overdue
    ? `${overdue} item(ns) atrasado(s). Resolva o primeiro antes de abrir novas frentes.`
    : tasks.length
      ? `${tasks.length} movimento(s) pedem atenção hoje. O essencial está nesta fila.`
      : "Agenda protegida. Use o ritmo do dia para manter volume comercial.";

  el("operationAgendaList").innerHTML = tasks.length ? tasks.slice(0, 5).map(task => `
    <button type="button" class="operation-agenda-item" data-operation-client="${escapeHtml(task.clientId || "")}">
      <span class="${task.date < today ? "late" : "today"}">${task.time ? escapeHtml(task.time) : task.date < today ? "Atrasado" : "Hoje"}</span>
      <div><strong>${escapeHtml(displaySystemText(task.title || "Movimento comercial"))}</strong><small>${escapeHtml(displaySystemText(task.notes || task.type || "Próximo passo"))}</small></div>
    </button>`).join("") : `<div class="mini-empty compact">Nenhuma pendência vencida ou para hoje.</div>`;

  el("operationRitualList").innerHTML = ritual.checks.map(check => {
    const progress = ritualProgress(check.actual, check.target);
    return `<button type="button" class="operation-ritual-item" data-operation-ritual-view="${escapeHtml(check.view || "input")}" data-operation-filter="${escapeHtml(check.filter || "")}">
      <span><strong>${escapeHtml(check.title)}</strong><small>${check.actual}/${check.target}</small></span>
      <i><b style="width:${progress}%"></b></i>
    </button>`;
  }).join("");

  document.querySelectorAll("[data-operation-client]").forEach(button => button.addEventListener("click", () => {
    const clientId = button.dataset.operationClient;
    if (clientId) openClientModal(clientId);
    else switchView("tasks");
  }));
  document.querySelectorAll("[data-operation-ritual-view]").forEach(button => button.addEventListener("click", () => {
    if (button.dataset.operationFilter) workFilter = button.dataset.operationFilter;
    const targetView = button.dataset.operationRitualView || "input";
    if (targetView === "clients") crmTab = "list";
    switchView(targetView);
    if (targetView === "clients") renderClients();
  }));
}

function renderGeneralSettings() {
  if (el("defaultAreaCodeInput")) el("defaultAreaCodeInput").value = defaultAreaCode(state.settings);
}

function policySaleDate(client) {
  return client.outcomeDate || client.cpfDate || client.pcDate || client.createdAt?.slice(0, 10) || "";
}

function policyMonthlyPremium(client) {
  return Number(client.value || 0);
}

function policyAnnualPremium(client) {
  return policyMonthlyPremium(client) * 12;
}

function policyMonthlyCommission(client) {
  return policyAnnualPremium(client) * 0.4 / 12;
}

function policiesData() {
  const month = offsetDate(0).slice(0, 7);
  const closed = state.clients.filter(client => client.outcome === "N" || client.cpfDone);
  const monthClosed = closed.filter(client => policySaleDate(client).startsWith(month));
  const monthCpfs = state.clients.filter(client => client.cpfDone && (client.cpfDate || policySaleDate(client)).startsWith(month));
  const issued = monthClosed.filter(client => ["Emitida", "Entregue"].includes(client.policyStatus) || client.policyDelivered);
  const toIssue = monthClosed.filter(client => !["Emitida", "Entregue"].includes(client.policyStatus) && !client.policyDelivered);
  const annualPremium = monthClosed.reduce((sum, client) => sum + policyAnnualPremium(client), 0);
  const issuedAnnualPremium = issued.reduce((sum, client) => sum + policyAnnualPremium(client), 0);
  const pendingAnnualPremium = toIssue.reduce((sum, client) => sum + policyAnnualPremium(client), 0);
  const estimatedCommission = annualPremium * 0.4 / 12;
  const estimatedCommission12 = annualPremium * 0.4;
  const pendingIssue = closed.filter(client => !["Emitida", "Entregue"].includes(client.policyStatus) && !client.policyDelivered);
  const issuedNotDelivered = closed.filter(client => client.policyStatus === "Emitida" && !client.policyDelivered);
  const delivered = monthClosed.filter(client => client.policyDelivered || client.policyStatus === "Entregue");
  return { closed, monthClosed, monthCpfs, issued, toIssue, annualPremium, issuedAnnualPremium, pendingAnnualPremium, estimatedCommission, estimatedCommission12, pendingIssue, issuedNotDelivered, delivered };
}

function policyStatusLabel(client) {
  if (client.policyDelivered || client.policyStatus === "Entregue") return { label: "Entregue", tone: "success" };
  if (client.policyStatus === "Emitida") return { label: "Emitida, falta entregar", tone: "warning" };
  if (client.policyStatus === "Em análise") return { label: "Em análise", tone: "info" };
  if (client.policyStatus === "Pendente") return { label: "Pendente", tone: "danger" };
  return { label: "Aguardando emissão", tone: "danger" };
}

function renderPolicies() {
  if (!el("policiesList")) return;
  const data = policiesData();
  el("policiesMonthSales").textContent = data.monthClosed.length;
  el("policiesMonthSalesDetail").textContent = data.monthClosed.length
    ? `${data.delivered.length} entregue(s) no mês.`
    : "Sem fechamentos no mês.";
  if (el("policiesMonthCpfs")) el("policiesMonthCpfs").textContent = data.monthCpfs.length;
  if (el("policiesMonthCpfsDetail")) el("policiesMonthCpfsDetail").textContent = data.monthCpfs.length
    ? "CPFs concluídos no mês."
    : "Nenhum CPF concluído no mês.";
  if (el("policiesIssued")) el("policiesIssued").textContent = data.issued.length;
  if (el("policiesIssuedDetail")) el("policiesIssuedDetail").textContent = `${currency.format(data.issuedAnnualPremium)} de PA emitido.`;
  el("policiesAnnualPremium").textContent = currency.format(data.annualPremium);
  if (el("policiesAnnualPremiumDetail")) el("policiesAnnualPremiumDetail").textContent = `${currency.format(data.pendingAnnualPremium)} ainda a emitir.`;
  el("policiesPendingIssue").textContent = data.toIssue.length;
  el("policiesPendingIssueDetail").textContent = data.toIssue.length
    ? `${currency.format(data.pendingAnnualPremium)} de PA pendente.`
    : "Nenhuma pendência crítica.";
  el("policiesEstimatedCommission").textContent = currency.format(data.estimatedCommission);
  if (el("policyCommissionRate")) el("policyCommissionRate").value = 40;
  el("policiesBoardStatus").textContent = data.pendingIssue.length ? "Atenção" : "Em dia";

  const rows = [...data.closed].sort((a, b) => (policySaleDate(b) || "").localeCompare(policySaleDate(a) || ""));
  el("policiesList").innerHTML = rows.length ? rows.map(client => {
    const status = policyStatusLabel(client);
    const monthly = policyMonthlyPremium(client);
    const annual = policyAnnualPremium(client);
    const commission = policyMonthlyCommission(client);
    return `<article class="policy-row ${status.tone}">
      <div>
        <small>Cliente</small>
        <strong>${escapeHtml(client.name)}</strong>
        <span>${formatDate(policySaleDate(client)) || "Data não informada"}</span>
      </div>
      <div>
        <small>PA / prêmio</small>
        <strong>${currency.format(annual)}</strong>
        <span>${currency.format(monthly)} por mês</span>
      </div>
      <div>
        <small>Comissão mensal</small>
        <strong>${currency.format(commission)}</strong>
        <span>PA x 40% ÷ 12</span>
      </div>
      <div>
        <small>Status da apólice</small>
        <strong>${escapeHtml(status.label)}</strong>
        <span>${client.deliveryDate ? `Entrega: ${formatDate(client.deliveryDate)}` : "Entrega não registrada"}</span>
      </div>
      <div class="policy-actions">
        <button data-policy-open="${client.id}">Abrir ficha</button>
        <button data-policy-status="${client.id}" data-status="Emitida">Marcar emitida</button>
        <button data-policy-delivered="${client.id}">Marcar entregue</button>
      </div>
    </article>`;
  }).join("") : `<div class="mini-empty">Quando um contato virar N, ele aparecerá aqui para acompanhar emissão, entrega e comissão.</div>`;

  document.querySelectorAll("[data-policy-open]").forEach(button => button.addEventListener("click", () => openClientModal(button.dataset.policyOpen)));
  document.querySelectorAll("[data-policy-status]").forEach(button => button.addEventListener("click", () => {
    const client = state.clients.find(item => item.id === button.dataset.policyStatus);
    if (!client) return;
    client.policyStatus = button.dataset.status;
    if (!client.cpfDate && client.cpfDone) client.cpfDate = offsetDate(0);
    saveState();
    renderAll();
    showToast("Status da apólice atualizado");
  }));
  document.querySelectorAll("[data-policy-delivered]").forEach(button => button.addEventListener("click", () => {
    const client = state.clients.find(item => item.id === button.dataset.policyDelivered);
    if (!client) return;
    client.policyStatus = "Entregue";
    client.policyDelivered = true;
    client.deliveryDate ||= offsetDate(0);
    saveState();
    renderAll();
    showToast("Apólice marcada como entregue");
  }));
}

function organizeIndicatorsView() {
  const target = el("indicatorsContent");
  if (!target || target.dataset.organized === "true") return;
  [
    ".god-mode-panel",
    ".board-panel",
    ".campaign-panel",
    ".annual-goal-card",
    ".weekly-cockpit",
    ".metrics-grid",
    ".performance-intelligence",
    ".referral-intelligence",
    ".vip-referrers",
    ".dashboard-grid"
  ].forEach(selector => {
    const node = document.querySelector(`#dashboardView ${selector}`);
    if (node) target.appendChild(node);
  });
  target.dataset.organized = "true";
}

function organizeResultsWorkspace() {
  const summary = el("resultsSummaryPanel");
  const performance = el("resultsPerformancePanel");
  const policies = el("resultsPoliciesPanel");
  const diagnosis = el("resultsDiagnosisPanel");
  if (!summary || summary.dataset.organized === "true") return;

  [document.querySelector(".ceo-command-panel"), document.querySelector(".executive-digest-panel")]
    .filter(Boolean).forEach(node => summary.appendChild(node));
  [document.querySelector(".funnel-ray-panel"), document.querySelector("#indicatorsView .ritual-panel"), el("indicatorsContent")]
    .filter(Boolean).forEach(node => performance.appendChild(node));
  [document.querySelector("#policiesView .policies-hero"), document.querySelector("#policiesView .policies-summary"), document.querySelector("#policiesView .policies-board")]
    .filter(Boolean).forEach(node => policies.appendChild(node));
  [document.querySelector("#managementView .management-hero"), document.querySelector("#managementView .leak-radar-panel"), el("managementContent"), document.querySelector(".data-vault-panel")]
    .filter(Boolean).forEach(node => diagnosis.appendChild(node));

  summary.dataset.organized = "true";
  switchResultsTab(resultsTab);
}

function switchResultsTab(tab = "summary") {
  const allowed = new Set(["summary", "performance", "policies", "diagnosis"]);
  resultsTab = allowed.has(tab) ? tab : "summary";
  document.querySelectorAll("[data-results-tab]").forEach(button => {
    const active = button.dataset.resultsTab === resultsTab;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", String(active));
  });
  document.querySelectorAll("[data-results-panel]").forEach(panel => {
    panel.classList.toggle("active", panel.dataset.resultsPanel === resultsTab);
  });
}

function switchCrmTab(tab = "list") {
  crmTab = tab === "funnel" ? "funnel" : "list";
  document.querySelectorAll("[data-crm-tab]").forEach(button => {
    const active = button.dataset.crmTab === crmTab;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", String(active));
  });
  document.querySelectorAll("[data-crm-panel]").forEach(panel => {
    panel.classList.toggle("active", panel.dataset.crmPanel === crmTab);
  });
  if (crmTab === "funnel") renderPipeline();
  else renderClients();
}

function dateInCurrentMonth(date) {
  return Boolean(date && date.startsWith(offsetDate(0).slice(0, 7)));
}

function funnelStats(scope = "week") {
  const week = selectedWeek();
  const inScope = date => scope === "week" ? inSelectedWeek(date, week) : dateInCurrentMonth(date);
  const ta = state.clients.filter(client => inScope(client.taDate)).length;
  const oiScheduled = state.clients.filter(client => inScope(client.oiScheduledDate)).length;
  const oiDone = state.clients.filter(client => client.oiDone && inScope(client.oiDate || client.oiScheduledDate)).length;
  const pcScheduled = state.clients.filter(client => inScope(client.pcScheduledDate)).length;
  const pcDone = state.clients.filter(client => client.pcDone && inScope(client.pcDate || client.pcScheduledDate)).length;
  const n = state.clients.filter(client => client.outcome === "N" && inScope(client.outcomeDate || client.pcDate)).length;
  const c2 = state.clients.filter(client => client.outcome === "C2" && inScope(client.outcomeDate || client.pcDate)).length;
  return { ta, oiScheduled, oiDone, pcScheduled, pcDone, n, c2 };
}

function rate(part, total) {
  if (!total) return null;
  return Math.round((part / total) * 100);
}

function rateLabel(part, total) {
  const value = rate(part, total);
  return value === null ? "Sem base" : `${value}%`;
}

function funnelDiagnosis(stats, scope = "week") {
  const goal = scope === "week" ? weeklyGoal(selectedWeek()) : Math.max(weeklyGoal(selectedWeek()) * 4, 80);
  if (scope === "week" && stats.oiScheduled < weeklyGoal(selectedWeek())) {
    return {
      status: "Volume",
      text: `Agenda abaixo da meta: ${stats.oiScheduled}/${weeklyGoal(selectedWeek())} OIs. A correção começa em TA e recomendações.`
    };
  }
  if (stats.ta && rate(stats.oiScheduled, stats.ta) < 35) {
    return {
      status: "TA → OI",
      text: "Você está abordando, mas pouca conversa vira OI. Revise abordagem, origem dos nomes e pedido de horário."
    };
  }
  if (stats.oiScheduled >= 3 && rate(stats.oiDone, stats.oiScheduled) < 65) {
    return {
      status: "Presença OI",
      text: "Tem OI marcada, mas realização baixa. Reforce confirmação no dia e mensagem de compromisso com link."
    };
  }
  if (stats.oiDone >= 2 && rate(stats.pcScheduled, stats.oiDone) < 60) {
    return {
      status: "OI → PC",
      text: "A OI aconteceu, mas o próximo compromisso não ficou protegido. Saia da OI com PC marcado sempre que possível."
    };
  }
  if (stats.pcDone >= 2 && rate(stats.n, stats.pcDone) < 40) {
    return {
      status: "PC → N",
      text: "O gargalo está no fechamento. Revise objeções, tese de venda e se o plano chegou pronto para decisão."
    };
  }
  if (!stats.ta && !stats.oiScheduled && !stats.n) {
    const periodLabel = scope === "week" ? "nesta semana" : "neste mês";
    return {
      status: "Sem leitura",
      text: `Ainda falta base ${periodLabel}. Registre TAs, OIs e PCs para o SITPLAN diagnosticar.`
    };
  }
  return {
    status: "Saudável",
    text: "O funil não mostra vazamento crítico. Mantenha volume de TA, plano no prazo e follow-up de C2."
  };
}

function funnelStepsHtml(stats) {
  const steps = [
    ["TAs", stats.ta, "Volume"],
    ["OI marcadas", stats.oiScheduled, rateLabel(stats.oiScheduled, stats.ta)],
    ["OI realizadas", stats.oiDone, rateLabel(stats.oiDone, stats.oiScheduled)],
    ["PC marcadas", stats.pcScheduled, rateLabel(stats.pcScheduled, stats.oiDone)],
    ["PC realizadas", stats.pcDone, rateLabel(stats.pcDone, stats.pcScheduled)],
    ["Ns", stats.n, rateLabel(stats.n, stats.pcDone)]
  ];
  return steps.map(([label, value, detail], index) => `
    <div class="funnel-step">
      <b>${index + 1}</b>
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(String(value))}</strong>
      <small>${escapeHtml(detail)}</small>
    </div>
  `).join("");
}

function renderFunnelRay() {
  if (!el("funnelRayWeek")) return;
  const week = funnelStats("week");
  const month = funnelStats("month");
  const diagnosis = funnelDiagnosis(week, "week");
  const c2 = c2PerformanceData();
  el("funnelRayWeek").innerHTML = funnelStepsHtml(week);
  el("funnelRayMonth").innerHTML = funnelStepsHtml(month);
  el("funnelRayStatus").textContent = diagnosis.status;
  el("funnelRayDiagnosis").textContent = diagnosis.text;
  if (el("c2PerformanceTouched")) el("c2PerformanceTouched").textContent = c2.touched;
  if (el("c2PerformanceRecovered")) el("c2PerformanceRecovered").textContent = c2.recovered;
  if (el("c2PerformanceRate")) {
    el("c2PerformanceRate").textContent = c2.recoveryRate === null
      ? "Sem C2 movimentado no mês."
      : `${c2.recoveryRate}% dos C2 movimentados viraram N.`;
  }
  if (el("c2PerformanceAverage")) el("c2PerformanceAverage").textContent = c2.average === null ? "—" : `${c2.average} dias`;
  if (el("c2PerformanceOpen")) el("c2PerformanceOpen").textContent = c2.open;
  if (el("c2PerformanceBalance")) {
    el("c2PerformanceBalance").textContent = c2.balance > 0
      ? `Carteira cresceu ${c2.balance} no mês.`
      : c2.balance < 0
        ? `Carteira reduziu ${Math.abs(c2.balance)} no mês.`
        : "Entradas e recuperações equilibradas.";
  }
}

function criticalConversionMetric(stats) {
  const stages = [
    { label: "TA -> OI", value: rate(stats.oiScheduled, stats.ta), hint: "abordagens que viraram OI marcada" },
    { label: "OI -> PC", value: rate(stats.pcScheduled, stats.oiDone), hint: "OIs realizadas que viraram PC marcada" },
    { label: "PC -> N", value: rate(stats.n, stats.pcDone), hint: "PCs realizados que viraram N" }
  ].filter(item => item.value !== null);
  if (!stages.length) return { label: "Sem base", value: null, text: "Registre TAs, OIs e PCs para formar leitura." };
  const worst = stages.sort((a, b) => a.value - b.value)[0];
  return { ...worst, text: `${worst.label}: ${worst.value}% (${worst.hint}).` };
}

function ceoToneFromRatio(value, target, warnRatio = .75) {
  if (!target) return "neutral";
  if (value >= target) return "good";
  if (value >= Math.ceil(target * warnRatio)) return "warn";
  return "risk";
}

function renderCeoCommandPanel() {
  if (!el("ceoCommandGrid")) return;
  const week = weekFromOffset(0);
  const nextWeek = weekFromOffset(1);
  const weekStats = funnelStats("week");
  const monthStats = funnelStats("month");
  const goal = weeklyGoal(week);
  const nextGoal = weeklyGoal(nextWeek);
  const clients = state.clients;
  const nextOi = clients.filter(client => inSelectedWeek(client.oiScheduledDate, nextWeek)).length;
  const nextPc = clients.filter(client => inSelectedWeek(client.pcScheduledDate, nextWeek)).length;
  const c2Open = clients.filter(client => client.outcome === "C2" && !client.doNotContact).length;
  const c2Due = clients.filter(client => client.outcome === "C2" && !client.doNotContact && (!client.followUp || client.followUp <= week.endKey || (daysSince(client.lastInteraction) ?? 0) >= 7)).length;
  const policies = policiesData();
  const diagnosis = funnelDiagnosis(weekStats, "week");
  const critical = criticalConversionMetric(weekStats);
  const ranking = referralRankingData();
  const topReferrer = ranking[0];
  const topConversion = topReferrer?.leads ? Math.round(referralConversion(topReferrer) * 100) : null;
  const weekMissing = Math.max(goal - weekStats.oiScheduled, 0);
  const nextMissing = Math.max(nextGoal - nextOi, 0);
  const weekTone = ceoToneFromRatio(weekStats.oiScheduled, goal);
  const nextTone = nextOi >= nextGoal ? "good" : nextOi >= Math.ceil(nextGoal * .4) ? "warn" : "risk";
  const conversionTone = critical.value === null ? "neutral" : critical.value >= 60 ? "good" : critical.value >= 35 ? "warn" : "risk";
  const bottleneckTone = diagnosis.status === "Saudável" ? "good" : diagnosis.status === "Sem leitura" ? "neutral" : "risk";
  const referrerTone = topReferrer?.n ? "good" : topReferrer?.leads ? "warn" : "neutral";
  const policyTone = policies.toIssue.length ? "warn" : policies.monthClosed.length ? "good" : "neutral";
  const overallTone = [weekTone, nextTone, conversionTone, bottleneckTone, policyTone].includes("risk")
    ? "risk"
    : [weekTone, nextTone, conversionTone, bottleneckTone, policyTone].includes("warn")
      ? "warn"
      : "good";
  const cards = [
    {
      tone: weekTone,
      label: "Ritmo da semana",
      value: `${weekStats.oiScheduled}/${goal} OIs`,
      detail: weekMissing
        ? `Faltam ${weekMissing}. PCs: ${weekStats.pcScheduled}. Ns: ${weekStats.n}. C2: ${c2Open} aberto(s), ${c2Due} para agir.`
        : `Semana protegida. PCs: ${weekStats.pcScheduled}. Ns: ${weekStats.n}. C2: ${c2Open} aberto(s).`
    },
    {
      tone: nextTone,
      label: "Próxima semana",
      value: `${nextOi}/${nextGoal} OIs`,
      detail: nextMissing
        ? `Faltam ${nextMissing}. PCs futuras: ${nextPc}.`
        : `Agenda futura protegida. PCs futuras: ${nextPc}.`
    },
    {
      tone: conversionTone,
      label: "Conversão crítica",
      value: critical.value === null ? "Sem base" : `${critical.label} ${critical.value}%`,
      detail: critical.value === null ? "Registre TA, OI e PC para formar leitura." : critical.hint
    },
    {
      tone: bottleneckTone,
      label: "Gargalo real",
      value: diagnosis.status,
      detail: diagnosis.text
    },
    {
      tone: referrerTone,
      label: "Origem de receita",
      value: topReferrer ? topReferrer.name : "Sem mapa",
      detail: topReferrer
        ? `${topReferrer.leads} indicado(s), ${topReferrer.n} N(s)${topConversion === null ? "" : `, ${topConversion}% conversão`}.`
        : "Cadastre recomendante para saber quem alimenta receita."
    },
    {
      tone: policyTone,
      label: "Apólices e comissão",
      value: currency.format(policies.estimatedCommission),
      detail: `${monthStats.n} N(s), ${policies.monthCpfs.length} CPF(s), ${policies.toIssue.length} a emitir.`
    }
  ];

  el("ceoCommandGrid").innerHTML = cards.map(card => `
    <article class="ceo-command-card ${card.tone}">
      <small>${escapeHtml(card.label)}</small>
      <strong>${escapeHtml(card.value)}</strong>
      <span>${escapeHtml(card.detail)}</span>
    </article>
  `).join("");
  const status = el("ceoOverallStatus");
  status.textContent = overallTone === "risk" ? "Risco" : overallTone === "warn" ? "Atenção" : "Protegido";
  status.classList.remove("good", "warn", "risk", "neutral");
  status.classList.add(overallTone);
  el("ceoCommandSummary").textContent = overallTone === "good"
    ? "Operação saudável: semana protegida, próxima semana em construção e sem gargalo crítico."
    : overallTone === "warn"
      ? `Atenção seletiva: ${weekMissing ? `faltam ${weekMissing} OI(s) nesta semana. ` : ""}${nextMissing ? `Próxima semana ainda precisa de ${nextMissing} OI(s). ` : ""}Use o gargalo para escolher a ação.`
      : `Risco comercial: ${diagnosis.text}`;
}

function dataVaultData() {
  const backupAge = state.settings.lastBackupAt ? daysSince(state.settings.lastBackupAt) : null;
  const cloud = isCloudConnected();
  const backupOk = backupAge !== null && backupAge <= 3;
  const backupWarn = backupAge !== null && backupAge <= 7;
  const privacyLocked = PRIVATE_ACCESS_REQUIRED;
  return {
    cloud,
    backupAge,
    backupOk,
    backupWarn,
    privacyLocked,
    cloudLabel: cloud ? "Nuvem ativa" : "Modo local",
    cloudHint: cloud
      ? "Base sincronizada entre computador e celular."
      : "Entre e sincronize para reduzir risco de perda.",
    backupLabel: backupAge === null
      ? "Sem backup"
      : backupAge === 0
        ? "Backup hoje"
        : `${backupAge} dia(s)`,
    backupHint: backupAge === null
      ? "Exporte um backup antes de usar forte no dia a dia."
      : backupAge <= 3
        ? "Cópia recente. Bom padrão de segurança."
        : backupAge <= 7
          ? "Faça novo backup esta semana."
          : "Backup antigo. Exporte uma cópia hoje.",
    privacyLabel: privacyLocked ? "Login obrigatório" : "Arquivo local",
    privacyHint: privacyLocked
      ? "Site publicado fica protegido por autenticação."
      : "Abrindo por arquivo local; não depende de internet."
  };
}

function executiveDigestData() {
  const week = selectedWeek();
  const weekStats = funnelStats("week");
  const monthStats = funnelStats("month");
  const goal = weeklyGoal(week);
  const policies = policiesData();
  const campaign = campaignData();
  const vault = dataVaultData();
  const diagnosis = funnelDiagnosis(weekStats, "week");
  const critical = criticalConversionMetric(weekStats);
  const topReferrer = referralRankingData()[0];
  const priorities = rankedNextMovements(3);
  const ritual = dailyRitualData();
  const doneChecks = ritual.checks.filter(item => item.actual >= item.target).length;
  const subject = `SITPLAN - Resumo executivo ${formatDate(offsetDate(0))}`;
  const lines = [
    subject,
    "",
    `1. Semana: ${weekStats.oiScheduled}/${goal} OIs agendadas. TA: ${weekStats.ta}. OI realizadas: ${weekStats.oiDone}. PC realizadas: ${weekStats.pcDone}. N: ${weekStats.n}.`,
    `2. Gargalo: ${diagnosis.status} - ${diagnosis.text}`,
    `3. Conversão crítica: ${critical.text}`,
    `4. Mês: ${monthStats.n} N(s), ${policies.monthCpfs.length} CPF(s), PA ${currency.format(policies.annualPremium)}, comissão mensal estimada ${currency.format(policies.estimatedCommission)}.`,
    `5. Campanhas: ${campaign.totalN}/${campaign.settings.octoberTarget} N até ${formatDate(campaign.settings.octoberDeadline)}. 3W: semana ${campaign.settings.threeWCurrentWeek}/${campaign.settings.threeWTotalWeeks}, meta ${campaign.settings.threeWWeeklyTarget} N por semana.`,
    `6. Top recomendante: ${topReferrer ? `${topReferrer.name} (${topReferrer.leads} indicado(s), ${topReferrer.n} N)` : "ainda sem recomendante mapeado"}.`,
    `7. Ritual diário: ${doneChecks}/${ritual.checks.length} blocos protegidos hoje.`,
    `8. Segurança: ${vault.cloudLabel}. Backup: ${vault.backupLabel}. ${vault.backupHint}`,
    "",
    "Top 3 movimentos:",
    ...(priorities.length ? priorities.map((item, index) => `${index + 1}. ${item.movement.action} - ${item.client.name}: ${item.movement.reason}.`) : ["1. Sem prioridade crítica no momento."]),
    "",
    "Regra de ouro do dia: executar a próxima ação antes de abrir novas ideias."
  ];
  return { weekStats, monthStats, goal, policies, campaign, vault, diagnosis, critical, subject, body: lines.join("\n") };
}

function setVaultCardTone(card, tone) {
  if (!card) return;
  card.classList.remove("good", "warn", "risk");
  card.classList.add(tone);
}

function renderExecutiveDigest() {
  if (!el("executiveDigestPreview")) return;
  const data = executiveDigestData();
  const oisMissing = Math.max(data.goal - data.weekStats.oiScheduled, 0);
  el("digestWeekOis").textContent = `${data.weekStats.oiScheduled}/${data.goal} OIs`;
  el("digestWeekHint").textContent = oisMissing ? `Faltam ${oisMissing} OI(s) para proteger a semana.` : "Meta semanal protegida.";
  el("digestCriticalRate").textContent = data.critical.value === null ? "Sem base" : `${data.critical.label} ${data.critical.value}%`;
  el("digestCriticalHint").textContent = data.diagnosis.status === "Saudável" ? "Sem vazamento crítico agora." : data.diagnosis.text;
  el("digestMonthMoney").textContent = currency.format(data.policies.estimatedCommission);
  el("digestMonthMoneyHint").textContent = `${data.monthStats.n} N(s), ${data.policies.monthCpfs.length} CPF(s), PA ${currency.format(data.policies.annualPremium)}.`;
  el("digestSecurityStatus").textContent = data.vault.cloudLabel;
  el("digestSecurityHint").textContent = `Backup: ${data.vault.backupLabel}. ${data.vault.backupHint}`;
  el("executiveDigestSubtitle").textContent = `Gerado com dados de ${formatDate(offsetDate(0))}. Use para abrir o dia sem procurar relatório.`;
  el("executiveDigestPreview").textContent = data.body;

  el("dataVaultSummary").textContent = data.vault.cloud
    ? "Nuvem conectada. Continue mantendo backup manual como segunda proteção."
    : "Você está em modo local. Para máxima segurança, conecte a nuvem e exporte backup regularmente.";
  el("vaultCloudStatus").textContent = data.vault.cloudLabel;
  el("vaultCloudHint").textContent = data.vault.cloudHint;
  el("vaultBackupStatus").textContent = data.vault.backupLabel;
  el("vaultBackupHint").textContent = data.vault.backupHint;
  el("vaultPrivacyStatus").textContent = data.vault.privacyLabel;
  el("vaultPrivacyHint").textContent = data.vault.privacyHint;
  setVaultCardTone(el("vaultCloudCard"), data.vault.cloud ? "good" : "warn");
  setVaultCardTone(el("vaultBackupCard"), data.vault.backupOk ? "good" : data.vault.backupWarn ? "warn" : "risk");
  setVaultCardTone(el("vaultPrivacyCard"), data.vault.privacyLocked ? "good" : "warn");
}

async function copyExecutiveDigest() {
  const data = executiveDigestData();
  try {
    await navigator.clipboard.writeText(data.body);
    showToast("Resumo executivo copiado");
  } catch (error) {
    showToast("Não consegui copiar automaticamente. Selecione o texto do resumo.", true);
  }
}

function openExecutiveDigestEmail() {
  const data = executiveDigestData();
  const email = state.settings.executiveDigestEmail || "";
  const mailto = `mailto:${encodeURIComponent(email)}?subject=${encodeURIComponent(data.subject)}&body=${encodeURIComponent(data.body)}`;
  window.location.href = mailto;
}

function truthModeData() {
  const week = selectedWeek();
  const today = offsetDate(0);
  const goal = weeklyGoal(week);
  const ta = state.clients.filter(client => inSelectedWeek(client.taDate, week)).length;
  const oiScheduled = state.clients.filter(client => inSelectedWeek(client.oiScheduledDate, week)).length;
  const oiDone = state.clients.filter(client => client.oiDone && inSelectedWeek(client.oiDate, week)).length;
  const pcScheduled = state.clients.filter(client => inSelectedWeek(client.pcScheduledDate, week)).length;
  const pcDone = state.clients.filter(client => client.pcDone && inSelectedWeek(client.pcDate, week)).length;
  const sales = state.clients.filter(client => client.outcome === "N" && inSelectedWeek(client.outcomeDate, week)).length;
  const c2 = state.clients.filter(client => client.outcome === "C2" && inSelectedWeek(client.outcomeDate, week)).length;
  const plansPending = state.clients.filter(client => client.oiDone && client.planStatus !== "Pronto" && client.outcome !== "N");
  const oiWithoutPc = state.clients.filter(client => client.oiDone && !client.pcScheduledDate && !client.pcDone && client.outcome !== "N");
  const staleC2 = state.clients.filter(client => client.outcome === "C2" && (!client.followUp || client.followUp <= today || (daysSince(client.lastInteraction) ?? 0) >= 6));
  const noNextStep = state.clients.filter(client => client.outcome !== "N" && (!client.nextAction || !client.followUp));
  const fridayQueue = fridayQueueClients();
  const oiNeeded = Math.max(goal - oiScheduled, 0);
  const oiShowRate = oiScheduled ? oiDone / oiScheduled : 0;
  const pcCloseRate = pcDone ? sales / pcDone : 0;

  if (oiScheduled < goal && ta < Math.max(10, Math.ceil(goal * 0.75))) {
    return {
      badge: "Gargalo: volume de TA",
      headline: "O problema principal ainda está no começo do funil.",
      text: `Você tem ${oiScheduled}/${goal} OIs agendadas na semana e apenas ${ta} TA(s) registrado(s). Sem abordagem suficiente, o resto vira torcida.`,
      action: `Faça ${Math.min(Math.max(oiNeeded, 5), 15)} abordagens antes de mexer em qualquer ajuste fino.`,
      primaryView: "friday",
      secondaryView: "clients",
      facts: [["TA semana", ta], ["OIs/meta", `${oiScheduled}/${goal}`], ["Fila sexta", fridayQueue.length], ["Sem próximo passo", noNextStep.length]]
    };
  }
  if (plansPending.length >= 3 || (plansPending.length && pcScheduled)) {
    return {
      badge: "Gargalo: plano parado",
      headline: "Você já conquistou atenção. Agora precisa transformar OI em material de PC.",
      text: `${plansPending.length} plano(s) estão pendentes. Plano atrasado mata timing, esfria confiança e empurra venda para C2.`,
      action: "Monte ou marque pronto o primeiro plano pendente hoje.",
      primaryView: "tasks",
      secondaryView: "management",
      facts: [["Planos pendentes", plansPending.length], ["PCs semana", pcScheduled], ["OIs feitas", oiDone], ["OI sem PC", oiWithoutPc.length]]
    };
  }
  if (oiWithoutPc.length >= 2 || (oiDone >= 2 && pcScheduled < oiDone * 0.6)) {
    return {
      badge: "Gargalo: OI sem PC",
      headline: "A conversa aconteceu, mas o próximo compromisso não está protegido.",
      text: `${oiWithoutPc.length} OI(s) realizadas ainda não têm PC marcado. O risco aqui é deixar interesse virar esquecimento.`,
      action: "Agende PC para as OIs sem próximo encontro antes de buscar novos ajustes.",
      primaryView: "clients",
      secondaryView: "tasks",
      facts: [["OI realizadas", oiDone], ["PCs marcados", pcScheduled], ["OI sem PC", oiWithoutPc.length], ["Presença OI", `${Math.round(oiShowRate * 100)}%`]]
    };
  }
  if (staleC2.length >= 2 || (pcDone >= 2 && pcCloseRate < .3 && c2 >= sales)) {
    return {
      badge: "Gargalo: fechamento C2",
      headline: "Você está deixando decisão aberta tempo demais.",
      text: `${staleC2.length} C2(s) pedem follow-up. C2 sem data e sem motivo vira arquivo morto com nome bonito.`,
      action: "Ataque o C2 mais recente com uma pergunta objetiva de decisão.",
      primaryView: "clients",
      secondaryView: "assistant",
      facts: [["PCs feitos", pcDone], ["Ns semana", sales], ["C2 semana", c2], ["C2 vencidos", staleC2.length]]
    };
  }
  if (noNextStep.length >= 5) {
    return {
      badge: "Gargalo: base sem direção",
      headline: "Sua base tem nomes, mas parte dela não tem comando.",
      text: `${noNextStep.length} contato(s) estão sem próximo passo claro. Contato sem próxima ação não é pipeline, é ruído.`,
      action: "Limpe 10 contatos sem próximo passo e coloque data, ação e prioridade.",
      primaryView: "clients",
      secondaryView: "management",
      facts: [["Sem próximo passo", noNextStep.length], ["Fila sexta", fridayQueue.length], ["OIs/meta", `${oiScheduled}/${goal}`], ["TA semana", ta]]
    };
  }
  return {
    badge: "Ritmo protegido",
    headline: "Hoje o jogo é consistência, não reinvenção.",
    text: `Você tem ${oiScheduled}/${goal} OIs na semana. Continue executando o básico: novas abordagens, planos no prazo e follow-up de C2.`,
    action: "Faça a próxima ação do Comando do Dia antes de abrir novas ideias.",
    primaryView: "dashboard",
    secondaryView: "tasks",
    facts: [["TA semana", ta], ["OIs/meta", `${oiScheduled}/${goal}`], ["Planos", plansPending.length], ["C2 vencidos", staleC2.length]]
  };
}

function renderTruthMode() {
  if (!el("truthHeadline")) return;
  const truth = truthModeData();
  el("truthBadge").textContent = truth.badge;
  el("truthHeadline").textContent = truth.headline;
  el("truthText").textContent = truth.text;
  el("truthAction").textContent = truth.action;
  el("truthFacts").innerHTML = truth.facts.map(([label, value]) => `<article><small>${escapeHtml(label)}</small><strong>${escapeHtml(String(value))}</strong></article>`).join("");
  el("truthPrimaryAction").dataset.truthView = truth.primaryView;
  el("truthSecondaryAction").dataset.truthView = truth.secondaryView;
}

function renderDecisionMode() {
  if (!el("decisionModeHeadline")) return;
  const truth = truthModeData();
  el("decisionModeBadge").textContent = truth.badge;
  el("decisionModeHeadline").textContent = truth.headline;
  el("decisionModeText").textContent = truth.text;
  el("decisionModeAction").textContent = truth.action;
  el("decisionModeButton").dataset.decisionView = truth.primaryView || "input";
}

function executionCallClients() {
  const today = offsetDate(0);
  return [...state.clients]
    .filter(client => !client.doNotContact && client.outcome !== "N")
    .filter(client => {
      if (client.followUp && client.followUp <= today) return true;
      if (["Base", "TA"].includes(client.stage) && !client.oiScheduledDate) return true;
      if (client.source === "Indicação" || client.referrer) return true;
      return false;
    })
    .sort((a, b) => assistantPriorityScore(b) - assistantPriorityScore(a))
    .slice(0, 4);
}

function executionPlanClients() {
  return [...state.clients]
    .filter(client => client.oiDone && client.planStatus !== "Pronto" && client.outcome !== "N")
    .sort((a, b) => (a.planDueDate || "9999").localeCompare(b.planDueDate || "9999"))
    .slice(0, 4);
}

function executionC2Clients() {
  return [...state.clients]
    .filter(client => client.outcome === "C2" && !client.doNotContact)
    .sort((a, b) => {
      const dateCompare = (a.followUp || "9999").localeCompare(b.followUp || "9999");
      return dateCompare || assistantPriorityScore(b) - assistantPriorityScore(a);
    })
    .slice(0, 4);
}

function executionClientLine(client, type) {
  const digits = normalizeWhatsapp(client.whatsapp || client.phone);
  const number = digits.startsWith("55") ? digits : `55${digits}`;
  const message = type === "c2"
    ? `Olá, ${client.name.split(/\s+/)[0]}! Pensei no nosso último encontro e queria retomar um ponto importante para te ajudar a decidir com segurança. Podemos falar rapidamente hoje?`
    : assistantWhatsappMessage(client, "short");
  const whatsapp = digits.length >= 10 ? `https://wa.me/${number}?text=${encodeURIComponent(message)}` : "";
  const detail = type === "plan"
    ? `${client.planDueDate ? `Prazo ${formatDate(client.planDueDate)}` : "Sem prazo"} · ${client.postOiCompletedAt ? "Pós-OI registrado" : "Revisar diagnóstico"}`
    : type === "c2"
      ? `${client.followUp ? `Follow-up ${formatDate(client.followUp)}` : "Sem data"} · ${assistantDays(client)}d sem contato`
      : `${assistantAction(client).title} · ${assistantDays(client)}d sem contato`;
  const primaryAction = type === "plan"
    ? `<button data-execution-plan-ready="${client.id}">Plano pronto</button>`
    : whatsapp ? `<a href="${whatsapp}" target="_blank" rel="noopener" data-execution-whatsapp="${client.id}">WhatsApp</a>` : "";
  const secondAction = type === "plan"
    ? `<button data-execution-pc="${client.id}">Preparar PC</button>`
    : "";
  return `<div class="execution-item">
    <div>
      <strong>${escapeHtml(client.name)}</strong>
      <small>${escapeHtml(detail)}</small>
    </div>
    <div class="execution-actions">
      ${primaryAction}
      <button data-execution-open="${client.id}">Abrir ficha</button>
      ${secondAction}
    </div>
  </div>`;
}

function renderExecutionCockpit() {
  if (!el("executionCallList")) return;
  const calls = executionCallClients();
  const plans = executionPlanClients();
  const c2s = executionC2Clients();
  const total = calls.length + plans.length + c2s.length;
  if (el("todayFocusCalls")) {
    el("todayFocusCalls").textContent = calls.length;
    el("todayFocusPlans").textContent = plans.length;
    el("todayFocusC2").textContent = c2s.length;
    el("todayFocusHeadline").textContent = total
      ? `${total} movimentos comerciais para proteger hoje.`
      : "Hoje esta limpo. Hora de criar novas oportunidades.";
    el("todayFocusSubline").textContent = total
      ? "Comece pelo maior gargalo e registre cada avanço no Lançar."
      : "Use o tempo livre para pedir RECs, fazer TA e alimentar a próxima semana.";
  }
  el("executionCallCount").textContent = calls.length;
  el("executionPlanCount").textContent = plans.length;
  el("executionC2Count").textContent = c2s.length;
  if (el("executionCockpitSummary")) el("executionCockpitSummary").textContent = total
    ? `${total} movimento(s) comerciais para proteger seu dia. Comece pela primeira fila.`
    : "Nenhuma frente crítica agora. Use o tempo para buscar novas OIs.";
  if (el("dailyCommandSummary")) el("dailyCommandSummary").textContent = total
    ? `${total} movimento(s) comerciais para proteger hoje. Comece pela primeira fila e registre o avanço no Lançar.`
    : "Nenhuma frente crítica agora. Use o tempo para pedir RECs, fazer TA e alimentar a próxima semana.";
  el("executionCallList").innerHTML = calls.length
    ? calls.map(client => executionClientLine(client, "call")).join("")
    : `<div class="mini-empty">Sem ligações críticas agora.</div>`;
  el("executionPlanList").innerHTML = plans.length
    ? plans.map(client => executionClientLine(client, "plan")).join("")
    : `<div class="mini-empty">Nenhum plano pendente.</div>`;
  el("executionC2List").innerHTML = c2s.length
    ? c2s.map(client => executionClientLine(client, "c2")).join("")
    : `<div class="mini-empty">Nenhum C2 pendente.</div>`;

  document.querySelectorAll("[data-execution-open]").forEach(button => button.addEventListener("click", () => openClientModal(button.dataset.executionOpen)));
  document.querySelectorAll("[data-execution-assistant]").forEach(button => button.addEventListener("click", () => {
    assistantClientId = button.dataset.executionAssistant;
    assistantMode = "next";
    switchView("assistant");
  }));
  document.querySelectorAll("[data-execution-whatsapp]").forEach(link => link.addEventListener("click", () => {
    const client = state.clients.find(item => item.id === link.dataset.executionWhatsapp);
    if (!client) return;
    client.lastInteraction = offsetDate(0);
    client.taDate ||= offsetDate(0);
    if (client.stage === "Base") client.stage = "TA";
    client.followUp = nextBusinessDate();
    client.nextAction = "Aguardar retorno no WhatsApp";
    client.completed = false;
    logClientEvent(client.id, "WhatsApp iniciado pelo Comando do Dia", "Contato priorizado na tela Hoje.", "WhatsApp");
    saveState();
    renderAll();
  }));
  document.querySelectorAll("[data-execution-plan-ready]").forEach(button => button.addEventListener("click", () => {
    const client = state.clients.find(item => item.id === button.dataset.executionPlanReady);
    if (!client) return;
    markPlanReady(client);
    saveState();
    renderAll();
    showToast("Plano marcado como pronto");
  }));
  document.querySelectorAll("[data-execution-pc]").forEach(button => button.addEventListener("click", () => {
    openMeetingModal("PC", button.dataset.executionPc);
  }));
}

function renderTodayWeekCommand() {
  if (!el("todayWeekOi")) return;
  const setText = (id, value) => {
    const node = el(id);
    if (node) node.textContent = value;
  };
  const clients = state.clients;
  const current = weekFromOffset(0);
  const next = weekFromOffset(1);
  const currentGoal = weeklyGoal(current);
  const nextGoal = weeklyGoal(next);
  const currentOi = clients.filter(client => inSelectedWeek(client.oiScheduledDate, current)).length;
  const currentPc = clients.filter(client => inSelectedWeek(client.pcScheduledDate, current)).length;
  const nextOi = clients.filter(client => inSelectedWeek(client.oiScheduledDate, next)).length;
  const nextPc = clients.filter(client => inSelectedWeek(client.pcScheduledDate, next)).length;
  const c2Open = clients.filter(client => client.outcome === "C2" && !client.doNotContact);
  const c2Due = c2Open.filter(client => {
    if (!client.followUp) return true;
    if (client.followUp <= current.endKey) return true;
    return (daysSince(client.lastInteraction) ?? 0) >= 7;
  }).length;
  const currentMissing = Math.max(currentGoal - currentOi, 0);
  const nextMissing = Math.max(nextGoal - nextOi, 0);
  const currentLabel = `${shortDate.format(current.start).replace(".", "")} a ${shortDate.format(current.end).replace(".", "")}`;
  const nextLabel = `${shortDate.format(next.start).replace(".", "")} a ${shortDate.format(next.end).replace(".", "")}`;

  setText("todayWeekOi", currentOi);
  setText("todayWeekGoal", currentGoal);
  setText("todayWeekPc", currentPc);
  setText("todayWeekC2", c2Open.length);
  setText("todayNextWeekOi", nextOi);
  setText("todayNextWeekGoal", nextGoal);
  setText("todayNextWeekPc", nextPc);
  setText("todayWeekOiHint", currentMissing
    ? `Faltam ${currentMissing} OI(s) para ${currentGoal}.`
    : `Semana atual protegida: ${currentLabel}.`);
  setText("todayWeekPcHint", currentPc
    ? `${currentPc} PC(s) marcadas. Plano precisa estar pronto.`
    : "Sem PC marcada nesta semana.");
  setText("todayWeekC2Hint", c2Open.length
    ? `${c2Due} pedem follow-up agora.`
    : "Nenhum C2 aberto.");
  setText("todayNextWeekHint", nextMissing
    ? `Faltam ${nextMissing} OI(s) para a próxima semana.`
    : `Próxima semana protegida: ${nextLabel}.`);
  setText("todayNextWeekPcHint", nextPc
    ? `${nextPc} PC(s) já exigem preparo de plano.`
    : "Ainda sem PC na próxima semana.");
  setText("todayWeekSummary", currentMissing || nextMissing
    ? `Esta semana: ${currentOi}/${currentGoal} OIs e ${currentPc} PC(s). Próxima semana: ${nextOi}/${nextGoal} OIs. O jogo agora é marcar um pouco todo dia para não depender só da sexta.`
    : `Agenda forte: ${currentOi}/${currentGoal} OIs nesta semana e ${nextOi}/${nextGoal} OIs na próxima. Agora proteja presença, planos e C2s.`);
}

function renderDashboard() {
  const clients = state.clients;
  const open = clients.filter(c => c.stage !== "N");
  const negotiation = clients.filter(c => ["OI realizada", "PC agendada", "PC realizada", "C2"].includes(c.stage));
  const today = offsetDate(0);
  const pending = clients.filter(c => !c.completed && c.followUp && c.followUp <= today);
  const week = selectedWeek();
  const goal = weeklyGoal(week);
  const ta = clients.filter(c => inSelectedWeek(c.taDate, week)).length;
  const oiScheduled = clients.filter(c => inSelectedWeek(c.oiScheduledDate, week)).length;
  const oiDone = clients.filter(c => c.oiDone && inSelectedWeek(c.oiDate, week)).length;
  const pcScheduled = clients.filter(c => inSelectedWeek(c.pcScheduledDate, week)).length;
  const pcDone = clients.filter(c => c.pcDone && inSelectedWeek(c.pcDate, week)).length;
  const sales = clients.filter(c => c.outcome === "N" && inSelectedWeek(c.outcomeDate, week)).length;
  const c2 = clients.filter(c => c.outcome === "C2" && inSelectedWeek(c.outcomeDate, week)).length;
  const recs = clients
    .filter(c => inSelectedWeek(c.recsDate, week))
    .reduce((total, client) => total + Number(client.recs || 0), 0);
  const oiToPcPending = clients.filter(c => c.oiDone && !c.pcScheduledDate && !c.pcDone && c.outcome !== "N");
  const pcWithoutSale = clients.filter(c => c.pcDone && c.outcome !== "N");
  const staleC2Clients = clients.filter(c => c.outcome === "C2" && (daysSince(c.lastInteraction) ?? 0) >= 6);
  const fridayEligible = clients.filter(c => ["Base", "TA"].includes(c.stage) && !c.oiScheduledDate && !c.doNotContact);
  const fridayRecommendations = fridayEligible.filter(c =>
    c.source === "Indicação" || Boolean(c.referrer) || Number(c.recs || 0) > 0
  );
  const fridayUncalled = fridayEligible.filter(c => !c.lastInteraction && !c.taDate);

  el("weekPeriod").textContent = `${new Intl.DateTimeFormat("pt-BR", { day: "2-digit", month: "short" }).format(week.start)} a ${new Intl.DateTimeFormat("pt-BR", { day: "2-digit", month: "short", year: "numeric" }).format(week.end)}`;
  el("weeklyTa").textContent = ta;
  el("weeklyOiGoal").textContent = oiScheduled;
  el("weeklyGoalInput").value = goal;
  el("weeklyOiGoalBar").style.width = `${Math.min(oiScheduled / goal, 1) * 100}%`;
  el("weeklyOiGoalMessage").textContent = oiScheduled >= goal
    ? `Meta atingida. ${oiScheduled - goal ? `${oiScheduled - goal} OI(s) acima da meta.` : "Agora proteja a presença nas reuniões."}`
    : `Faltam ${goal - oiScheduled} OI(s) agendadas para a meta desta semana.`;
  el("weeklyOiScheduled").textContent = oiScheduled;
  el("weeklyOiDone").textContent = oiDone;
  el("weeklyPcScheduled").textContent = pcScheduled;
  el("weeklyPcDone").textContent = pcDone;
  el("weeklySales").textContent = sales;
  el("weeklyC2").textContent = c2;
  el("weeklyRecs").textContent = recs;
  el("oiShowRate").textContent = `${oiScheduled ? Math.round((oiDone / oiScheduled) * 100) : 0}% presença`;
  el("pcShowRate").textContent = `${pcScheduled ? Math.round((pcDone / pcScheduled) * 100) : 0}% presença`;
  el("salesRate").textContent = `${pcDone ? Math.round((sales / pcDone) * 100) : 0}% conversão`;
  el("pendingOiToPc").textContent = oiToPcPending.length;
  el("pendingPcSale").textContent = pcWithoutSale.length;
  el("staleC2").textContent = staleC2Clients.length;
  el("fridayRecommendations").textContent = fridayRecommendations.length;
  el("fridayUncalled").textContent = fridayUncalled.length;
  el("fridayQueue").textContent = fridayEligible.length;

  el("metricClients").textContent = clients.filter(c => c.stage === "N").length;
  el("metricClientsDetail").textContent = `${clients.length} contatos na base`;
  el("metricNegotiations").textContent = negotiation.length;
  el("metricPending").textContent = pending.length;
  el("metricPendingDetail").textContent = pending.length ? "Precisam da sua atenção" : "Tudo em dia";
  el("metricValue").textContent = currency.format(open.reduce((sum, c) => sum + Number(c.value || 0), 0));
  el("welcomeSummary").textContent = pending.length
    ? `Você tem ${pending.length} ${pending.length === 1 ? "retorno importante" : "retornos importantes"} para cuidar hoje.`
    : "Seus retornos estão em dia. Um ótimo momento para cultivar novas conversas.";

  const upcoming = clients
    .filter(c => c.followUp && !c.completed)
    .sort((a, b) => a.followUp.localeCompare(b.followUp))
    .slice(0, 5);

  el("upcomingList").innerHTML = upcoming.length ? upcoming.map(client => {
    const date = parseLocalDate(client.followUp);
    const status = dateStatus(client.followUp);
    return `
      <div class="upcoming-item">
        <div class="date-chip"><strong>${date.getDate()}</strong><small>${date.toLocaleDateString("pt-BR", { month: "short" }).replace(".", "")}</small></div>
        <div><h4>${escapeHtml(client.name)}</h4><p>${escapeHtml(client.stage)} · ${escapeHtml(client.notes || "Realizar contato")}</p></div>
        <span class="status-label ${status.className}">${status.text}</span>
      </div>`;
  }).join("") : `<div class="mini-empty">Nenhum retorno agendado.</div>`;

  const max = Math.max(...STAGES.map(stage => clients.filter(c => c.stage === stage).length), 1);
  el("funnelChart").innerHTML = STAGES.map(stage => {
    const count = clients.filter(c => c.stage === stage).length;
    return `<div class="funnel-row"><span>${stage}</span><div class="bar-track"><div class="bar-fill" style="width:${(count / max) * 100}%"></div></div><strong>${count}</strong></div>`;
  }).join("");
  renderWeeklyBreakdown(week);
  renderExecutionCockpit();
  renderTodayWeekCommand();
}

function pipelineCardDetail(client, stage) {
  if (stage === "C2") {
    const reason = client.c2Objection || "Motivo não registrado";
    const date = client.followUp ? `retomar ${formatDate(client.followUp)}` : "sem próxima data";
    return `${reason} · ${date}`;
  }
  if (stage === "N") return client.policyStatus || "Acompanhar emissão";
  return `${client.source || "Origem não informada"} · ${client.followUp ? formatDate(client.followUp) : "Sem retorno"}`;
}

function pipelineStageColumn(stage, clients, modifier = "") {
  const stageClients = clients.filter(client => client.stage === stage);
  const readOnly = isCloudReadOnly();
  return `
    <section class="pipeline-column ${modifier}" data-stage="${stage}">
      <div class="column-header"><h3>${stage}</h3><span>${stageClients.length}</span></div>
      <div class="pipeline-cards">
        ${stageClients.map(client => `
          <article class="pipeline-card" draggable="${readOnly ? "false" : "true"}" data-client-id="${client.id}">
            <div class="card-top"><h4>${escapeHtml(client.name)}</h4><i class="priority-dot ${priorityClass(client.priority)}" title="Prioridade ${escapeHtml(client.priority)}"></i></div>
            <p>${escapeHtml(pipelineCardDetail(client, stage))}</p>
            <div class="card-bottom"><strong>${currency.format(client.value || 0)}</strong><button class="text-button edit-card" data-id="${client.id}">Ficha 360</button></div>
          </article>`).join("") || `<div class="mini-empty">Nenhum contato</div>`}
      </div>
    </section>`;
}

const C2_UPDATE_OUTCOMES = {
  advanced: { label: "Avançou", action: "Confirmar fechamento e próximos passos", days: 2 },
  decision_maker: { label: "Aguardando decisor", action: "Retomar com o decisor envolvido", days: 3 },
  asked_time: { label: "Pediu prazo", action: "Retomar decisão no prazo combinado", days: 5 },
  no_reply: { label: "Sem resposta", action: "Fazer nova tentativa de contato", days: 2 },
  won: { label: "Virou N", action: "Acompanhar emissão e entrega da apólice", days: 0 }
};

function isLegacyLightC2(client) {
  if (!client || client.outcome !== "C2" || client.c2LegacyReview) return false;
  return state.activities.some(activity =>
    activity.clientId === client.id
    && normalizedPersonName(String(activity.title || "")).includes("sem interesse virou c2 leve")
  );
}

function c2Priority(client) {
  const today = offsetDate(0);
  const days = assistantDays(client);
  const followUpGap = client.followUp ? Math.round((parseLocalDate(client.followUp) - parseLocalDate(today)) / 86400000) : null;
  let score = 0;
  if (followUpGap === null) score += 42;
  else if (followUpGap < 0) score += Math.min(52, 38 + Math.abs(followUpGap) * 3);
  else if (followUpGap === 0) score += 38;
  else if (followUpGap <= 3) score += 24;
  else if (followUpGap <= 7) score += 12;
  score += days >= 21 ? 24 : days >= 14 ? 19 : days >= 7 ? 12 : days >= 3 ? 5 : 0;
  const reason = normalizedPersonName(client.c2Objection);
  if (!reason) score += 10;
  else if (reason.includes("decisor") || reason.includes("resposta")) score += 12;
  else if (reason.includes("preco") || reason.includes("timing")) score += 8;
  if (client.c2LastOutcome === "advanced") score += 12;
  if (client.c2LastOutcome === "decision_maker") score += 8;
  if (Number(client.value || 0) >= 1000) score += 8;
  else if (Number(client.value || 0) > 0) score += 4;
  score = Math.min(100, score);
  return {
    score,
    label: score >= 70 ? "Alta" : score >= 45 ? "Média" : "Protegida",
    tone: score >= 70 ? "high" : score >= 45 ? "medium" : "low"
  };
}

function c2RecoveryData() {
  const today = offsetDate(0);
  const weekLimit = offsetDate(7);
  const clients = state.clients
    .filter(client => client.outcome === "C2" && !client.doNotContact)
    .sort((a, b) => c2Priority(b).score - c2Priority(a).score || (a.followUp || "0000").localeCompare(b.followUp || "0000"));
  const actNow = clients.filter(client => !client.followUp || client.followUp <= today);
  const nextWeek = clients.filter(client => client.followUp > today && client.followUp <= weekLimit);
  const future = clients.filter(client => client.followUp > weekLimit);
  return {
    clients,
    actNow,
    nextWeek,
    future,
    overdue: clients.filter(client => client.followUp && client.followUp < today),
    noDate: clients.filter(client => !client.followUp),
    legacy: clients.filter(isLegacyLightC2)
  };
}

function c2PerformanceData(data = c2RecoveryData()) {
  const month = offsetDate(0).slice(0, 7);
  const created = state.clients.filter(client => String(client.c2EnteredAt || (client.outcome === "C2" ? client.outcomeDate : "") || "").startsWith(month)).length;
  const recoveredClients = state.clients.filter(client => String(client.c2RecoveredAt || "").startsWith(month));
  const touchActivities = state.activities.filter(activity =>
    String(activity.date || "").startsWith(month)
    && (String(activity.title || "").startsWith("C2: ") || activity.title === "C2 recuperado em N")
  );
  const touchedIds = new Set([
    ...touchActivities.map(activity => activity.clientId).filter(Boolean),
    ...state.clients.filter(client => String(client.c2LastOutcomeDate || "").startsWith(month)).map(client => client.id),
    ...recoveredClients.map(client => client.id)
  ]);
  const recoveryDays = recoveredClients
    .map(client => client.c2EnteredAt ? Math.max(0, Math.round((parseLocalDate(client.c2RecoveredAt) - parseLocalDate(client.c2EnteredAt)) / 86400000)) : null)
    .filter(value => value !== null);
  const average = recoveryDays.length ? Math.round(recoveryDays.reduce((sum, value) => sum + value, 0) / recoveryDays.length) : null;
  const reasons = new Map();
  data.clients.forEach(client => {
    const reason = client.c2Objection || "Não informado";
    reasons.set(reason, (reasons.get(reason) || 0) + 1);
  });
  const mainReason = [...reasons.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || "sem base";
  const recovered = recoveredClients.length;
  const touched = touchedIds.size;
  return {
    created,
    recovered,
    touched,
    recoveryRate: touched ? Math.round((recovered / touched) * 100) : null,
    average,
    open: data.clients.length,
    balance: created - recovered,
    mainReason,
    updates: touchActivities.length
  };
}

function c2RecoveryInsight(data) {
  const performance = c2PerformanceData(data);
  return `Mês: ${performance.created} novo(s) C2 · ${performance.recovered} recuperado(s) em N · ${performance.updates} atualização(ões)${performance.average === null ? "" : ` · ${performance.average}d em média`} · motivo recorrente: ${performance.mainReason}`;
}

function c2RecoveryCard(client) {
  const days = assistantDays(client);
  const followUp = client.followUp ? formatDate(client.followUp) : "Definir data";
  const priority = c2Priority(client);
  const lastOutcome = C2_UPDATE_OUTCOMES[client.c2LastOutcome]?.label;
  return `
    <article class="c2-recovery-card ${isLegacyLightC2(client) ? "legacy-review" : ""}">
      <div class="c2-recovery-card-top">
        <strong>${escapeHtml(client.name)}</strong>
        <span class="c2-priority ${priority.tone}" title="Prioridade calculada por prazo, tempo sem contato, motivo e potencial">${priority.label} ${priority.score}</span>
      </div>
      <p>${escapeHtml(client.c2NextAction || client.nextAction || "Definir próxima ação de fechamento")}</p>
      <small>${escapeHtml(client.c2Objection || "Motivo não registrado")} · ${followUp} · ${days}d sem contato${client.c2TouchCount ? ` · ${client.c2TouchCount} tentativa(s)` : ""}</small>
      ${lastOutcome ? `<em>Último retorno: ${escapeHtml(lastOutcome)}</em>` : ""}
      ${isLegacyLightC2(client) ? `<em class="legacy-label">Revisar: registro antigo de TA</em>` : ""}
      <div class="c2-recovery-actions">
        <button type="button" data-c2-update="${client.id}">Registrar avanço</button>
        <button type="button" data-c2-open="${client.id}">Ficha 360</button>
      </div>
    </article>`;
}

function c2RecoveryLane(key, title, subtitle, clients, tone = "") {
  return `
    <section class="c2-recovery-lane ${tone} ${c2RecoveryView === key ? "active" : ""}" data-c2-lane="${key}">
      <header><div><strong>${title}</strong><small>${subtitle}</small></div><span>${clients.length}</span></header>
      <div class="c2-recovery-list">
        ${clients.length ? clients.map(c2RecoveryCard).join("") : `<div class="mini-empty">Nenhum C2 aqui.</div>`}
      </div>
    </section>`;
}

function setC2UpdateDate(outcome) {
  const config = C2_UPDATE_OUTCOMES[outcome] || C2_UPDATE_OUTCOMES.advanced;
  const dateField = el("c2UpdateNextDate");
  const dateLabel = el("c2UpdateDateLabel");
  if (outcome === "won") {
    dateField.required = false;
    dateField.value = "";
    dateLabel.classList.add("hidden");
    return;
  }
  dateLabel.classList.remove("hidden");
  dateField.required = true;
  dateField.value = businessDateAfterToday(config.days);
}

function openC2UpdateModal(clientId) {
  const client = state.clients.find(item => item.id === clientId);
  if (!client || isCloudReadOnly()) return;
  if (!el("c2UpdateModal") || !el("c2UpdateForm")) {
    openClientModal(client.id);
    switchFormTab("management");
    if (el("commercialProcessSection")) el("commercialProcessSection").open = true;
    showToast("Atualize o C2 pela Ficha 360. Depois, use Atualizar sistema para carregar a interface mais recente.");
    return;
  }
  el("c2UpdateForm").reset();
  el("c2UpdateClientId").value = client.id;
  el("c2UpdateClientName").textContent = client.name;
  const suggested = C2_UPDATE_OUTCOMES[client.c2LastOutcome] ? client.c2LastOutcome : "advanced";
  const option = document.querySelector(`input[name="c2UpdateOutcome"][value="${suggested}"]`);
  if (option) option.checked = true;
  setC2UpdateDate(suggested);
  el("c2LegacyReviewBox").classList.toggle("hidden", !isLegacyLightC2(client));
  el("c2UpdateModal").classList.remove("hidden");
  setTimeout(() => option?.focus(), 60);
}

function closeC2UpdateModal() {
  el("c2UpdateModal")?.classList.add("hidden");
}

function submitC2Update(event) {
  event.preventDefault();
  const client = state.clients.find(item => item.id === el("c2UpdateClientId").value);
  const outcome = new FormData(event.currentTarget).get("c2UpdateOutcome");
  const config = C2_UPDATE_OUTCOMES[outcome];
  if (!client || !config || isCloudReadOnly()) return;
  const today = offsetDate(0);
  const note = el("c2UpdateNote").value.trim();
  const nextDate = el("c2UpdateNextDate").value;
  if (outcome !== "won" && !nextDate) {
    el("c2UpdateNextDate").reportValidity();
    return;
  }

  client.c2EnteredAt ||= client.outcomeDate || today;
  client.c2LastOutcome = outcome;
  client.c2LastOutcomeDate = today;
  client.c2TouchCount = Number(client.c2TouchCount || 0) + 1;
  client.lastInteraction = today;
  if (outcome === "won") {
    client.outcome = "N";
    client.stage = "N";
    client.outcomeDate = today;
    client.c2RecoveredAt = today;
    client.followUp = "";
    client.nextAction = config.action;
    logClientEvent(client.id, "C2 recuperado em N", note || "Venda fechada após acompanhamento C2.", "Historico", today);
  } else {
    client.followUp = nextDate;
    client.c2NextAction = config.action;
    client.nextAction = config.action;
    client.completed = false;
    if (outcome === "decision_maker" && !client.c2Objection) client.c2Objection = "Cônjuge / decisor";
    if (outcome === "no_reply" && !client.c2Objection) client.c2Objection = "Sem resposta";
    const details = [note, `Próximo contato em ${formatDate(nextDate)}.`].filter(Boolean).join(" ");
    logClientEvent(client.id, `C2: ${config.label}`, details, "Follow-up", today);
  }
  saveState();
  closeC2UpdateModal();
  renderAll();
  showToast(outcome === "won" ? "C2 recuperado e convertido em N" : "C2 atualizado com próximo passo");
}

function reviewLegacyC2(action) {
  const client = state.clients.find(item => item.id === el("c2UpdateClientId").value);
  if (!client || !isLegacyLightC2(client) || isCloudReadOnly()) return;
  const today = offsetDate(0);
  if (action === "confirm") {
    client.c2LegacyReview = "confirmed";
    client.c2EnteredAt ||= client.outcomeDate || today;
    logClientEvent(client.id, "C2 antigo revisado", "Confirmado como acompanhamento legítimo após PC.", "Historico", today);
  } else if (action === "ta") {
    client.c2LegacyReview = "ta";
    client.outcome = "";
    client.outcomeDate = "";
    client.stage = client.taDate ? "TA" : "Base";
    client.followUp = nextBusinessDate();
    client.nextAction = "Retomar abordagem de TA";
    logClientEvent(client.id, "C2 antigo reclassificado", "Registro devolvido para TA; histórico preservado.", "Historico", today);
  } else if (action === "do_not_contact") {
    client.c2LegacyReview = "do_not_contact";
    client.doNotContact = true;
    client.followUp = "";
    client.nextAction = "Não contatar";
    logClientEvent(client.id, "C2 antigo revisado", "Contato marcado como não contatar; histórico preservado.", "Historico", today);
  }
  saveState();
  closeC2UpdateModal();
  renderAll();
  showToast("Registro antigo revisado sem apagar o histórico");
}

function renderC2Recovery() {
  if (!el("c2RecoveryBoard")) return;
  const data = c2RecoveryData();
  if (el("c2RecoveryTotal")) el("c2RecoveryTotal").textContent = data.clients.length;
  if (el("c2RecoveryOverdue")) el("c2RecoveryOverdue").textContent = data.overdue.length;
  if (el("c2RecoveryWeek")) el("c2RecoveryWeek").textContent = data.nextWeek.length;
  if (el("c2RecoveryNoDate")) el("c2RecoveryNoDate").textContent = data.noDate.length;
  if (el("c2RecoveryInsight")) el("c2RecoveryInsight").textContent = c2RecoveryInsight(data);
  if (el("c2LegacyReview")) el("c2LegacyReview").classList.toggle("hidden", !data.legacy.length);
  if (el("c2LegacyReviewCount")) el("c2LegacyReviewCount").textContent = data.legacy.length;
  el("c2RecoveryBoard").innerHTML = [
    c2RecoveryLane("now", "Agir agora", "Atrasados ou sem data", data.actNow, "danger"),
    c2RecoveryLane("week", "Próximos 7 dias", "Follow-ups já combinados", data.nextWeek, "warn"),
    c2RecoveryLane("future", "Futuro", "Carteira protegida", data.future, "calm")
  ].join("");
  document.querySelectorAll("[data-c2-open]").forEach(button => button.addEventListener("click", () => {
    openClientModal(button.dataset.c2Open);
    switchFormTab("management");
    el("commercialProcessSection").open = true;
  }));
  document.querySelectorAll("[data-c2-update]").forEach(button => button.addEventListener("click", () => openC2UpdateModal(button.dataset.c2Update)));
  document.querySelectorAll("[data-c2-view]").forEach(button => {
    button.classList.toggle("active", button.dataset.c2View === c2RecoveryView);
  });
}

function renderPipeline() {
  const priority = el("pipelineFilter").value;
  const clients = filteredClients().filter(client => priority === "all" || client.priority === priority);
  const processStages = STAGES.slice(0, 6);
  el("pipelineBoard").innerHTML = `
    <section class="pipeline-process">
      <div class="pipeline-flow-label"><span>PROCESSO COMERCIAL</span><small>Cada avanço abre o próximo compromisso.</small></div>
      <div class="pipeline-flow-columns">${processStages.map(stage => pipelineStageColumn(stage, clients)).join("")}</div>
    </section>
    <section class="pipeline-decision">
      <div class="pipeline-decision-heading"><span>DECISÃO APÓS O PC</span><small>Venda fechada ou acompanhamento com motivo, ação e data.</small></div>
      <div class="pipeline-outcome-columns">
        ${pipelineStageColumn("N", clients, "outcome-n")}
        ${pipelineStageColumn("C2", clients, "outcome-c2")}
      </div>
    </section>`;
  bindPipelineEvents();
  renderC2Recovery();
}

function renderClients() {
  const operationalViewer = isCloudReadOnly();
  const stage = el("stageFilter").value;
  const attention = el("attentionFilter").value;
  let referrerFilter = el("referrerFilter")?.value || "all";
  let sort = el("sortClients").value;
  if (operationalViewer && sort === "score") {
    sort = "name";
    el("sortClients").value = "name";
  }
  const today = offsetDate(0);
  const currentReferrerValue = el("referrerFilter")?.value || "all";
  const referrers = [...new Set(state.clients.map(client => client.referrer).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, "pt-BR"));
  if (el("referrerFilter")) {
    el("referrerFilter").innerHTML = `<option value="all">Todos os recomendantes</option>` +
      referrers.map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
    el("referrerFilter").value = referrers.includes(currentReferrerValue) ? currentReferrerValue : "all";
    referrerFilter = el("referrerFilter").value;
  }
  let clients = filteredClients().filter(client => {
    if (stage !== "all" && client.stage !== stage) return false;
    if (referrerFilter !== "all" && client.referrer !== referrerFilter) return false;
    if (workFilter === "oi-pending") return client.oiDone && !client.pcScheduledDate && !client.pcDone && client.outcome !== "N";
    if (workFilter === "pc-pending") return client.pcDone && client.outcome !== "N";
    if (workFilter === "c2-all") return client.outcome === "C2";
    if (workFilter === "c2-stale") return client.outcome === "C2" && (daysSince(client.lastInteraction) ?? 0) >= 6;
    if (workFilter === "no-next-step") return client.outcome !== "N" && (!client.nextAction || !client.followUp);
    if (workFilter === "plan-overdue") return client.outcome !== "N" && client.oiDone && client.planStatus !== "Pronto" && client.planDueDate && client.planDueDate < today;
    if (workFilter === "pc-no-outcome") return client.pcDone && !client.outcome;
    if (workFilter === "friday-recs") return ["Base", "TA"].includes(client.stage) && !client.oiScheduledDate && !client.doNotContact
      && (String(client.source || "").includes("Indica") || Boolean(client.referrer) || Number(client.recs || 0) > 0);
    if (workFilter === "friday-uncalled") return ["Base", "TA"].includes(client.stage) && !client.oiScheduledDate && !client.doNotContact
      && !client.lastInteraction && !client.taDate;
    if (workFilter === "friday-all") return ["Base", "TA"].includes(client.stage) && !client.oiScheduledDate && !client.doNotContact;
    if (attention === "today") return !client.completed && client.followUp === today;
    if (attention === "overdue") return !client.completed && client.followUp && client.followUp < today;
    if (attention === "no-action") return !client.nextAction || !client.followUp;
    if (attention === "high") return client.priority === "Alta";
    return true;
  });
  clients.sort((a, b) => {
    if (workFilter.startsWith("friday-")) {
      const fridayScore = client => {
        const recommendation = String(client.source || "").includes("Indica") || Boolean(client.referrer) || Number(client.recs || 0) > 0;
        if (recommendation && !client.lastInteraction) return 0;
        if (recommendation) return 1;
        if (!client.lastInteraction && !client.taDate) return 2;
        return 3;
      };
      const scoreDifference = fridayScore(a) - fridayScore(b);
      if (scoreDifference) return scoreDifference;
    }
    if (sort === "name") return a.name.localeCompare(b.name, "pt-BR");
    if (sort === "followup") return (a.followUp || "9999").localeCompare(b.followUp || "9999");
    if (sort === "score") return sitplanOpportunityScore(b).score - sitplanOpportunityScore(a).score;
    return b.createdAt.localeCompare(a.createdAt);
  });

  const totalClients = clients.length;
  const totalPages = Math.max(1, Math.ceil(totalClients / CLIENTS_PAGE_SIZE));
  clientsPage = Math.min(Math.max(clientsPage, 1), totalPages);
  const pageStart = (clientsPage - 1) * CLIENTS_PAGE_SIZE;
  const pageClients = clients.slice(pageStart, pageStart + CLIENTS_PAGE_SIZE);

  const overdueCount = state.clients.filter(c => !c.completed && c.followUp && c.followUp < today).length;
  el("baseTotal").textContent = state.clients.length;
  el("baseDueToday").textContent = state.clients.filter(c => !c.completed && c.followUp === today).length;
  el("baseOverdue").textContent = overdueCount;
  el("baseNoAction").textContent = state.clients.filter(c => !c.nextAction || !c.followUp).length;
  const workLabels = {
    "friday-recs": "Recomendações pendentes para sexta-feira",
    "friday-uncalled": "Contatos ainda sem ligação",
    "friday-all": "Fila completa de TA para conquistar OIs",
    "oi-pending": "OI realizada sem PC marcada",
    "pc-pending": "PC realizada sem conversão em N",
    "c2-all": "Todos os C2 abertos",
    "c2-stale": "C2 sem contato há 6+ dias",
    "no-next-step": "Contatos sem próxima ação",
    "plan-overdue": "Planos com prazo vencido",
    "pc-no-outcome": "PC realizada sem decisão"
  };
  el("clientCountLabel").textContent = workFilter
    ? `${totalClients} - ${workLabels[workFilter]}`
    : `${totalClients} ${totalClients === 1 ? "lead exibido" : "leads exibidos"}`;
  el("clientsTableBody").innerHTML = pageClients.map(client => {
    const days = daysSince(client.lastInteraction);
    const alert = contactAlert(client);
    const directWhatsapp = whatsappLink(client);
    const missedWhatsapp = whatsappLink(client, true);
    const opportunity = operationalViewer ? null : sitplanOpportunityScore(client);
    const movement = operationalViewer ? null : nextBestMovement(client);
    return `
    <tr class="${!client.completed && client.followUp && client.followUp < today ? "attention-row" : client.followUp === today ? "today-row" : ""}">
      <td class="client-cell">
        <strong>${escapeHtml(client.name)}</strong>
        <small>${escapeHtml(client.referrer ? `Indicado por ${client.referrer}` : client.source || "Origem não informada")}</small>
        ${client.referrerRelationship ? `<em class="relationship-mini">${escapeHtml(client.referrerRelationship)}</em>` : ""}
      </td>
      <td class="contact-lines">
        <strong>${escapeHtml(client.whatsapp || client.phone || "-")}</strong>
        <small>${escapeHtml(client.city || client.profession || "Contexto a completar")}</small>
        ${directWhatsapp
          ? `<a class="whatsapp-action" href="${directWhatsapp}" target="_blank" rel="noopener">Chamar no WhatsApp</a>
             ${operationalViewer ? "" : `<button class="missed-call-action" data-missed-call="${client.id}" data-whatsapp-url="${missedWhatsapp}">Não atendeu</button>`}`
          : `<small class="phone-warning">Falta DDD para abrir o WhatsApp</small>`}
      </td>
      <td class="crm-stage-cell"><span class="stage-readonly">${escapeHtml(client.stage)}</span><span class="temperature-pill ${temperatureClass(client.temperature)}">${escapeHtml(client.temperature)}</span>${operationalViewer ? "" : `<small class="score-mini">Score ${opportunity.score}</small>`}</td>
      <td class="next-action-cell"><strong>${escapeHtml(client.nextAction || movement?.action || "Definir próximo passo")}</strong><small>${client.followUp ? formatDate(client.followUp) : "Sem data definida"}</small></td>
      <td class="crm-signal-cell"><span class="alert-pill ${alert.className}">${alert.text}</span><small>${days === null ? "Sem contato registrado" : `${days} dia(s) sem contato`}</small></td>
      <td>${operationalViewer ? "" : `<div class="row-actions"><button data-input-client="${client.id}">Registrar</button><button data-edit="${client.id}">Ficha 360</button></div>`}</td>
    </tr>`;
  }).join("");
  el("clientsMobileList").innerHTML = pageClients.map(client => {
    const movement = operationalViewer ? null : nextBestMovement(client);
    const alert = contactAlert(client);
    const directWhatsapp = whatsappLink(client);
    return `<article class="client-mobile-card">
      <div class="client-mobile-heading">
        <div><strong>${escapeHtml(client.name)}</strong><small>${escapeHtml(client.referrer ? `Indicado por ${client.referrer}` : client.source || "Origem não informada")}</small></div>
        <span class="alert-pill ${alert.className}">${alert.text}</span>
      </div>
      ${client.referrerRelationship || client.referralContext ? `<p class="client-mobile-origin">${escapeHtml([client.referrerRelationship, client.referralContext].filter(Boolean).join(" · "))}</p>` : ""}
      <div class="client-mobile-meta">
        <span>${escapeHtml(client.stage)}</span>
        <span>${client.followUp ? `Retorno ${formatDate(client.followUp)}` : "Sem retorno"}</span>
        <span>${escapeHtml(client.temperature || "Morno")}</span>
      </div>
      ${operationalViewer ? "" : `<p>${escapeHtml(movement.action)} · ${escapeHtml(movement.reason)}</p>`}
      <div class="client-mobile-actions">
        ${directWhatsapp ? `<a href="${directWhatsapp}" target="_blank" rel="noopener">WhatsApp</a>` : ""}
        ${operationalViewer ? "" : `<button data-input-client="${client.id}">Lançar</button><button data-edit="${client.id}">Abrir ficha</button>`}
      </div>
    </article>`;
  }).join("");
  el("clientsEmpty").classList.toggle("hidden", totalClients > 0);
  el("clientsPageStatus").textContent = totalClients
    ? `Página ${clientsPage} de ${totalPages} · ${pageStart + 1}-${Math.min(pageStart + CLIENTS_PAGE_SIZE, totalClients)} de ${totalClients}`
    : "Nenhum contato";
  el("clientsPreviousPage").disabled = clientsPage <= 1;
  el("clientsNextPage").disabled = clientsPage >= totalPages;
  if (operationalViewer) return;
  document.querySelectorAll("[data-edit]").forEach(button => button.addEventListener("click", () => openClientModal(button.dataset.edit)));
  document.querySelectorAll("[data-input-client]").forEach(button => button.addEventListener("click", () => {
    switchView("input");
    setQuickInputClient(button.dataset.inputClient);
  }));
  document.querySelectorAll("[data-missed-call]").forEach(button => button.addEventListener("click", () => {
    const client = state.clients.find(item => item.id === button.dataset.missedCall);
    if (!client) return;
    const today = offsetDate(0);
    client.taDate = today;
    client.lastInteraction = today;
    client.stage = "TA";
    client.followUp = nextBusinessDate();
    client.nextAction = "Aguardar resposta no WhatsApp e tentar nova ligação";
    client.completed = false;
    logClientEvent(client.id, "Ligação sem atendimento", "WhatsApp de retorno preparado e nova tentativa agendada.", "Ligacao");
    saveState();
    window.open(button.dataset.whatsappUrl, "_blank", "noopener");
    renderAll();
    showToast("TA registrada e WhatsApp aberto");
  }));
}
function renderTasks() {
  let clients = state.clients.filter(c => c.followUp);
  if (taskFilter === "pending") clients = clients.filter(c => !c.completed);
  if (taskFilter === "completed") clients = clients.filter(c => c.completed);
  clients.sort((a, b) => a.followUp.localeCompare(b.followUp));

  el("tasksList").innerHTML = clients.length ? clients.map(client => {
    const status = dateStatus(client.followUp);
    return `
      <article class="task-item ${client.completed ? "completed" : ""}">
        <input class="task-check" type="checkbox" data-complete="${client.id}" ${client.completed ? "checked" : ""} aria-label="Concluir contato com ${escapeHtml(client.name)}">
        <div><h4>Contato com ${escapeHtml(client.name)}</h4><p>${escapeHtml(client.nextAction || client.notes || `Próximo passo da etapa ${client.stage}`)}</p></div>
        <div class="task-date"><strong class="${status.className}">${status.text}</strong><br>${formatDate(client.followUp)}</div>
      </article>`;
  }).join("") : `<div class="mini-empty">Nenhuma atividade nesta categoria.</div>`;

  document.querySelectorAll("[data-complete]").forEach(check => check.addEventListener("change", () => {
    const client = state.clients.find(c => c.id === check.dataset.complete);
    client.completed = check.checked;
    saveState();
    renderAll();
    showToast(check.checked ? "Atividade concluída" : "Atividade reaberta");
  }));

  const completedToday = state.clients.filter(c => c.completed && c.followUp === offsetDate(0)).length;
  el("focusText").textContent = `${Math.min(completedToday, 3)} de 3 concluídos hoje`;
  el("focusBar").style.width = `${Math.min(completedToday / 3, 1) * 100}%`;
}

renderTasks = function renderAgenda() {
  const contactTasks = state.clients.filter(client => client.followUp).map(client => ({
    id: client.id,
    source: "contact",
    title: `Contato com ${client.name}`,
    notes: client.nextAction || client.notes || `Proximo passo da etapa ${client.stage}`,
    type: client.stage,
    date: client.followUp,
    time: "",
    completed: Boolean(client.completed)
  }));
  const customTasks = state.activities.map(activity => ({ ...activity, source: "activity" }));
  let tasks = [...contactTasks, ...customTasks];
  if (taskFilter === "pending") tasks = tasks.filter(task => !task.completed);
  if (taskFilter === "completed") tasks = tasks.filter(task => task.completed);
  tasks.sort((a, b) => `${a.date}${a.time || ""}`.localeCompare(`${b.date}${b.time || ""}`));

  el("tasksList").innerHTML = tasks.length ? tasks.map(task => {
    const status = dateStatus(task.date);
    return `
      <article class="task-item ${task.completed ? "completed" : ""}">
        <input class="task-check" type="checkbox" data-task-source="${task.source}" data-task-complete="${task.id}" ${task.completed ? "checked" : ""} aria-label="Concluir ${escapeHtml(task.title)}">
        <div><h4>${escapeHtml(displaySystemText(task.title))}</h4><p><span class="activity-kind">${escapeHtml(displayActivityType(task.type || "Atividade"))}</span>${escapeHtml(displaySystemText(task.notes || "Sem observações"))}</p></div>
        <div class="task-date"><strong class="${status.className}">${status.text}</strong><br>${task.time ? `${escapeHtml(task.time)} - ` : ""}${formatDate(task.date)}</div>
        ${task.source === "activity"
          ? `<div class="task-row-actions">${["OI", "PC"].includes(task.type) && task.clientId ? `<button class="task-edit" data-meeting-task="${task.clientId}" data-meeting-kind="${task.type}">Confirmar</button>` : ""}<button class="task-edit" data-edit-activity="${task.id}">Editar</button></div>`
          : `<button class="task-edit" data-edit="${task.id}">Contato</button>`}
      </article>`;
  }).join("") : `<div class="mini-empty">Nenhuma atividade nesta categoria.</div>`;

  document.querySelectorAll("[data-task-complete]").forEach(check => check.addEventListener("change", () => {
    const collection = check.dataset.taskSource === "activity" ? state.activities : state.clients;
    const item = collection.find(entry => entry.id === check.dataset.taskComplete);
    if (!item) return;
    item.completed = check.checked;
    if (check.dataset.taskSource === "activity" && item.type === "Plano" && item.clientId) {
      const client = state.clients.find(entry => entry.id === item.clientId);
      if (client && check.checked) markPlanReady(client);
      if (client && !check.checked) {
        client.planStatus = "Em montagem";
        client.planReadyDate = "";
      }
    }
    saveState();
    renderAll();
    showToast(check.checked ? "Atividade concluída" : "Atividade reaberta");
  }));
  document.querySelectorAll("[data-edit-activity]").forEach(button => button.addEventListener("click", () => openActivityModal(button.dataset.editActivity)));
  document.querySelectorAll("[data-meeting-task]").forEach(button => button.addEventListener("click", () => {
    openClientModal(button.dataset.meetingTask);
    switchFormTab("management");
    openMeetingModal(button.dataset.meetingKind);
  }));
  document.querySelectorAll("#tasksList [data-edit]").forEach(button => button.addEventListener("click", () => openClientModal(button.dataset.edit)));

  const completedToday = tasks.filter(task => task.completed && task.date === offsetDate(0)).length;
  el("focusText").textContent = `${Math.min(completedToday, 3)} de 3 concluídos hoje`;
  el("focusBar").style.width = `${Math.min(completedToday / 3, 1) * 100}%`;
};

function selectedAgendaWeek() {
  const start = startOfWeek(agendaWeekOffset);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  return { start, end, startKey: dateKey(start), endKey: dateKey(end) };
}

function agendaTasksData() {
  const actionableActivities = state.activities
    .filter(activity => activity.type !== "Historico")
    .filter(activity => !(activity.completed && ["Ligacao", "WhatsApp"].includes(activity.type)))
    .map(activity => ({ ...activity, source: "activity" }));
  const activityKeys = new Set(actionableActivities
    .filter(activity => !activity.completed && activity.clientId)
    .map(activity => `${activity.clientId}|${activity.date}`));
  const contactTasks = state.clients
    .filter(client => client.followUp)
    .filter(client => !activityKeys.has(`${client.id}|${client.followUp}`))
    .map(client => ({
      id: client.id,
      source: "contact",
      clientId: client.id,
      title: `Contato com ${client.name}`,
      notes: client.nextAction || client.notes || `Proximo passo da etapa ${client.stage}`,
      type: client.stage,
      date: client.followUp,
      time: "",
      completed: Boolean(client.completed)
    }));
  return [...contactTasks, ...actionableActivities];
}

function agendaPremiumData() {
  const today = offsetDate(0);
  const nextBusiness = nextBusinessDate();
  const tasks = agendaTasksData();
  const todayTasks = tasks.filter(task => !task.completed && task.date === today);
  const overdueTasks = tasks.filter(task => !task.completed && task.date < today);
  const meetingsToday = tasks.filter(task => !task.completed && task.date === today && ["OI", "PC"].includes(task.type));
  const calls = executionCallClients();
  const plans = executionPlanClients();
  const c2s = executionC2Clients();
  const policyPending = state.clients.filter(client => client.outcome === "N" && !["Emitida", "Entregue"].includes(client.policyStatus) && !client.policyDelivered);
  const aftercare = state.clients.filter(client => client.outcome === "N" && client.policyStatus === "Emitida" && !client.policyDelivered);
  const todayMetrics = [
    {
      key: "meetings",
      label: "Reuniões hoje",
      value: meetingsToday.length,
      detail: meetingsToday.length ? "Confirmar presença e preparar contexto." : "Sem OI/PC pendente hoje.",
      tone: meetingsToday.length ? "good" : "neutral"
    },
    {
      key: "calls",
      label: "Ligar / TA",
      value: calls.length,
      detail: calls.length ? "Prioridade comercial para abrir agenda." : "Sem TA crítica agora.",
      tone: calls.length ? "warn" : "neutral"
    },
    {
      key: "plans",
      label: "Planos",
      value: plans.length,
      detail: plans.length ? "Montar antes do PC esfriar." : "Nenhum plano pendente.",
      tone: plans.length ? "warn" : "good"
    },
    {
      key: "c2",
      label: "C2",
      value: c2s.length,
      detail: c2s.length ? "Follow-up para buscar decisão." : "Nenhum C2 crítico.",
      tone: c2s.length ? "risk" : "good"
    },
    {
      key: "policy",
      label: "Apólices",
      value: policyPending.length + aftercare.length,
      detail: policyPending.length ? `${policyPending.length} a emitir, ${aftercare.length} a entregar.` : `${aftercare.length} a entregar.`,
      tone: policyPending.length ? "warn" : aftercare.length ? "neutral" : "good"
    }
  ];
  const lanes = [
    {
      key: "calls",
      title: "Ligar / TA",
      count: calls.length,
      empty: "Sem ligação crítica agora.",
      items: calls.slice(0, 5).map(client => ({
        id: client.id,
        title: client.name,
        detail: assistantAction(client).action,
        action: "Abrir",
        view: "client"
      }))
    },
    {
      key: "meetings",
      title: "Reuniões OI/PC",
      count: meetingsToday.length,
      empty: "Sem reunião pendente hoje.",
      items: meetingsToday.slice(0, 5).map(task => ({
        id: task.clientId || task.id,
        title: displaySystemText(task.title),
        detail: `${displayActivityType(task.type)} ${task.time ? `às ${task.time}` : "hoje"}`,
        action: task.clientId ? "Abrir" : "Editar",
        view: task.clientId ? "client" : "activity",
        activityId: task.id
      }))
    },
    {
      key: "plans",
      title: "Montar plano",
      count: plans.length,
      empty: "Nenhum plano pendente.",
      items: plans.slice(0, 5).map(client => ({
        id: client.id,
        title: client.name,
        detail: client.pcScheduledDate ? `PC ${formatDate(client.pcScheduledDate)}` : "OI realizada, plano pendente",
        action: "Abrir",
        view: "client"
      }))
    },
    {
      key: "c2",
      title: "C2 / Follow-up",
      count: c2s.length,
      empty: "Nenhum C2 crítico.",
      items: c2s.slice(0, 5).map(client => ({
        id: client.id,
        title: client.name,
        detail: client.followUp ? `Follow-up ${formatDate(client.followUp)}` : "Sem data definida",
        action: "Abrir",
        view: "client"
      }))
    },
    {
      key: "policy",
      title: "Apólice / Pós-venda",
      count: policyPending.length + aftercare.length,
      empty: "Nada crítico em apólices.",
      items: [...policyPending.slice(0, 3), ...aftercare.slice(0, 2)].map(client => ({
        id: client.id,
        title: client.name,
        detail: client.policyStatus === "Emitida" ? "Emitida, falta entregar" : "N fechado, falta emissão",
        action: "Abrir",
        view: "client"
      }))
    }
  ];
  return { today, nextBusiness, tasks, todayTasks, overdueTasks, meetingsToday, todayMetrics, lanes };
}

function renderAgendaPremium() {
  if (!el("agendaTodayGrid")) return;
  const data = agendaPremiumData();
  const total = data.todayMetrics.reduce((sum, item) => sum + Number(item.value || 0), 0);
  el("agendaTodaySummary").textContent = total
    ? `${total} frente(s) comerciais para proteger hoje. Comece pelo item com alerta mais forte.`
    : "Hoje está limpo. Use a janela para criar OIs da próxima semana e pedir recomendações.";
  el("agendaTodayGrid").innerHTML = data.todayMetrics.map(item => `
    <article class="agenda-today-card ${item.tone}">
      <small>${escapeHtml(item.label)}</small>
      <strong>${item.value}</strong>
      <span>${escapeHtml(item.detail)}</span>
    </article>
  `).join("");
  el("agendaKanbanGrid").innerHTML = data.lanes.map(lane => `
    <section class="agenda-kanban-lane ${lane.key}">
      <div class="agenda-kanban-head">
        <strong>${escapeHtml(lane.title)}</strong>
        <span>${lane.count}</span>
      </div>
      <div class="agenda-kanban-list">
        ${lane.items.length ? lane.items.map(item => `
          <article class="agenda-kanban-item">
            <div>
              <strong>${escapeHtml(item.title)}</strong>
              <small>${escapeHtml(item.detail)}</small>
            </div>
            <button data-agenda-premium-open="${escapeHtml(item.id)}" data-agenda-premium-view="${escapeHtml(item.view)}" ${item.activityId ? `data-agenda-premium-activity="${escapeHtml(item.activityId)}"` : ""}>${escapeHtml(item.action)}</button>
          </article>
        `).join("") : `<div class="mini-empty">${escapeHtml(lane.empty)}</div>`}
      </div>
    </section>
  `).join("");
  document.querySelectorAll("[data-agenda-premium-open]").forEach(button => button.addEventListener("click", () => {
    if (button.dataset.agendaPremiumView === "activity") {
      openActivityModal(button.dataset.agendaPremiumActivity || button.dataset.agendaPremiumOpen);
      return;
    }
    openClientModal(button.dataset.agendaPremiumOpen);
  }));
}

function agendaGroupTitle(date, overdue = false) {
  if (overdue) return "Atrasados - resolver primeiro";
  const parsed = parseLocalDate(date);
  const today = offsetDate(0);
  const tomorrow = offsetDate(1);
  if (date === today) return "Hoje";
  if (date === tomorrow) return "Amanha";
  return parsed.toLocaleDateString("pt-BR", { weekday: "long", day: "2-digit", month: "short" });
}

function planChecklistProgress(client) {
  const items = [client.planProfile, client.planGoals, client.planProtection, client.planBudget, client.planAlternatives];
  return items.filter(Boolean).length;
}

function renderDailyProductionQueues() {
  if (!el("dailyPlansQueue")) return;
  const today = offsetDate(0);
  const plans = state.clients
    .filter(client => client.oiDone && client.planStatus !== "Pronto" && client.outcome !== "N")
    .map(client => {
      client.planStatus ||= "Em montagem";
      client.planDueDate ||= planDueDateFor(client);
      ensurePlanTask(client);
      return { client, due: client.planDueDate };
    })
    .sort((a, b) =>
      (a.client.pcScheduledDate || "9999").localeCompare(b.client.pcScheduledDate || "9999")
      || a.due.localeCompare(b.due)
      || (a.client.oiDate || "9999").localeCompare(b.client.oiDate || "9999")
    );
  const c2s = state.clients
    .filter(client => client.outcome === "C2")
    .sort((a, b) =>
      (a.followUp || "0000").localeCompare(b.followUp || "0000")
      || Number(b.value || 0) - Number(a.value || 0)
      || assistantDays(b) - assistantDays(a)
    );

  el("dailyPlansCount").textContent = plans.length;
  el("dailyC2Count").textContent = c2s.length;
  el("dailyPlansQueue").innerHTML = plans.length ? plans.slice(0, 6).map(({ client, due }) => {
    const progress = planChecklistProgress(client);
    const pcRisk = client.pcScheduledDate && client.planStatus !== "Pronto";
    const alert = pcRisk && due <= today ? "PC próximo sem plano pronto" : due < today ? "Plano atrasado" : `Prazo ${formatDate(due)}`;
    return `<article class="daily-production-item ${due <= today ? "urgent" : ""}">
      <div>
        <strong>${escapeHtml(client.name)}</strong>
        <small>${escapeHtml(alert)} · checklist ${progress}/5${client.pcScheduledDate ? ` · PC ${formatDate(client.pcScheduledDate)}` : ""}</small>
      </div>
      <div class="daily-production-actions">
        <button data-open-plan="${client.id}">Abrir</button>
        <button class="complete" data-plan-ready="${client.id}">Plano pronto</button>
      </div>
    </article>`;
  }).join("") : `<div class="mini-empty">Nenhum plano pendente.</div>`;

  el("dailyC2Queue").innerHTML = c2s.length ? c2s.slice(0, 6).map(client => {
    const overdue = !client.followUp || client.followUp <= today;
    const whatsapp = whatsappLink(client);
    return `<article class="daily-production-item ${overdue ? "urgent" : ""}">
      <div>
        <strong>${escapeHtml(client.name)}</strong>
        <small>${client.followUp ? `Follow-up ${formatDate(client.followUp)}` : "Sem próxima data"} · ${assistantDays(client)}d sem contato</small>
      </div>
      <div class="daily-production-actions">
        ${whatsapp ? `<a href="${whatsapp}" target="_blank" rel="noopener">WhatsApp</a>` : ""}
        <button data-open-c2="${client.id}">Abrir</button>
        <button class="complete" data-c2-next="${client.id}">Registrar avanço</button>
      </div>
    </article>`;
  }).join("") : `<div class="mini-empty">Nenhum C2 pendente.</div>`;

  document.querySelectorAll("[data-open-plan], [data-open-c2]").forEach(button => button.addEventListener("click", () => {
    openClientModal(button.dataset.openPlan || button.dataset.openC2);
    switchFormTab("management");
  }));
  document.querySelectorAll("[data-plan-ready]").forEach(button => button.addEventListener("click", () => {
    const client = state.clients.find(item => item.id === button.dataset.planReady);
    if (!client) return;
    markPlanReady(client);
    saveState();
    renderAll();
    showToast("Plano marcado como pronto");
  }));
  document.querySelectorAll("[data-c2-next]").forEach(button => button.addEventListener("click", () => {
    openC2UpdateModal(button.dataset.c2Next);
  }));
}

renderTasks = function renderWeeklyAgenda() {
  renderDailyProductionQueues();
  renderAgendaPremium();
  const week = selectedAgendaWeek();
  const today = offsetDate(0);
  const allTasks = agendaTasksData();
  const weekTasks = allTasks.filter(task => task.date >= week.startKey && task.date <= week.endKey);
  const overdueTasks = agendaWeekOffset === 0
    ? allTasks.filter(task => !task.completed && task.date < today)
    : [];
  const pendingWeek = weekTasks.filter(task => !task.completed);
  const completedWeek = weekTasks.filter(task => task.completed);
  const meetingsWeek = weekTasks.filter(task => ["OI", "PC"].includes(task.type));

  el("agendaWeekPeriod").textContent = `${shortDate.format(week.start)} a ${new Intl.DateTimeFormat("pt-BR", { day: "2-digit", month: "short", year: "numeric" }).format(week.end)}`;
  el("agendaOverdue").textContent = overdueTasks.length;
  el("agendaPendingWeek").textContent = pendingWeek.length;
  el("agendaMeetingsWeek").textContent = meetingsWeek.length;
  el("agendaCompletedWeek").textContent = completedWeek.length;

  let visible = [...overdueTasks, ...weekTasks];
  const unique = new Map(visible.map(task => [`${task.source}-${task.id}`, task]));
  visible = [...unique.values()];
  if (taskFilter === "pending") visible = visible.filter(task => !task.completed);
  if (taskFilter === "completed") visible = visible.filter(task => task.completed);
  visible.sort((a, b) => `${a.date}${a.time || ""}`.localeCompare(`${b.date}${b.time || ""}`));

  const groups = [];
  if (agendaWeekOffset === 0) {
    const overdue = visible.filter(task => task.date < today);
    if (overdue.length) groups.push({ key: "overdue", title: agendaGroupTitle("", true), tasks: overdue });
  }
  for (let date = new Date(week.start); date <= week.end; date.setDate(date.getDate() + 1)) {
    const key = dateKey(date);
    const tasks = visible.filter(task => task.date === key);
    if (tasks.length) groups.push({ key, title: agendaGroupTitle(key), tasks });
  }

  const taskMarkup = task => {
    const status = dateStatus(task.date);
    const relatedClient = task.clientId ? state.clients.find(client => client.id === task.clientId) : null;
    const pcWithoutPlan = task.type === "PC" && relatedClient && relatedClient.planStatus !== "Pronto";
    return `<article class="task-item ${task.completed ? "completed" : ""}">
      <input class="task-check" type="checkbox" data-task-source="${task.source}" data-task-complete="${task.id}" ${task.completed ? "checked" : ""} aria-label="Concluir ${escapeHtml(task.title)}">
        <div><h4>${escapeHtml(displaySystemText(task.title))}</h4><p><span class="activity-kind">${escapeHtml(displayActivityType(task.type || "Atividade"))}</span>${escapeHtml(displaySystemText(task.notes || "Sem observações"))}${pcWithoutPlan ? `<strong class="plan-warning">PC sem plano pronto</strong>` : ""}</p></div>
      <div class="task-date"><strong class="${status.className}">${status.text}</strong><br>${task.time ? `${escapeHtml(task.time)} - ` : ""}${formatDate(task.date)}</div>
      <div class="task-row-actions">
        ${["OI", "PC"].includes(task.type) && task.clientId ? `<button class="task-edit" data-meeting-task="${task.clientId}" data-meeting-kind="${task.type}">Confirmar</button>` : ""}
        <label class="task-reschedule">Remarcar<input type="date" value="${task.date}" data-task-reschedule="${task.id}" data-task-source="${task.source}"></label>
        ${task.source === "activity"
          ? `<button class="task-edit" data-edit-activity="${task.id}">Editar</button>`
          : `<button class="task-edit" data-edit="${task.id}">Contato</button>`}
      </div>
    </article>`;
  };

  el("tasksList").innerHTML = groups.length
    ? groups.map(group => `<section class="agenda-day-group ${group.key === "overdue" ? "overdue-group" : ""}">
        <div class="agenda-day-heading"><h4>${escapeHtml(group.title)}</h4><span>${group.tasks.length}</span></div>
        <div class="agenda-day-tasks">${group.tasks.map(taskMarkup).join("")}</div>
      </section>`).join("")
    : `<div class="mini-empty">Nenhum to do nesta semana e categoria.</div>`;

  document.querySelectorAll("[data-task-complete]").forEach(check => check.addEventListener("change", () => {
    const collection = check.dataset.taskSource === "activity" ? state.activities : state.clients;
    const item = collection.find(entry => entry.id === check.dataset.taskComplete);
    if (!item) return;
    item.completed = check.checked;
    if (check.dataset.taskSource === "activity" && item.type === "Plano" && item.clientId) {
      const client = state.clients.find(entry => entry.id === item.clientId);
      if (client && check.checked) markPlanReady(client);
      if (client && !check.checked) {
        client.planStatus = "Em montagem";
        client.planReadyDate = "";
      }
    }
    saveState();
    renderAll();
    showToast(check.checked ? "Atividade concluída" : "Atividade reaberta");
  }));
  document.querySelectorAll("[data-task-reschedule]").forEach(input => input.addEventListener("change", () => {
    const collection = input.dataset.taskSource === "activity" ? state.activities : state.clients;
    const item = collection.find(entry => entry.id === input.dataset.taskReschedule);
    if (!item || !input.value) return;
    if (input.dataset.taskSource === "activity") item.date = input.value;
    else item.followUp = input.value;
    item.completed = false;
    saveState();
    renderAll();
    showToast("To do remarcado");
  }));
  document.querySelectorAll("[data-edit-activity]").forEach(button => button.addEventListener("click", () => openActivityModal(button.dataset.editActivity)));
  document.querySelectorAll("[data-meeting-task]").forEach(button => button.addEventListener("click", () => openMeetingModal(button.dataset.meetingKind, button.dataset.meetingTask)));
  document.querySelectorAll("#tasksList [data-edit]").forEach(button => button.addEventListener("click", () => openClientModal(button.dataset.edit)));

  const totalWeek = weekTasks.length;
  const completionRate = totalWeek ? Math.round((completedWeek.length / totalWeek) * 100) : 0;
  el("focusText").textContent = `${completionRate}% concluído na semana`;
  el("focusBar").style.width = `${completionRate}%`;
};

function openActivityModal(id = null, presetClientId = "") {
  activityInteractionMode = false;
  el("activityForm").reset();
  el("activityId").value = "";
  el("activityDate").value = offsetDate(0);
  el("activityModalTitle").textContent = "Nova atividade";
  el("deleteActivity").classList.add("hidden");
  el("activityClient").innerHTML = `<option value="">Sem contato relacionado</option>${state.clients
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name, "pt-BR"))
    .map(client => `<option value="${client.id}">${escapeHtml(client.name)}</option>`)
    .join("")}`;

  if (presetClientId) {
    const client = state.clients.find(item => item.id === presetClientId);
    activityInteractionMode = true;
    el("activityModalTitle").textContent = "Registrar interação";
    el("activityTitle").value = client ? `Contato com ${client.name}` : "Interação comercial";
    el("activityType").value = "Ligacao";
    el("activityClient").value = presetClientId;
    el("activityNotes").placeholder = "O que aconteceu, objeções, interesse demonstrado e próximo passo";
  }

  if (id) {
    activityInteractionMode = false;
    const activity = state.activities.find(item => item.id === id);
    if (!activity) return;
    el("activityModalTitle").textContent = "Editar atividade";
    el("activityId").value = activity.id;
    el("activityTitle").value = activity.title;
    el("activityDate").value = activity.date;
    el("activityTime").value = activity.time || "";
    el("activityType").value = activity.type || "Ligacao";
    el("activityClient").value = activity.clientId || "";
    el("activityNotes").value = activity.notes || "";
    el("deleteActivity").classList.remove("hidden");
  }

  el("activityModal").classList.remove("hidden");
  setTimeout(() => el("activityTitle").focus(), 50);
}

function closeActivityModal() {
  el("activityModal").classList.add("hidden");
  activityInteractionMode = false;
}

function submitActivity(event) {
  event.preventDefault();
  const id = el("activityId").value;
  const existing = state.activities.find(item => item.id === id);
  const wasInteraction = activityInteractionMode;
  const activity = {
    id: id || crypto.randomUUID(),
    title: el("activityTitle").value.trim(),
    date: el("activityDate").value,
    time: el("activityTime").value,
    type: el("activityType").value,
    clientId: el("activityClient").value,
    notes: el("activityNotes").value.trim(),
    completed: existing?.completed || activityInteractionMode,
    createdAt: existing?.createdAt || new Date().toISOString()
  };
  if (existing) Object.assign(existing, activity);
  else state.activities.push(activity);
  if (activityInteractionMode && activity.clientId) {
    const client = state.clients.find(item => item.id === activity.clientId);
    if (client) {
      client.lastInteraction = activity.date;
      client.completed = false;
      if (!client.followUp || client.followUp <= activity.date) client.followUp = nextBusinessDate();
      if (!client.nextAction) client.nextAction = "Definir próximo passo após a interação";
    }
  }
  saveState();
  closeActivityModal();
  renderAll();
  if (wasInteraction && activity.clientId) renderClientHistory(activity.clientId);
  showToast(existing ? "Atividade atualizada" : wasInteraction ? "Interação registrada no histórico" : "Atividade criada");
}

function deleteCurrentActivity() {
  const id = el("activityId").value;
  if (!id || !confirm("Excluir esta atividade?")) return;
  state.activities = state.activities.filter(activity => activity.id !== id);
  saveState();
  closeActivityModal();
  renderAll();
  showToast("Atividade excluída");
}

function openPipelineProcessReview(client, fieldId, message, outcome = "") {
  openClientModal(client.id);
  switchFormTab("management");
  el("commercialProcessSection").open = true;
  if (outcome) {
    el("clientOutcome").value = outcome;
    el("clientOutcomeDate").value ||= offsetDate(0);
  }
  setTimeout(() => el(fieldId)?.focus(), 80);
  showToast(message, true);
}

function requestPipelineStageMove(client, targetStage) {
  if (!client || isCloudReadOnly() || client.stage === targetStage) return;
  if (client.stage === "C2" && targetStage === "N") {
    prepareC2Win(client.id);
    return;
  }
  const currentIndex = STAGES.indexOf(client.stage);
  const targetIndex = STAGES.indexOf(targetStage);
  if (currentIndex >= 0 && targetIndex < currentIndex) {
    openPipelineProcessReview(client, "clientTaDate", "Para corrigir uma etapa anterior, revise a Ficha 360 sem apagar o histórico.");
    return;
  }
  if (targetStage === "TA") {
    applyProcessStage(client, targetStage);
    saveState();
    renderAll();
    showToast(`Movido para ${targetStage}`);
    return;
  }
  if (targetStage === "OI agendada") {
    if (!client.oiScheduledDate || !client.oiTime) {
      openPipelineProcessReview(client, "clientOiScheduledDate", "Para avançar, registre data e horário da OI.");
      return;
    }
  }
  if (targetStage === "OI realizada") {
    if (!client.oiDone || !client.oiDate || !hasPostOiDiagnosis(client)) {
      openPipelineProcessReview(client, "clientOiDone", "OI realizada exige data e diagnóstico Pós-OI.");
      return;
    }
  }
  if (targetStage === "PC agendada") {
    if (client.planStatus !== "Pronto" || !client.pcScheduledDate || !client.pcTime) {
      openPipelineProcessReview(client, "clientPlanStatus", "PC agendada exige plano pronto, data e horário.");
      return;
    }
  }
  if (targetStage === "PC realizada") {
    openPipelineProcessReview(client, "clientPcDone", "Registre a PC e sua decisão: N ou C2.");
    return;
  }
  if (targetStage === "N") {
    openPipelineProcessReview(client, "clientOutcomeDate", "Revise a data e salve para confirmar o N.", "N");
    return;
  }
  if (targetStage === "C2") {
    openPipelineProcessReview(client, "clientC2Objection", "C2 exige motivo, próxima ação e data de follow-up.", "C2");
    return;
  }
  applyProcessStage(client, targetStage);
  saveState();
  renderAll();
  showToast(`Movido para ${targetStage}`);
}

function bindPipelineEvents() {
  document.querySelectorAll(".edit-card").forEach(button => button.addEventListener("click", () => openClientModal(button.dataset.id)));
  document.querySelectorAll(".pipeline-card").forEach(card => {
    card.addEventListener("dragstart", () => card.classList.add("dragging"));
    card.addEventListener("dragend", () => card.classList.remove("dragging"));
  });
  document.querySelectorAll(".pipeline-column").forEach(column => {
    column.addEventListener("dragover", event => {
      event.preventDefault();
      column.classList.add("drag-over");
    });
    column.addEventListener("dragleave", () => column.classList.remove("drag-over"));
    column.addEventListener("drop", event => {
      event.preventDefault();
      column.classList.remove("drag-over");
      const card = document.querySelector(".pipeline-card.dragging");
      if (!card) return;
      const client = state.clients.find(c => c.id === card.dataset.clientId);
      requestPipelineStageMove(client, column.dataset.stage);
    });
  });
}

function switchView(view) {
  let requestedCrmTab = "";
  let requestedResultsTab = "";
  if (view === "pipeline") {
    view = "clients";
    requestedCrmTab = "funnel";
  }
  if (view === "policies") {
    view = "indicators";
    requestedResultsTab = "policies";
  }
  if (view === "management") {
    view = "indicators";
    requestedResultsTab = "diagnosis";
  }
  if (isCloudReadOnly() && !VIEWER_ALLOWED_VIEWS.has(view)) view = VIEWER_DEFAULT_VIEW;
  if (!el(`${view}View`)) view = VIEWER_DEFAULT_VIEW;
  currentView = view;
  document.body.dataset.currentView = view;
  const navigationView = ["tasks", "ritual", "assistant"].includes(view) ? "dashboard" : view;
  document.querySelectorAll(".view").forEach(item => item.classList.remove("active"));
  document.querySelectorAll(".nav-item").forEach(item => item.classList.toggle("active", item.dataset.view === navigationView));
  document.querySelectorAll(".mobile-nav-item").forEach(item => item.classList.toggle("active", item.dataset.mobileView === navigationView));
  el(`${view}View`).classList.add("active");
  const titles = { dashboard: "Operação", indicators: "Resultados", friday: "Modo TA", input: "Registro", clients: "CRM", tasks: "Agenda da operação", assistant: "Assistente comercial", ritual: "Ritmo da operação" };
  el("pageTitle").textContent = titles[view];
  document.querySelector(".sidebar").classList.remove("open");
  if (view === "clients") switchCrmTab(requestedCrmTab || crmTab);
  if (view === "indicators") switchResultsTab(requestedResultsTab || resultsTab);
  if (view === "assistant") renderAssistant();
  if (view === "tasks") renderTasks();
  if (view === "ritual") renderDailyRitual();
  if (view === "friday") renderFridayMode();
  if (view === "input") renderQuickInput();
}

function organizeMainViews() {
  const managementContent = el("managementContent");
  if (!managementContent || managementContent.children.length) return;
  [
    ".truth-mode",
    ".god-mode-panel"
  ].forEach(selector => {
    const section = el("dashboardView").querySelector(selector);
    if (section) managementContent.appendChild(section);
  });
}

function switchFormTab(tabName) {
  document.querySelectorAll("[data-form-tab]").forEach(button => {
    button.classList.toggle("active", button.dataset.formTab === tabName);
  });
  document.querySelectorAll("[data-form-panel]").forEach(panel => {
    panel.classList.toggle("active", panel.dataset.formPanel === tabName);
  });
}

function openClientModal(id = null) {
  if (isCloudReadOnly()) {
    showToast("Modo operacional: ficha completa protegida.", true);
    return;
  }
  el("clientForm").reset();
  document.querySelectorAll("#clientForm details").forEach(section => section.removeAttribute("open"));
  switchFormTab("profile");
  el("clientId").value = "";
  el("deleteClient").classList.add("hidden");
  el("modalTitle").textContent = "Novo contato";
  el("clientStage").value = "Base";
  el("clientPriority").value = "Média";
  el("clientTemperature").value = "Morno";
  el("clientFollowUp").value = offsetDate(1);
  renderClientHistory("");
  renderClient360(null);
  renderClientCareStatus(null);
  loadCitiesForState("");

  if (id) {
    const client = state.clients.find(c => c.id === id);
    if (!client) return;
    el("modalTitle").textContent = "Editar contato";
    el("clientId").value = client.id;
    el("clientName").value = client.name;
    el("clientWhatsapp").value = client.whatsapp || client.phone || "";
    el("clientEmail").value = client.email;
    el("clientState").value = client.state || "";
    el("clientProfession").value = client.profession || "";
    el("clientIncomeRange").value = client.incomeRange || "";
    el("clientMaritalStatus").value = client.maritalStatus || "";
    el("clientChildren").value = client.children || 0;
    el("clientReferrer").value = client.referrer || "";
    el("clientReferrerRelationship").value = client.referrerRelationship || "";
    el("clientReferralContext").value = client.referralContext || "";
    el("clientTemperature").value = client.temperature || "Morno";
    el("clientStage").value = client.stage;
    el("clientPriority").value = client.priority;
    el("clientScore").value = client.score || 0;
    el("clientFollowUp").value = client.followUp;
    el("clientLastInteraction").value = client.lastInteraction || "";
    el("clientNextAction").value = client.nextAction || "";
    el("clientValue").value = client.value;
    el("clientSource").value = client.source;
    el("clientNotes").value = client.notes;
    el("clientTaDate").value = client.taDate || "";
    el("clientOiScheduledDate").value = client.oiScheduledDate || "";
    el("clientOiTime").value = client.oiTime || "";
    el("clientOiMeetingLink").value = client.oiMeetingLink || "";
    el("clientOiDone").checked = Boolean(client.oiDone);
    el("clientOiDate").value = client.oiDate || "";
    el("clientPlanStatus").value = client.planStatus || "";
    el("clientPlanDueDate").value = client.planDueDate || "";
    el("clientPlanProfile").checked = Boolean(client.planProfile);
    el("clientPlanGoals").checked = Boolean(client.planGoals);
    el("clientPlanProtection").checked = Boolean(client.planProtection);
    el("clientPlanBudget").checked = Boolean(client.planBudget);
    el("clientPlanAlternatives").checked = Boolean(client.planAlternatives);
    el("clientPlanNotes").value = client.planNotes || "";
    el("clientPlanMonthlyBudget").value = client.planMonthlyBudget || "";
    el("clientPlanFixedCost").value = client.planFixedCost || "";
    el("clientPlanLifestyle").value = client.planLifestyle || "";
    el("clientPlanRiskConcerns").value = client.planRiskConcerns || "";
    el("clientPcScheduledDate").value = client.pcScheduledDate || "";
    el("clientPcTime").value = client.pcTime || "";
    el("clientPcMeetingLink").value = client.pcMeetingLink || "";
    el("clientPcDone").checked = Boolean(client.pcDone);
    el("clientPcDate").value = client.pcDate || "";
    el("clientOutcome").value = client.outcome || "";
    el("clientOutcomeDate").value = client.outcomeDate || "";
    el("clientC2NextAction").value = client.c2NextAction || "";
    el("clientC2Objection").value = client.c2Objection || "";
    el("clientCpfDone").checked = Boolean(client.cpfDone);
    el("clientCpfDate").value = client.cpfDate || "";
    el("clientPolicyDelivered").checked = Boolean(client.policyDelivered);
    el("clientDeliveryDate").value = client.deliveryDate || "";
    el("clientPolicyStatus").value = client.policyStatus || "";
    el("clientNextRelationship").value = client.nextRelationship || "";
    el("clientRelationshipStatus").value = client.relationshipStatus || "";
    el("clientRecs").value = client.recs || 0;
    el("clientRecsDate").value = client.recsDate || "";
    el("clientD7Done").checked = Boolean(client.d7Done);
    el("clientD30Done").checked = Boolean(client.d30Done);
    el("clientD60Done").checked = Boolean(client.d60Done);
    el("clientD180Done").checked = Boolean(client.d180Done);
    el("deleteClient").classList.remove("hidden");
    renderClientHistory(client.id);
    renderClient360(client);
    renderClientCareStatus(client);
    renderPcPlanner(client);
    if (client.state) loadCitiesForState(client.state, client.city || "");
  } else {
    renderPcPlanner(plannerClientFromForm());
  }

  el("clientModal").classList.remove("hidden");
  setTimeout(() => el("clientName").focus(), 50);
}

function createRecommendedContact() {
  const referrerId = el("clientId").value;
  openRecommendedContactFrom(referrerId);
}

function openRecommendedContactFrom(referrerId) {
  const referrer = state.clients.find(client => client.id === referrerId);
  if (!referrer) {
    showToast("Salve primeiro o contato que fez a recomendação", true);
    return;
  }
  openClientModal();
  el("clientSource").value = "Indicação";
  el("clientReferrer").value = referrer.name;
  el("clientTemperature").value = "Quente";
  el("clientPriority").value = "Alta";
  el("clientFollowUp").value = nextBusinessDate();
  el("clientNextAction").value = `Primeiro contato - indicação de ${referrer.name.split(/\s+/)[0]}`;
  el("clientNotes").value = `Contato recomendado por ${referrer.name}.`;
  el("clientName").focus();
  showToast("Recomendação vinculada. Complete nome e WhatsApp");
}

function openOiRecommendationImport() {
  const referrerId = el("clientId").value;
  if (!referrerId) {
    showToast("Salve primeiro o contato que realizou a OI", true);
    return;
  }
  closeClientModal();
  openPdfModal("recommendations", referrerId);
}

function closeClientModal() {
  el("clientModal").classList.add("hidden");
}

function submitClient(event) {
  event.preventDefault();
  const name = el("clientName").value.trim();
  const whatsapp = el("clientWhatsapp").value.trim();
  const whatsappDigits = normalizeWhatsapp(whatsapp);
  el("clientName").setCustomValidity(name ? "" : "Informe o nome do contato.");
  el("clientWhatsapp").setCustomValidity(
    !whatsapp ? "Informe o WhatsApp do contato." :
    whatsappDigits.length < 10 || whatsappDigits.length > 11 ? "Informe um WhatsApp válido com DDD." : ""
  );
  if (!name || !whatsapp || whatsappDigits.length < 10 || whatsappDigits.length > 11) {
    switchFormTab("profile");
    el("clientForm").reportValidity();
    return;
  }

  const id = el("clientId").value;
  const existing = state.clients.find(c => c.id === id);
  const previousStage = existing?.stage || "";
  const previousOutcome = existing?.outcome || "";
  const previousRecs = Number(existing?.recs || 0);
  const duplicate = contactDuplicate({
    name,
    phone: whatsappDigits,
    email: el("clientEmail").value
  }, id);
  if (duplicate) {
    switchFormTab("profile");
    showToast(`Contato já cadastrado: ${duplicate.name}`, true);
    return;
  }
  const outcome = el("clientOutcome").value;
  let taDate = el("clientTaDate").value;
  const oiScheduledDate = el("clientOiScheduledDate").value;
  const oiDone = el("clientOiDone").checked;
  const pcScheduledDate = el("clientPcScheduledDate").value;
  const pcDone = el("clientPcDone").checked;
  if (oiScheduledDate && !taDate) taDate = offsetDate(0);
  let processStage = el("clientStage").value;
  if (taDate) processStage = "TA";
  if (oiScheduledDate) processStage = "OI agendada";
  if (oiDone) processStage = "OI realizada";
  if (pcScheduledDate) processStage = "PC agendada";
  if (pcDone) processStage = "PC realizada";
  if (outcome) processStage = outcome;

  const data = {
    id: id || crypto.randomUUID(),
    name,
    phone: whatsappDigits,
    whatsapp: whatsappDigits,
    email: el("clientEmail").value.trim(),
    state: el("clientState").value,
    city: el("clientCity").value,
    profession: el("clientProfession").value.trim(),
    incomeRange: el("clientIncomeRange").value,
    maritalStatus: el("clientMaritalStatus").value,
    children: Number(el("clientChildren").value || 0),
    referrer: el("clientReferrer").value.trim(),
    referrerRelationship: el("clientReferrerRelationship").value.trim(),
    referralContext: el("clientReferralContext").value.trim(),
    temperature: el("clientTemperature").value,
    stage: processStage,
    priority: el("clientPriority").value,
    score: Number(el("clientScore").value || 0),
    followUp: el("clientFollowUp").value,
    lastInteraction: el("clientLastInteraction").value,
    nextAction: el("clientNextAction").value.trim(),
    value: Number(el("clientValue").value || 0),
    source: el("clientSource").value,
    notes: el("clientNotes").value.trim(),
    taDate,
    oiScheduledDate,
    oiTime: el("clientOiTime").value,
    oiMeetingLink: el("clientOiMeetingLink").value.trim(),
    oiDone,
    oiDate: el("clientOiDate").value,
    planStatus: el("clientPlanStatus").value,
    planDueDate: el("clientPlanDueDate").value,
    planReadyDate: el("clientPlanStatus").value === "Pronto" ? (existing?.planReadyDate || offsetDate(0)) : "",
    planProfile: el("clientPlanProfile").checked,
    planGoals: el("clientPlanGoals").checked,
    planProtection: el("clientPlanProtection").checked,
    planBudget: el("clientPlanBudget").checked,
    planAlternatives: el("clientPlanAlternatives").checked,
    planNotes: el("clientPlanNotes").value.trim(),
    planMonthlyBudget: Number(el("clientPlanMonthlyBudget").value || 0),
    planFixedCost: Number(el("clientPlanFixedCost").value || 0),
    planLifestyle: el("clientPlanLifestyle").value.trim(),
    planRiskConcerns: el("clientPlanRiskConcerns").value.trim(),
    pcScheduledDate,
    pcTime: el("clientPcTime").value,
    pcMeetingLink: el("clientPcMeetingLink").value.trim(),
    pcDone,
    pcDate: el("clientPcDate").value,
    outcome,
    outcomeDate: el("clientOutcomeDate").value,
    c2NextAction: el("clientC2NextAction").value.trim(),
    c2Objection: el("clientC2Objection").value,
    c2EnteredAt: existing?.c2EnteredAt || "",
    c2LastOutcome: existing?.c2LastOutcome || "",
    c2LastOutcomeDate: existing?.c2LastOutcomeDate || "",
    c2TouchCount: Number(existing?.c2TouchCount || 0),
    c2RecoveredAt: existing?.c2RecoveredAt || "",
    c2LegacyReview: existing?.c2LegacyReview || "",
    cpfDone: el("clientCpfDone").checked,
    cpfDate: el("clientCpfDate").value,
    policyDelivered: el("clientPolicyDelivered").checked,
    deliveryDate: el("clientDeliveryDate").value,
    policyStatus: el("clientPolicyStatus").value,
    nextRelationship: el("clientNextRelationship").value,
    relationshipStatus: el("clientRelationshipStatus").value,
    recs: Number(el("clientRecs").value || 0),
    recsDate: el("clientRecsDate").value,
    d7Done: el("clientD7Done").checked,
    d30Done: el("clientD30Done").checked,
    d60Done: el("clientD60Done").checked,
    d180Done: el("clientD180Done").checked,
    taPortfolioWeek: existing?.taPortfolioWeek || "",
    taPortfolioPinnedAt: existing?.taPortfolioPinnedAt || "",
    completed: existing?.completed || false,
    createdAt: existing?.createdAt || new Date().toISOString()
  };

  const isProcessAdvance = !existing
    || STAGES.indexOf(data.stage) > STAGES.indexOf(previousStage)
    || previousOutcome !== data.outcome;
  const processBlock = isProcessAdvance ? validateClientProcessGate(data, existing) : null;
  if (processBlock) {
    focusProcessField(processBlock.field, processBlock.message);
    return;
  }

  if (existing) {
    Object.assign(existing, data);
    if (previousStage !== data.stage) logClientEvent(data.id, "Etapa atualizada", `${previousStage} → ${data.stage}`, "Historico");
    if (previousOutcome !== data.outcome && data.outcome) {
      logClientEvent(
        data.id,
        data.outcome === "N" ? "Venda fechada" : "Oportunidade em C2",
        data.outcome === "N" ? "Resultado N registrado no CRM." : "Acompanhamento de fechamento iniciado.",
        "Historico"
      );
    }
  } else {
    state.clients.push(data);
    logClientEvent(data.id, "Contato cadastrado", `Origem: ${data.source || "não informada"}.`, "Historico");
  }
  const savedClient = existing || data;
  const recIncrease = Number(savedClient.recs || 0) - previousRecs;
  if (savedClient.outcome === "C2") {
    savedClient.c2EnteredAt = previousOutcome === "C2"
      ? (savedClient.c2EnteredAt || savedClient.outcomeDate || offsetDate(0))
      : (savedClient.outcomeDate || offsetDate(0));
    if (previousOutcome !== "C2") savedClient.c2RecoveredAt = "";
  }
  if (previousOutcome === "C2" && savedClient.outcome === "N") {
    savedClient.c2RecoveredAt ||= savedClient.outcomeDate || offsetDate(0);
    logClientEvent(savedClient.id, "C2 recuperado em N", "Conversão confirmada pela Ficha 360.", "Historico", savedClient.outcomeDate || offsetDate(0));
  }
  if (savedClient.outcome === "C2" && !savedClient.followUp) {
    savedClient.followUp = nextBusinessDate();
    savedClient.nextAction ||= savedClient.c2NextAction || "Realizar follow-up de fechamento";
  }
  if (savedClient.oiDone && savedClient.planStatus !== "Pronto") ensurePlanTask(savedClient);
  if (savedClient.planStatus === "Pronto") {
    state.activities
      .filter(activity => activity.clientId === savedClient.id && activity.type === "Plano")
      .forEach(activity => { activity.completed = true; });
  }
  saveState();
  closeClientModal();
  renderAll();
  if (recIncrease > 0) {
    const label = recIncrease === 1 ? "1 REC nova" : `${recIncrease} RECs novas`;
    setTimeout(() => {
      if (confirm(`${savedClient.name} registrou ${label}. Quer cadastrar a indicação agora?`)) {
        openRecommendedContactFrom(savedClient.id);
      } else {
        showToast(`${label} registrada. Você pode cadastrar depois pelo botão de recomendação.`);
      }
    }, 120);
  } else {
    showToast(existing ? "Contato atualizado" : "Contato cadastrado");
  }
}

function deleteCurrentClient() {
  const id = el("clientId").value;
  if (!id || !confirm("Excluir este contato e todo o histórico local dele?")) return;
  state.clients = state.clients.filter(c => c.id !== id);
  state.activities = state.activities.filter(activity => activity.clientId !== id);
  saveState();
  closeClientModal();
  renderAll();
  showToast("Contato excluído");
}

function exportData() {
  state.settings.lastBackupAt = offsetDate(0);
  saveState();
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `sitplan-crm-backup-${offsetDate(0)}.json`;
  link.click();
  URL.revokeObjectURL(link.href);
  showToast("Backup exportado");
}

function importData(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const imported = JSON.parse(reader.result);
      if (!Array.isArray(imported.clients)) throw new Error();
      storeSafetyBackup(serializeStateForStorage(state), "antes de importar backup");
      activeSettingsForNormalization = imported.settings || {};
      state = {
        ...imported,
        clients: imported.clients.map(normalizeClient),
        activities: Array.isArray(imported.activities) ? imported.activities : [],
        weeklyGoals: imported.weeklyGoals || {},
        settings: imported.settings || {}
      };
      activateStateSettings();
      saveState();
      renderAll();
      showToast("Backup importado");
    } catch (_) {
      showToast("Arquivo de backup inválido", true);
    }
    event.target.value = "";
  };
  reader.readAsText(file);
}

function parseCsvLine(line, delimiter) {
  const fields = [];
  let current = "";
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    if (char === '"' && quoted && line[index + 1] === '"') {
      current += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === delimiter && !quoted) {
      fields.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }
  fields.push(current.trim());
  return fields;
}

function normalizeHeader(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

const IMPORT_COLUMN_ALIASES = {
  name: ["nome", "nome completo", "cliente", "contato"],
  phone: ["telefone", "whatsapp", "celular", "tel contato", "tels de contato", "telefone contato"],
  email: ["email", "e mail"],
  source: ["origem", "fonte"],
  segment: ["segmento", "profissao segmento"],
  referrer: ["indicador", "recomendante", "indicado por", "quem indicou", "info client", "infoclient", "cliente indicador", "referrer"],
  profession: ["profissao", "profissao segmento", "cargo"],
  notes: ["observacoes", "observacao", "notas", "contexto", "memoria", "resumo"],
  temperature: ["temperatura"],
  priority: ["prioridade"]
};

function findImportColumn(headers, aliases) {
  const normalizedAliases = (Array.isArray(aliases) ? aliases : IMPORT_COLUMN_ALIASES[aliases] || [aliases]).map(normalizeHeader);
  return headers.findIndex(header => normalizedAliases.includes(header));
}

function importColumnValue(values, headers, aliases, fallback = "") {
  const index = findImportColumn(headers, aliases);
  return index >= 0 ? String(values[index] || "").trim() : fallback;
}

function importCsv(event) {
  const file = event.target.files[0];
  if (!file) return;
  if (!confirmOfficialImportSource(file)) {
    event.target.value = "";
    showToast("Importação cancelada para proteger a base");
    return;
  }
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const text = String(reader.result || "").replace(/^\uFEFF/, "").trim();
      const lines = text.split(/\r?\n/).filter(line => line.trim());
      if (lines.length < 2) throw new Error("Planilha vazia");
      const delimiter = (lines[0].match(/;/g) || []).length >= (lines[0].match(/,/g) || []).length ? ";" : ",";
      const headers = parseCsvLine(lines[0], delimiter).map(normalizeHeader);
      const column = key => findImportColumn(headers, key);
      const nameIndex = column("name");
      const phoneIndex = column("phone");
      const emailIndex = column("email");
      const sourceIndex = column("source");
      const segmentIndex = column("segment");
      const referrerIndex = column("referrer");
      const professionIndex = column("profession");
      const notesIndex = column("notes");
      const temperatureIndex = column("temperature");
      const priorityIndex = column("priority");
      if (nameIndex < 0 || phoneIndex < 0) throw new Error("Use as colunas Nome e Telefone");

      let imported = 0;
      let duplicates = 0;
      let missingAreaCode = 0;
      let invalid = 0;
      const referrerImportCounts = new Map();
      const configuredAreaCode = defaultAreaCode();
      const previousState = serializeStateForStorage(state);
      lines.slice(1).forEach(line => {
        const values = parseCsvLine(line, delimiter);
        const name = (values[nameIndex] || "").trim();
        let whatsapp = (values[phoneIndex] || "").replace(/\.0$/, "").trim();
        if (!name || !whatsapp) return;
        const originalDigits = whatsapp.replace(/\D/g, "").replace(/^0+/, "");
        const areaCodeAdded = originalDigits.length === 8 || originalDigits.length === 9;
        let digits = normalizeWhatsapp(whatsapp);
        whatsapp = digits;
        if (digits.length < 10 || digits.length > 11) {
          invalid += 1;
          return;
        }
        const email = emailIndex >= 0 ? (values[emailIndex] || "").trim() : "";
        const duplicate = contactDuplicate({ name, phone: digits, email });
        if (duplicate) {
          duplicates += 1;
          return;
        }
        const source = sourceIndex >= 0 ? (values[sourceIndex] || "Lista importada").trim() : "Lista importada";
        const segment = segmentIndex >= 0 ? (values[segmentIndex] || "").trim() : "";
        const referrer = referrerIndex >= 0 ? (values[referrerIndex] || "").trim() : "";
        const importedProfession = professionIndex >= 0 ? (values[professionIndex] || "").trim() : "";
        const importedNotes = notesIndex >= 0 ? (values[notesIndex] || "").trim() : "";
        const importedTemperature = temperatureIndex >= 0 ? (values[temperatureIndex] || "").trim() : "";
        const importedPriority = priorityIndex >= 0 ? (values[priorityIndex] || "").trim() : "";
        if (referrer) referrerImportCounts.set(referrer, (referrerImportCounts.get(referrer) || 0) + 1);
        const hasAreaCode = digits.length >= 10;
        if (!hasAreaCode) missingAreaCode += 1;
        const profession = importedProfession || (/medic/i.test(segment)
          ? "Médico(a)"
          : /empres|invest/i.test(segment) ? "Empresário(a)" : "");
        const noteParts = [
          segment ? `Segmento importado: ${segment}` : "",
          referrer ? `Indicador/recomendante: ${referrer}.` : "",
          importedNotes,
          areaCodeAdded ? `DDD ${configuredAreaCode} adicionado automaticamente.` : hasAreaCode ? "" : "Telefone importado sem DDD."
        ].filter(Boolean);
        state.clients.push(normalizeClient({
          id: crypto.randomUUID(),
          name,
          phone: whatsapp,
          whatsapp,
          email,
          profession,
          stage: "Base",
          priority: importedPriority || (referrer ? "Alta" : "Média"),
          temperature: importedTemperature || (referrer ? "Quente" : "Morno"),
          followUp: "",
          lastInteraction: "",
          nextAction: hasAreaCode
            ? (referrer ? `Fazer TA - indicação de ${referrer.split(/\s+/)[0]}` : "Fazer primeiro contato")
            : "Corrigir WhatsApp: falta DDD",
          value: 0,
          referrer,
          source: source || "Lista importada",
          notes: noteParts.join(" "),
          completed: false,
          createdAt: new Date().toISOString()
        }));
        imported += 1;
      });

      if (imported) storeSafetyBackup(previousState, "antes de importar planilha");
      saveState();
      renderAll();
      switchView("clients");
      const referrerSummary = Array.from(referrerImportCounts.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 4)
        .map(([name, count]) => `${name.split(/\s+/)[0]} ${count}`)
        .join(", ");
      const details = [
        `${imported} contato(s) importado(s)`,
        referrerSummary ? `Indicadores: ${referrerSummary}` : "",
        duplicates ? `${duplicates} duplicado(s) ignorado(s)` : "",
        missingAreaCode ? `${missingAreaCode} sem DDD` : "",
        invalid ? `${invalid} número(s) inválido(s)` : ""
      ].filter(Boolean).join(" · ");
      showToast(details || "Nenhum contato novo encontrado", !imported);
    } catch (error) {
      showToast(error.message || "Não foi possível importar a planilha", true);
    }
    event.target.value = "";
  };
  reader.readAsText(file, "UTF-8");
}

function openPdfModal(mode = "contacts", referrerId = "") {
  resetPdfImport();
  pdfImportMode = mode === "recommendations" ? "recommendations" : "contacts";
  pdfReferrerId = referrerId || "";
  el("pdfImportMode").value = pdfImportMode;
  el("pdfReferrer").innerHTML = state.clients.length
    ? state.clients.map(client => `<option value="${client.id}">${escapeHtml(client.name)}</option>`).join("")
    : `<option value="">Nenhum contato cadastrado</option>`;
  if (pdfReferrerId && state.clients.some(client => client.id === pdfReferrerId)) {
    el("pdfReferrer").value = pdfReferrerId;
  } else {
    pdfReferrerId = el("pdfReferrer").value;
  }
  updatePdfImportMode();
  el("pdfDefaultFollowUp").value = offsetDate(1);
  el("pdfModal").classList.remove("hidden");
}

function closePdfModal() {
  el("pdfModal").classList.add("hidden");
  resetPdfImport();
}

function resetPdfImport() {
  importedPdfContacts = [];
  el("pdfFile").value = "";
  el("pdfUploadStep").classList.remove("hidden");
  el("pdfReviewStep").classList.add("hidden");
  el("pdfLoadingStep").classList.add("hidden");
  el("pdfContactsList").innerHTML = "";
}

function updatePdfImportMode() {
  pdfImportMode = el("pdfImportMode").value;
  const recommendations = pdfImportMode === "recommendations";
  el("pdfReferrerField").classList.toggle("hidden", !recommendations);
  el("pdfSourceField").classList.toggle("hidden", recommendations);
  el("pdfModalTitle").textContent = recommendations ? "Importar recomendações da OI" : "Importar contatos de PDF";
  el("pdfDropHelp").textContent = recommendations
    ? "A IA lê o PDF, inclusive páginas escaneadas, e procura somente as recomendações registradas na OI."
    : "O arquivo é processado neste navegador. PDFs exportados de planilhas funcionam melhor.";
  el("pdfHelpText").textContent = recommendations
    ? "O PDF será enviado à OpenAI pela conexão protegida do Supabase. Confira nomes e WhatsApps antes de cadastrar. Nenhum contato será criado sem sua confirmação."
    : "Use PDFs com colunas de nome e telefone. Para bases grandes, CSV é mais rápido e confiável.";
  el("pdfLoadingStep").querySelector("strong").textContent = recommendations ? "A IA está lendo a OI..." : "Lendo o PDF...";
  el("pdfLoadingStep").querySelector("p").textContent = recommendations
    ? "Identificando apenas as recomendações e os respectivos contatos."
    : "Extraindo nomes e informações de contato.";
}

function fileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Não foi possível preparar o PDF."));
    reader.readAsDataURL(file);
  });
}

function recommendationDuplicate(name, phone, email = "") {
  const regularDuplicate = contactDuplicate({ name, phone, email });
  if (regularDuplicate) return regularDuplicate;
  const normalizedName = normalizedPersonName(name);
  return normalizedName
    ? state.clients.find(client => normalizedPersonName(client.name) === normalizedName)
    : null;
}

async function extractPdfLines(file) {
  const pdfLibrary = await getPdfLibrary();
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfLibrary.getDocument({ data }).promise;
  const lines = [];
  let extractedItems = 0;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const rows = [];
    content.items.forEach(item => {
      const value = String(item.str || "").trim();
      if (!value) return;
      extractedItems += 1;
      const x = Number(item.transform?.[4] || 0);
      const y = Number(item.transform?.[5] || 0);
      let row = rows.find(candidate => Math.abs(candidate.y - y) <= 3);
      if (!row) {
        row = { y, cells: [] };
        rows.push(row);
      }
      row.cells.push({ x, value });
    });
    rows
      .sort((a, b) => b.y - a.y)
      .forEach(row => {
        const line = row.cells
          .sort((a, b) => a.x - b.x)
          .map(cell => cell.value)
          .join("\t")
          .trim();
        if (line) lines.push(line);
      });
  }

  return { lines, extractedItems };
}

async function processPdfFile(file) {
  if (!file || (!file.name.toLowerCase().endsWith(".pdf") && file.type !== "application/pdf")) {
    showToast("Selecione um arquivo PDF", true);
    return;
  }

  el("pdfUploadStep").classList.add("hidden");
  el("pdfLoadingStep").classList.remove("hidden");

  try {
    const { lines, extractedItems } = await extractPdfLines(file);

    if (!extractedItems && pdfImportMode !== "recommendations") throw new Error("PDF_SCAN");

    if (pdfImportMode === "recommendations") {
      pdfReferrerId = el("pdfReferrer").value;
      if (!pdfReferrerId) throw new Error("REFERRER_REQUIRED");
      if (file.size > 4.5 * 1024 * 1024) throw new Error("PDF_TOO_LARGE");
      let recommendations = [];
      let aiResult = null;
      try {
        aiResult = await extractOIRecommendations({
          fileName: file.name,
          fileData: await fileAsDataUrl(file),
          extractedText: lines.join("\n")
        });
        recommendations = aiResult.recommendations;
      } catch (aiError) {
        const localContacts = extractContactsFromLines(lines);
        if (!localContacts.length) throw aiError;
        recommendations = localContacts.map(contact => ({
          name: contact.name,
          phone: contact.phone,
          referrer: contact.referrer || "",
          source: contact.source || "Indicação",
          profession: contact.profession || "",
          relationship: "",
          notes: contact.notes || "Extraído localmente; confirme se faz parte das recomendações.",
          confidence: "low"
        }));
        showToast("IA indisponível. Fiz uma leitura local para você conferir.", true);
      }
      importedPdfContacts = recommendations.map(item => {
        const name = String(item.name || "").trim().replace(/\s+/g, " ");
        const phone = normalizeWhatsapp(item.phone || "");
        const duplicate = recommendationDuplicate(name, phone);
        const validPhone = phone.length >= 10 && phone.length <= 11;
        return {
          name,
          phone,
          email: "",
          referrer: String(item.referrer || item.indicador || item.recomendante || item.indicatedBy || "").trim(),
          source: String(item.source || item.origem || "Indicação").trim(),
          profession: String(item.profession || item.profissao || "").trim(),
          relationship: String(item.relationship || "").trim(),
          notes: String(item.notes || "").trim(),
          confidence: String(item.confidence || "medium"),
          selected: Boolean(name && validPhone && !duplicate),
          duplicate: Boolean(duplicate),
          invalid: !validPhone
        };
      }).filter(contact => contact.name);
      if (aiResult) registerPaidApiUsage(aiResult, "Importação de PDF da OI");
    } else {
      importedPdfContacts = extractContactsFromLines(lines);
    }
    el("pdfLoadingStep").classList.add("hidden");

    if (!importedPdfContacts.length) {
      el("pdfUploadStep").classList.remove("hidden");
      showToast(pdfImportMode === "recommendations" ? "A IA não encontrou recomendações explícitas nessa OI" : "Não encontrei nomes legíveis nesse PDF", true);
      return;
    }

    renderPdfReview(file.name);
    el("pdfReviewStep").classList.remove("hidden");
  } catch (error) {
    console.error(error);
    el("pdfLoadingStep").classList.add("hidden");
    el("pdfUploadStep").classList.remove("hidden");
    showToast(
      error.message === "REFERRER_REQUIRED"
        ? "Selecione primeiro quem fez as recomendações."
        : error.message === "PDF_TOO_LARGE"
          ? "Este PDF passa de 4,5 MB. Comprima o arquivo e tente novamente."
      : error.message === "PDF_SCAN" && pdfImportMode !== "recommendations"
        ? "Este PDF parece escaneado como imagem. Exporte a planilha como CSV ou PDF com texto."
        : error.message || "Não foi possível ler este PDF.",
      true
    );
  }
}

function extractContactsFromLines(lines) {
  const emailPattern = /[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}/i;
  const phonePattern = /(?:\+?55\s*)?(?:\(?\d{2}\)?\s*)?(?:9?\d{4})[-.\s]?\d{4}/;
  const ignored = /^(nome|nomes|contato|contatos|telefone|e-?mail|lista|cliente|clientes|página|page|data|cpf|endereço|relatório|cadastro|total)\b/i;
  const seen = new Set();
  const contacts = [];
  let tableHeaders = [];

  const looksLikeName = name => {
    const words = name.split(" ").filter(Boolean);
    return words.length >= 2
      && words.length <= 8
      && name.length <= 80
      && !/\d/.test(name)
      && !ignored.test(name)
      && words.every(word => /^[a-z'.-]+$/i.test(normalizedPersonName(word)));
  };

  const addContact = rawContact => {
    const name = String(rawContact.name || "").trim().replace(/\s+/g, " ");
    const rawPhone = String(rawContact.phone || "").trim();
    const email = String(rawContact.email || "").trim();
    if (!looksLikeName(name) || !rawPhone) return false;
    const normalizedPhone = normalizeWhatsapp(rawPhone);
    const normalizedEmail = email.toLocaleLowerCase("pt-BR");
    const key = normalizedPhone || normalizedEmail || name.toLocaleLowerCase("pt-BR");
    if (seen.has(key)) return true;
    seen.add(key);
    const validPhone = normalizedPhone.length >= 10 && normalizedPhone.length <= 11;
    const duplicate = contactDuplicate({ name, phone: normalizedPhone, email });
    contacts.push({
      name,
      phone: normalizedPhone,
      email,
      referrer: String(rawContact.referrer || "").trim(),
      source: String(rawContact.source || "").trim(),
      profession: String(rawContact.profession || "").trim(),
      notes: String(rawContact.notes || "").trim(),
      temperature: String(rawContact.temperature || "").trim(),
      priority: String(rawContact.priority || "").trim(),
      selected: Boolean(validPhone && !duplicate),
      duplicate: Boolean(duplicate),
      invalid: !validPhone
    });
    return true;
  };

  const contactFromTable = (cells, headers) => {
    const name = importColumnValue(cells, headers, "name");
    const phone = importColumnValue(cells, headers, "phone");
    if (!name || !phone) return null;
    return {
      name,
      phone,
      email: importColumnValue(cells, headers, "email"),
      source: importColumnValue(cells, headers, "source"),
      referrer: importColumnValue(cells, headers, "referrer"),
      profession: importColumnValue(cells, headers, "profession"),
      notes: importColumnValue(cells, headers, "notes"),
      temperature: importColumnValue(cells, headers, "temperature"),
      priority: importColumnValue(cells, headers, "priority")
    };
  };

  lines.forEach(rawLine => {
    const raw = String(rawLine || "");
    const cells = raw.split(/\t+| {2,}/).map(cell => cell.trim()).filter(Boolean);
    if (cells.length >= 2) {
      const normalizedCells = cells.map(normalizeHeader);
      if (findImportColumn(normalizedCells, "name") >= 0 && findImportColumn(normalizedCells, "phone") >= 0) {
        tableHeaders = normalizedCells;
        return;
      }
      if (tableHeaders.length) {
        const tableContact = contactFromTable(cells, tableHeaders);
        if (tableContact && addContact(tableContact)) return;
      }
      const phoneCellIndex = cells.findIndex(cell => phonePattern.test(cell));
      if (phoneCellIndex > 0) {
        const afterPhone = cells.slice(phoneCellIndex + 1);
        const firstAfterPhone = normalizeHeader(afterPhone[0] || "");
        const hasSourceCell = ["indicacao", "lista importada", "networking", "prospeccao", "outro"].includes(firstAfterPhone);
        const inferredContact = {
          name: cells.slice(0, phoneCellIndex).join(" "),
          phone: cells[phoneCellIndex],
          source: hasSourceCell ? afterPhone[0] : "",
          referrer: hasSourceCell ? afterPhone[1] : afterPhone[0],
          profession: hasSourceCell ? afterPhone[2] : afterPhone[1],
          notes: (hasSourceCell ? afterPhone.slice(3) : afterPhone.slice(2)).join(" ")
        };
        if (addContact(inferredContact)) return;
      }
    }

    const normalized = raw.replace(/\s+/g, " ").replace(/[|;]+/g, " | ").trim();
    const email = normalized.match(emailPattern)?.[0] || "";
    const phoneMatch = normalized.match(phonePattern);
    const phone = phoneMatch?.[0] || "";
    const textBeforePhone = phoneMatch ? normalized.slice(0, phoneMatch.index) : normalized;
    const name = textBeforePhone
      .replace(emailPattern, " ")
      .replace(/\b(?:CPF|RG|Tel(?:efone)?|Cel(?:ular)?|E-?mail|Nome)\s*:?\s*/gi, " ")
      .replace(/\s*[|,:-]\s*/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    addContact({ name: name.replace(/^\d+[.)\s-]+/, "").trim(), phone, email });
  });

  return contacts;
}

function renderPdfReview(fileName) {
  const duplicates = importedPdfContacts.filter(contact => contact.duplicate).length;
  const invalid = importedPdfContacts.filter(contact => contact.invalid).length;
  const label = pdfImportMode === "recommendations" ? "recomendações encontradas" : "contatos encontrados";
  const defaultReferrer = pdfImportMode === "recommendations"
    ? (state.clients.find(client => client.id === pdfReferrerId)?.name || "")
    : "";
  if (defaultReferrer) {
    importedPdfContacts.forEach(contact => {
      if (!contact.referrer) contact.referrer = defaultReferrer;
    });
  }
  el("pdfResultTitle").textContent = `${importedPdfContacts.length} ${label}`;
  el("pdfResultSummary").textContent = [
    fileName,
    duplicates ? `${duplicates} possível(is) duplicado(s)` : "",
    invalid ? `${invalid} sem WhatsApp válido` : ""
  ].filter(Boolean).join(" · ");
  el("pdfContactsList").innerHTML = importedPdfContacts.map((contact, index) => `
    <label class="import-row ${contact.duplicate || contact.invalid ? "duplicate" : ""}">
      <input type="checkbox" data-import-select="${index}" ${contact.selected ? "checked" : ""}>
      <input type="text" value="${escapeHtml(contact.name)}" data-import-field="name" data-import-index="${index}" aria-label="Nome importado ${index + 1}">
      <input type="tel" value="${escapeHtml(contact.phone)}" data-import-field="phone" data-import-index="${index}" placeholder="Telefone" aria-label="Telefone importado ${index + 1}">
      ${pdfImportMode === "recommendations"
        ? `<input type="text" value="${escapeHtml(contact.relationship || "")}" data-import-field="relationship" data-import-index="${index}" placeholder="Relação com o cliente" aria-label="Relação importada ${index + 1}">`
        : `<input type="email" value="${escapeHtml(contact.email)}" data-import-field="email" data-import-index="${index}" placeholder="E-mail" aria-label="E-mail importado ${index + 1}">`}
      <input type="text" value="${escapeHtml(contact.referrer || "")}" data-import-field="referrer" data-import-index="${index}" placeholder="Indicador" aria-label="Indicador importado ${index + 1}">
      ${contact.duplicate ? `<span class="duplicate-note">Possível duplicado: desmarcado para evitar cadastro repetido.</span>` : ""}
      ${contact.invalid ? `<span class="duplicate-note">WhatsApp ausente ou incompleto. Corrija o número para importar.</span>` : ""}
      ${(contact.profession || contact.source || contact.notes) ? `<span class="import-context">${[
        contact.referrer ? `Indicado por: ${contact.referrer}` : "",
        contact.source ? `Origem: ${contact.source}` : "",
        contact.profession ? `Profissão: ${contact.profession}` : "",
        contact.notes || ""
      ].filter(Boolean).map(escapeHtml).join(" · ")}</span>` : ""}
    </label>`).join("");

  document.querySelectorAll("[data-import-select]").forEach(checkbox => checkbox.addEventListener("change", () => {
    importedPdfContacts[Number(checkbox.dataset.importSelect)].selected = checkbox.checked;
  }));
  document.querySelectorAll("[data-import-field]").forEach(input => input.addEventListener("input", () => {
    const contact = importedPdfContacts[Number(input.dataset.importIndex)];
    contact[input.dataset.importField] = input.value.trim();
    if (input.dataset.importField === "phone") {
      const phone = normalizeWhatsapp(input.value);
      contact.invalid = phone.length < 10 || phone.length > 11;
    }
  }));
}

function confirmPdfImport() {
  const selected = importedPdfContacts.filter(contact => contact.selected && contact.name.trim());
  if (!selected.length) {
    showToast("Selecione pelo menos um contato", true);
    return;
  }

  const validContacts = [];
  let duplicates = 0;
  let invalid = 0;
  const batchKeys = new Set();
  selected.forEach(contact => {
    const name = contact.name.trim().replace(/\s+/g, " ");
    const phone = normalizeWhatsapp(contact.phone);
    const email = String(contact.email || "").trim();
    const key = phone || email.toLocaleLowerCase("pt-BR") || name.toLocaleLowerCase("pt-BR");
    if (phone.length < 10 || !name) {
      invalid += 1;
      return;
    }
    const existingDuplicate = pdfImportMode === "recommendations"
      ? recommendationDuplicate(name, phone, email)
      : contactDuplicate({ name, phone, email });
    if (batchKeys.has(key) || existingDuplicate) {
      duplicates += 1;
      return;
    }
    batchKeys.add(key);
    validContacts.push({
      name,
      phone,
      email,
      referrer: String(contact.referrer || "").trim(),
      relationship: String(contact.relationship || "").trim(),
      source: String(contact.source || "").trim(),
      profession: String(contact.profession || "").trim(),
      notes: String(contact.notes || "").trim(),
      temperature: String(contact.temperature || "").trim(),
      priority: String(contact.priority || "").trim()
    });
  });
  if (!validContacts.length) {
    showToast(duplicates ? "Todos os selecionados já estão cadastrados" : "Revise nome e WhatsApp dos contatos", true);
    return;
  }

  const source = pdfImportMode === "recommendations" ? "Indicação" : el("pdfDefaultSource").value;
  const followUp = el("pdfDefaultFollowUp").value;
  const referrer = state.clients.find(client => client.id === pdfReferrerId);
  const referrerCounts = new Map();
  validContacts.forEach(contact => {
    const imported = importedPdfContacts.find(item =>
      normalizeWhatsapp(item.phone) === contact.phone || normalizedPersonName(item.name) === normalizedPersonName(contact.name)
    );
    const contactReferrer = (contact.referrer || imported?.referrer || (pdfImportMode === "recommendations" ? referrer?.name || "" : "")).trim();
    const isRecommendation = pdfImportMode === "recommendations" || Boolean(contactReferrer);
    const contactSource = contact.source || imported?.source || (contactReferrer ? "Indicação" : source);
    const contactProfession = contact.profession || imported?.profession || "";
    const contactTemperature = contact.temperature || imported?.temperature || (isRecommendation ? "Quente" : "Morno");
    const contactPriority = contact.priority || imported?.priority || (isRecommendation ? "Alta" : "Média");
    const contactRelationship = contact.relationship || imported?.relationship || "";
    const contactNotes = contact.notes || imported?.notes || "";
    const referrerFirstName = contactReferrer ? contactReferrer.split(/\s+/)[0] : "";
    const matchedReferrer = contactReferrer
      ? state.clients.find(clientItem => normalizedPersonName(clientItem.name) === normalizedPersonName(contactReferrer))
      : null;
    if (matchedReferrer) {
      referrerCounts.set(matchedReferrer.id, (referrerCounts.get(matchedReferrer.id) || 0) + 1);
    }
    const client = {
      id: crypto.randomUUID(),
      name: contact.name.trim(),
      phone: contact.phone,
      whatsapp: contact.phone,
      email: contact.email,
      state: "",
      city: "",
      profession: contactProfession,
      referrer: contactReferrer,
      referrerRelationship: contactRelationship,
      referralContext: contactNotes,
      temperature: contactTemperature,
      stage: "Base",
      priority: contactPriority,
      followUp,
      lastInteraction: "",
      nextAction: isRecommendation
        ? `Fazer TA - indicação de ${referrerFirstName || "cliente"}`
        : "Fazer primeiro contato",
      value: 0,
      source: contactSource || source,
      notes: isRecommendation
        ? [
          `Recomendado por ${contactReferrer || "cliente da OI"}.`,
          contactRelationship ? `Relação: ${contactRelationship}.` : "",
          contactNotes
        ].filter(Boolean).join(" ")
        : "Contato importado de PDF.",
      completed: false,
      createdAt: new Date().toISOString()
    };
    state.clients.push(normalizeClient(client));
    if (isRecommendation) {
      state.activities.push({
        id: crypto.randomUUID(),
        title: `TA - ${client.name}`,
        date: followUp || nextBusinessDate(),
        time: "",
        type: "Ligacao",
        clientId: client.id,
        notes: contactReferrer
          ? `Primeira abordagem. Recomendação recebida de ${contactReferrer}.`
          : "Primeira abordagem. Recomendação recebida na OI.",
        completed: false,
        createdAt: new Date().toISOString()
      });
    }
  });

  referrerCounts.forEach((count, referrerId) => {
    const owner = state.clients.find(client => client.id === referrerId);
    if (!owner) return;
    owner.recs = Number(owner.recs || 0) + count;
    owner.recsDate = offsetDate(0);
    logClientEvent(
      owner.id,
      "Recomendações importadas da OI",
      `${count} contato(s) cadastrado(s) e incluído(s) na fila de TA.`,
      "Historico"
    );
  });

  saveState();
  closePdfModal();
  renderAll();
  switchView("clients");
  const details = [
    `${validContacts.length} ${pdfImportMode === "recommendations" ? (validContacts.length === 1 ? "recomendação cadastrada" : "recomendações cadastradas") : (validContacts.length === 1 ? "contato importado" : "contatos importados")}`,
    duplicates ? `${duplicates} duplicado(s) ignorado(s)` : "",
    invalid ? `${invalid} inválido(s)` : ""
  ].filter(Boolean).join(" · ");
  showToast(details);
  if (pdfImportMode === "recommendations") pendingPostOiRecommendations = [];
}

function showToast(message, error = false) {
  el("toast").textContent = message;
  el("toast").style.background = error ? "var(--red)" : "var(--green)";
  if (el("liveRegion")) el("liveRegion").textContent = message;
  el("toast").classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => el("toast").classList.remove("show"), 2300);
}

function bindCloudAccessEvents() {
  if (window.__sitplanCloudEventsBound) return;
  window.__sitplanCloudEventsBound = true;

  el("cloudAccessButton")?.addEventListener("click", async () => {
    if (el("cloudAccessButton").dataset.connected === "true") {
      await signOut();
      secureEntryConfirmed = false;
      setSyncStatus();
      applyPrivateAccessLock(true);
      showToast("Você saiu da nuvem");
    } else {
      openAuthModal();
    }
  });
  el("closeAuthModal")?.addEventListener("click", closeAuthModal);
  el("authModal")?.addEventListener("click", event => {
    if (event.target === el("authModal")) closeAuthModal();
  });
  el("authForm")?.addEventListener("submit", submitAuth);
  el("createAccountButton")?.addEventListener("click", createAccount);
  el("continueStoredSessionButton")?.addEventListener("click", confirmStoredSession);
  el("forgetStoredSessionButton")?.addEventListener("click", () => {
    forgetStoredSession().catch(error => showAuthError(authConnectionError(error)));
  });
  el("runCloudDiagnostic")?.addEventListener("click", runCloudDiagnostic);
  el("cloudDiagnosticButton")?.addEventListener("click", runCloudDiagnostic);
  el("toggleAuthPassword")?.addEventListener("click", () => {
    const input = el("authPassword");
    const button = el("toggleAuthPassword");
    const visible = input.type === "text";
    input.type = visible ? "password" : "text";
    button.setAttribute("aria-label", visible ? "Mostrar senha" : "Ocultar senha");
    button.setAttribute("title", visible ? "Mostrar senha" : "Ocultar senha");
  });
}

async function hardRefreshApp() {
  showToast("Atualizando arquivos do SITPLAN...");
  try {
    if ("serviceWorker" in navigator) {
      const registrations = await navigator.serviceWorker.getRegistrations();
      await Promise.all(registrations.map(registration => registration.unregister()));
    }
    if ("caches" in window) {
      const keys = await caches.keys();
      await Promise.all(keys.filter(key => key.startsWith("sitplan-crm-")).map(key => caches.delete(key)));
    }
  } catch (error) {
    console.warn("Não foi possível limpar todo o cache:", error);
  }
  window.location.reload();
}

let cloudSyncStarted = false;
let cloudStatusEventsBound = false;
let secureEntryConfirmed = !PRIVATE_ACCESS_REQUIRED;
let manualAuthInProgress = false;

function storedCloudSession() {
  try {
    return JSON.parse(localStorage.getItem(SUPABASE_SESSION_KEY) || "null");
  } catch (error) {
    return null;
  }
}

function storedSessionUser() {
  return storedCloudSession()?.user || null;
}

function shouldHoldStoredSession(user) {
  return PRIVATE_ACCESS_REQUIRED
    && Boolean(user)
    && Boolean(storedCloudSession()?.access_token)
    && !secureEntryConfirmed
    && !manualAuthInProgress;
}

function toggleStoredSessionActions(show = false) {
  el("storedSessionActions")?.classList.toggle("hidden", !show);
}

function showStoredSessionGate(user) {
  setSyncStatus(user, "Sessão salva");
  applyPrivateAccessLock(true);
  openAuthModal({ locked: true, storedSession: true, user });
}

function confirmStoredSession() {
  const user = storedSessionUser() || { id: "connected" };
  secureEntryConfirmed = true;
  connectAuthenticatedUser(user).catch(error => {
    console.error(error);
    secureEntryConfirmed = false;
    setSyncStatus(null, "Nuvem não conectou");
    applyPrivateAccessLock(true);
    showAuthError(error);
  });
}

async function forgetStoredSession() {
  secureEntryConfirmed = false;
  manualAuthInProgress = false;
  await signOut();
  el("authPassword").value = "";
  setSyncStatus();
  applyPrivateAccessLock(true);
  openAuthModal({ locked: true });
  el("authMessage").textContent = "Sessão salva encerrada. Digite e-mail e senha para entrar novamente.";
}

async function grantViewerAccess() {
  const email = el("viewerEmailInput")?.value.trim();
  if (!email) {
    showToast("Informe o e-mail do sócio operacional.", true);
    return;
  }
  try {
    await grantViewer(email);
    showToast(`Acesso TA + CRM básico liberado para ${email}`);
  } catch (error) {
    showToast(error.message || "Não foi possível liberar o acesso TA.", true);
  }
}

async function createViewerAccess() {
  const email = el("viewerEmailInput")?.value.trim();
  const password = el("viewerPasswordInput")?.value || "";
  if (!email) {
    showToast("Informe o e-mail do sócio operacional.", true);
    return;
  }
  if (password.length < 6) {
    showToast("Crie uma senha provisória com pelo menos 6 caracteres.", true);
    return;
  }
  try {
    const { data, error } = await createViewerAccount(email, password);
    if (error) throw error;
    el("viewerPasswordInput").value = "";
    showToast(data?.created
      ? `Acesso criado e liberado para ${email}`
      : `Conta existente liberada para ${email}`);
  } catch (error) {
    showToast(error.message || "Não foi possível criar o acesso do sócio.", true);
  }
}

async function revokeViewerAccess() {
  const email = el("viewerEmailInput")?.value.trim();
  if (!email) {
    showToast("Informe o e-mail do sócio para revogar.", true);
    return;
  }
  try {
    await revokeViewer(email);
    showToast(`Acesso revogado para ${email}`);
  } catch (error) {
    showToast(error.message || "Não foi possível revogar o acesso operacional.", true);
  }
}

function bindCloudStatusEvents() {
  if (cloudStatusEventsBound) return;
  cloudStatusEventsBound = true;

  window.addEventListener("cloud-sync-error", event => {
    setSyncStatus({ id: "connected" }, "Falha ao salvar");
    showToast(event.detail || "Não foi possível sincronizar com a nuvem", true);
  });
  window.addEventListener("cloud-sync-blocked", event => {
    const detail = event.detail || {};
    const amount = detail.missingCount ? `${detail.missingCount} registro(s)` : "alguns registros";
    setSyncStatus({ id: "connected" }, "Proteção ativada");
    showToast(`Sincronização pausada por segurança: ${amount} não foram apagados da nuvem.`, true);
  });
  window.addEventListener("cloud-sync-saving", () => {
    setSyncStatus({ id: "connected" }, "Salvando...");
  });
  window.addEventListener("cloud-sync-saved", event => {
    const savedAt = new Date(event.detail?.savedAt || Date.now());
    const time = savedAt.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
    setSyncStatus({ id: "connected" }, `Salvo na nuvem às ${time}`);
  });
}

function startCloudSync() {
  bindCloudStatusEvents();
  if (cloudSyncStarted) return;
  cloudSyncStarted = true;
  applyPrivateAccessLock();

  setTimeout(() => {
    initializeCloud(
      user => {
        if (user) {
          if (shouldHoldStoredSession(user)) {
            showStoredSessionGate(user);
            return;
          }
          secureEntryConfirmed = true;
          connectAuthenticatedUser(user).catch(error => {
            console.error(error);
            secureEntryConfirmed = false;
            setSyncStatus(null, "Nuvem não conectou");
            applyPrivateAccessLock(true);
            showToast(cloudErrorMessage(error), true);
          });
        } else {
          secureEntryConfirmed = false;
          setSyncStatus();
          applyPrivateAccessLock(true);
        }
      },
      refreshFromCloud
    ).catch(error => {
      console.error(error);
      setSyncStatus(null, "Nuvem indisponivel");
      applyPrivateAccessLock(true);
      showToast(cloudErrorMessage(error), true);
    });
  }, 250);
}

async function init() {
  window.__sitplanAppReady = false;
  bindCloudAccessEvents();
  applyPrivateAccessLock();
  document.body.dataset.currentView = currentView;
  organizeMainViews();
  const today = new Intl.DateTimeFormat("pt-BR", { weekday: "long", day: "2-digit", month: "long" }).format(new Date());
  el("todayLabel").textContent = today.toUpperCase();
  el("clientStage").innerHTML = STAGES.map(stage => `<option>${stage}</option>`).join("");
  el("stageFilter").innerHTML += STAGES.map(stage => `<option>${stage}</option>`).join("");
  el("clientState").innerHTML += BRAZIL_STATES.map(([code, name]) => `<option value="${code}">${name} (${code})</option>`).join("");
  el("professionOptions").innerHTML = PROFESSION_SUGGESTIONS.map(name => `<option value="${name}"></option>`).join("");

  document.querySelectorAll(".nav-item").forEach(button => button.addEventListener("click", () => {
    if (button.dataset.view === "clients") {
      workFilter = "";
      clientsPage = 1;
    }
    switchView(button.dataset.view);
    if (button.dataset.view === "clients") renderClients();
  }));
  document.querySelectorAll("[data-mobile-view]").forEach(button => button.addEventListener("click", () => {
    if (button.dataset.mobileView === "clients") {
      workFilter = "";
      clientsPage = 1;
    }
    switchView(button.dataset.mobileView);
  }));
  document.querySelectorAll("[data-crm-tab]").forEach(button => button.addEventListener("click", () => switchCrmTab(button.dataset.crmTab)));
  document.querySelectorAll("[data-results-tab]").forEach(button => button.addEventListener("click", () => switchResultsTab(button.dataset.resultsTab)));
  document.querySelectorAll("[data-view-jump]").forEach(button => button.addEventListener("click", () => {
    switchView(button.dataset.viewJump);
  }));
  el("ritualOpenFriday")?.addEventListener("click", () => switchView("friday"));
  el("ritualOpenInput")?.addEventListener("click", () => switchView("input"));
  document.querySelectorAll("[data-action='show-tasks']").forEach(button => button.addEventListener("click", () => {
    switchView("dashboard");
    window.setTimeout(() => el("operationWorkspace")?.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
  }));
  el("operationOpenAgenda")?.addEventListener("click", () => switchView("tasks"));
  el("operationOpenRitual")?.addEventListener("click", () => switchView("ritual"));
  el("clientsPreviousPage")?.addEventListener("click", () => {
    clientsPage = Math.max(1, clientsPage - 1);
    renderClients();
    el("clientsView")?.scrollIntoView({ behavior: "smooth", block: "start" });
  });
  el("clientsNextPage")?.addEventListener("click", () => {
    clientsPage += 1;
    renderClients();
    el("clientsView")?.scrollIntoView({ behavior: "smooth", block: "start" });
  });
  el("fridayLoadMore")?.addEventListener("click", () => {
    fridayQueueVisibleLimit += FRIDAY_QUEUE_PAGE_SIZE;
    renderFridayMode();
  });
  el("decisionModeButton")?.addEventListener("click", event => {
    const view = event.currentTarget.dataset.decisionView || "input";
    switchView(view);
  });
  document.querySelectorAll("[data-work-filter]").forEach(button => button.addEventListener("click", () => {
    workFilter = button.dataset.workFilter;
    clientsPage = 1;
    crmTab = "list";
    el("attentionFilter").value = "all";
    el("stageFilter").value = "all";
    switchView("clients");
    renderClients();
  }));
  document.querySelectorAll(".task-tab").forEach(button => button.addEventListener("click", () => {
    taskFilter = button.dataset.taskFilter;
    document.querySelectorAll(".task-tab").forEach(tab => tab.classList.toggle("active", tab === button));
    renderTasks();
  }));
  el("agendaPreviousWeek").addEventListener("click", () => {
    agendaWeekOffset -= 1;
    renderTasks();
  });
  el("agendaCurrentWeek").addEventListener("click", () => {
    agendaWeekOffset = 0;
    renderTasks();
  });
  el("agendaNextWeek").addEventListener("click", () => {
    agendaWeekOffset += 1;
    renderTasks();
  });
  document.querySelectorAll("[data-form-tab]").forEach(button => button.addEventListener("click", () => {
    switchFormTab(button.dataset.formTab);
  }));

  el("newClientButton").addEventListener("click", () => openClientModal());
  el("openMonthlyDiagnosis").addEventListener("click", openMonthlyDiagnosis);
  el("closeMonthlyDiagnosis").addEventListener("click", closeMonthlyDiagnosis);
  el("closeMonthlyDiagnosisAction").addEventListener("click", closeMonthlyDiagnosis);
  el("monthlyDiagnosisModal").addEventListener("click", event => {
    if (event.target === el("monthlyDiagnosisModal")) closeMonthlyDiagnosis();
  });
  el("diagnosisMonth").addEventListener("change", renderMonthlyDiagnosis);
  el("annualGoalTarget").addEventListener("change", event => {
    state.settings.annualGoalTarget = Math.max(1, Math.min(1000, Number(event.target.value) || 134));
    saveState();
    renderAnnualGoal();
    showToast("Meta anual atualizada");
  });
  el("annualGoalDeadline").addEventListener("change", event => {
    state.settings.annualGoalDeadline = event.target.value || "2026-10-31";
    saveState();
    renderAnnualGoal();
    showToast("Prazo da meta atualizado");
  });
  el("fridayRefreshQueue").addEventListener("click", () => {
    fridaySosActive = false;
    fridaySosFocus = "";
    fridayReferrerFilter = "all";
    fridayPinnedClientId = "";
    fridaySessionActive = false;
    fridayQueueVisibleLimit = FRIDAY_QUEUE_PAGE_SIZE;
    clearFridayQueueFineFilters();
    resetFridayStableQueue();
    renderFridayMode();
    focusFridayQueue("Fila completa de sexta atualizada");
  });
  el("taPortfolioBuildCurrent")?.addEventListener("click", () => buildTaPortfolio(0));
  el("taPortfolioBuildNext")?.addEventListener("click", () => buildTaPortfolio(1));
  el("taPortfolioTarget")?.addEventListener("change", event => updateTaPortfolioTarget(event.target.value));
  document.querySelectorAll("[data-ta-portfolio-filter]").forEach(button => button.addEventListener("click", () => {
    fridayPortfolioFilter = button.dataset.taPortfolioFilter || "all";
    fridayPinnedClientId = "";
    fridaySessionActive = false;
    clearFridayQueueFineFilters();
    resetFridayStableQueue();
    renderFridayMode();
  }));
  el("runAutomationNow").addEventListener("click", () => {
    runOperationalAutomation({ notify: true });
    renderAll();
  });
  el("runDailyReview").addEventListener("click", () => runDailyReview());
  el("smartDailyReviewRun")?.addEventListener("click", () => runDailyReview());
  el("smartDailyReviewDismiss")?.addEventListener("click", () => {
    sessionStorage.setItem(DAILY_REVIEW_PROMPT_KEY, offsetDate(0));
    renderSmartDailyReviewPrompt();
  });
  el("copyExecutiveDigest")?.addEventListener("click", copyExecutiveDigest);
  el("emailExecutiveDigest")?.addEventListener("click", openExecutiveDigestEmail);
  el("saveWeeklyContract").addEventListener("click", () => {
    const week = selectedWeek();
    state.settings.weeklyContracts ||= {};
    state.settings.weeklyContracts[week.startKey] = {
      oi: Math.max(0, Number(el("contractOi").value || 0)),
      pc: Math.max(0, Number(el("contractPc").value || 0)),
      c2: Math.max(0, Number(el("contractC2").value || 0)),
      plans: Math.max(0, Number(el("contractPlans").value || 0)),
      hub: el("contractHub").value.trim()
    };
    state.weeklyGoals[week.startKey] = Math.max(1, Number(el("contractOi").value || weeklyGoal(week)));
    saveState();
    renderAll();
    showToast("Contrato semanal salvo");
  });
  el("policyCommissionRate")?.addEventListener("change", event => {
    state.settings.policyCommissionRate = Math.max(0, Math.min(100, Number(event.target.value || 0)));
    saveState();
    renderPolicies();
    showToast("Taxa de comissão atualizada");
  });
  el("defaultAreaCodeInput")?.addEventListener("change", event => {
    const digits = event.target.value.replace(/\D/g, "").slice(0, 2);
    state.settings.defaultAreaCode = digits.length === 2 ? digits : "11";
    activeSettingsForNormalization = state.settings;
    event.target.value = state.settings.defaultAreaCode;
    saveState();
    showToast(`DDD padrão atualizado para ${state.settings.defaultAreaCode}`);
  });
  el("grantViewerAccess")?.addEventListener("click", grantViewerAccess);
  el("createViewerAccess")?.addEventListener("click", createViewerAccess);
  el("revokeViewerAccess")?.addEventListener("click", revokeViewerAccess);
  ["truthPrimaryAction", "truthSecondaryAction"].forEach(id => {
    el(id).addEventListener("click", event => {
      const view = event.currentTarget.dataset.truthView || "dashboard";
      switchView(view);
      if (view === "clients") {
        workFilter = "";
        renderClients();
      }
    });
  });
  el("importCsvButton").addEventListener("click", () => el("csvFile").click());
  el("csvFile").addEventListener("change", importCsv);
  el("microsoftAccessButton").addEventListener("click", openMicrosoftModal);
  el("closeMicrosoftModal").addEventListener("click", closeMicrosoftModal);
  el("microsoftModal").addEventListener("click", event => {
    if (event.target === el("microsoftModal")) closeMicrosoftModal();
  });
  el("connectMicrosoftButton").addEventListener("click", () => {
    submitMicrosoftConnection().catch(error => {
      el("microsoftMessage").textContent = error.message || "Falha ao iniciar a conexão Microsoft.";
      el("microsoftMessage").classList.add("error");
    });
  });
  el("disconnectMicrosoftButton").addEventListener("click", () => {
    disconnectMicrosoft();
    closeMicrosoftModal();
    updateMicrosoftStatus();
    showToast("Microsoft 365 desconectado");
  });
  el("undoLastChange").addEventListener("click", undoLastChange);
  el("newActivityButton").addEventListener("click", () => openActivityModal());
  el("quickInputClient").addEventListener("change", updateQuickInputPreview);
  el("quickInputAction").addEventListener("change", () => {
    if (el("quickInputAction").value === "recs" && Number(el("quickInputRecs").value || 0) === 0) el("quickInputRecs").value = "1";
    updateQuickInputFields();
  });
  el("quickInputForm").addEventListener("submit", submitQuickInput);
  el("recordNewContact")?.addEventListener("click", () => openClientModal());
  el("quickInputOpenContact").addEventListener("click", () => {
    const clientId = el("quickInputClient").value;
    if (clientId) openClientModal(clientId);
  });
  el("closeActivityModal").addEventListener("click", closeActivityModal);
  el("cancelActivityModal").addEventListener("click", closeActivityModal);
  el("activityForm").addEventListener("submit", submitActivity);
  el("deleteActivity").addEventListener("click", deleteCurrentActivity);
  el("activityModal").addEventListener("click", event => {
    if (event.target === el("activityModal")) closeActivityModal();
  });
  el("closeC2UpdateModal")?.addEventListener("click", closeC2UpdateModal);
  el("cancelC2UpdateModal")?.addEventListener("click", closeC2UpdateModal);
  el("c2UpdateForm")?.addEventListener("submit", submitC2Update);
  el("c2UpdateModal")?.addEventListener("click", event => {
    if (event.target === el("c2UpdateModal")) closeC2UpdateModal();
  });
  document.querySelectorAll('input[name="c2UpdateOutcome"]').forEach(input => input.addEventListener("change", () => setC2UpdateDate(input.value)));
  document.querySelectorAll("[data-c2-legacy-action]").forEach(button => button.addEventListener("click", () => reviewLegacyC2(button.dataset.c2LegacyAction)));
  el("reviewLegacyC2")?.addEventListener("click", () => {
    const client = c2RecoveryData().legacy[0];
    if (client) openC2UpdateModal(client.id);
  });
  document.querySelectorAll("[data-c2-view]").forEach(button => button.addEventListener("click", () => {
    c2RecoveryView = button.dataset.c2View;
    renderC2Recovery();
  }));
  el("prepareOiMeeting").addEventListener("click", () => openMeetingModal("OI"));
  el("preparePcMeeting").addEventListener("click", () => openMeetingModal("PC"));
  el("closeMeetingModal").addEventListener("click", closeMeetingModal);
  el("cancelMeetingModal").addEventListener("click", closeMeetingModal);
  el("meetingModal").addEventListener("click", event => {
    if (event.target === el("meetingModal")) closeMeetingModal();
  });
  ["meetingDate", "meetingTime", "meetingDuration", "meetingLink"].forEach(id => {
    el(id).addEventListener("input", updateMeetingPreview);
    el(id).addEventListener("change", updateMeetingPreview);
  });
  el("meetingOpenTeams").addEventListener("click", openTeamsScheduler);
  el("meetingCheckAvailability").addEventListener("click", checkMeetingAvailability);
  el("meetingCreateMicrosoft").addEventListener("click", createMeetingInMicrosoft);
  el("meetingGoogle").addEventListener("click", openGoogleCalendar);
  el("meetingOutlook").addEventListener("click", openOutlookCalendar);
  el("meetingIcs").addEventListener("click", downloadMeetingIcs);
  el("meetingWhatsapp").addEventListener("click", openMeetingWhatsapp);
  el("meetingDayWhatsapp").addEventListener("click", openMeetingDayWhatsapp);
  el("saveMeetingDetails").addEventListener("click", () => saveMeetingDetails());
  el("meetingCompleteOi").addEventListener("click", () => registerMeetingOutcome("OI_DONE"));
  el("meetingOiNoShow").addEventListener("click", () => registerMeetingOutcome("OI_NO_SHOW"));
  el("meetingCloseN").addEventListener("click", () => registerMeetingOutcome("N"));
  el("meetingCloseC2").addEventListener("click", () => registerMeetingOutcome("C2"));
  el("reviewPostOi").addEventListener("click", () => openPostOiModal(el("clientId").value));
  el("closePostOiModal").addEventListener("click", closePostOiModal);
  el("postOiModal").addEventListener("click", event => {
    if (event.target === el("postOiModal")) closePostOiModal();
  });
  el("postOiForm").addEventListener("submit", event => {
    event.preventDefault();
    savePostOi("save");
  });
  el("postOiPreparePc").addEventListener("click", () => savePostOi("pc"));
  el("postOiImportRecs").addEventListener("click", () => savePostOi("recs"));
  el("postOiReadDocument").addEventListener("click", () => el("postOiDocumentFile").click());
  el("postOiDocumentFile").addEventListener("change", event => processPostOiDocumentFile(event.target.files[0]));
  el("importPdfButton").addEventListener("click", () => openPdfModal());
  el("closeModal").addEventListener("click", closeClientModal);
  el("cancelModal").addEventListener("click", closeClientModal);
  el("clientForm").addEventListener("submit", submitClient);
  ["clientPlanMonthlyBudget", "clientPlanFixedCost", "clientPlanLifestyle", "clientPlanRiskConcerns", "clientValue", "clientIncomeRange", "clientName"].forEach(id => {
    el(id)?.addEventListener("input", () => renderPcPlanner(plannerClientFromForm()));
    el(id)?.addEventListener("change", () => renderPcPlanner(plannerClientFromForm()));
  });
  el("createRecommendedContact").addEventListener("click", createRecommendedContact);
  el("importOiRecommendations").addEventListener("click", openOiRecommendationImport);
  el("clientLogInteraction").addEventListener("click", () => {
    const clientId = el("clientId").value;
    if (clientId) openActivityModal(null, clientId);
  });
  el("deleteClient").addEventListener("click", deleteCurrentClient);
  el("clientModal").addEventListener("click", event => {
    if (event.target === el("clientModal")) closeClientModal();
  });
  el("mobileMenu").addEventListener("click", () => document.querySelector(".sidebar").classList.toggle("open"));
  el("assistantClient").addEventListener("change", event => {
    assistantClientId = event.target.value;
    renderAssistantResult();
  });
  document.querySelectorAll("[data-assistant-mode]").forEach(button => button.addEventListener("click", () => {
    assistantMode = button.dataset.assistantMode;
    document.querySelectorAll("[data-assistant-mode]").forEach(tab => tab.classList.toggle("active", tab === button));
    renderAssistantResult();
  }));
  document.querySelectorAll("[data-assistant-tone]").forEach(button => button.addEventListener("click", () => {
    assistantTone = button.dataset.assistantTone;
    document.querySelectorAll("[data-assistant-tone]").forEach(option => option.classList.toggle("active", option === button));
    renderAssistantResult();
  }));
  el("assistantLocalAnalyze")?.addEventListener("click", refreshLocalAssistant);
  el("assistantAnalyze").addEventListener("click", () => copyAssistantBriefing());

  el("assistantChatSend")?.addEventListener("click", sendAssistantChat);
  el("assistantChatInput")?.addEventListener("keydown", event => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      sendAssistantChat();
    }
  });

  async function sendAssistantChat() {
    const input = el("assistantChatInput");
    const question = input?.value?.trim();
    const sendBtn = el("assistantChatSend");
    if (!question) {
      await copyAssistantBriefing();
      return;
    }
    appendAssistantChat(question, "user");
    sendBtn.disabled = true;
    try {
      await copyAssistantBriefing(question);
      input.value = "";
    } catch (err) {
      showToast("Não foi possível copiar o briefing", true);
    } finally {
      sendBtn.disabled = false;
    }
  }
  el("assistantEditContact").addEventListener("click", () => {
    if (assistantClientId) openClientModal(assistantClientId);
  });
  el("pipelineFilter").addEventListener("change", renderPipeline);
  el("showAllC2")?.addEventListener("click", () => {
    workFilter = "c2-all";
    clientsPage = 1;
    el("attentionFilter").value = "all";
    el("stageFilter").value = "all";
    switchCrmTab("list");
    renderClients();
  });
  el("attentionFilter").addEventListener("change", () => {
    workFilter = "";
    clientsPage = 1;
    renderClients();
  });
  el("stageFilter").addEventListener("change", () => {
    workFilter = "";
    clientsPage = 1;
    renderClients();
  });
  el("referrerFilter").addEventListener("change", () => {
    workFilter = "";
    clientsPage = 1;
    renderClients();
  });
  el("sortClients").addEventListener("change", () => {
    clientsPage = 1;
    renderClients();
  });
  el("previousWeek").addEventListener("click", () => {
    viewedWeekOffset -= 1;
    renderDashboard();
  });
  el("currentWeek").addEventListener("click", () => {
    viewedWeekOffset = 0;
    renderDashboard();
  });
  el("nextWeek").addEventListener("click", () => {
    viewedWeekOffset += 1;
    renderDashboard();
  });
  el("weeklyGoalInput").addEventListener("change", event => {
    const goal = Math.max(1, Math.min(200, Number(event.target.value) || DEFAULT_WEEKLY_OI_GOAL));
    state.weeklyGoals[selectedWeek().startKey] = goal;
    saveState();
    renderDashboard();
    showToast(`Meta semanal atualizada para ${goal} OIs`);
  });
  el("clientOiDone").addEventListener("change", event => {
    if (event.target.checked && !el("clientOiDate").value) el("clientOiDate").value = offsetDate(0);
  });
  el("clientPcDone").addEventListener("change", event => {
    if (event.target.checked && !el("clientPcDate").value) el("clientPcDate").value = offsetDate(0);
  });
  el("clientPlanStatus").addEventListener("change", event => {
    if (event.target.value === "Pronto") {
      ["clientPlanProfile", "clientPlanGoals", "clientPlanProtection", "clientPlanBudget", "clientPlanAlternatives"]
        .forEach(id => { el(id).checked = true; });
    }
  });
  el("clientOutcome").addEventListener("change", event => {
    if (event.target.value && !el("clientOutcomeDate").value) el("clientOutcomeDate").value = offsetDate(0);
  });
  el("clientRecs").addEventListener("input", event => {
    if (Number(event.target.value) > 0 && !el("clientRecsDate").value) el("clientRecsDate").value = offsetDate(0);
  });
  el("clientState").addEventListener("change", event => {
    el("clientCity").value = "";
    loadCitiesForState(event.target.value);
  });
  el("clientName").addEventListener("input", event => event.target.setCustomValidity(""));
  el("clientWhatsapp").addEventListener("input", event => event.target.setCustomValidity(""));
  el("globalSearch").addEventListener("input", event => {
    searchTerm = event.target.value;
    if (el("crmSearch") && el("crmSearch").value !== searchTerm) el("crmSearch").value = searchTerm;
    clientsPage = 1;
    if (searchTerm) crmTab = "list";
    renderPipeline();
    renderClients();
    if (searchTerm && currentView === "dashboard") switchView("clients");
  });
  el("crmSearch")?.addEventListener("input", event => {
    searchTerm = event.target.value;
    if (el("globalSearch") && el("globalSearch").value !== searchTerm) el("globalSearch").value = searchTerm;
    clientsPage = 1;
    crmTab = "list";
    renderClients();
  });
  el("exportData").addEventListener("click", exportData);
  el("importData").addEventListener("change", importData);
  el("hardRefreshApp")?.addEventListener("click", hardRefreshApp);
  el("closePdfModal").addEventListener("click", closePdfModal);
  el("cancelPdfImport").addEventListener("click", closePdfModal);
  el("chooseAnotherPdf").addEventListener("click", resetPdfImport);
  el("pdfImportMode").addEventListener("change", () => {
    resetPdfImport();
    updatePdfImportMode();
  });
  el("pdfReferrer").addEventListener("change", event => {
    pdfReferrerId = event.target.value;
  });
  el("confirmPdfImport").addEventListener("click", confirmPdfImport);
  el("selectAllImports").addEventListener("click", () => {
    const shouldSelect = importedPdfContacts.some(contact => !contact.selected);
    importedPdfContacts.forEach(contact => { contact.selected = shouldSelect; });
    renderPdfReview(el("pdfResultSummary").textContent.split(" · ")[0]);
  });
  el("pdfFile").addEventListener("change", event => processPdfFile(event.target.files[0]));
  ["dragenter", "dragover"].forEach(type => el("pdfDropZone").addEventListener(type, event => {
    event.preventDefault();
    el("pdfDropZone").classList.add("dragging");
  }));
  ["dragleave", "drop"].forEach(type => el("pdfDropZone").addEventListener(type, event => {
    event.preventDefault();
    el("pdfDropZone").classList.remove("dragging");
  }));
  el("pdfDropZone").addEventListener("drop", event => processPdfFile(event.dataTransfer.files[0]));
  document.addEventListener("keydown", event => {
    if (event.key === "Escape") {
      closeClientModal();
      closeActivityModal();
      closeC2UpdateModal();
      closeMeetingModal();
      closeMicrosoftModal();
      closeMonthlyDiagnosis();
      closeAuthModal();
      closePdfModal();
    }
  });

  runOperationalAutomation();
  renderAll();
  try {
    const result = await initializeMicrosoft();
    updateMicrosoftStatus(result.connected ? "Microsoft 365 conectado" : "");
    if (result.justConnected) {
      showToast("Microsoft 365 conectado com sucesso");
    }
  } catch (error) {
    console.error(error);
    updateMicrosoftStatus("Falha na conexão Microsoft");
    showToast(error.message || "Não foi possível conectar o Microsoft 365", true);
  }
  window.__sitplanAppReady = true;
  startCloudSync();
}

init();

