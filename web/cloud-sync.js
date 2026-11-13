const SUPABASE_URL = "https://bfkfuymbjhkyqyscyszu.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_eT5ZvYChk41Ac9vfsPDLkg_vNqgg93-";
const SESSION_KEY = "sitplan-supabase-session";

let session = readSession();
let authChangeHandler = () => {};
let remoteChangeHandler = () => {};
let pollingTimer;
let saveQueue = Promise.resolve();
let accessMode = { role: "owner", ownerUserId: session?.user?.id || null };

export async function initializeCloud(onAuthChange, onRemoteChange) {
  authChangeHandler = onAuthChange;
  remoteChangeHandler = onRemoteChange;
  let notifiedStoredSession = false;

  if (session?.access_token) {
    const valid = await validateOrRefreshSession();
    if (!valid) {
      window.dispatchEvent(new CustomEvent("cloud-sync-error", {
        detail: "Sessão salva, mas o Supabase não confirmou o login. Entre novamente."
      }));
      clearSession();
    } else {
      await determineAccessMode();
      if (session.user) {
        notifiedStoredSession = true;
        authChangeHandler(session.user);
      }
    }
  }

  if (!notifiedStoredSession || !session?.user) authChangeHandler(session?.user || null);
  startPolling();
}

export async function signIn(email, password) {
  const result = await authRequest("/auth/v1/token?grant_type=password", {
    method: "POST",
    body: JSON.stringify({ email, password })
  });
  if (result.error) return result;
  storeSession(result.data);
  await determineAccessMode();
  authChangeHandler(session.user);
  startPolling();
  return { data: result.data, error: null };
}

export async function signUp(email, password) {
  const redirectTo = encodeURIComponent(window.location.origin);
  const result = await authRequest(`/auth/v1/signup?redirect_to=${redirectTo}`, {
    method: "POST",
    body: JSON.stringify({ email, password })
  });
  if (result.error) return result;
  if (result.data.access_token) {
    storeSession(result.data);
    await determineAccessMode();
    authChangeHandler(session.user);
    startPolling();
  }
  return {
    data: {
      ...result.data,
      session: result.data.access_token ? result.data : null
    },
    error: null
  };
}

export async function createViewerAccount(email, password) {
  if (!isCloudConnected()) throw new Error("Entre na nuvem com o usuário dono antes de criar o acesso.");
  if (isCloudReadOnly()) throw new Error("Modo operacional não pode criar acesso.");
  const cleanEmail = String(email || "").trim();
  if (!cleanEmail) throw new Error("Informe o e-mail do sócio.");
  if (!password || String(password).length < 6) throw new Error("A senha provisória precisa ter pelo menos 6 caracteres.");

  await ensureFreshSession();
  const redirectTo = encodeURIComponent(window.location.origin);
  const signup = await authRequest(`/auth/v1/signup?redirect_to=${redirectTo}`, {
    method: "POST",
    body: JSON.stringify({ email: cleanEmail, password })
  });

  const alreadyExists = signup.error && /already|registered|exists|ja existe|cadastrado/i.test(signup.error.message || "");
  if (signup.error && !alreadyExists) return signup;

  await grantViewer(cleanEmail);
  return {
    data: {
      created: !alreadyExists,
      email: cleanEmail
    },
    error: null
  };
}

export async function signOut() {
  if (session?.access_token) {
    await fetch(`${SUPABASE_URL}/auth/v1/logout`, {
      method: "POST",
      headers: authHeaders()
    }).catch(() => {});
  }
  clearSession();
  accessMode = { role: "owner", ownerUserId: null };
  stopPolling();
  authChangeHandler(null);
  return { error: null };
}

export function isCloudConnected() {
  return Boolean(session?.access_token && session?.user);
}

export function isCloudReadOnly() {
  return accessMode.role === "viewer";
}

export function cloudAccessLabel() {
  return isCloudReadOnly() ? "TA + CRM básico" : "Nuvem sincronizada";
}

export async function grantViewer(email) {
  if (!isCloudConnected()) throw new Error("Entre na nuvem com o usuário dono antes de compartilhar.");
  if (isCloudReadOnly()) throw new Error("Modo operacional não pode conceder acesso.");
  await ensureFreshSession();
  return restRequest("/rest/v1/rpc/sitplan_grant_viewer", {
    method: "POST",
    body: JSON.stringify({ viewer_email: email })
  });
}

export async function revokeViewer(email) {
  if (!isCloudConnected()) throw new Error("Entre na nuvem com o usuário dono antes de revogar acesso.");
  if (isCloudReadOnly()) throw new Error("Modo operacional não pode revogar acesso.");
  await ensureFreshSession();
  return restRequest("/rest/v1/rpc/sitplan_revoke_viewer", {
    method: "POST",
    body: JSON.stringify({ viewer_email: email })
  });
}

export async function diagnoseCloud() {
  const checks = [];
  const add = (name, ok, detail) => checks.push({ name, ok, detail });

  try {
    const auth = await fetch(`${SUPABASE_URL}/auth/v1/settings`, {
      headers: { apikey: SUPABASE_PUBLISHABLE_KEY }
    });
    add("Auth Supabase", auth.ok, auth.ok ? "Login por e-mail acessível." : `Auth respondeu erro ${auth.status}.`);
  } catch (error) {
    add("Auth Supabase", false, "Não consegui acessar o serviço de login do Supabase.");
  }

  for (const table of ["contacts", "activities", "weekly_goals"]) {
    try {
      const response = await fetch(`${SUPABASE_URL}/rest/v1/${table}?select=*&limit=1`, {
        headers: authHeaders()
      });
      const text = await response.text();
      const data = text ? JSON.parse(text) : null;
      if (response.ok) {
        add(`Tabela ${table}`, true, "Tabela existe e respondeu.");
      } else {
        add(`Tabela ${table}`, false, data?.message || data?.hint || data?.code || `Erro ${response.status}.`);
      }
    } catch (error) {
      add(`Tabela ${table}`, false, "Não consegui testar esta tabela.");
    }
  }

  return {
    ok: checks.every(check => check.ok),
    checks
  };
}

export async function invokeCloudFunction(name, body = {}, attempt = 0) {
  const maxRetries = 2;
  const retryDelay = () => new Promise(resolve => setTimeout(resolve, 350 * Math.pow(2, attempt)));

  if (!isCloudConnected()) {
    throw new Error("Entre e sincronize para usar a IA.");
  }

  await ensureFreshSession();
  try {
    const response = await fetch(`${SUPABASE_URL}/functions/v1/${name}`, {
      method: "POST",
      headers: {
        ...authHeaders(),
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      const error = new Error(data.error || data.message || `Erro ${response.status} ao acessar a IA.`);
      error.status = response.status;
      const transient = response.status >= 500 || response.status === 429;
      if (transient && attempt < maxRetries) {
        await retryDelay();
        return invokeCloudFunction(name, body, attempt + 1);
      }
      throw error;
    }
    return data;
  } catch (error) {
    if (error?.status) throw error;
    if (attempt < maxRetries) {
      await retryDelay();
      return invokeCloudFunction(name, body, attempt + 1);
    }
    const networkError = new Error("Não consegui falar com a IA na nuvem. Verifique sua internet e tente novamente.");
    networkError.code = "NETWORK";
    throw networkError;
  }
}

export async function loadCloudState() {
  if (!isCloudConnected()) return null;
  await ensureFreshSession();
  await determineAccessMode();
  if (isCloudReadOnly()) return loadTaViewerCloudState();

  const ownerFilter = isCloudReadOnly() && accessMode.ownerUserId ? `&user_id=eq.${accessMode.ownerUserId}` : "";
  const [contacts, activities, goals] = await Promise.all([
    restRequest(`/rest/v1/contacts?select=id,data,created_at,updated_at,user_id${ownerFilter}&order=created_at`),
    restRequest(`/rest/v1/activities?select=*${ownerFilter}&order=activity_date,activity_time`),
    restRequest(`/rest/v1/weekly_goals?select=week_start,target,user_id${ownerFilter}`)
  ]);

  const settings = contacts.find(row => row.data?.crmSettings)?.data.crmSettings || {};
  const remoteUpdatedAt = [...contacts, ...activities]
    .map(row => row.updated_at || "")
    .reduce((latest, value) => (value > latest ? value : latest), "");
  return {
    remoteUpdatedAt,
    clients: contacts.map(row => ({ ...row.data, id: row.id })),
    activities: activities.map(row => ({
      id: row.id,
      title: row.title,
      date: row.activity_date,
      time: row.activity_time?.slice(0, 5) || "",
      type: row.activity_type,
      clientId: row.contact_id || "",
      notes: row.notes || "",
      completed: row.completed,
      createdAt: row.created_at
    })),
    weeklyGoals: Object.fromEntries(goals.map(row => [row.week_start, row.target])),
    settings
  };
}

async function loadTaViewerCloudState() {
  const contacts = await restRequest("/rest/v1/rpc/sitplan_ta_viewer_contacts", {
    method: "POST",
    body: JSON.stringify({})
  });

  return {
    clients: contacts.map(row => ({
      id: row.id,
      name: row.name || "Contato",
      phone: row.phone || row.whatsapp || "",
      whatsapp: row.whatsapp || row.phone || "",
      stage: row.stage || "Base",
      priority: row.priority || "Media",
      followUp: row.follow_up || "",
      source: row.source || "",
      referrer: row.referrer || "",
      temperature: row.temperature || "Morno",
      lastInteraction: row.last_interaction || "",
      taDate: row.ta_date || "",
      oiScheduledDate: row.oi_scheduled_date || "",
      nextAction: row.next_action || "Ligar para agendar OI",
      createdAt: row.created_at || "",
      doNotContact: false,
      outcome: ""
    })),
    activities: [],
    weeklyGoals: {},
    settings: { defaultAreaCode: "11", viewerScope: "ta" }
  };
}

export function queueCloudSave(state) {
  if (!isCloudConnected()) return;
  if (isCloudReadOnly()) {
    window.dispatchEvent(new CustomEvent("cloud-sync-error", {
      detail: "Modo operacional: alterações bloqueadas para proteger a base."
    }));
    return;
  }
  const snapshot = JSON.parse(JSON.stringify(state));
  saveQueue = saveQueue.then(async () => {
    window.dispatchEvent(new CustomEvent("cloud-sync-saving"));
    try {
      await persistSnapshot(snapshot);
      window.dispatchEvent(new CustomEvent("cloud-sync-saved", {
        detail: { savedAt: new Date().toISOString() }
      }));
    } catch (error) {
      console.error("Falha ao sincronizar:", error);
      window.dispatchEvent(new CustomEvent("cloud-sync-error", { detail: friendlyError(error) }));
    }
  });
}

async function persistSnapshot(state) {
  await ensureFreshSession();
  if (isCloudReadOnly()) throw new Error("Modo operacional: alterações bloqueadas.");
  const userId = session.user.id;
  const now = new Date().toISOString();
  const base = readSyncBase(userId);
  const next = snapshotKeys(state);
  const settingsChanged = !base || base.settings !== next.settings;

  // Upsert only rows that changed since the last successful sync, and delete
  // only rows this device had synced before and the user removed here. Rows
  // created on another device are never deleted just because this device has
  // not seen them yet.
  const contacts = state.clients
    .filter(client => settingsChanged || base.contacts[client.id] !== next.contacts[client.id])
    .map(client => ({
      id: client.id,
      user_id: userId,
      data: { ...client, crmSettings: state.settings || {} },
      updated_at: now
    }));
  const activities = state.activities
    .filter(activity => !base || base.activities[activity.id] !== next.activities[activity.id])
    .map(activity => ({
      id: activity.id,
      user_id: userId,
      contact_id: activity.clientId || null,
      title: activity.title,
      activity_date: activity.date,
      activity_time: activity.time || null,
      activity_type: activity.type || "Follow-up",
      notes: activity.notes || "",
      completed: Boolean(activity.completed),
      created_at: activity.createdAt || now,
      updated_at: now
    }));
  const goals = Object.entries(state.weeklyGoals)
    .filter(([weekStart]) => !base || base.goals[weekStart] !== next.goals[weekStart])
    .map(([weekStart, target]) => ({
      user_id: userId,
      week_start: weekStart,
      target,
      updated_at: now
    }));

  const deletedContacts = base ? Object.keys(base.contacts).filter(id => !(id in next.contacts)) : [];
  const deletedActivities = base ? Object.keys(base.activities).filter(id => !(id in next.activities)) : [];
  const deletedGoals = base ? Object.keys(base.goals).filter(week => !(week in next.goals)) : [];
  guardMassDeletion("contacts", deletedContacts.length, base ? Object.keys(base.contacts).length : 0);
  guardMassDeletion("activities", deletedActivities.length, base ? Object.keys(base.activities).length : 0);
  guardMassDeletion("weekly_goals", deletedGoals.length, base ? Object.keys(base.goals).length : 0);

  if (contacts.length) await upsertRows("contacts", contacts);
  if (activities.length) await upsertRows("activities", activities);
  if (goals.length) await upsertRows("weekly_goals", goals);
  // Activities first: they reference contacts.
  if (deletedActivities.length) await deleteRows("activities", "id", deletedActivities);
  if (deletedContacts.length) await deleteRows("contacts", "id", deletedContacts);
  if (deletedGoals.length) await deleteRows("weekly_goals", "week_start", deletedGoals);

  writeSyncBase(userId, next);
}

async function upsertRows(table, rows) {
  for (let index = 0; index < rows.length; index += 200) {
    await restRequest(`/rest/v1/${table}?on_conflict=${table === "weekly_goals" ? "user_id,week_start" : "id"}`, {
      method: "POST",
      headers: { Prefer: "resolution=merge-duplicates,return=minimal" },
      body: JSON.stringify(rows.slice(index, index + 200))
    });
  }
}

async function deleteRows(table, column, values) {
  for (let index = 0; index < values.length; index += 100) {
    const chunk = values.slice(index, index + 100).map(value => encodeURIComponent(value));
    await restRequest(`/rest/v1/${table}?${column}=in.(${chunk.join(",")})`, { method: "DELETE" });
  }
}

function guardMassDeletion(table, missingCount, knownCount) {
  if (!missingCount) return;
  const looksLikeMassDeletion = knownCount >= 5 && missingCount / knownCount >= 0.5;
  if (!looksLikeMassDeletion) return;
  window.dispatchEvent(new CustomEvent("cloud-sync-blocked", {
    detail: { table, missingCount, remoteCount: knownCount }
  }));
  throw new Error(`Sincronização bloqueada por segurança: ${missingCount} de ${knownCount} registro(s) seriam apagados da nuvem.`);
}

// ---------------------------------------------------------------------------
// Sync base: fingerprints of what this device last synced successfully.
// It is the common ancestor for a 3-way merge between this device and the
// cloud, so local edits that have not reached the cloud yet are never
// overwritten by polling or by signing in again.
// ---------------------------------------------------------------------------

const SYNC_BASE_PREFIX = "sitplan-sync-base-v1:";

function stableStringify(value) {
  if (value === undefined) return "null";
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  return `{${Object.keys(value).sort()
    .filter(key => value[key] !== undefined)
    .map(key => `${JSON.stringify(key)}:${stableStringify(value[key])}`)
    .join(",")}}`;
}

function fingerprint(value) {
  const text = stableStringify(value);
  let h1 = 0x811c9dc5;
  let h2 = 0x01000193;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    h1 = Math.imul(h1 ^ code, 0x01000193) >>> 0;
    h2 = Math.imul(h2 ^ code, 0x5bd1e995) >>> 0;
  }
  return `${h1.toString(36)}.${h2.toString(36)}.${text.length.toString(36)}`;
}

function snapshotKeys(state) {
  const keyed = (items, keyOf, valueOf) => Object.fromEntries(items.map(item => [keyOf(item), fingerprint(valueOf(item))]));
  return {
    contacts: keyed(state.clients || [], client => client.id, client => client),
    activities: keyed(state.activities || [], activity => activity.id, activity => activity),
    goals: keyed(Object.entries(state.weeklyGoals || {}), ([week]) => week, ([, target]) => target),
    settings: fingerprint(state.settings || {})
  };
}

function readSyncBase(userId) {
  try {
    const base = JSON.parse(localStorage.getItem(SYNC_BASE_PREFIX + userId));
    return base?.contacts && base?.activities && base?.goals ? base : null;
  } catch (_) {
    return null;
  }
}

function writeSyncBase(userId, keys) {
  try {
    localStorage.setItem(SYNC_BASE_PREFIX + userId, JSON.stringify({ ...keys, syncedAt: new Date().toISOString() }));
  } catch (error) {
    console.warn("Não foi possível guardar a base de sincronização:", error);
  }
}

/**
 * Merge the cloud state into this device's state without losing local work.
 *
 * - Row changed or created here since the last sync -> keep this device's version.
 * - Row deleted here since the last sync -> stays deleted.
 * - Anything else -> take the cloud version (new, edited or deleted elsewhere).
 *
 * Without a sync base (first run after this update, or a new device) there is
 * no common ancestor: rows that exist on only one side are kept, and for rows
 * that differ the newer side wins (this device's last local save vs the
 * newest change in the cloud).
 *
 * Returns { state, hasLocalChanges }. When hasLocalChanges is true the caller
 * should queue a save so the kept local changes reach the cloud.
 */
export function mergeCloudState(localState, remoteState, { localSavedAt = "" } = {}) {
  const userId = session?.user?.id;
  const base = userId ? readSyncBase(userId) : null;
  const local = snapshotKeys(localState);
  const remote = snapshotKeys(remoteState);
  const localWinsWithoutBase = Boolean(localSavedAt) && localSavedAt > (remoteState.remoteUpdatedAt || "");
  let hasLocalChanges = !base;

  const mergeKeyed = (localItems, remoteItems, keyOf, localKeys, remoteKeys, baseKeys) => {
    const localById = new Map(localItems.map(item => [keyOf(item), item]));
    const remoteById = new Map(remoteItems.map(item => [keyOf(item), item]));
    const result = [];
    const seen = new Set();
    const take = (id, item) => {
      if (seen.has(id)) return;
      seen.add(id);
      result.push(item);
    };

    for (const [id, remoteItem] of remoteById) {
      const localItem = localById.get(id);
      if (baseKeys) {
        const changedHere = localItem ? localKeys[id] !== baseKeys[id] : false;
        const deletedHere = !localItem && id in baseKeys;
        if (deletedHere) {
          hasLocalChanges = true;
          continue;
        }
        if (changedHere) {
          hasLocalChanges = true;
          take(id, localItem);
        } else {
          take(id, remoteItem);
        }
      } else {
        const differs = localItem && localKeys[id] !== remoteKeys[id];
        take(id, differs && localWinsWithoutBase ? localItem : remoteItem);
      }
    }
    for (const [id, localItem] of localById) {
      if (seen.has(id)) continue;
      if (baseKeys && id in baseKeys && localKeys[id] === baseKeys[id]) continue; // deleted in the cloud
      hasLocalChanges = true;
      take(id, localItem);
    }
    return result;
  };

  const clients = mergeKeyed(
    localState.clients || [], remoteState.clients || [], client => client.id,
    local.contacts, remote.contacts, base?.contacts
  );
  const activities = mergeKeyed(
    localState.activities || [], remoteState.activities || [], activity => activity.id,
    local.activities, remote.activities, base?.activities
  );
  const goals = mergeKeyed(
    Object.entries(localState.weeklyGoals || {}), Object.entries(remoteState.weeklyGoals || {}), ([week]) => week,
    local.goals, remote.goals, base?.goals
  );

  let settings = remoteState.settings || {};
  const settingsChangedHere = base ? local.settings !== base.settings : localWinsWithoutBase;
  if (settingsChangedHere && Object.keys(localState.settings || {}).length) {
    settings = localState.settings;
    if (local.settings !== remote.settings) hasLocalChanges = true;
  }

  return {
    state: {
      clients,
      activities,
      weeklyGoals: Object.fromEntries(goals),
      settings
    },
    hasLocalChanges
  };
}

async function authRequest(path, options = {}) {
  try {
    const response = await fetch(`${SUPABASE_URL}${path}`, {
      ...options,
      headers: {
        apikey: SUPABASE_PUBLISHABLE_KEY,
        "Content-Type": "application/json",
        ...(options.headers || {})
      }
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      const error = new Error(data.msg || data.error_description || data.message || data.error || `Erro ${response.status}`);
      error.status = response.status;
      error.code = data.error_code || data.code || data.error || "";
      return { data: null, error };
    }
    return { data, error: null };
  } catch (error) {
    const networkError = new Error("Falha de conexao com o Supabase. Verifique sua internet e tente novamente.");
    networkError.code = "NETWORK";
    return { data: null, error: networkError };
  }
}

async function restRequest(path, options = {}, attempt = 0) {
  const maxRetries = 2;
  const retryDelay = () => new Promise(resolve => setTimeout(resolve, 300 * Math.pow(2, attempt)));
  try {
    const response = await fetch(`${SUPABASE_URL}${path}`, {
      ...options,
      headers: {
        ...authHeaders(),
        "Content-Type": "application/json",
        ...(options.headers || {})
      }
    });
    const text = await response.text();
    const data = text ? JSON.parse(text) : [];
    if (!response.ok) {
      const error = new Error(data.message || data.hint || data.code || `Erro ${response.status}`);
      error.status = response.status;
      error.code = data.code || "";
      const transient = response.status >= 500 || response.status === 429;
      if (transient && attempt < maxRetries) {
        await retryDelay();
        return restRequest(path, options, attempt + 1);
      }
      throw error;
    }
    return data;
  } catch (error) {
    if (error?.status) throw error;
    if (attempt < maxRetries) {
      await retryDelay();
      return restRequest(path, options, attempt + 1);
    }
    const networkError = new Error("Não consegui falar com o Supabase. Verifique se o projeto está ativo e se a internet está liberando o acesso.");
    networkError.code = "NETWORK";
    throw networkError;
  }
}

function authHeaders() {
  return {
    apikey: SUPABASE_PUBLISHABLE_KEY,
    Authorization: `Bearer ${session?.access_token || SUPABASE_PUBLISHABLE_KEY}`
  };
}

async function validateOrRefreshSession() {
  const response = await fetch(`${SUPABASE_URL}/auth/v1/user`, { headers: authHeaders() }).catch(() => null);
  if (response?.ok) {
    session.user = await response.json();
    storeSession(session);
    return true;
  }
  return refreshSession();
}

async function ensureFreshSession() {
  const stored = readSession();
  if (stored?.access_token && Number(stored.expires_at || 0) > Number(session?.expires_at || 0)) session = stored;
  if (!session?.expires_at || session.expires_at * 1000 > Date.now() + 60000) return;
  const refreshed = await refreshSession();
  if (!refreshed) throw new Error("Sua sessão expirou. Entre novamente.");
}

async function refreshSession() {
  const stored = readSession();
  if (stored?.access_token && stored.refresh_token !== session?.refresh_token
      && Number(stored.expires_at || 0) * 1000 > Date.now() + 60000) {
    session = stored;
    return true;
  }
  if (!session?.refresh_token) return false;
  const result = await authRequest("/auth/v1/token?grant_type=refresh_token", {
    method: "POST",
    body: JSON.stringify({ refresh_token: session.refresh_token })
  });
  if (result.error) return false;
  storeSession(result.data);
  await determineAccessMode();
  return true;
}

async function determineAccessMode() {
  accessMode = { role: "owner", ownerUserId: session?.user?.id || null };
  if (!session?.access_token || !session?.user?.id) return accessMode;
  try {
    const response = await fetch(`${SUPABASE_URL}/rest/v1/sitplan_shares?select=owner_user_id,role&viewer_user_id=eq.${session.user.id}&active=eq.true&limit=1`, {
      headers: authHeaders()
    });
    if (!response.ok) return accessMode;
    const rows = await response.json().catch(() => []);
    const share = rows[0];
    if (share?.owner_user_id) {
      accessMode = {
        role: share.role || "viewer",
        ownerUserId: share.owner_user_id
      };
    }
  } catch (_) {
    accessMode = { role: "owner", ownerUserId: session?.user?.id || null };
  }
  return accessMode;
}

function storeSession(value) {
  session = {
    ...value,
    expires_at: value.expires_at || Math.floor(Date.now() / 1000) + Number(value.expires_in || 3600)
  };
  localStorage.setItem(SESSION_KEY, JSON.stringify(session));
}

function readSession() {
  try {
    return JSON.parse(localStorage.getItem(SESSION_KEY));
  } catch (_) {
    return null;
  }
}

function clearSession() {
  session = null;
  localStorage.removeItem(SESSION_KEY);
}

function startPolling() {
  stopPolling();
  if (!isCloudConnected()) return;
  pollingTimer = setInterval(() => remoteChangeHandler(), 15000);
}

function stopPolling() {
  clearInterval(pollingTimer);
}

function friendlyError(error) {
  if (error?.code === "NETWORK") return error.message;
  if (error?.status === 401 || error?.status === 403) return "Sessão expirada ou permissão negada. Saia da nuvem e entre novamente.";
  if (error?.status === 404 || error?.code === "42P01") return "Banco do Supabase sem as tabelas do SITPLAN. Execute o arquivo supabase-schema.sql.";
  return error?.message || "Falha ao sincronizar.";
}
