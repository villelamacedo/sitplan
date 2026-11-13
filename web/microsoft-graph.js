const CLIENT_ID_KEY = "sitplan-microsoft-client-id";
const TOKEN_KEY = "sitplan-microsoft-token";
const VERIFIER_KEY = "sitplan-microsoft-verifier";
const STATE_KEY = "sitplan-microsoft-state";
const SCOPES = "openid profile offline_access User.Read Calendars.ReadWrite";
const AUTHORITY = "https://login.microsoftonline.com/common/oauth2/v2.0";

let token = readJson(sessionStorage.getItem(TOKEN_KEY));

export function microsoftClientId() {
  return localStorage.getItem(CLIENT_ID_KEY) || "";
}

export function configureMicrosoft(clientId) {
  const value = String(clientId || "").trim();
  if (value) localStorage.setItem(CLIENT_ID_KEY, value);
  else localStorage.removeItem(CLIENT_ID_KEY);
}

export function isMicrosoftConnected() {
  return Boolean(token?.accessToken && token.expiresAt > Date.now() + 60000);
}

export async function initializeMicrosoft() {
  const params = new URLSearchParams(window.location.search);
  const code = params.get("code");
  const returnedState = params.get("state");
  const error = params.get("error");
  if (!code && !error) return { connected: isMicrosoftConnected() };

  const cleanUrl = `${window.location.origin}${window.location.pathname}${window.location.hash}`;
  window.history.replaceState({}, document.title, cleanUrl);
  if (error) throw new Error(params.get("error_description") || "A Microsoft não autorizou a conexão.");
  if (returnedState !== sessionStorage.getItem(STATE_KEY)) throw new Error("A validação da conexão Microsoft falhou.");

  const clientId = microsoftClientId();
  const verifier = sessionStorage.getItem(VERIFIER_KEY);
  if (!clientId || !verifier) throw new Error("Configuração Microsoft incompleta.");
  const body = new URLSearchParams({
    client_id: clientId,
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri(),
    code_verifier: verifier,
    scope: SCOPES
  });
  const response = await fetch(`${AUTHORITY}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error_description || "Não foi possível concluir a conexão Microsoft.");
  storeToken(data);
  sessionStorage.removeItem(VERIFIER_KEY);
  sessionStorage.removeItem(STATE_KEY);
  return { connected: true, justConnected: true };
}

export async function connectMicrosoft() {
  const clientId = microsoftClientId();
  if (!clientId) throw new Error("Informe primeiro o ID do aplicativo Microsoft.");
  const verifier = randomString(64);
  const state = crypto.randomUUID();
  const challenge = base64Url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  sessionStorage.setItem(VERIFIER_KEY, verifier);
  sessionStorage.setItem(STATE_KEY, state);
  const url = new URL(`${AUTHORITY}/authorize`);
  url.searchParams.set("client_id", clientId);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", redirectUri());
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", SCOPES);
  url.searchParams.set("state", state);
  url.searchParams.set("code_challenge", challenge);
  url.searchParams.set("code_challenge_method", "S256");
  window.location.assign(url.toString());
}

export function disconnectMicrosoft() {
  token = null;
  sessionStorage.removeItem(TOKEN_KEY);
}

export async function createTeamsEvent({
  subject,
  body,
  startDateTime,
  endDateTime,
  attendeeEmail,
  attendeeName,
  transactionId
}) {
  await ensureToken();
  const event = {
    subject,
    body: { contentType: "HTML", content: escapeHtml(body).replace(/\n/g, "<br>") },
    start: { dateTime: startDateTime, timeZone: "E. South America Standard Time" },
    end: { dateTime: endDateTime, timeZone: "E. South America Standard Time" },
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness",
    allowNewTimeProposals: true,
    reminderMinutesBeforeStart: 30,
    isReminderOn: true,
    transactionId: transactionId || crypto.randomUUID(),
    attendees: attendeeEmail ? [{
      emailAddress: { address: attendeeEmail, name: attendeeName || attendeeEmail },
      type: "required"
    }] : []
  };
  const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token.accessToken}`,
      "Content-Type": "application/json",
      Prefer: 'outlook.timezone="E. South America Standard Time"'
    },
    body: JSON.stringify(event)
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error?.message || "Não foi possível criar o evento no Microsoft 365.");
  return {
    id: data.id,
    webLink: data.webLink || "",
    joinUrl: data.onlineMeeting?.joinUrl || data.onlineMeetingUrl || ""
  };
}

export async function checkCalendarAvailability({ startDateTime, endDateTime, excludeEventId = "" }) {
  await ensureToken();
  const url = new URL("https://graph.microsoft.com/v1.0/me/calendarView");
  url.searchParams.set("startDateTime", startDateTime);
  url.searchParams.set("endDateTime", endDateTime);
  url.searchParams.set("$select", "id,subject,start,end,isCancelled,showAs");
  url.searchParams.set("$orderby", "start/dateTime");
  const response = await fetch(url.toString(), {
    headers: {
      Authorization: `Bearer ${token.accessToken}`,
      Prefer: 'outlook.timezone="E. South America Standard Time"'
    }
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error?.message || "Não foi possível consultar sua agenda Microsoft.");
  const busy = (data.value || []).filter(event =>
    event.id !== excludeEventId
    && !event.isCancelled
    && !["free", "workingElsewhere"].includes(String(event.showAs || "").toLowerCase())
  );
  return {
    available: busy.length === 0,
    conflicts: busy.map(event => ({
      id: event.id,
      subject: event.subject || "Compromisso",
      start: event.start?.dateTime || "",
      end: event.end?.dateTime || "",
      showAs: event.showAs || ""
    }))
  };
}

async function ensureToken() {
  if (isMicrosoftConnected()) return;
  if (!token?.refreshToken) throw new Error("Conecte sua conta Microsoft 365.");
  const response = await fetch(`${AUTHORITY}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: microsoftClientId(),
      grant_type: "refresh_token",
      refresh_token: token.refreshToken,
      redirect_uri: redirectUri(),
      scope: SCOPES
    })
  });
  const data = await response.json();
  if (!response.ok) {
    disconnectMicrosoft();
    throw new Error("Sua conexão Microsoft expirou. Entre novamente.");
  }
  storeToken(data);
}

function storeToken(data) {
  token = {
    accessToken: data.access_token,
    refreshToken: data.refresh_token || token?.refreshToken || "",
    expiresAt: Date.now() + Number(data.expires_in || 3600) * 1000
  };
  sessionStorage.setItem(TOKEN_KEY, JSON.stringify(token));
}

function redirectUri() {
  return `${window.location.origin}${window.location.pathname}`;
}

function randomString(length) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
  const bytes = crypto.getRandomValues(new Uint8Array(length));
  return Array.from(bytes, value => chars[value % chars.length]).join("");
}

function base64Url(buffer) {
  return btoa(String.fromCharCode(...new Uint8Array(buffer)))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

function readJson(value) {
  try {
    return JSON.parse(value);
  } catch (_) {
    return null;
  }
}

function escapeHtml(value = "") {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
