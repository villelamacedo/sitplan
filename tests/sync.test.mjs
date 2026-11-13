// Two "devices" share one fake Supabase. Each device loads its own copy of
// cloud-sync.js (fresh module instance + its own localStorage).
import assert from "node:assert/strict";
import { pathToFileURL } from "node:url";
// Run: node tests/sync.test.mjs
const SRC = process.argv[2] || new URL("../web/cloud-sync.js", import.meta.url).pathname;

const db = { contacts: new Map(), activities: new Map(), weekly_goals: new Map() };
const calls = [];
let failWrites = false;
const USER = { id: "u1", email: "a@b.c" };

function makeStorage() { const m = new Map(); return { getItem: k => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, String(v)), removeItem: k => m.delete(k) }; }

async function fakeFetch(url, opts = {}) {
  const u = new URL(url); const method = opts.method || "GET";
  const ok = (body, status = 200) => ({ ok: true, status, json: async () => body, text: async () => (body === undefined ? "" : JSON.stringify(body)) });
  if (u.pathname.startsWith("/auth/v1/user")) return ok(USER);
  if (u.pathname.includes("sitplan_shares")) return { ok: false, status: 404, json: async () => ({}), text: async () => "{}" };
  const table = u.pathname.split("/").pop();
  calls.push({ method, table, n: opts.body ? JSON.parse(opts.body).length : undefined, q: u.search });
  if (method !== "GET" && failWrites) return { ok: false, status: 400, json: async () => ({ message: "boom" }), text: async () => '{"message":"boom"}' };
  if (method === "GET") {
    const now = new Date().toISOString();
    if (table === "weekly_goals") return ok([...db.weekly_goals.values()]);
    return ok([...db[table].values()]);
  }
  if (method === "POST") {
    for (const row of JSON.parse(opts.body)) {
      const key = table === "weekly_goals" ? row.week_start : row.id;
      db[table].set(key, { ...row, updated_at: new Date().toISOString() });
    }
    return ok(undefined, 201);
  }
  if (method === "DELETE") {
    const [col, expr] = [...u.searchParams.entries()][0];
    const ids = decodeURIComponent(expr).replace(/^in\.\(|\)$/g, "").split(",");
    for (const id of ids) db[table].delete(id);
    return ok(undefined, 204);
  }
}

let instance = 0;
async function device(name) {
  const storage = makeStorage();
  storage.setItem("sitplan-supabase-session", JSON.stringify({ access_token: "t", refresh_token: "r", expires_at: Math.floor(Date.now() / 1000) + 3600, user: USER }));
  globalThis.localStorage = storage;
  const mod = await import(pathToFileURL(SRC).href + "?dev=" + (++instance));
  const d = { name, storage, mod, state: { clients: [], activities: [], weeklyGoals: {}, settings: { defaultAreaCode: "11" } } };
  d.use = () => { globalThis.localStorage = storage; };
  d.save = async () => { d.use(); d.mod.queueCloudSave(d.state); await new Promise(r => setTimeout(r, 30)); };
  d.pull = async () => { d.use(); const remote = await d.mod.loadCloudState(); const m = d.mod.mergeCloudState(d.state, remote, { localSavedAt: d.localSavedAt || "" }); d.state = m.state; if (m.hasLocalChanges) await d.save(); return m; };
  return d;
}

const events = [];
globalThis.window = { dispatchEvent: e => events.push(e.type + ":" + JSON.stringify(e.detail ?? null)), location: { origin: "http://x" } };
globalThis.CustomEvent = class { constructor(type, init) { this.type = type; this.detail = init?.detail; } };
globalThis.fetch = fakeFetch;

const c = (id, name, extra = {}) => ({ id, name, stage: "Base", ...extra });
const names = d => d.state.clients.map(x => x.name).sort();

// --- Setup: phone creates 6 contacts, laptop pulls them
const phone = await device("phone");
const laptop = await device("laptop");
phone.state.clients = ["a", "b", "c", "d", "e", "f"].map(id => c(id, id.toUpperCase()));
phone.state.activities = [{ id: "act1", title: "Ligar", date: "2026-11-14", clientId: "a" }];
await phone.save();
assert.equal(db.contacts.size, 6);
await laptop.pull();
assert.deepEqual(names(laptop), ["A", "B", "C", "D", "E", "F"]);

// --- 1) Bug fixed: contact created on phone is NOT deleted by a stale laptop save
phone.state.clients.push(c("g", "G")); await phone.save();
laptop.state.clients.find(x => x.id === "a").name = "A (laptop)"; await laptop.save();
assert.ok(db.contacts.has("g"), "contact created on the phone survived the laptop save");
assert.equal(db.contacts.get("a").data.name, "A (laptop)");

// --- 2) Delta: laptop only sent the 1 changed contact, not all 7
const lastPost = calls.filter(x => x.method === "POST" && x.table === "contacts").at(-1);
assert.equal(lastPost.n, 1, "only the changed contact is uploaded");

// --- 3) Edits on different contacts on both devices are both kept
phone.state.clients.find(x => x.id === "b").name = "B (phone)";
await phone.pull(); // pull before save: local edit must survive the merge
assert.equal(phone.state.clients.find(x => x.id === "b").name, "B (phone)", "unsynced local edit survives polling");
assert.equal(phone.state.clients.find(x => x.id === "a").name, "A (laptop)", "remote edit arrives");
assert.equal(db.contacts.get("b").data.name, "B (phone)", "merge queued the local edit to the cloud");
await laptop.pull();
assert.equal(laptop.state.clients.find(x => x.id === "b").name, "B (phone)");

// --- 4) Deletion on one device propagates, and is not resurrected
laptop.state.clients = laptop.state.clients.filter(x => x.id !== "c"); await laptop.save();
assert.ok(!db.contacts.has("c"));
await phone.pull();
assert.ok(!phone.state.clients.some(x => x.id === "c"), "deleted elsewhere -> removed here");

// --- 5) Failed save: local edits stay and are not wiped by polling
failWrites = true;
phone.state.clients.find(x => x.id === "d").name = "D offline edit"; await phone.save();
assert.ok(events.some(e => e.startsWith("cloud-sync-error")));
const m = await phone.pull();
assert.equal(phone.state.clients.find(x => x.id === "d").name, "D offline edit", "poll after failed save keeps local edit");
failWrites = false;
await phone.save();
assert.equal(db.contacts.get("d").data.name, "D offline edit", "edit reaches cloud once writes work again");

// --- 6) Mass deletion guard still works
const before = db.contacts.size;
laptop.use(); await laptop.pull();
laptop.state.clients = laptop.state.clients.slice(0, 1); await laptop.save();
assert.equal(db.contacts.size, before, "mass deletion blocked");
assert.ok(events.some(e => e.startsWith("cloud-sync-blocked")));

// --- 7) Upgrade path (no sync base): device with newer local work keeps it
const old = await device("old-install");
old.state.clients = [...db.contacts.values()].map(r => ({ ...r.data, crmSettings: undefined }));
old.state.clients = old.state.clients.map(({ crmSettings, ...rest }) => rest);
old.state.clients.find(x => x.id === "e").name = "E edited during 6 weeks offline";
old.state.clients.push(c("h", "H only local"));
old.localSavedAt = new Date(Date.now() + 1000).toISOString(); // newer than cloud
await old.pull();
assert.equal(db.contacts.get("e").data.name, "E edited during 6 weeks offline");
assert.ok(db.contacts.has("h"));
// and a stale device (older local save) does NOT overwrite newer cloud data
const stale = await device("stale");
stale.state.clients = [c("a", "A ancient")];
stale.localSavedAt = "2020-01-01T00:00:00Z";
await stale.pull();
assert.equal(db.contacts.get("a").data.name, "A (laptop)", "stale device does not overwrite");
assert.ok(stale.state.clients.length >= 6, "stale device receives cloud contacts");

// --- 8) Settings change propagates
phone.use(); await phone.pull();
phone.state.settings = { defaultAreaCode: "21" }; await phone.save();
await laptop.pull();
assert.equal(laptop.state.settings.defaultAreaCode, "21");

console.log("ALL SYNC TESTS PASSED");
