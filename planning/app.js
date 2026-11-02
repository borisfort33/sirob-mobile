/* SIRØB Planning : ta journée au même endroit.
 *
 * - Rendez-vous : Google Agenda (lecture seule), connexion Google sans serveur.
 * - Tâches : les tiennes, enregistrées dans OneDrive (planning.json), synchronisées partout.
 *   Le Mac n'a pas besoin d'être allumé pour en ajouter.
 * - Morceaux en attente : le catalogue.json écrit par l'app du Mac (statut, deadline, client…).
 */
"use strict";

const CFG = window.SIROB_CONFIG;
const AUTH = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPE = "Files.ReadWrite offline_access User.Read";
const GAUTH = "https://accounts.google.com/o/oauth2/v2/auth";
const GCAL = "https://www.googleapis.com/calendar/v3";
const GMAIL = "https://gmail.googleapis.com/gmail/v1/users/me";
const GSCOPE = "https://www.googleapis.com/auth/calendar.readonly https://www.googleapis.com/auth/gmail.readonly";
const DIGEST_SUBJECT = "Résumé mails SIRØB";
const REDIRECT = new URL("./", location.href).href;
const VERSION = "1.1";
const DAYS_BEFORE = 7;
const DAYS_AFTER = 30;

const $ = (id) => document.getElementById(id);

// --- Petit stockage local tolérant (navigation privée, stockage plein…) -----------
const store = {
  get(key, fallback = null) {
    try { const v = localStorage.getItem("sirob." + key); return v == null ? fallback : JSON.parse(v); }
    catch { return fallback; }
  },
  set(key, value) {
    try { localStorage.setItem("sirob." + key, JSON.stringify(value)); } catch { /* plein */ }
  },
  del(key) { try { localStorage.removeItem("sirob." + key); } catch { /* rien */ } },
};

// --- Dates (toujours en heure locale) -------------------------------------------------

const pad = (n) => String(n).padStart(2, "0");
const keyOf = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const dateOf = (key) => new Date(key + "T00:00:00");
const todayKey = () => keyOf(new Date());
function addDays(key, n) { const d = dateOf(key); d.setDate(d.getDate() + n); return keyOf(d); }
const fmtDay = (key, opts) => dateOf(key).toLocaleDateString("fr-FR", opts);
const shortDay = (key) => fmtDay(key, { weekday: "short", day: "numeric", month: "short" });
const hm = (d) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;

function relDay(key) {
  const n = Math.round((dateOf(key) - dateOf(todayKey())) / 86_400_000);
  if (n === 0) return "Aujourd'hui";
  if (n === 1) return "Demain";
  if (n === -1) return "Hier";
  return "";
}

// --- Connexion Microsoft (comme SIRØB Mobile) ----------------------------------------

class NeedLogin extends Error {}

function b64url(bytes) {
  let s = "";
  for (const b of new Uint8Array(bytes)) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
const randomString = () => b64url(crypto.getRandomValues(new Uint8Array(32)));

async function login() {
  const verifier = randomString();
  const state = randomString();
  const challenge = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  store.set("pkce", { verifier, state });
  const params = new URLSearchParams({
    client_id: CFG.clientId, response_type: "code", redirect_uri: REDIRECT, scope: SCOPE,
    code_challenge: challenge, code_challenge_method: "S256", state, response_mode: "query",
  });
  const hint = store.get("account");
  if (hint) params.set("login_hint", hint);
  location.href = `${AUTH}/authorize?${params}`;
}

async function tokenRequest(fields) {
  const resp = await fetch(`${AUTH}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CFG.clientId, scope: SCOPE, ...fields }),
  });
  const data = await resp.json().catch(() => ({}));
  if (!data.access_token) throw new Error(data.error_description || data.error || "Connexion refusée");
  store.set("tok", {
    access: data.access_token,
    refresh: data.refresh_token,
    exp: Date.now() + (data.expires_in || 3600) * 1000,
  });
}

/** Retour de Microsoft (?code=…) ou de Google (#access_token=…). */
async function handleRedirect() {
  const hash = new URLSearchParams(location.hash.slice(1));
  if (hash.has("access_token") || (hash.has("error") && (hash.get("state") || "").startsWith("g."))) {
    history.replaceState(null, "", REDIRECT);
    if (hash.get("state") !== store.get("plan.gstate")) return;
    store.del("plan.gstate");
    if (hash.has("access_token")) {
      store.set("plan.gtok", {
        access: hash.get("access_token"),
        exp: Date.now() + (Number(hash.get("expires_in")) || 3600) * 1000,
        scope: hash.get("scope") || "",
      });
      store.set("plan.gcal", true);
      store.del("plan.gNeedsClick");
    } else {
      // Connexion silencieuse refusée : il faudra une touche sur « Connecter Google Agenda ».
      store.set("plan.gNeedsClick", true);
    }
    return;
  }
  const params = new URLSearchParams(location.search);
  if (!params.has("code") && !params.has("error")) return;
  history.replaceState(null, "", REDIRECT);
  if (params.has("error")) throw new Error(params.get("error_description") || params.get("error"));
  const pkce = store.get("pkce");
  store.del("pkce");
  if (!pkce || pkce.state !== params.get("state")) throw new Error("Connexion interrompue, réessaie.");
  await tokenRequest({
    grant_type: "authorization_code", code: params.get("code"),
    redirect_uri: REDIRECT, code_verifier: pkce.verifier,
  });
}

let refreshing = null;
async function accessToken(force = false) {
  const tok = store.get("tok");
  if (!tok) throw new NeedLogin();
  if (!force && Date.now() < tok.exp - 60_000) return tok.access;
  if (!tok.refresh) throw new NeedLogin();
  refreshing ??= tokenRequest({ grant_type: "refresh_token", refresh_token: tok.refresh })
    .catch((e) => {
      if (/invalid_grant|expired|AADSTS700084|AADSTS70008/i.test(e.message)) {
        store.del("tok");
        throw new NeedLogin();
      }
      throw e;
    })
    .finally(() => { refreshing = null; });
  await refreshing;
  return store.get("tok").access;
}

function logout() {
  for (const k of ["tok", "account", "catalogue", "plan.data", "plan.etag", "plan.dirty",
                   "plan.gtok", "plan.gcal", "plan.events"]) store.del(k);
  location.replace(REDIRECT);
}

// --- Microsoft Graph ----------------------------------------------------------------

const encPath = (p) => p.split("/").map(encodeURIComponent).join("/");

async function graph(path, options = {}, retry = true) {
  let resp;
  try {
    resp = await fetch(GRAPH + path, {
      ...options,
      headers: {
        Authorization: `Bearer ${await accessToken()}`,
        ...(options.body ? { "Content-Type": "application/json" } : {}),
        ...(options.headers || {}),
      },
    });
  } catch (e) {
    if (e instanceof NeedLogin) throw e;
    const err = new Error("Pas de connexion internet");
    err.offline = true;
    throw err;
  }
  if (resp.status === 401 && retry) {
    await accessToken(true);
    return graph(path, options, false);
  }
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const err = new Error(resp.status === 404 ? "Fichier introuvable sur OneDrive"
      : (data.error && data.error.message) || `Erreur OneDrive ${resp.status}`);
    err.status = resp.status;
    throw err;
  }
  return data;
}

async function rememberAccount() {
  if (store.get("account")) return;
  try {
    const me = await graph("/me?$select=userPrincipalName,mail");
    store.set("account", me.userPrincipalName || me.mail || "");
  } catch { /* pas grave */ }
}

// --- Tâches : planning.json dans OneDrive ----------------------------------------------
// Chaque tâche : {id, t: texte, d: "AAAA-MM-JJ", h: "HH:MM" | "", done, song?, c: créée, u: modifiée, del?}
// Fusion par tâche (la plus récemment modifiée gagne) : iPhone, iPad et Mac peuvent écrire
// le même fichier sans s'écraser.

let plan = store.get("plan.data", { tasks: [] });
let planETag = store.get("plan.etag");
const planPath = () => store.get("plan.path", CFG.planning);

function mergeTasks(a, b) {
  const byId = new Map();
  for (const t of [...a, ...b]) {
    const cur = byId.get(t.id);
    if (!cur || (t.u || 0) > (cur.u || 0)) byId.set(t.id, t);
  }
  const limit = Date.now() - 30 * 86_400_000;   // les suppressions sont oubliées après 30 jours
  return [...byId.values()].filter((t) => !(t.del && t.u < limit));
}

function keepPlan() {
  store.set("plan.data", plan);
  store.set("plan.etag", planETag);
}

let syncing = null;
let saveTimer = null;

function changed() {
  store.set("plan.dirty", true);
  keepPlan();
  render();
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => syncPlan(), 700);
}

/** Lit planning.json, fusionne avec l'iPhone et réécrit si besoin. */
function syncPlan() {
  syncing ??= doSync().finally(() => { syncing = null; });
  return syncing;
}

async function doSync(attempt = 0) {
  try {
    let remote = null;
    let etag = null;
    try {
      const item = await graph(`/me/drive/root:/${encPath(planPath())}`);
      etag = item.eTag;
      if (etag !== planETag || !plan.tasks) {
        const resp = await fetch(item["@microsoft.graph.downloadUrl"]);
        remote = await resp.json();
      }
    } catch (e) {
      if (e.status !== 404) throw e;
    }
    if (remote) {
      plan = { ...remote, tasks: mergeTasks(plan.tasks || [], remote.tasks || []) };
      planETag = etag;
      keepPlan();
      render();
    } else if (etag === null) {
      planETag = null;   // pas encore de fichier : on le crée
    }
    if (!store.get("plan.dirty") && etag !== null) return;
    if (!store.get("plan.dirty") && !(plan.tasks || []).length) return;
    plan.updated = new Date().toISOString();
    plan.app = `SIRØB Planning ${VERSION}`;
    const saved = await graph(`/me/drive/root:/${encPath(planPath())}:/content`, {
      method: "PUT",
      body: JSON.stringify(plan, null, 1),
      headers: planETag ? { "If-Match": planETag } : {},
    });
    planETag = saved.eTag;
    store.set("plan.dirty", false);
    keepPlan();
    setStatusLine();
  } catch (e) {
    if (e instanceof NeedLogin) return showLogin();
    if ((e.status === 412 || e.status === 409) && attempt < 3) {
      planETag = null;   // modifié ailleurs entre-temps : on relit, on fusionne, on réécrit
      return doSync(attempt + 1);
    }
    setStatusLine(e.offline ? "Hors ligne · tes tâches seront envoyées à OneDrive au retour du réseau"
      : `Tâches : ${e.message}`);
  }
}

const liveTasks = () => (plan.tasks || []).filter((t) => !t.del);
const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

function addTask(text, day, h = "", song) {
  const now = Date.now();
  const t = { id: newId(), t: text, d: day, h, done: false, c: now, u: now };
  if (song) t.song = song;
  plan.tasks = [...(plan.tasks || []), t];
  changed();
  return t;
}

function updateTask(id, patch) {
  plan.tasks = plan.tasks.map((t) => (t.id === id ? { ...t, ...patch, u: Date.now() } : t));
  changed();
}

// --- Catalogue du Mac : morceaux en attente --------------------------------------------

let catalogue = store.get("catalogue");
const STATUS_COLORS = { "Idée": "#b67ce8", "Active": "#4aa3ff", "Mix": "#ff9f43", "Done": "#4cd964" };
const STATUS_ORDER = { "Mix": 0, "Active": 1, "Idée": 2 };

async function loadCatalogue() {
  const path = store.get("catPath", CFG.catalogue);
  try {
    const item = await graph(`/me/drive/root:/${encPath(path)}`);
    if (catalogue && catalogue.eTag === item.eTag) return;
    const resp = await fetch(item["@microsoft.graph.downloadUrl"]);
    const data = await resp.json();
    data.eTag = item.eTag;
    catalogue = data;
    store.set("catalogue", data);
    render();
  } catch (e) {
    if (e instanceof NeedLogin) return showLogin();
    if (e.status === 404) setStatusLine("Catalogue introuvable : vérifie son emplacement dans ⚙︎.");
  }
}

/** Morceaux sur lesquels il reste du travail : statut Idée / Active / Mix, ou une deadline. */
function pendingSongs() {
  const out = [];
  for (const s of (catalogue && catalogue.songs) || []) {
    const i = s.i || {};
    if (i.st === "Done") continue;
    if (!i.st && !i.deadline) continue;
    const last = Math.max(s.m || 0, ...(s.b || []).map((f) => f[1]));
    out.push({ name: s.n, info: i, last, where: (s.d || "").split("/").slice(-2, -1)[0] || "" });
  }
  out.sort((a, b) =>
    (a.info.deadline || "9999").localeCompare(b.info.deadline || "9999")
    || (STATUS_ORDER[a.info.st] ?? 3) - (STATUS_ORDER[b.info.st] ?? 3)
    || b.last - a.last);
  return out;
}

const songInfo = (name) => {
  const s = ((catalogue && catalogue.songs) || []).find((x) => x.n === name);
  return s && s.i;
};

// --- Google Agenda ---------------------------------------------------------------------

let events = store.get("plan.events", { items: [], at: 0 });
let gcalState = "";   // "", "off", "expired", "error", "ok"

function googleLogin(silent) {
  const state = "g." + randomString();
  store.set("plan.gstate", state);
  if (silent) store.set("plan.gTry", Date.now());
  const params = new URLSearchParams({
    client_id: CFG.googleClientId, redirect_uri: REDIRECT, response_type: "token",
    scope: GSCOPE, include_granted_scopes: "true", state,
  });
  const hint = store.get("plan.gemail");
  if (hint) params.set("login_hint", hint);
  if (silent) params.set("prompt", "none");
  location.href = `${GAUTH}?${params}`;
}

function googleToken() {
  const tok = store.get("plan.gtok");
  return tok && Date.now() < tok.exp - 60_000 ? tok.access : null;
}

async function gget(path, params = {}, base = GCAL) {
  const tok = googleToken();
  if (!tok) { const e = new Error("expired"); e.expired = true; throw e; }
  const resp = await fetch(`${base}${path}?${new URLSearchParams(params)}`, {
    headers: { Authorization: `Bearer ${tok}` },
  });
  if (resp.status === 401) { store.del("plan.gtok"); const e = new Error("expired"); e.expired = true; throw e; }
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const e = new Error((data.error && data.error.message) || `Erreur Google ${resp.status}`);
    e.status = resp.status;
    throw e;
  }
  return data;
}

async function loadEvents() {
  if (!CFG.googleClientId) { gcalState = "off"; return renderBanner(); }
  if (!store.get("plan.gcal")) { gcalState = "off"; return renderBanner(); }
  try {
    const cals = await gget("/users/me/calendarList", { minAccessRole: "reader" });
    const primary = (cals.items || []).find((c) => c.primary);
    if (primary) store.set("plan.gemail", primary.id);
    const off = new Set(store.get("plan.calOff", []));
    const list = (cals.items || []).map((c) => ({ id: c.id, name: c.summaryOverride || c.summary,
      color: c.backgroundColor || "#4aa3ff", primary: !!c.primary }));
    store.set("plan.cals", list);
    const timeMin = dateOf(addDays(todayKey(), -DAYS_BEFORE)).toISOString();
    const timeMax = dateOf(addDays(todayKey(), DAYS_AFTER + 1)).toISOString();
    const items = [];
    await Promise.all(list.filter((c) => !off.has(c.id)).map(async (c) => {
      const data = await gget(`/calendars/${encodeURIComponent(c.id)}/events`, {
        timeMin, timeMax, singleEvents: "true", orderBy: "startTime", maxResults: "250",
      });
      for (const ev of data.items || []) {
        if (ev.status === "cancelled") continue;
        const allDay = !!(ev.start && ev.start.date);
        items.push({
          id: ev.id, cal: c.id, color: c.color, title: ev.summary || "(sans titre)",
          loc: ev.location || "", link: ev.htmlLink || "", allDay,
          start: allDay ? ev.start.date : ev.start.dateTime,
          end: allDay ? ev.end.date : ev.end.dateTime,
        });
      }
    }));
    events = { items, at: Date.now() };
    store.set("plan.events", events);
    gcalState = "ok";
  } catch (e) {
    if (e.expired) {
      gcalState = "expired";
      // Jeton Google expiré (1 h) : on repasse chez Google sans rien demander, au plus une
      // fois toutes les 10 minutes. Si Google veut une confirmation, un bouton s'affiche.
      if (!store.get("plan.gNeedsClick") && Date.now() - store.get("plan.gTry", 0) > 10 * 60_000
          && !document.hidden && !store.get("plan.dirty")) {
        return googleLogin(true);
      }
    } else {
      gcalState = "error";
      setStatusLine(navigator.onLine === false ? "Hors ligne · agenda en cache" : `Agenda : ${e.message}`);
    }
  }
  renderBanner();
  render();
}

/** Rendez-vous d'un jour (y compris ceux sur plusieurs jours). */
function eventsOn(key) {
  const start = dateOf(key);
  const end = dateOf(addDays(key, 1));
  return (events.items || []).filter((ev) => {
    const s = ev.allDay ? dateOf(ev.start) : new Date(ev.start);
    const e = ev.allDay ? dateOf(ev.end) : new Date(ev.end);
    return s < end && e > start;
  }).sort((a, b) => (b.allDay - a.allDay) || new Date(a.start) - new Date(b.start));
}

function renderBanner() {
  const b = $("gcal-banner");
  if (gcalState === "off" && !CFG.googleClientId) {
    $("gcal-text").textContent = "Google Agenda n'est pas encore relié : il manque l'identifiant Google dans config.js (voir le guide).";
    $("gcal-btn").hidden = true;
    b.hidden = !!store.get("plan.hideSetup");
    return;
  }
  $("gcal-btn").hidden = false;
  if (gcalState === "off") {
    $("gcal-text").textContent = "Affiche tes rendez-vous Google ici (lecture seule).";
    b.hidden = !!store.get("plan.gDismissed");
  } else if (gcalState === "expired") {
    $("gcal-text").textContent = "Google demande de confirmer la connexion pour actualiser l'agenda.";
    b.hidden = false;
  } else {
    b.hidden = true;
  }
}

// --- Affichage ------------------------------------------------------------------------

let day = todayKey();
let pendingFilter = store.get("plan.filter", "");

function el(tag, cls, text) {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text != null) n.textContent = text;
  return n;
}

function show(id) {
  for (const s of document.querySelectorAll(".screen")) s.hidden = s.id !== id;
}

function showLogin(message) {
  show("login");
  const err = $("login-error");
  err.hidden = !message;
  err.textContent = message || "";
}

let statusOverride = "";
function setStatusLine(text) {
  if (text !== undefined) statusOverride = text;
  if (statusOverride) { $("status").textContent = statusOverride; return; }
  const parts = [];
  if (store.get("plan.dirty")) parts.push("envoi des tâches…");
  if (catalogue && catalogue.generated) parts.push(`catalogue du Mac ${ago(catalogue.generated)}`);
  if (events.at && gcalState !== "ok" && store.get("plan.gcal")) parts.push(`agenda ${ago(events.at / 1000)}`);
  $("status").textContent = parts.join(" · ");
}

function ago(ts) {
  const d = (Date.now() / 1000 - ts) / 60;
  if (d < 2) return "à l'instant";
  if (d < 60) return `il y a ${Math.round(d)} min`;
  if (d < 48 * 60) return `il y a ${Math.round(d / 60)} h`;
  return `il y a ${Math.round(d / 1440)} j`;
}

function daysLeft(iso) {
  const n = Math.round((dateOf(iso) - dateOf(todayKey())) / 86_400_000);
  return isNaN(n) ? null : n;
}

function prodBadges(info) {
  if (!info) return [];
  const out = [];
  if (info.st) {
    const b = el("span", "pill-st", info.st);
    b.style.setProperty("--c", STATUS_COLORS[info.st] || "var(--dim)");
    out.push(b);
  }
  if (info.deadline && info.st !== "Done") {
    const n = daysLeft(info.deadline);
    if (n !== null) {
      const text = n < 0 ? `en retard de ${-n} j` : n === 0 ? "aujourd'hui" : n === 1 ? "demain" : `dans ${n} j`;
      out.push(el("span", `pill-dl${n <= 2 ? " urgent" : ""}`, `⏰ ${fmtDay(info.deadline, { day: "numeric", month: "short" })} · ${text}`));
    }
  }
  const paid = info.paid === true ? "yes" : info.paid;
  if (paid === "no") out.push(el("span", "pill-pay", "💶 Non payé"));
  if (info.client) out.push(el("span", "pill-client", `👤 ${info.client}`));
  if (info.style) out.push(el("span", "pill-style", `🎵 ${info.style}`));
  return out;
}

function renderDays() {
  const box = $("days");
  box.textContent = "";
  const today = todayKey();
  const tasks = liveTasks();
  const songs = pendingSongs();
  for (let i = -DAYS_BEFORE; i <= DAYS_AFTER; i++) {
    const k = addDays(today, i);
    const b = el("button", "day" + (k === day ? " on" : "") + (k === today ? " today" : "") + (i < 0 ? " past" : ""));
    b.append(el("span", "dw", fmtDay(k, { weekday: "short" }).replace(".", "")));
    b.append(el("span", "dn", String(dateOf(k).getDate())));
    const dots = el("span", "dots");
    if (eventsOn(k).some((e) => !e.allDay || !/holiday/.test(e.cal))) dots.append(el("i", "d-ev"));
    if (tasks.some((t) => t.d === k && !t.done)) dots.append(el("i", "d-task"));
    if (songs.some((s) => s.info.deadline === k)) dots.append(el("i", "d-dl"));
    b.append(dots);
    b.addEventListener("click", () => goDay(k));
    box.append(b);
  }
  const on = box.querySelector(".on");
  if (on) on.scrollIntoView({ inline: "center", block: "nearest" });
}

function renderEvents() {
  const ul = $("events");
  ul.textContent = "";
  const list = eventsOn(day);
  $("agenda-block").hidden = !store.get("plan.gcal") && !list.length;
  if (!list.length) {
    ul.append(el("li", "empty", store.get("plan.gcal") ? "Aucun rendez-vous." : "Agenda non connecté."));
    return;
  }
  for (const ev of list) {
    const li = el("li", "row ev");
    li.style.setProperty("--c", ev.color);
    const time = el("div", "ev-time");
    if (ev.allDay) time.textContent = "Journée";
    else {
      const s = new Date(ev.start);
      const e = new Date(ev.end);
      time.append(el("span", "", keyOf(s) === day ? hm(s) : "…"));
      time.append(el("span", "dim", keyOf(e) === day ? hm(e) : "…"));
    }
    const main = el("div", "row-main");
    main.append(el("div", "row-title", ev.title));
    if (ev.loc) main.append(el("div", "row-sub", `📍 ${ev.loc}`));
    li.append(time, main);
    if (!ev.allDay && day === todayKey() && new Date(ev.end) < new Date()) li.classList.add("over");
    if (ev.link) li.addEventListener("click", () => window.open(ev.link, "_blank"));
    ul.append(li);
  }
}

function taskRow(t, carried) {
  const li = el("li", "row task" + (t.done ? " done" : ""));
  const check = el("button", "check", t.done ? "✓" : "");
  check.setAttribute("aria-label", t.done ? "Marquer à faire" : "Marquer fait");
  check.addEventListener("click", (e) => {
    e.stopPropagation();
    updateTask(t.id, { done: !t.done });
  });
  const main = el("div", "row-main");
  const title = el("div", "row-title wrap");
  if (t.h) title.append(el("span", "task-time", t.h));
  title.append(document.createTextNode((t.song ? "🎛 " : "") + t.t));
  main.append(title);
  if (carried) main.append(el("div", "row-sub late", `Prévu ${shortDay(t.d)} · pas encore fait`));
  if (t.song) {
    const prod = el("div", "prod");
    prod.append(...prodBadges(songInfo(t.song)).filter((b) => !b.classList.contains("pill-style")));
    if (prod.childNodes.length) main.append(prod);
  }
  li.append(check, main);
  li.addEventListener("click", () => openTask(t));
  return li;
}

function renderTasks() {
  const ul = $("tasks");
  ul.textContent = "";
  const today = todayKey();
  const all = liveTasks();
  const mine = all.filter((t) => t.d === day);
  const carried = day === today ? all.filter((t) => !t.done && t.d < today) : [];
  const order = (a, b) => (a.done - b.done) || (a.h || "99").localeCompare(b.h || "99") || a.c - b.c;
  carried.sort((a, b) => a.d.localeCompare(b.d) || order(a, b));
  mine.sort(order);
  for (const t of carried) ul.append(taskRow(t, true));
  for (const t of mine) ul.append(taskRow(t, false));
  if (!carried.length && !mine.length) ul.append(el("li", "empty", day < today ? "Rien de noté ce jour-là." : "Rien pour l'instant."));
}

function songRow(s, planned) {
  const li = el("li", "row song");
  const main = el("div", "row-main");
  main.append(el("div", "row-title", s.name));
  const prod = el("div", "prod");
  prod.append(...prodBadges(s.info));
  if (planned) prod.append(el("span", "pill-plan", `📅 ${relDay(planned.d) || shortDay(planned.d)}`));
  main.append(prod);
  li.append(main);
  li.append(el("span", "chev", "›"));
  li.addEventListener("click", () => openSong(s, planned));
  return li;
}

function plannedFor(name) {
  return liveTasks().filter((t) => t.song === name && !t.done && t.d >= todayKey())
    .sort((a, b) => a.d.localeCompare(b.d))[0];
}

function renderSongs() {
  const today = todayKey();
  const songs = pendingSongs();
  // Deadlines du jour (et, aujourd'hui, celles déjà dépassées)
  const due = songs.filter((s) => s.info.deadline && (s.info.deadline === day || (day === today && s.info.deadline < today)));
  $("deadlines-block").hidden = !due.length;
  $("deadlines-title").textContent = day === today ? "Deadlines (aujourd'hui et en retard)" : "Deadlines";
  const ul = $("deadlines");
  ul.textContent = "";
  for (const s of due) ul.append(songRow(s, plannedFor(s.name)));

  // Tous les morceaux en attente : affichés sur la journée d'aujourd'hui
  const block = $("pending-block");
  block.hidden = day !== today || !catalogue;
  if (block.hidden) return;
  const rest = songs.filter((s) => !due.includes(s));
  const counts = {};
  for (const s of rest) counts[s.info.st || "Deadline"] = (counts[s.info.st || "Deadline"] || 0) + 1;
  const chips = $("pending-filters");
  chips.textContent = "";
  const opts = [["", `Tous ${rest.length}`], ...["Mix", "Active", "Idée", "Deadline"]
    .filter((k) => counts[k]).map((k) => [k, `${k} ${counts[k]}`])];
  if (pendingFilter && !counts[pendingFilter]) pendingFilter = "";
  for (const [k, label] of opts) {
    const b = el("button", k === pendingFilter ? "on" : "", label);
    b.addEventListener("click", () => { pendingFilter = k; store.set("plan.filter", k); renderSongs(); });
    chips.append(b);
  }
  chips.hidden = opts.length <= 2;
  $("pending-title").textContent = `Morceaux en attente (${rest.length})`;
  const list = $("pending");
  list.textContent = "";
  const shown = rest.filter((s) => !pendingFilter || (s.info.st || "Deadline") === pendingFilter);
  for (const s of shown) list.append(songRow(s, plannedFor(s.name)));
  if (!shown.length) {
    list.append(el("li", "empty", (catalogue.songs || []).some((s) => s.i && s.i.st)
      ? "Tout est à jour 🎉"
      : "Aucun statut renseigné. Sur le Mac, donne un statut (Idée, Active, Mix) à tes morceaux en cours pour les voir ici."));
  }
}

function render() {
  const today = todayKey();
  const rel = relDay(day);
  $("day-title").textContent = rel || fmtDay(day, { weekday: "long", day: "numeric", month: "long" }).replace(/^./, (c) => c.toUpperCase());
  $("day-sub").textContent = rel ? fmtDay(day, { weekday: "long", day: "numeric", month: "long" }) : "";
  $("today-btn").hidden = day === today;
  renderDays();
  renderEvents();
  renderTasks();
  renderSongs();
  renderDigestCard();
  setStatusLine();
}

function goDay(k) {
  const min = addDays(todayKey(), -DAYS_BEFORE);
  const max = addDays(todayKey(), DAYS_AFTER);
  day = k < min ? min : k > max ? max : k;
  render();
  window.scrollTo({ top: 0 });
}

// --- Feuilles d'actions -------------------------------------------------------------

function openSheet(title, sub, build) {
  $("sheet-title").textContent = title;
  $("sheet-sub").textContent = sub || "";
  $("sheet-sub").hidden = !sub;
  const body = $("sheet-body");
  body.textContent = "";
  build(body);
  $("sheet").hidden = false;
}
const closeSheet = () => { $("sheet").hidden = true; };

function button(text, cls, onClick) {
  const b = el("button", cls, text);
  b.type = "button";
  b.addEventListener("click", onClick);
  return b;
}

function dayChoices(body, current, onPick) {
  const today = todayKey();
  const row = el("div", "choice-row");
  row.append(
    button("Aujourd'hui", current === today ? "ghost on" : "ghost", () => onPick(today)),
    button("Demain", current === addDays(today, 1) ? "ghost on" : "ghost", () => onPick(addDays(today, 1))),
  );
  body.append(row);
  const pick = el("label", "date-pick");
  pick.append(el("span", "", "Autre jour"));
  const input = el("input");
  input.type = "date";
  input.value = current || "";
  input.addEventListener("change", () => input.value && onPick(input.value));
  pick.append(input);
  body.append(pick);
}

function openTask(t) {
  openSheet(t.song ? "Morceau à travailler" : "Tâche", t.d < todayKey() && !t.done ? `Prévue ${shortDay(t.d)}` : "", (body) => {
    const text = el("input", "field");
    text.type = "text";
    text.value = t.t;
    const time = el("input", "field time");
    time.type = "time";
    time.value = t.h || "";
    const line = el("div", "edit-line");
    line.append(text, time);
    body.append(line);
    const save = () => {
      const v = text.value.trim();
      if (v && (v !== t.t || time.value !== (t.h || ""))) updateTask(t.id, { t: v, h: time.value });
    };
    body.append(button(t.done ? "Remettre à faire" : "✓ Fait", "primary big", () => {
      save(); updateTask(t.id, { done: !t.done }); closeSheet();
    }));
    body.append(el("p", "dim small sheet-label", "Déplacer à"));
    dayChoices(body, t.d, (k) => {
      save(); updateTask(t.id, { d: k, done: false }); closeSheet();
      toast(`Déplacée à ${relDay(k).toLowerCase() || shortDay(k)}`);
    });
    body.append(button("Enregistrer", "ghost big", () => { save(); closeSheet(); }));
    body.append(button("Supprimer", "ghost danger", () => {
      updateTask(t.id, { del: true }); closeSheet(); toast("Tâche supprimée");
    }));
  });
}

function openSong(s, planned) {
  const sub = [s.info.client && `👤 ${s.info.client}`, s.where !== s.name && s.where].filter(Boolean).join(" · ");
  openSheet(s.name, sub, (body) => {
    const prod = el("div", "prod big");
    prod.append(...prodBadges(s.info));
    if (prod.childNodes.length) body.append(prod);
    body.append(el("p", "dim small sheet-label",
      planned ? `Prévu ${relDay(planned.d).toLowerCase() || shortDay(planned.d)}. Déplacer à :` : "Travailler dessus :"));
    dayChoices(body, planned && planned.d, (k) => {
      if (planned) updateTask(planned.id, { d: k });
      else addTask(s.name, k, "", s.name);
      closeSheet();
      toast(`Ajouté à ${relDay(k).toLowerCase() || shortDay(k)}`);
    });
  });
}

let toastTimer;
function toast(text) {
  const t = $("toast");
  t.textContent = text;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, 2500);
}

// --- Mails : résumé du matin (écrit par Claude) + non lus des dernières 24 h ------------
// Le résumé arrive chaque matin dans Gmail, envoyé par une tâche planifiée de Claude, avec
// l'objet « Résumé mails SIRØB · <date> ». L'app le retrouve et l'affiche ici.

const hasGmail = () => ((store.get("plan.gtok") || {}).scope || "").includes("gmail.readonly");

function b64utf8(data) {
  const bin = atob(data.replace(/-/g, "+").replace(/_/g, "/"));
  return new TextDecoder().decode(Uint8Array.from(bin, (c) => c.charCodeAt(0)));
}

function textOf(payload) {
  const parts = [payload];
  let html = "";
  while (parts.length) {
    const p = parts.shift();
    if (p.parts) parts.push(...p.parts);
    if (p.mimeType === "text/plain" && p.body && p.body.data) return b64utf8(p.body.data);
    if (p.mimeType === "text/html" && p.body && p.body.data && !html) html = b64utf8(p.body.data);
  }
  if (!html) return "";
  const doc = new DOMParser().parseFromString(html.replace(/<br\s*\/?>/gi, "\n").replace(/<\/(p|div|li|h\d)>/gi, "\n"), "text/html");
  return doc.body.textContent || "";
}

const header = (m, name) => ((m.payload && m.payload.headers) || [])
  .find((h) => h.name.toLowerCase() === name.toLowerCase())?.value || "";
const senderName = (from) => (from.match(/^\s*"?([^"<]+?)"?\s*</) || [, from])[1].trim();
const gmailLink = (threadId) => `https://mail.google.com/mail/u/0/#inbox/${threadId}`;

/** Dernier résumé du matin (au plus 3 jours). */
async function fetchDigest() {
  const list = await gget("/messages", { q: `subject:"${DIGEST_SUBJECT}" newer_than:3d`, maxResults: "1" }, GMAIL);
  const id = list.messages && list.messages[0] && list.messages[0].id;
  if (!id) return null;
  const m = await gget(`/messages/${id}`, { format: "full" }, GMAIL);
  return { id, thread: m.threadId, at: Number(m.internalDate), subject: header(m, "Subject"), text: textOf(m.payload).trim() };
}

async function checkDigest() {
  if (!hasGmail() || !googleToken()) return;
  try {
    const d = await fetchDigest();
    store.set("plan.digest", d);
    renderDigestCard();
  } catch { /* le bouton ✉︎ affichera l'erreur */ }
}

function renderDigestCard() {
  const d = store.get("plan.digest");
  const fresh = d && keyOf(new Date(d.at)) === todayKey() && store.get("plan.digestSeen") !== d.id;
  $("digest-card").hidden = !(fresh && day === todayKey());
}

function renderDigest(d) {
  const box = $("digest");
  box.textContent = "";
  if (!d) {
    box.append(el("p", "dim small", "Pas encore de résumé ces derniers jours. Il arrive chaque matin vers 7 h."));
    return;
  }
  const when = new Date(d.at);
  box.append(el("p", "digest-when dim small",
    `${relDay(keyOf(when)) || shortDay(keyOf(when))} à ${hm(when)}`));
  for (const raw of d.text.split(/\n/)) {
    const line = raw.trimEnd();
    if (!line.trim()) continue;
    if (/^[-–—_=]{3,}$/.test(line.trim())) continue;
    const t = line.trim();
    if (/^#+\s/.test(t) || (/^[^a-zà-ÿ]*[A-ZÀ-Ÿ].*:$/.test(t) && t.length < 60)) {
      box.append(el("h4", "digest-h", t.replace(/^#+\s*/, "").replace(/:$/, "")));
    } else if (/^[-•*·]\s/.test(t)) {
      box.append(el("p", "digest-li", t.replace(/^[-•*·]\s*/, "")));
    } else {
      box.append(el("p", "digest-p", t));
    }
  }
  const open = el("a", "digest-open small", "Ouvrir dans Gmail ›");
  open.href = gmailLink(d.thread);
  open.target = "_blank";
  box.append(open);
}

function mailRow(m) {
  const li = el("li", "row mail");
  const main = el("div", "row-main");
  const top = el("div", "mail-top");
  top.append(el("span", "row-title", senderName(header(m, "From"))));
  top.append(el("span", "mail-time dim small", hm(new Date(Number(m.internalDate)))));
  main.append(top);
  main.append(el("div", "mail-subj", header(m, "Subject") || "(sans objet)"));
  if (m.snippet) main.append(el("div", "row-sub", m.snippet.replace(/&#39;/g, "'").replace(/&quot;/g, '"').replace(/&amp;/g, "&")));
  li.append(main);
  li.addEventListener("click", () => window.open(gmailLink(m.threadId), "_blank"));
  return li;
}

async function loadMails(force = false) {
  const ul = $("unread");
  const info = $("mails-info");
  if (!CFG.googleClientId || !store.get("plan.gcal") || !hasGmail()) {
    $("mails-auth").hidden = false;
    $("mails-content").hidden = true;
    return;
  }
  $("mails-auth").hidden = true;
  $("mails-content").hidden = false;
  if (!googleToken()) {
    if (force || Date.now() - store.get("plan.gTry", 0) > 60_000) return googleLogin(true);
    info.textContent = "Connexion Google expirée : touche ↻.";
    return;
  }
  info.textContent = "Chargement…";
  try {
    const [d, list] = await Promise.all([
      fetchDigest(),
      gget("/messages", { q: `is:unread in:inbox newer_than:1d -subject:"${DIGEST_SUBJECT}"`, maxResults: "40" }, GMAIL),
    ]);
    store.set("plan.digest", d);
    if (d) store.set("plan.digestSeen", d.id);
    renderDigest(d);
    renderDigestCard();
    const full = (await Promise.all((list.messages || []).map((x) => gget(`/messages/${x.id}`, [
      ["format", "metadata"], ["metadataHeaders", "From"], ["metadataHeaders", "Subject"],
    ], GMAIL).catch(() => null)))).filter(Boolean);
    const primary = full.filter((m) => (m.labelIds || []).includes("CATEGORY_PERSONAL") || !(m.labelIds || []).some((l) => l.startsWith("CATEGORY_")));
    const others = full.length - primary.length;
    ul.textContent = "";
    for (const m of primary) ul.append(mailRow(m));
    if (!primary.length) ul.append(el("li", "empty", "Aucun mail non lu important depuis hier 🎉"));
    info.textContent = others ? `+ ${others} non lu${others > 1 ? "s" : ""} dans Promotions, Réseaux sociaux ou Notifications` : "";
  } catch (e) {
    if (e.expired) return googleLogin(true);
    if (e.status === 403 && /has not been used|disabled/i.test(e.message)) {
      info.textContent = "L'API Gmail n'est pas encore activée dans ton projet Google Cloud (voir le guide, étape 4).";
      return;
    }
    if (e.status === 403) {
      $("mails-auth").hidden = false;
      $("mails-content").hidden = true;
      return;
    }
    info.textContent = `Mails : ${e.message}`;
  }
}

function openMails() {
  const d = store.get("plan.digest");
  if (d) renderDigest(d);
  show("mails");
  history.pushState({ view: "mails" }, "");
  loadMails();
}

// --- Réglages -----------------------------------------------------------------------

function openSettings() {
  $("cat-path").value = store.get("catPath", CFG.catalogue);
  $("cat-info").textContent = catalogue
    ? `${pendingSongs().length} morceaux en attente sur ${(catalogue.songs || []).length}. Catalogue écrit par le Mac ${ago(catalogue.generated)}.`
    : "Pas encore de catalogue.";
  $("account").textContent = store.get("account") ? `OneDrive : ${store.get("account")}` : "Connecté à OneDrive";
  $("version").textContent = `SIRØB Planning ${VERSION}`;
  const connected = !!store.get("plan.gcal");
  $("gcal-info").textContent = !CFG.googleClientId
    ? "Pas encore configuré : il manque l'identifiant Google dans config.js."
    : connected ? `Connecté${store.get("plan.gemail") ? " : " + store.get("plan.gemail") : ""}. Agendas affichés :`
      : "Non connecté.";
  $("gcal-toggle").hidden = !CFG.googleClientId;
  $("gcal-toggle").textContent = connected ? "Déconnecter Google Agenda" : "Connecter Google Agenda";
  const list = $("cal-list");
  list.textContent = "";
  if (connected) {
    const off = new Set(store.get("plan.calOff", []));
    for (const c of store.get("plan.cals", [])) {
      const lab = el("label", "cal");
      const cb = el("input");
      cb.type = "checkbox";
      cb.checked = !off.has(c.id);
      cb.addEventListener("change", () => {
        const s = new Set(store.get("plan.calOff", []));
        cb.checked ? s.delete(c.id) : s.add(c.id);
        store.set("plan.calOff", [...s]);
        loadEvents();
      });
      const dot = el("span", "dot");
      dot.style.background = c.color;
      lab.append(cb, dot, el("span", "", c.name));
      list.append(lab);
    }
  }
  show("settings");
  history.pushState({ view: "settings" }, "");
}

// --- Démarrage ----------------------------------------------------------------------

function wire() {
  $("login-btn").addEventListener("click", login);
  $("add").addEventListener("submit", (e) => {
    e.preventDefault();
    const v = $("add-text").value.trim();
    if (!v) return;
    addTask(v, day, $("add-time").value);
    $("add-text").value = "";
    $("add-time").value = "";
  });
  $("today-btn").addEventListener("click", () => goDay(todayKey()));
  $("refresh-btn").addEventListener("click", () => refreshAll(true));
  $("settings-btn").addEventListener("click", openSettings);
  $("settings-back").addEventListener("click", () => history.back());
  window.addEventListener("popstate", () => { show("main"); render(); });
  $("cat-path").addEventListener("change", (e) => {
    const v = e.target.value.trim().replace(/^\/+/, "");
    store.set("catPath", v || CFG.catalogue);
    store.del("catalogue");
    catalogue = null;
    loadCatalogue();
  });
  $("logout").addEventListener("click", logout);
  $("gcal-btn").addEventListener("click", () => googleLogin(false));
  $("mail-btn").addEventListener("click", openMails);
  $("digest-card").addEventListener("click", openMails);
  $("mails-back").addEventListener("click", () => history.back());
  $("mails-refresh").addEventListener("click", () => loadMails(true));
  $("mails-auth-btn").addEventListener("click", () => googleLogin(false));
  $("gcal-toggle").addEventListener("click", () => {
    if (store.get("plan.gcal")) {
      for (const k of ["plan.gcal", "plan.gtok", "plan.events", "plan.cals"]) store.del(k);
      events = { items: [], at: 0 };
      store.set("plan.gDismissed", true);
      history.back();
      loadEvents();
    } else googleLogin(false);
  });
  $("sheet-close").addEventListener("click", closeSheet);
  $("sheet").addEventListener("click", (e) => { if (e.target.id === "sheet") closeSheet(); });

  // Glisser à gauche / à droite pour changer de jour
  let sx = 0, sy = 0;
  const main = $("main");
  main.addEventListener("touchstart", (e) => { sx = e.touches[0].clientX; sy = e.touches[0].clientY; }, { passive: true });
  main.addEventListener("touchend", (e) => {
    if (e.target.closest("#days, .chips, input")) return;
    const dx = e.changedTouches[0].clientX - sx;
    const dy = e.changedTouches[0].clientY - sy;
    if (Math.abs(dx) > 70 && Math.abs(dy) < 45) goDay(addDays(day, dx < 0 ? 1 : -1));
  });

  // Retour dans l'app : on actualise (et on passe au jour suivant après minuit).
  let lastToday = todayKey();
  document.addEventListener("visibilitychange", () => {
    if (document.hidden || !store.get("tok")) return;
    if (todayKey() !== lastToday) { if (day === lastToday) day = todayKey(); lastToday = todayKey(); }
    if (Date.now() - store.get("plan.checkedAt", 0) > 2 * 60_000) refreshAll();
    else render();
  });
  window.addEventListener("online", () => syncPlan());
}

async function refreshAll(manual = false) {
  store.set("plan.checkedAt", Date.now());
  statusOverride = "";
  if (manual) {
    store.del("plan.gNeedsClick");
    store.set("plan.gTry", 0);
    $("status").textContent = "Actualisation…";
  }
  await Promise.all([syncPlan(), loadCatalogue(), loadEvents().then(checkDigest)]);
  render();
}

async function start() {
  wire();
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
  try {
    await handleRedirect();
  } catch (e) {
    return showLogin(e.message);
  }
  if (!store.get("tok")) return showLogin();
  show("main");
  render();
  rememberAccount();
  refreshAll();
}

start();
