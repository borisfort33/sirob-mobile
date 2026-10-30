/* SIRØB Mobile : chercher un morceau dans le catalogue et envoyer son bounce OneDrive.
 *
 * - Connexion Microsoft (OAuth « code + PKCE », sans serveur), jetons gardés sur le téléphone.
 * - Catalogue : JSON écrit par l'app du Mac dans OneDrive (mobile_index.py), mis en cache ici.
 * - Liens : Microsoft Graph createLink, « toute personne disposant du lien », lecture seule.
 */
"use strict";

const CFG = window.SIROB_CONFIG;
const AUTH = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPE = "Files.ReadWrite offline_access User.Read";
const REDIRECT = new URL("./", location.href).href;
const VERSION = "1.1";
const AUDIO_EXT = /\.(wav|mp3|aif|aiff|flac|m4a)$/i;

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

// --- Connexion Microsoft ------------------------------------------------------------

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

/** Retour de la page Microsoft : ?code=… dans l'adresse. */
async function handleRedirect() {
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
      // Microsoft limite les jetons des applis web à 24 h : on repasse par la page de connexion.
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
  for (const k of ["tok", "links", "account", "catalogue"]) store.del(k);
  catalogue = null;
  entries = [];
  history.back();
  showLogin();
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
      },
    });
  } catch (e) {
    if (e instanceof NeedLogin) throw e;
    throw new Error("Pas de connexion internet");
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

/** Adresse Graph d'un fichier : par id (recherche OneDrive) ou par chemin (catalogue). */
const itemUrl = (file) => file.id ? `/me/drive/items/${file.id}` : `/me/drive/root:/${encPath(file.p)}:`;

async function createLink(file) {
  const links = store.get("links", {});
  const key = file.id || file.p;
  if (links[key]) return links[key];
  const data = await graph(`${itemUrl(file)}/createLink`, {
    method: "POST", body: JSON.stringify({ type: "view", scope: "anonymous" }),
  });
  links[key] = data.link.webUrl;
  store.set("links", links);
  return links[key];
}

// --- Catalogue ----------------------------------------------------------------------

let catalogue = store.get("catalogue");
let entries = [];      // morceaux (+ exports isolés) prêts pour la recherche

const norm = (s) => (s || "").replace(/[øØ]/g, "o").normalize("NFD")
  .replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[_\-.()[\],]+/g, " ");
const baseName = (p) => p.split("/").pop();
const stripExt = (n) => n.replace(/\.[^.]+$/, "");
const parentName = (p) => { const parts = p.split("/"); return parts[parts.length - 2] || ""; };
const isMaster = (f) => /master/i.test(f[0]) && !/pre[\s_-]?master/i.test(baseName(f[0]));

function fileObj(f) {
  return { p: f[0], m: f[1], dir: !!f[2], name: baseName(f[0]) };
}

function buildEntries() {
  entries = [];
  if (!catalogue) return;
  for (const s of catalogue.songs || []) {
    const all = [...s.b, ...(s.v || []).flatMap((v) => v.b)];
    const last = Math.max(s.m || 0, ...all.map((f) => f[1]));
    const hay = norm([s.n, s.d, ...all.map((f) => stripExt(baseName(f[0]))), ...(s.v || []).map((v) => v.n)].join(" "));
    entries.push({ song: s, title: norm(s.n), hay, last, count: all.length });
  }
  for (const f of catalogue.loose || []) {
    const name = stripExt(baseName(f[0]));
    entries.push({ loose: fileObj(f), title: norm(name), hay: norm(name), last: f[1], count: 1 });
  }
}

async function loadCatalogue(force = false) {
  const path = store.get("catPath", CFG.catalogue);
  setStatus(catalogue ? statusText() + " · mise à jour…" : "Chargement du catalogue…");
  try {
    const item = await graph(`/me/drive/root:/${encPath(path)}`);
    if (!force && catalogue && catalogue.eTag === item.eTag) { setStatus(statusText()); return; }
    const resp = await fetch(item["@microsoft.graph.downloadUrl"]);
    const data = await resp.json();
    data.eTag = item.eTag;
    catalogue = data;
    store.set("catalogue", data);
    buildEntries();
    renderFilters();
    render();
  } catch (e) {
    if (e instanceof NeedLogin) return showLogin();
    setStatus(e.status === 404
      ? "Catalogue introuvable : ouvre l'app du Mac une fois (elle l'écrit dans OneDrive), ou vérifie l'emplacement dans ⚙︎."
      : `${e.message}${catalogue ? " · catalogue en cache affiché" : ""}`);
    return;
  }
  setStatus(statusText());
  store.set("checkedAt", Date.now());
}

function ago(ts) {
  const d = (Date.now() / 1000 - ts) / 60;
  if (d < 2) return "à l'instant";
  if (d < 60) return `il y a ${Math.round(d)} min`;
  if (d < 48 * 60) return `il y a ${Math.round(d / 60)} h`;
  return `il y a ${Math.round(d / 1440)} j`;
}
const fmtDate = (ts) => new Date(ts * 1000).toLocaleDateString("fr-FR", { day: "numeric", month: "short", year: "2-digit" });

function statusText() {
  if (!catalogue) return "";
  const n = (catalogue.songs || []).length;
  return `${n} morceaux · catalogue du Mac ${ago(catalogue.generated)}`;
}
function setStatus(text) { $("status").textContent = text; }

// --- Recherche ----------------------------------------------------------------------

// --- Menus : filtres du Mac (SIRØB, M6, Ghost…), années, tags ---------------------
// Un choix par groupe ; les groupes se combinent (ex. M6 + 2025).

let filter = store.get("filter", {});   // {q: "M6", y: 2025, t: "Rouge"}

function passes(e) {
  if (!filter.q && !filter.y && !filter.t) return true;
  if (!e.song) return false;            // exports isolés : seulement sans filtre
  const s = e.song;
  return (!filter.q || (s.q || []).includes(filter.q))
    && (!filter.y || s.y === filter.y)
    && (!filter.t || (s.t || []).includes(filter.t));
}

function renderFilters() {
  const box = $("filters");
  box.replaceChildren();
  if (!catalogue) { box.hidden = true; return; }
  const songs = catalogue.songs || [];
  const count = (pred) => songs.filter(pred).length;
  const groups = [
    ["q", (catalogue.filters || []).filter((n) => count((s) => (s.q || []).includes(n)))],
    ["y", [...new Set(songs.map((s) => s.y).filter(Boolean))].sort((a, b) => b - a)],
    ["t", Object.keys(catalogue.tags || {}).filter((n) => count((s) => (s.t || []).includes(n)))],
  ];
  // Un filtre qui n'existe plus dans le catalogue est oublié.
  for (const [key, values] of groups) if (filter[key] && !values.includes(filter[key])) delete filter[key];
  const all = el("button", !filter.q && !filter.y && !filter.t ? "on" : "", "Tout");
  all.addEventListener("click", () => setFilter({}));
  box.append(all);
  for (const [key, values] of groups) {
    if (!values.length) continue;
    box.append(el("span", "sep"));
    for (const v of values) {
      const b = el("button", filter[key] === v ? "on" : "");
      const color = key === "t" && catalogue.tags[v];
      if (color) {
        const dot = el("span", "dot");
        dot.style.background = color;
        b.append(dot);
      }
      b.append(String(v));
      b.addEventListener("click", () => setFilter({ ...filter, [key]: filter[key] === v ? undefined : v }));
      box.append(b);
    }
  }
  box.hidden = false;
}

function setFilter(f) {
  filter = Object.fromEntries(Object.entries(f).filter(([, v]) => v !== undefined));
  store.set("filter", filter);
  renderFilters();
  render();
  window.scrollTo(0, 0);
}

function search(q) {
  const tokens = norm(q).split(/\s+/).filter(Boolean);
  const filtered = filter.q || filter.y || filter.t;
  if (!tokens.length) {
    return entries.filter((e) => passes(e) && (filtered || e.count))
      .sort((a, b) => b.last - a.last).slice(0, filtered ? 400 : 40);
  }
  const nq = tokens.join(" ");
  const out = [];
  for (const e of entries) {
    if (!passes(e)) continue;
    if (!tokens.every((t) => e.hay.includes(t))) continue;
    const rank = e.title.startsWith(nq) ? 0 : e.title.includes(nq) ? 1
      : tokens.every((t) => e.title.includes(t)) ? 2 : 3;
    out.push([rank, e]);
  }
  out.sort((a, b) => a[0] - b[0] || (b[1].count > 0) - (a[1].count > 0) || b[1].last - a[1].last);
  return out.slice(0, 80).map((x) => x[1]);
}

function highlight(text, q) {
  const span = document.createElement("span");
  const t = norm(q).trim();
  const i = t ? norm(text).indexOf(t) : -1;
  // norm() garde la longueur pour les lettres courantes, sinon on n'éclaire rien.
  if (i < 0 || norm(text).length !== text.length) { span.textContent = text; return span; }
  span.append(text.slice(0, i));
  const mark = document.createElement("mark");
  mark.textContent = text.slice(i, i + t.length);
  span.append(mark, text.slice(i + t.length));
  return span;
}

/** Le fichier à envoyer d'un coup : le master le plus récent, sinon le bounce le plus récent. */
function bestFile(song) {
  const files = [...song.b].sort((a, b) => b[1] - a[1]);
  const pick = files.find(isMaster) || files[0];
  return pick ? fileObj(pick) : null;
}

function render() {
  const q = $("q").value;
  $("clear").hidden = !q;
  const list = $("results");
  list.replaceChildren();
  $("live").replaceChildren();
  if (!catalogue) { $("more").hidden = !q.trim(); return; }
  const found = search(q);
  for (const e of found) list.append(e.song ? songRow(e, q) : looseRow(e.loose, q));
  if (q.trim() && !found.length) {
    list.append(Object.assign(document.createElement("li"), { className: "empty", textContent: "Rien dans le catalogue." }));
  }
  $("more").hidden = !q.trim();
  $("live-btn").textContent = `Chercher « ${q.trim()} » dans tout OneDrive`;
  const active = [filter.q, filter.y, filter.t].filter(Boolean).join(" · ");
  if (active) setStatus(`${found.length} morceau${found.length > 1 ? "x" : ""} · ${active}`);
  else if (!q.trim()) setStatus(statusText() + " · derniers bounces");
  else setStatus(statusText());
}

function el(tag, cls, text) {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text != null) n.textContent = text;
  return n;
}

function songRow(e, q) {
  const s = e.song;
  const li = el("li", "row");
  const main = el("div", "row-main");
  const title = el("div", "row-title");
  title.append(highlight(s.n, q));
  if (s.v) title.append(el("span", "badge", `${s.v.length + 1} formats`));
  const where = s.d.split("/").slice(-2, -1)[0] || "";
  const sub = e.count
    ? `${e.count} fichier${e.count > 1 ? "s" : ""} · ${fmtDate(e.last)}${where ? " · " + where : ""}`
    : `Aucun bounce trouvé${where ? " · " + where : ""}`;
  main.append(title, el("div", "row-sub", sub));
  li.append(main);
  const best = bestFile(s);
  if (best) {
    const btn = el("button", "send", "Envoyer");
    btn.addEventListener("click", (ev) => { ev.stopPropagation(); send(best, s.n, btn); });
    li.append(btn);
  }
  li.addEventListener("click", () => openSong(s));
  return li;
}

function looseRow(file, q) {
  const li = el("li", "row");
  const main = el("div", "row-main");
  const title = el("div", "row-title");
  title.append(highlight(stripExt(file.name), q));
  main.append(title, el("div", "row-sub", `${parentName(file.p)}${file.m ? " · " + fmtDate(file.m) : ""}`));
  li.append(playButton(file), main, sendButton(file, stripExt(file.name)));
  return li;
}

// --- Recherche directe dans OneDrive (fichier récent pas encore dans le catalogue) -----

async function liveSearch() {
  const q = $("q").value.trim();
  if (!q) return;
  const list = $("live");
  list.replaceChildren(el("li", "empty", "Recherche dans OneDrive…"));
  $("more").hidden = true;
  try {
    const term = encodeURIComponent(q.replace(/'/g, "''"));
    const data = await graph(`/me/drive/root/search(q='${term}')?$top=200`);
    const files = (data.value || [])
      .filter((it) => it.file && (AUDIO_EXT.test(it.name) || (/stem/i.test(it.name) && /\.zip$/i.test(it.name))))
      .sort((a, b) => b.lastModifiedDateTime.localeCompare(a.lastModifiedDateTime));
    list.replaceChildren(el("li", "section", `OneDrive · ${files.length} fichier${files.length > 1 ? "s" : ""} audio`));
    for (const it of files.slice(0, 60)) {
      let folder = (it.parentReference && it.parentReference.path || "").replace(/^\/drive\/root:?\/?/, "");
      try { folder = decodeURIComponent(folder); } catch { /* déjà décodé */ }
      const file = { id: it.id, name: it.name, p: folder ? `${folder}/${it.name}` : it.name,
                     m: Date.parse(it.lastModifiedDateTime) / 1000 };
      const li = el("li", "row");
      const main = el("div", "row-main");
      main.append(el("div", "row-title", stripExt(it.name)),
                  el("div", "row-sub", `${folder.split("/").pop() || "OneDrive"} · ${fmtDate(file.m)}`));
      li.append(playButton(file), main, sendButton(file, stripExt(it.name)));
      list.append(li);
    }
    if (!files.length) list.append(el("li", "empty", "Aucun fichier audio trouvé."));
  } catch (e) {
    if (e instanceof NeedLogin) return showLogin();
    list.replaceChildren(el("li", "empty error", e.message));
    $("more").hidden = false;
  }
}

// --- Page d'un morceau -------------------------------------------------------------

let current = null;

function openSong(song) {
  current = song;
  $("song-title").textContent = song.n;
  $("song-sub").textContent = song.d;
  const chips = $("variants");
  chips.replaceChildren();
  chips.hidden = !song.v;
  if (song.v) {
    const formats = [{ l: song.l || "Principal", n: song.n, b: song.b, s: song.s }, ...song.v];
    formats.forEach((v, i) => {
      const b = el("button", i === 0 ? "on" : "", v.l);
      b.addEventListener("click", () => {
        chips.querySelectorAll("button").forEach((x) => x.classList.toggle("on", x === b));
        fillFiles(v);
      });
      chips.append(b);
    });
  }
  fillFiles(song);
  show("song");
  $("song").scrollTop = 0;
  history.pushState({ view: "song" }, "");
}

function fillFiles(v) {
  const title = v.n;
  const bounces = [...v.b].sort((a, b) => isMaster(b) - isMaster(a) || b[1] - a[1]).map(fileObj);
  const list = $("bounces");
  list.replaceChildren();
  for (const f of bounces) list.append(fileRow(f, title));
  if (!bounces.length) list.append(el("li", "empty", "Aucun bounce. Essaie « Chercher dans tout OneDrive »."));
  const stems = v.s.map(fileObj);
  $("stems").replaceChildren(...stems.map((f) => fileRow(f, `${title} (stems)`)));
  $("stems-title").hidden = !stems.length;
}

function fileRow(file, title) {
  const li = el("li", "row");
  const main = el("div", "row-main");
  const kind = file.dir ? "Dossier" : (file.name.split(".").pop() || "").toUpperCase();
  main.append(el("div", "row-title", file.dir ? file.name : stripExt(file.name)),
              el("div", "row-sub", `${parentName(file.p)} · ${file.m ? fmtDate(file.m) : ""} · ${kind}`));
  li.append(file.dir ? el("span", "play", "▤") : playButton(file), main, sendButton(file, title));
  return li;
}

function sendButton(file, title) {
  const btn = el("button", "send light", "Envoyer");
  btn.addEventListener("click", (ev) => { ev.stopPropagation(); send(file, title, btn); });
  return btn;
}

// --- Envoi : lien OneDrive en lecture seule, puis partage iOS / mail / copie -------

let pending = null;    // {url, title, file}

async function send(file, title, btn) {
  if (btn.getAttribute("aria-busy") === "true") return;
  btn.setAttribute("aria-busy", "true");
  const label = btn.textContent;
  btn.textContent = "Lien…";
  try {
    const url = await createLink(file);
    pending = { url, title, file };
    $("share-name").textContent = file.name;
    $("share-native").hidden = !navigator.share;
    $("share").hidden = false;
  } catch (e) {
    if (e instanceof NeedLogin) return showLogin();
    toast(e.message);
  } finally {
    btn.removeAttribute("aria-busy");
    btn.textContent = label;
  }
}

async function shareNative() {
  const { url, title } = pending;
  try {
    await navigator.share({ title, text: title, url });
    closeShare();
  } catch (e) {
    if (e.name !== "AbortError") toast("Partage impossible : " + e.message);
  }
}

async function copyLink() {
  try {
    await navigator.clipboard.writeText(pending.url);
    toast("Lien copié");
    closeShare();
  } catch {
    prompt("Copie le lien :", pending.url);
  }
}

function closeShare() { $("share").hidden = true; }

function fileLabel(file) {
  if (file.dir || /stem/i.test(file.name)) return "Stems";
  if (/pre[\s_-]?master/i.test(file.name)) return "Pre-master";
  return isMaster([file.p]) ? "Master" : "Bounce";
}

let mailFrom = "search";

// Mail : mêmes champs que sur le Mac, envoyé par l'app Mail de l'iPhone.
function contacts() {
  const seen = new Set();
  return [...store.get("recipients", []), ...((catalogue && catalogue.contacts) || [])]
    .filter((a) => a && !seen.has(a.toLowerCase()) && seen.add(a.toLowerCase()));
}

function openMail() {
  const { url, title, file } = pending;
  closeShare();
  $("mail-to").value = "";
  $("mail-subject").value = title;
  const sig = (catalogue && catalogue.signature || "").trim();
  $("mail-body").value = `Salut,\n\nVoici le lien pour « ${title} » :\n\n${fileLabel(file)} : ${url}`
    + (sig ? `\n\n-- \n${sig}` : "") + "\n";
  const all = contacts();
  $("mail-contacts").replaceChildren(...all.map((a) => Object.assign(document.createElement("option"), { value: a })));
  const chips = $("mail-chips");
  chips.replaceChildren();
  for (const a of all.slice(0, 12)) {
    const b = el("button", "", a);
    b.type = "button";
    b.addEventListener("click", () => {
      const cur = $("mail-to").value.split(",").map((x) => x.trim()).filter(Boolean);
      if (!cur.includes(a)) cur.push(a);
      $("mail-to").value = cur.join(", ");
    });
    chips.append(b);
  }
  chips.hidden = !all.length;
  mailFrom = $("song").hidden ? "search" : "song";
  show("mail");
  history.pushState({ view: "mail" }, "");
}

function sendMail() {
  const to = $("mail-to").value.split(/[,;\s]+/).filter((a) => a.includes("@"));
  if (!to.length) { toast("Ajoute un destinataire"); $("mail-to").focus(); return; }
  const recents = store.get("recipients", []).filter((a) => !to.includes(a));
  store.set("recipients", [...to, ...recents].slice(0, 30));
  const params = `subject=${encodeURIComponent($("mail-subject").value)}&body=${encodeURIComponent($("mail-body").value)}`;
  location.href = `mailto:${to.map(encodeURIComponent).join(",")}?${params}`;
  history.back();
}

// --- Écoute -------------------------------------------------------------------------

let playing = null;

function playButton(file) {
  const b = el("button", "play", "▶");
  b.dataset.key = file.id || file.p;
  if (playing && playing === b.dataset.key) { b.classList.add("on"); b.textContent = "❚❚"; }
  b.addEventListener("click", (ev) => { ev.stopPropagation(); togglePlay(file); });
  return b;
}

function syncPlayButtons() {
  const audio = $("audio");
  document.querySelectorAll(".play[data-key]").forEach((b) => {
    const on = b.dataset.key === playing;
    b.classList.toggle("on", on);
    b.textContent = on && !audio.paused ? "❚❚" : "▶";
  });
  $("play-toggle").textContent = audio.paused ? "▶" : "❚❚";
}

async function togglePlay(file) {
  const audio = $("audio");
  const key = file.id || file.p;
  if (playing === key) {
    audio.paused ? audio.play() : audio.pause();
    return;
  }
  playing = key;
  $("play-name").textContent = "Chargement… " + file.name;
  $("player").hidden = false;
  syncPlayButtons();
  try {
    const item = await graph(itemUrl(file).replace(/:$/, ""));
    if (playing !== key) return;
    audio.src = item["@microsoft.graph.downloadUrl"];
    $("play-name").textContent = file.name;
    await audio.play();
  } catch (e) {
    if (e instanceof NeedLogin) return showLogin();
    $("play-name").textContent = file.name;
    if (e.name !== "AbortError") toast("Lecture impossible : " + e.message);
  }
}

function clock(sec) {
  if (!isFinite(sec) || sec < 0) return "0:00";
  const m = Math.floor(sec / 60);
  return `${m}:${String(Math.floor(sec % 60)).padStart(2, "0")}`;
}

function stopPlay() {
  const audio = $("audio");
  audio.pause();
  audio.removeAttribute("src");
  audio.load();
  playing = null;
  $("player").hidden = true;
  syncPlayButtons();
}

// --- Navigation, réglages ------------------------------------------------------------

const SHEETS = ["song", "settings", "mail"];
function show(name) {
  for (const s of SHEETS) $(s).hidden = s !== name;
  $("search").hidden = name === "login";
  $("login").hidden = name !== "login";
  document.body.style.overflow = SHEETS.includes(name) ? "hidden" : "";
}

window.addEventListener("popstate", () => {
  if ($("login").hidden) show(!$("mail").hidden ? mailFrom : "search");
});

function showLogin(message) {
  show("login");
  $("login-error").hidden = !message;
  $("login-error").textContent = message || "";
}

function openSettings() {
  $("cat-path").value = store.get("catPath", CFG.catalogue);
  $("cat-info").textContent = catalogue
    ? `${(catalogue.songs || []).length} morceaux, ${(catalogue.loose || []).length} exports isolés. `
      + `Écrit par le Mac ${ago(catalogue.generated)}.`
    : "Pas encore de catalogue.";
  $("account").textContent = store.get("account") ? `Connecté : ${store.get("account")}` : "Connecté à OneDrive";
  $("version").textContent = `SIRØB Mobile ${VERSION}`;
  show("settings");
  history.pushState({ view: "settings" }, "");
}

let toastTimer;
function toast(text) {
  const t = $("toast");
  t.textContent = text;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, 3000);
}

async function rememberAccount() {
  if (store.get("account")) return;
  try {
    const me = await graph("/me?$select=userPrincipalName,mail");
    store.set("account", me.userPrincipalName || me.mail || "");
  } catch { /* pas grave */ }
}

function wire() {
  $("login-btn").addEventListener("click", login);
  let timer;
  $("q").addEventListener("input", () => { clearTimeout(timer); timer = setTimeout(render, 60); });
  $("q").addEventListener("keydown", (e) => { if (e.key === "Enter") e.target.blur(); });
  $("clear").addEventListener("click", () => { $("q").value = ""; render(); $("q").focus(); });
  $("live-btn").addEventListener("click", liveSearch);
  $("back").addEventListener("click", () => history.back());
  $("settings-btn").addEventListener("click", openSettings);
  $("settings-back").addEventListener("click", () => history.back());
  $("mail-back").addEventListener("click", () => history.back());
  $("mail-send").addEventListener("click", sendMail);
  $("reload").addEventListener("click", () => { history.back(); loadCatalogue(true); });
  $("cat-path").addEventListener("change", (e) => {
    const v = e.target.value.trim().replace(/^\/+/, "");
    store.set("catPath", v || CFG.catalogue);
    loadCatalogue(true);
  });
  $("logout").addEventListener("click", logout);
  $("share-native").addEventListener("click", shareNative);
  $("share-mail").addEventListener("click", openMail);
  $("share-copy").addEventListener("click", copyLink);
  $("share-close").addEventListener("click", closeShare);
  $("share").addEventListener("click", (e) => { if (e.target.id === "share") closeShare(); });
  $("play-toggle").addEventListener("click", () => { const a = $("audio"); a.paused ? a.play() : a.pause(); });
  $("play-close").addEventListener("click", stopPlay);
  const audio = $("audio");
  for (const ev of ["play", "pause", "ended"]) audio.addEventListener(ev, syncPlayButtons);
  // Barre de lecture : on peut la faire glisser pour avancer ou reculer.
  const pos = $("play-pos");
  let seeking = false;
  const paint = (ratio) => pos.style.setProperty("--p", `${(ratio * 100).toFixed(2)}%`);
  const showTime = () => {
    const d = audio.duration;
    if (!seeking) {
      pos.value = d ? Math.round((audio.currentTime / d) * 1000) : 0;
      paint(d ? audio.currentTime / d : 0);
    }
    $("play-time").textContent = clock(seeking && d ? (pos.value / 1000) * d : audio.currentTime);
    $("play-dur").textContent = clock(d);
  };
  for (const ev of ["timeupdate", "loadedmetadata", "durationchange", "emptied"]) audio.addEventListener(ev, showTime);
  pos.addEventListener("input", () => { seeking = true; paint(pos.value / 1000); showTime(); });
  pos.addEventListener("change", () => {
    if (audio.duration) audio.currentTime = (pos.value / 1000) * audio.duration;
    seeking = false;
    showTime();
  });
  // Retour dans l'app après un moment : catalogue rafraîchi si le Mac l'a réécrit.
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden && store.get("tok") && Date.now() - store.get("checkedAt", 0) > 5 * 60_000) {
      loadCatalogue();
    }
  });
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
  show("search");
  buildEntries();
  renderFilters();
  render();
  loadCatalogue();
  rememberAccount();
}

start();
