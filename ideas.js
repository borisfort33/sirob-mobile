/* SIRØB Mobile : idées de prod, façon Notes, synchronisées avec l'app du Mac.
 *
 * Fichier partagé : OneDrive « SIROB Browser/idees.json » (même format que flbrowser/ideas.py).
 * Chaque appareil relit le fichier avant d'écrire et fusionne idée par idée (la modification la
 * plus récente gagne ; une idée supprimée devient {"id", "u", "d": 1} pendant 90 jours).
 * L'écriture utilise l'eTag (If-Match) : si le Mac a écrit entre-temps, on refusionne et on réessaie.
 *
 * Les liens (YouTube, SoundCloud, Spotify, Deezer, Apple Music, Shazam) se jouent ici même
 * dans le lecteur officiel « intégrable » de chaque site.
 */
"use strict";

const IDEAS_PATH = "SIROB Browser/idees.json";
const TOMBSTONE_MS = 90 * 86400_000;
const URL_RE = /https?:\/\/[^\s<>"')\]]+/gi;
const MONTHS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
  "septembre", "octobre", "novembre", "décembre"];

let ideas = store.get("ideas", []);
let ideasETag = store.get("ideasETag", null);
let ideasDirty = store.get("ideasDirty", false);
let ideaTab = store.get("tab", "songs") === "ideas";
let currentIdea = null;     // id de l'idée ouverte
let ideasSyncing = null;
let ideaSaveTimer;

// --- Données --------------------------------------------------------------------------

const nowMs = () => Date.now();
const newId = () => Array.from(crypto.getRandomValues(new Uint8Array(6)), (b) => b.toString(16).padStart(2, "0")).join("");

function mergeIdeas(...lists) {
  const best = new Map();
  for (const items of lists) {
    for (const it of items || []) {
      if (!it || !it.id) continue;
      const cur = best.get(it.id);
      if (!cur || (it.u || 0) > (cur.u || 0) || ((it.u || 0) === (cur.u || 0) && it.d && !cur.d)) best.set(it.id, it);
    }
  }
  const limit = nowMs() - TOMBSTONE_MS;
  return [...best.values()].filter((it) => !(it.d && (it.u || 0) < limit));
}

const sameIdeas = (a, b) => JSON.stringify([...a].sort((x, y) => (x.id < y.id ? -1 : 1)))
  === JSON.stringify([...b].sort((x, y) => (x.id < y.id ? -1 : 1)));

const ideaTitle = (it) => ((it.x || "").split("\n").map((l) => l.trim()).find(Boolean)) || "Nouvelle idée";
const ideaPreview = (it) => (it.x || "").split("\n").map((l) => l.trim()).filter(Boolean).slice(1).join(" ");
function ideaLinks(it) {
  const out = [];
  for (const m of (it.x || "").matchAll(URL_RE)) {
    const u = m[0].replace(/[.,;:!?]+$/, "");
    if (!out.includes(u)) out.push(u);
  }
  return out;
}
const liveIdeas = () => ideas.filter((it) => !it.d)
  .sort((a, b) => (b.p ? 1 : 0) - (a.p ? 1 : 0) || (b.u || 0) - (a.u || 0));
const getIdea = (id) => ideas.find((it) => it.id === id && !it.d);

function saveLocal() {
  store.set("ideas", ideas);
  store.set("ideasDirty", ideasDirty);
}

function touch(id, fields) {
  const it = getIdea(id);
  if (!it) return null;
  if (Object.entries(fields).every(([k, v]) => it[k] === v)) return it;
  Object.assign(it, fields, { u: Math.max(nowMs(), (it.u || 0) + 1) });
  ideasDirty = true;
  saveLocal();
  scheduleSync();
  return it;
}

// --- Synchro OneDrive -----------------------------------------------------------------

async function fetchRemote() {
  try {
    const item = await graph(`/me/drive/root:/${encPath(IDEAS_PATH)}`);
    const url = item["@microsoft.graph.downloadUrl"];
    if (!url) throw new Error("OneDrive n'a pas donné le fichier des idées");
    const resp = await fetch(url, { cache: "no-store" });
    if (!resp.ok) throw new Error(`Téléchargement des idées refusé (${resp.status})`);
    const data = await resp.json().catch(() => ({}));
    return { eTag: item.eTag, items: Array.isArray(data) ? data : (data.ideas || []) };
  } catch (e) {
    if (e.status === 404) return { eTag: null, items: [] };
    throw e;
  }
}

async function putRemote(items, eTag) {
  const body = JSON.stringify({ version: 1, updated: nowMs(), ideas: items }, null, 1);
  const headers = eTag ? { "If-Match": eTag } : {};
  const path = `/me/drive/root:/${encPath(IDEAS_PATH)}:/content` + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
  const item = await graph(path, { method: "PUT", body, headers });
  return item.eTag;
}

function scheduleSync(delay = 1500) {
  clearTimeout(ideaSaveTimer);
  ideaSaveTimer = setTimeout(syncIdeas, delay);
}

async function syncIdeas() {
  if (!store.get("tok")) return;
  if (ideasSyncing) { await ideasSyncing; if (!ideasDirty) return; }
  ideasSyncing = (async () => {
    for (let attempt = 0; attempt < 4; attempt++) {
      let remote;
      try { remote = await fetchRemote(); } catch (e) {
        if (e instanceof NeedLogin) return;
        setIdeasStatus(`Synchro impossible : ${e.message}${ideasDirty ? " · modifications gardées sur l'iPhone" : ""}`);
        return;
      }
      const merged = mergeIdeas(ideas, remote.items);
      const changedHere = !sameIdeas(merged, ideas);
      ideas = merged;
      if (!sameIdeas(merged, remote.items)) {
        try {
          ideasETag = await putRemote(merged, remote.eTag);
        } catch (e) {
          if (e.status === 412 || e.status === 409) continue;   // le Mac vient d'écrire : on refusionne
          setIdeasStatus(`Synchro impossible (${e.message}) · modifications gardées sur l'iPhone`);
          saveLocal();
          return;
        }
      } else {
        ideasETag = remote.eTag;
      }
      ideasDirty = false;
      saveLocal();
      store.set("ideasETag", ideasETag);
      store.set("ideasAt", nowMs());
      if (changedHere) refreshIdeasUi();
      setIdeasStatus();
      return;
    }
  })();
  try { await ideasSyncing; } finally { ideasSyncing = null; }
}

// --- Liste (onglet Idées) ---------------------------------------------------------------

function ideasActive() { return ideaTab; }

function setTab(tab) {
  ideaTab = tab === "ideas";
  store.set("tab", ideaTab ? "ideas" : "songs");
  document.querySelectorAll("#tabs button").forEach((b) => b.classList.toggle("on", b.dataset.tab === (ideaTab ? "ideas" : "songs")));
  $("results").hidden = ideaTab;
  $("live").hidden = ideaTab;
  if (ideaTab) $("more").hidden = true;
  $("filters").hidden = ideaTab || !$("filters").childElementCount;
  $("ideas-list").hidden = !ideaTab;
  $("idea-new").hidden = !ideaTab;
  $("q").placeholder = ideaTab ? "Chercher une idée, un nom…" : "Titre, artiste, client…";
  if (ideaTab) {
    renderIdeas();
    syncIdeas();   // à chaque ouverture de l'onglet : on récupère ce que le Mac a écrit
  } else {
    if (typeof renderFilters === "function") renderFilters();
    render();
  }
}

function setIdeasStatus(extra) {
  if (!ideaTab) return;
  const n = liveIdeas().length;
  const base = `${n} idée${n > 1 ? "s" : ""} de prod`;
  $("status").textContent = extra ? `${base} · ${extra}` : `${base} · synchronisées avec le Mac${ideasDirty ? " · envoi…" : ""}`;
}

function shortDate(ms) {
  const d = new Date(ms);
  const today = new Date();
  const days = Math.floor((new Date(today.toDateString()) - new Date(d.toDateString())) / 86400_000);
  if (days === 0) return d.toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" });
  if (days === 1) return "Hier";
  if (days < 7) return d.toLocaleDateString("fr-FR", { weekday: "long" });
  return d.toLocaleDateString("fr-FR");
}
function longDate(ms) {
  const d = new Date(ms);
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()} à ${d.toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}`;
}

function renderIdeas() {
  const box = $("ideas-list");
  box.replaceChildren();
  const q = norm($("q").value.trim());
  $("clear").hidden = !$("q").value;
  const list = liveIdeas().filter((it) => !q || norm(`${it.x || ""} ${it.f || ""}`).includes(q));
  const groups = [["Épinglées", list.filter((it) => it.p)], ["Idées", list.filter((it) => !it.p)]];
  for (const [label, items] of groups) {
    if (!items.length) continue;
    if (groups[0][1].length) box.append(el("h3", "notes-head", label));
    const card = el("ul", "notes-card");
    for (const it of items) {
      const li = el("li", "note" + (it.k ? " done" : ""));
      li.append(el("div", "note-title", (it.k ? "✓ " : "") + ideaTitle(it)));
      const sub = el("div", "note-sub");
      sub.append(el("span", "note-date", shortDate(it.u || it.c || 0)));
      sub.append(document.createTextNode(" " + (ideaPreview(it) || (ideaLinks(it).length ? linkKind(ideaLinks(it)[0])[1] : "Aucun texte supplémentaire"))));
      li.append(sub);
      if (it.f) li.append(el("div", "note-from", `de la part de ${it.f}`));
      li.addEventListener("click", () => openIdea(it.id));
      card.append(li);
    }
    box.append(card);
  }
  if (!list.length) {
    box.append(el("p", "empty center", q ? "Aucune idée ne correspond." : "Aucune idée pour l'instant.\nAppuie sur ✎ pour en noter une."));
  }
  setIdeasStatus();
}

function refreshIdeasUi() {
  if (ideaTab) renderIdeas();
  if (currentIdea && !$("idea").hidden) {
    const it = getIdea(currentIdea);
    if (!it) { history.back(); toast("Idée supprimée sur le Mac"); return; }
    if (document.activeElement !== $("idea-text") && $("idea-text").value !== (it.x || "")) $("idea-text").value = it.x || "";
    if (document.activeElement !== $("idea-from") && $("idea-from").value !== (it.f || "")) $("idea-from").value = it.f || "";
    fillIdeaMeta(it);
  }
}

// --- Note ouverte ------------------------------------------------------------------------

function openIdea(id, focus = false) {
  const it = getIdea(id);
  if (!it) return;
  currentIdea = id;
  $("idea-text").value = it.x || "";
  $("idea-from").value = it.f || "";
  const names = new Set([...ideas.filter((x) => x.f).map((x) => x.f), ...contacts().map((c) => c.split("@")[0])]);
  $("idea-senders").replaceChildren(...[...names].sort().map((n) => Object.assign(document.createElement("option"), { value: n })));
  closeEmbed();
  fillIdeaMeta(it);
  show("idea");
  $("idea").scrollTop = 0;
  history.pushState({ view: "idea" }, "");
  autoGrow();
  if (focus) setTimeout(() => $("idea-text").focus(), 50);
}

function fillIdeaMeta(it) {
  $("idea-date").textContent = longDate(it.u || it.c || nowMs());
  $("idea-pin").classList.toggle("on", !!it.p);
  $("idea-done").classList.toggle("on", !!it.k);
  const box = $("idea-links");
  box.replaceChildren();
  for (const url of ideaLinks(it)) {
    const [icon, label] = linkKind(url);
    const playable = !!embedFor(url) || isShazam(url);
    const row = el("div", "link-row");
    const b = el("button", "link-btn", `${playable ? "▶" : icon}  ${label}`);
    b.addEventListener("click", () => (playable ? playLink(url, row) : window.open(url, "_blank")));
    const open = el("button", "link-open", "↗");
    open.setAttribute("aria-label", "Ouvrir dans l'app");
    open.addEventListener("click", () => window.open(url, "_blank"));
    const top = el("div", "link-top");
    top.append(b, open);
    row.append(top);
    box.append(row);
  }
  $("idea-links-title").hidden = !box.childElementCount;
}

function flushIdea() {
  clearTimeout(flushIdea.t);
  if (!currentIdea) return;
  touch(currentIdea, { x: $("idea-text").value, f: $("idea-from").value.trim() });
}

function newIdea() {
  const t = nowMs();
  const it = { id: newId(), x: "", f: "", c: t, u: t, p: 0, k: 0 };
  ideas.push(it);
  ideasDirty = true;
  saveLocal();
  openIdea(it.id, true);
}

function closeIdea() {
  if (!currentIdea) return;
  flushIdea();
  const it = getIdea(currentIdea);
  // Note laissée vide (comme Notes) : elle disparaît.
  if (it && !(it.x || "").trim() && !(it.f || "").trim()) deleteIdea(it.id, true);
  currentIdea = null;
  closeEmbed();
  if (ideaTab) renderIdeas();
  syncIdeas();
}

function deleteIdea(id, silent = false) {
  const i = ideas.findIndex((x) => x.id === id);
  if (i < 0) return;
  ideas[i] = { id, u: Math.max(nowMs(), (ideas[i].u || 0) + 1), d: 1 };
  ideasDirty = true;
  saveLocal();
  if (!silent) scheduleSync(10);
}

function autoGrow() {
  const t = $("idea-text");
  t.style.height = "auto";
  t.style.height = Math.max(240, t.scrollHeight + 4) + "px";
}

// --- Lecture des liens -------------------------------------------------------------------

function linkKind(url) {
  const u = url.toLowerCase();
  if (u.includes("youtu")) return ["▶", "YouTube"];
  if (u.includes("shazam") || u.includes("shz.am")) return ["Ⓢ", "Shazam"];
  if (u.includes("spotify")) return ["●", "Spotify"];
  if (u.includes("music.apple") || u.includes("itunes.apple")) return ["♫", "Apple Music"];
  if (u.includes("soundcloud")) return ["☁", "SoundCloud"];
  if (u.includes("deezer")) return ["≋", "Deezer"];
  if (u.includes("tiktok")) return ["♪", "TikTok"];
  if (u.includes("instagram")) return ["◎", "Instagram"];
  return ["🔗", url.replace(/^https?:\/\/(www\.)?/i, "").split("/")[0] || "Lien"];
}

const isShazam = (url) => /shazam\.com|shz\.am/i.test(url);

/** [adresse du lecteur intégrable, hauteur en px (0 = 16/9)] ou null. Même logique que linkplay.py. */
function embedFor(url) {
  let m;
  if (/youtu/i.test(url)) {
    m = url.match(/(?:youtu\.be\/|[?&]v=|\/shorts\/|\/embed\/|\/live\/)([A-Za-z0-9_-]{11})/);
    if (m) {
      const t = url.match(/[?&#]t=(?:(\d+)h)?(?:(\d+)m)?(\d+)s?/);
      const start = t ? (+(t[1] || 0)) * 3600 + (+(t[2] || 0)) * 60 + (+(t[3] || 0)) : 0;
      return [`https://www.youtube-nocookie.com/embed/${m[1]}?autoplay=1&playsinline=1&rel=0${start ? `&start=${start}` : ""}`, 0];
    }
    m = url.match(/[?&]list=([A-Za-z0-9_-]+)/);
    return m ? [`https://www.youtube-nocookie.com/embed/videoseries?list=${m[1]}&playsinline=1`, 0] : null;
  }
  if (/soundcloud\.com/i.test(url)) {
    return [`https://w.soundcloud.com/player/?url=${encodeURIComponent(url)}&auto_play=true&visual=false&color=%23e5b937&show_comments=false&hide_related=true`, 166];
  }
  m = url.match(/open\.spotify\.com\/(?:intl-[a-z]+\/)?(track|album|playlist|episode|artist)\/([A-Za-z0-9]+)/);
  if (m) return [`https://open.spotify.com/embed/${m[1]}/${m[2]}?theme=0`, m[1] === "track" ? 152 : 352];
  m = url.match(/deezer\.com\/(?:[a-z]{2}\/)?(track|album|playlist)\/(\d+)/);
  if (m) return [`https://widget.deezer.com/widget/dark/${m[1]}/${m[2]}?autoplay=true`, m[1] === "track" ? 150 : 300];
  if (/music\.apple\.com/i.test(url)) {
    return [url.split("&uo=")[0].replace(/https?:\/\/(geo\.)?music\.apple\.com/i, "https://embed.music.apple.com"), 175];
  }
  return null;
}

/** iTunes Search en JSONP (l'API n'autorise pas les appels directs depuis une page web). */
function itunes(params) {
  return new Promise((resolve, reject) => {
    const cb = "sirobItunes" + Math.random().toString(36).slice(2);
    const s = document.createElement("script");
    const timer = setTimeout(() => { cleanup(); reject(new Error("délai dépassé")); }, 10000);
    function cleanup() { clearTimeout(timer); delete window[cb]; s.remove(); }
    window[cb] = (data) => { cleanup(); resolve(data.results || []); };
    s.onerror = () => { cleanup(); reject(new Error("iTunes injoignable")); };
    s.src = `https://itunes.apple.com/${params}&country=fr&callback=${cb}`;
    document.head.append(s);
  });
}

/** Lien Shazam -> lecteur Apple Music (extrait 30 s) du même morceau. */
async function shazamEmbed(url, context) {
  let m = url.match(/\/song\/(\d+)(?:\/([^/?#]+))?/);
  let results = [];
  if (m) results = await itunes(`lookup?id=${m[1]}&entity=song`);
  let words = m ? (m[2] || "").replace(/-/g, " ") : "";
  if (!m) {
    m = url.match(/\/track\/\d+(?:\/([^/?#]+))?/);
    words = m ? (m[1] || "").replace(/-/g, " ") : "";
  }
  if (!words && context) {
    const t = context.match(/trouvé (.+?) (?:par|de) (.+?) (?:avec|sur|grâce|via)/) || context.match(/(?:found|discovered) (.+?) by (.+?) (?:with|on|using)/i);
    if (t) words = `${t[1]} ${t[2]}`;
  }
  if (!results.length && words.trim()) results = await itunes(`search?term=${encodeURIComponent(words)}&entity=song&limit=1`);
  const r = results.find((x) => x.trackViewUrl);
  return r ? { src: embedFor(r.trackViewUrl)[0], title: `${r.artistName} – ${r.trackName}` } : null;
}

function closeEmbed() {
  document.querySelectorAll("#idea-links .embed").forEach((n) => n.remove());
}

async function playLink(url, row) {
  const already = row.querySelector(".embed");
  closeEmbed();
  if (already) return;          // 2e appui : on ferme le lecteur
  const holder = el("div", "embed");
  row.append(holder);
  let src, height, title = "";
  if (isShazam(url)) {
    holder.append(el("p", "dim small", "Recherche du morceau…"));
    try {
      const found = await shazamEmbed(url, $("idea-text").value);
      if (!found) throw new Error("introuvable");
      [src, height, title] = [found.src, 175, found.title];
    } catch {
      holder.replaceChildren(el("p", "dim small", "Morceau introuvable : ouverture de Shazam…"));
      window.open(url, "_blank");
      return;
    }
  } else {
    [src, height] = embedFor(url);
  }
  if (!holder.isConnected) return;
  holder.replaceChildren();
  if (title) holder.append(el("p", "dim small", title));
  const frame = document.createElement("iframe");
  frame.src = src;
  frame.allow = "autoplay; encrypted-media; clipboard-write; fullscreen; picture-in-picture";
  frame.allowFullscreen = true;
  frame.referrerPolicy = "strict-origin-when-cross-origin";
  frame.className = height ? "" : "video";
  if (height) frame.style.height = height + "px";
  holder.append(frame);
  // Un seul son à la fois : le lecteur OneDrive se met en pause.
  const a = $("audio");
  if (a && !a.paused) a.pause();
}

// --- Branchements ------------------------------------------------------------------------

function wireIdeas() {
  document.querySelectorAll("#tabs button").forEach((b) => b.addEventListener("click", () => setTab(b.dataset.tab)));
  $("idea-new").addEventListener("click", newIdea);
  $("idea-back").addEventListener("click", () => history.back());
  $("idea-text").addEventListener("input", () => { autoGrow(); clearTimeout(flushIdea.t); flushIdea.t = setTimeout(flushIdea, 800); });
  $("idea-text").addEventListener("blur", () => { flushIdea(); const it = getIdea(currentIdea); if (it) fillIdeaMeta(it); });
  $("idea-from").addEventListener("input", () => { clearTimeout(flushIdea.t); flushIdea.t = setTimeout(flushIdea, 800); });
  $("idea-from").addEventListener("change", flushIdea);
  $("idea-pin").addEventListener("click", () => { flushIdea(); const it = getIdea(currentIdea); if (it) fillIdeaMeta(touch(it.id, { p: it.p ? 0 : 1 })); });
  $("idea-done").addEventListener("click", () => {
    flushIdea();
    const it = getIdea(currentIdea);
    if (it) { fillIdeaMeta(touch(it.id, { k: it.k ? 0 : 1 })); toast(it.k ? "Idée marquée comme faite" : "Idée remise à faire"); }
  });
  $("idea-delete").addEventListener("click", () => {
    const it = getIdea(currentIdea);
    if (!it || ((it.x || "").trim() && !confirm(`Supprimer « ${ideaTitle(it)} » ?\nElle disparaît aussi du Mac.`))) return;
    deleteIdea(it.id);
    currentIdea = null;
    history.back();
  });
  // Coller le lien reçu (WhatsApp, Shazam…) d'un appui, en bas de la note.
  $("idea-paste").addEventListener("click", async () => {
    let text = "";
    try { text = await navigator.clipboard.readText(); } catch { /* refusé */ }
    if (!text) { toast("Presse-papiers vide ou refusé : appui long dans le texte › Coller"); return; }
    const t = $("idea-text");
    t.value = (t.value.trimEnd() ? t.value.trimEnd() + "\n" : "") + text.trim();
    autoGrow();
    flushIdea();
    fillIdeaMeta(getIdea(currentIdea));
  });
  window.addEventListener("popstate", () => { if (currentIdea && $("idea").hidden) closeIdea(); });
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) { flushIdea(); if (ideasDirty) syncIdeas(); return; }
    if (store.get("tok") && (ideasDirty || nowMs() - store.get("ideasAt", 0) > 30_000)) syncIdeas();
  });
  window.addEventListener("online", () => { if (ideasDirty) syncIdeas(); });
}

wireIdeas();
if (ideaTab) setTab("ideas");
if (store.get("tok")) syncIdeas();
