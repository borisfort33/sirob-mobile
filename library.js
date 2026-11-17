/* SIRØB Mobile : bibliothèque façon Musique (dossiers + playlists), synchronisée avec le Mac.
 *
 * Fichier partagé : OneDrive « SIROB Browser/bibliotheque.json » (même format que
 * flbrowser/library.py). Un nœud = {id, t: "f" dossier | "p" playlist, n nom, in dossier parent,
 * o ordre, c créé, u modifié, s: [{k clé du morceau, b export choisi?}]} ; supprimé = {id, u, d: 1}.
 * Fusion nœud par nœud (la modification la plus récente gagne), écriture avec l'eTag (If-Match).
 * Les morceaux sont retrouvés dans catalogue.json par leur clé « k ».
 * Partage : uniquement les exports audio (jamais les stems).
 */
"use strict";

const LIB_PATH = "SIROB Browser/bibliotheque.json";
const LIB_TOMBSTONE_MS = 90 * 86400_000;

let libNodes = store.get("lib", []);
let libETag = store.get("libETag", null);
let libDirty = store.get("libDirty", false);
let libSyncing = null;
let libSaveTimer;
let libAt = "";            // dossier ou playlist affiché ("" = racine)
let libSelect = null;      // Set des positions sélectionnées (mode sélection), sinon null
let libError = "";

const libNow = () => Date.now();
const libId = () => Array.from(crypto.getRandomValues(new Uint8Array(6)), (b) => b.toString(16).padStart(2, "0")).join("");

// --- Données --------------------------------------------------------------------------

function mergeLib(...lists) {
  const best = new Map();
  for (const items of lists) {
    for (const it of items || []) {
      if (!it || !it.id) continue;
      const cur = best.get(it.id);
      if (!cur || (it.u || 0) > (cur.u || 0) || ((it.u || 0) === (cur.u || 0) && it.d && !cur.d)) best.set(it.id, it);
    }
  }
  const limit = libNow() - LIB_TOMBSTONE_MS;
  return [...best.values()].filter((it) => !(it.d && (it.u || 0) < limit));
}
const sortedJson = (a) => JSON.stringify([...a].sort((x, y) => (x.id < y.id ? -1 : 1)));

const libLive = () => libNodes.filter((n) => !n.d);
const libGet = (id) => libNodes.find((n) => n.id === id && !n.d) || null;
function libParent(n) {
  const p = n.in ? libGet(n.in) : null;
  return p && p.t === "f" ? p.id : "";
}
function libChildren(parent = "") {
  return libLive().filter((n) => libParent(n) === parent)
    .sort((a, b) => (a.t !== "f") - (b.t !== "f") || (a.o || 0) - (b.o || 0) || (a.n || "").localeCompare(b.n || "", "fr"));
}
function libWalk(parent = "", depth = 0, out = []) {
  for (const n of libChildren(parent)) {
    out.push([n, depth]);
    if (n.t === "f") libWalk(n.id, depth + 1, out);
  }
  return out;
}
function libPathOf(n) {
  const names = [n.n];
  const seen = new Set([n.id]);
  let pid = libParent(n);
  while (pid && !seen.has(pid)) {
    seen.add(pid);
    const p = libGet(pid);
    if (!p) break;
    names.unshift(p.n);
    pid = libParent(p);
  }
  return names.join(" / ");
}

function libSaveLocal() {
  store.set("lib", libNodes);
  store.set("libDirty", libDirty);
}
function libTouch(n) {
  n.u = Math.max(libNow(), (n.u || 0) + 1);
  libDirty = true;
  libSaveLocal();
  clearTimeout(libSaveTimer);
  libSaveTimer = setTimeout(syncLibrary, 1200);
  return n;
}
function libAdd(t, name, parent = "") {
  const siblings = libChildren(parent);
  const n = { id: libId(), t, n: name.trim() || "Sans titre", in: parent, c: libNow(), u: libNow(),
              o: siblings.length ? Math.max(...siblings.map((x) => x.o || 0)) + 1 : 0 };
  if (t === "p") n.s = [];
  libNodes.push(n);
  libTouch(n);
  return n;
}
function libDelete(id) {
  const doomed = new Set([id]);
  if ((libGet(id) || {}).t === "f") libWalk(id).forEach(([n]) => doomed.add(n.id));
  libNodes = libNodes.map((n) => (doomed.has(n.id) && !n.d ? { id: n.id, u: Math.max(libNow(), (n.u || 0) + 1), d: 1 } : n));
  libDirty = true;
  libSaveLocal();
  clearTimeout(libSaveTimer);
  libSaveTimer = setTimeout(syncLibrary, 600);
}
function libAddSongs(pid, songs) {
  const n = libGet(pid);
  if (!n || n.t !== "p") return 0;
  n.s = n.s || [];
  const have = new Set(n.s.map((e) => `${e.k}|${e.b || ""}`));
  const fresh = songs.filter((e) => e.k && !have.has(`${e.k}|${e.b || ""}`));
  if (!fresh.length) return 0;
  n.s.push(...fresh);
  libTouch(n);
  return fresh.length;
}

// --- Synchro OneDrive -----------------------------------------------------------------

async function libFetch() {
  try {
    const item = await graph(`/me/drive/root:/${encPath(LIB_PATH)}`);
    const url = item["@microsoft.graph.downloadUrl"];
    if (!url) throw new Error("OneDrive n'a pas donné le fichier de la bibliothèque");
    const resp = await fetch(url, { cache: "no-store" });
    if (!resp.ok) throw new Error(`Téléchargement de la bibliothèque refusé (${resp.status})`);
    const data = await resp.json().catch(() => ({}));
    return { eTag: item.eTag, items: Array.isArray(data) ? data : (data.nodes || []) };
  } catch (e) {
    if (e.status === 404) return { eTag: null, items: [] };
    throw e;
  }
}

async function libPut(items, eTag) {
  const body = JSON.stringify({ version: 1, updated: libNow(), nodes: items }, null, 1);
  const headers = eTag ? { "If-Match": eTag } : {};
  const path = `/me/drive/root:/${encPath(LIB_PATH)}:/content` + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
  const item = await graph(path, { method: "PUT", body, headers });
  return item.eTag;
}

async function syncLibrary() {
  if (!store.get("tok")) return;
  if (libSyncing) { await libSyncing; if (!libDirty) return; }
  libSyncing = (async () => {
    for (let attempt = 0; attempt < 4; attempt++) {
      let remote;
      try { remote = await libFetch(); } catch (e) {
        if (e instanceof NeedLogin) return;
        libError = `Synchro impossible : ${e.message}${libDirty ? " · modifications gardées sur l'iPhone" : ""}`;
        setLibStatus();
        return;
      }
      const merged = mergeLib(libNodes, remote.items);
      const changedHere = sortedJson(merged) !== sortedJson(libNodes);
      libNodes = merged;
      if (sortedJson(merged) !== sortedJson(remote.items)) {
        try {
          libETag = await libPut(merged, remote.eTag);
        } catch (e) {
          if (e.status === 412 || e.status === 409) continue;   // le Mac vient d'écrire : on refusionne
          libError = `Synchro impossible (${e.message}) · modifications gardées sur l'iPhone`;
          libSaveLocal();
          setLibStatus();
          return;
        }
      } else {
        libETag = remote.eTag;
      }
      libDirty = false;
      libError = "";
      libSaveLocal();
      store.set("libETag", libETag);
      store.set("libAt", libNow());
      if (changedHere && libActive()) renderLibrary();
      setLibStatus();
      return;
    }
  })();
  try { await libSyncing; } finally { libSyncing = null; }
}

// --- Morceaux (catalogue du Mac) ---------------------------------------------------------

function songsByKey() {
  const map = new Map();
  for (const s of (catalogue && catalogue.songs) || []) if (s.k) map.set(s.k, s);
  return map;
}
/** Export joué / envoyé : celui choisi dans la playlist (sur le Mac), sinon le master le plus récent. */
function libFile(song, entry) {
  if (!song) return null;
  if (entry && entry.b) {
    const all = [...song.b, ...(song.v || []).flatMap((v) => v.b)];
    const f = all.find((x) => x[0] === entry.b);
    if (f) return fileObj(f);
  }
  return bestFile(song);
}
function libTracks(node) {
  const map = songsByKey();
  if (node.t === "f") {   // dossier : morceaux de toutes ses playlists, sans doublon (comme Musique)
    const seen = new Set();
    const out = [];
    for (const [sub] of libWalk(node.id)) {
      for (const e of sub.s || []) {
        if (seen.has(e.k)) continue;
        seen.add(e.k);
        out.push({ e, i: null, song: map.get(e.k) || null });
      }
    }
    return out;
  }
  return (node.s || []).map((e, i) => ({ e, i, song: map.get(e.k) || null }));
}
const keyName = (k) => stripExt(baseName(k || ""));

// --- Affichage --------------------------------------------------------------------------

function libActive() { return store.get("tab", "songs") === "lib"; }

function setLibStatus() {
  if (!libActive()) return;
  const pls = libLive().filter((n) => n.t === "p").length;
  const base = `${pls} playlist${pls > 1 ? "s" : ""}`;
  $("status").textContent = libError ? `${base} · ${libError}`
    : `${base} · synchronisées avec le Mac${libDirty ? " · envoi…" : ""}`;
}

function artColor(name) {
  let h = 0;
  for (const ch of name || "") h = (h * 31 + ch.charCodeAt(0)) >>> 0;
  return [h % 360, (h >> 9) % 360];
}
function art(name, glyph, size = "") {
  const d = el("div", "lib-art" + (size ? " " + size : ""), glyph);
  const [a, b] = artColor(name);
  d.style.background = `linear-gradient(135deg, hsl(${a} 70% 58%), hsl(${b} 65% 30%))`;
  return d;
}

function renderLibrary() {
  const box = $("lib-view");
  box.replaceChildren();
  $("clear").hidden = !$("q").value;
  const at = libAt ? libGet(libAt) : null;
  if (libAt && !at) libAt = "";
  if (at) {
    const back = el("button", "lib-back", "‹ " + (libParent(at) ? libGet(libParent(at)).n : "Playlists"));
    back.addEventListener("click", () => openLib(libParent(at)));
    box.append(back);
  }
  if (!at || at.t === "f") {
    if (at) box.append(libHeader(at));
    const q = norm($("q").value.trim());
    const kids = q ? libWalk(at ? at.id : "").map(([n]) => n).filter((n) => norm(n.n).includes(q))
      : libChildren(at ? at.id : "");
    const list = el("ul", "list lib-list");
    for (const n of kids) list.append(libNodeRow(n));
    if (!kids.length) {
      list.append(el("li", "empty", q ? "Aucune playlist ne correspond."
        : at ? "Dossier vide." : "Aucune playlist pour l'instant.\nAppuie sur ＋ pour en créer une, ou crée-la sur le Mac."));
    }
    box.append(list);
    if (at) box.append(...libTrackList(at));
  } else {
    box.append(libHeader(at), ...libTrackList(at));
  }
  // ＋ (créer) : à la racine et dans un dossier ; dans une playlist il cacherait les boutons.
  $("lib-new").hidden = !libActive() || !!(at && at.t === "p");
  setLibStatus();
  if (typeof markDownloaded === "function") markDownloaded();
}

function libNodeRow(n) {
  const li = el("li", "row lib-row");
  li.append(art(n.n, n.t === "f" ? "📁" : "♫"));
  const main = el("div", "row-main");
  main.append(el("div", "row-title", n.n));
  const sub = n.t === "f"
    ? `${libChildren(n.id).length} élément${libChildren(n.id).length > 1 ? "s" : ""}`
    : `${(n.s || []).length} morceau${(n.s || []).length > 1 ? "x" : ""}`;
  const where = libParent(n) && libAt !== libParent(n) ? " · " + libPathOf(libGet(libParent(n))) : "";
  main.append(el("div", "row-sub", sub + where));
  li.append(main, el("span", "chev", "›"));
  li.addEventListener("click", () => openLib(n.id));
  return li;
}

function openLib(id) {
  libAt = id || "";
  libSelect = null;
  if ($("q").value) $("q").value = "";
  renderLibrary();
  window.scrollTo(0, 0);
}

function libHeader(n) {
  const head = el("div", "lib-head");
  head.append(art(n.n, n.t === "f" ? "📁" : "♫", "big"));
  const info = el("div", "lib-info");
  info.append(el("h2", "", n.n));
  const tracks = libTracks(n);
  const path = libParent(n) ? libPathOf(libGet(libParent(n))) + " · " : "";
  info.append(el("p", "dim small", `${path}${tracks.length} morceau${tracks.length > 1 ? "x" : ""}`));
  head.append(info);
  const more = el("button", "icon-btn lib-more", "⋯");
  more.addEventListener("click", () => libNodeMenu(n));
  head.append(more);
  const actions = el("div", "lib-actions");
  const playable = tracks.filter((t) => libFile(t.song, t.e));
  const play = el("button", "primary", "▶ Lire");
  play.disabled = !playable.length;
  play.addEventListener("click", () => libPlay(playable, false));
  const shuffle = el("button", "ghost", "🔀 Aléatoire");
  shuffle.disabled = !playable.length;
  shuffle.addEventListener("click", () => libPlay(playable, true));
  const sendAll = el("button", "ghost", "Envoyer");
  sendAll.disabled = !playable.length;
  sendAll.addEventListener("click", (ev) => libShare(playable, n.n, ev.currentTarget));
  actions.append(play, shuffle, sendAll);
  if (n.t === "p" && tracks.length) {
    const sel = el("button", "ghost" + (libSelect ? " on" : ""), libSelect ? "OK" : "Sélect.");
    sel.addEventListener("click", () => { libSelect = libSelect ? null : new Set(); renderLibrary(); });
    actions.append(sel);
  }
  const wrap = el("div", "lib-headwrap");
  wrap.append(head, actions);
  return wrap;
}

function libTrackList(n) {
  const tracks = libTracks(n);
  const q = norm($("q").value.trim());
  const shown = tracks.filter((t) => !q || norm(`${t.song ? t.song.n : keyName(t.e.k)} ${t.song ? styleOf(t.song) : ""}`).includes(q));
  const queueFiles = shown.map((t) => libFile(t.song, t.e)).filter(Boolean);
  const list = el("ul", "list lib-tracks");
  shown.forEach((t) => {
    const file = libFile(t.song, t.e);
    const li = el("li", "row" + (t.song ? "" : " missing"));
    const title = t.song ? t.song.n : keyName(t.e.k);
    if (libSelect && t.i !== null) {
      const box = el("button", "check" + (libSelect.has(t.i) ? " on" : ""), libSelect.has(t.i) ? "✓" : "");
      li.append(box);
      li.addEventListener("click", () => {
        libSelect.has(t.i) ? libSelect.delete(t.i) : libSelect.add(t.i);
        renderLibrary();
      });
    } else if (file) {
      li.append(playButton(file, queueFiles, title));
    } else {
      li.append(el("span", "play", "–"));
    }
    const main = el("div", "row-main");
    main.append(el("div", "row-title", title));
    const sub = !t.song ? (catalogue && (catalogue.songs || []).some((s) => s.k) ? "Introuvable dans le catalogue du Mac"
      : "Mets à jour l'app du Mac (2.4) pour voir ce morceau")
      : file ? `${t.e.b ? "📌 " : ""}${stripExt(file.name)}${styleOf(t.song) ? " · " + styleOf(t.song) : ""}`
        : "Aucun export";
    main.append(el("div", "row-sub", sub));
    li.append(main);
    if (!libSelect && file) li.append(sendButton(file, title));
    if (!libSelect && t.song) li.addEventListener("click", () => openSong(t.song));
    list.append(li);
  });
  const out = [list];
  if (libSelect) {
    const bar = el("div", "lib-selbar");
    const picked = tracks.filter((t) => libSelect.has(t.i));
    const send = el("button", "primary", `Envoyer (${picked.length})`);
    send.disabled = !picked.some((t) => libFile(t.song, t.e));
    send.addEventListener("click", (ev) => libShare(picked.filter((t) => libFile(t.song, t.e)), picked.length === 1 ? (picked[0].song || {}).n : n.n, ev.currentTarget));
    const up = el("button", "ghost", "↑");
    const down = el("button", "ghost", "↓");
    up.disabled = down.disabled = picked.length !== 1;
    up.addEventListener("click", () => libMove(n, picked[0].i, -1));
    down.addEventListener("click", () => libMove(n, picked[0].i, 1));
    const rm = el("button", "ghost danger", "Retirer");
    rm.disabled = !picked.length;
    rm.addEventListener("click", () => {
      n.s = (n.s || []).filter((_, i) => !libSelect.has(i));
      libSelect = new Set();
      libTouch(n);
      renderLibrary();
    });
    bar.append(send, up, down, rm);
    out.push(bar);
  }
  return out;
}

function libMove(n, i, delta) {
  const j = i + delta;
  if (!n.s || j < 0 || j >= n.s.length) return;
  [n.s[i], n.s[j]] = [n.s[j], n.s[i]];
  libSelect = new Set([j]);
  libTouch(n);
  renderLibrary();
}

function libPlay(tracks, shuffle) {
  let files = tracks.map((t) => ({ f: libFile(t.song, t.e), title: t.song.n }));
  if (shuffle) files = files.map((x) => [Math.random(), x]).sort((a, b) => a[0] - b[0]).map((x) => x[1]);
  if (!files.length) return;
  togglePlay(files[0].f, files.map((x) => x.f), files[0].title);
}

// --- Envoi de plusieurs morceaux (liens des exports, jamais les stems) ----------------------

async function libShare(tracks, title, btn) {
  if (!tracks.length) return;
  if (tracks.length === 1) { send(libFile(tracks[0].song, tracks[0].e), tracks[0].song.n, btn); return; }
  if (btn.getAttribute("aria-busy") === "true") return;
  btn.setAttribute("aria-busy", "true");
  const label = btn.textContent;
  btn.textContent = "Liens…";
  try {
    const items = [];
    for (const t of tracks) {
      const file = libFile(t.song, t.e);
      items.push({ title: t.song.n, file, url: await createLink(file) });
    }
    const text = items.map((x) => {
      const stem = stripExt(x.file.name);
      const name = norm(stem).includes(norm(x.title)) ? stem : `${x.title} (${stem})`;
      return `🎵 ${name} \n${x.url} `;
    }).join("\n\n");
    pending = { url: items[0].url, title, file: items[0].file, multi: items, text };
    $("share-name").textContent = `${items.length} morceaux · ${title}`;
    $("share-native").hidden = !navigator.share;
    $("share-dl").hidden = true;
    $("share").hidden = false;
  } catch (e) {
    if (e instanceof NeedLogin) return showLogin();
    toast(e.message);
  } finally {
    btn.removeAttribute("aria-busy");
    btn.textContent = label;
  }
}

// --- Créer, renommer, ranger, supprimer -----------------------------------------------------

function libNodeMenu(n) {
  const choices = [["Renommer", () => {
    const name = prompt("Nouveau nom :", n.n);
    if (name && name.trim() && name.trim() !== n.n) { n.n = name.trim(); libTouch(n); renderLibrary(); }
  }]];
  if (n.t === "f") {
    choices.push(["Nouvelle playlist ici", () => libCreate("p", n.id)], ["Nouveau dossier ici", () => libCreate("f", n.id)]);
  }
  choices.push(["Ranger dans un dossier", () => libPick("Ranger dans…", (target) => {
    if (target === n.id || libIsInside(n.id, target)) { toast("Impossible de ranger un dossier dans lui-même"); return; }
    n.in = target;
    const sib = libChildren(target).filter((x) => x.id !== n.id);
    n.o = sib.length ? Math.max(...sib.map((x) => x.o || 0)) + 1 : 0;
    libTouch(n);
    renderLibrary();
  }, "f")]);
  choices.push(["Supprimer", () => {
    const what = n.t === "f" ? `le dossier « ${n.n} » et son contenu` : `la playlist « ${n.n} »`;
    if (!confirm(`Supprimer ${what} ?\nLes morceaux ne sont pas touchés. Supprimé aussi sur le Mac.`)) return;
    libDelete(n.id);
    openLib(libParent(n));
  }]);
  libSheet(n.n, choices);
}

function libIsInside(nodeId, folderId) {
  const seen = new Set();
  let cur = folderId;
  while (cur && !seen.has(cur)) {
    if (cur === nodeId) return true;
    seen.add(cur);
    const n = libGet(cur);
    cur = n ? libParent(n) : "";
  }
  return false;
}

function libCreate(t, parent = libAt && (libGet(libAt) || {}).t === "f" ? libAt : "", songs = null) {
  const name = prompt(t === "f" ? "Nom du dossier :" : "Nom de la playlist :", "");
  if (!name || !name.trim()) return null;
  const n = libAdd(t, name, parent);
  if (songs) libAddSongs(n.id, songs);
  if (libActive()) openLib(t === "p" && !songs ? n.id : libAt);
  return n;
}

/** Feuille de choix (style iOS) réutilisant l'overlay #lib-sheet. */
function libSheet(title, choices) {
  const box = $("lib-sheet-box");
  box.replaceChildren(el("p", "share-title", title));
  for (const [label, fn] of choices) {
    const b = el("button", "ghost big", label);
    b.addEventListener("click", () => { $("lib-sheet").hidden = true; fn(); });
    box.append(b);
  }
  const cancel = el("button", "ghost", "Annuler");
  cancel.addEventListener("click", () => { $("lib-sheet").hidden = true; });
  box.append(cancel);
  $("lib-sheet").hidden = false;
}

/** Choisir une playlist (ou un dossier, kind "f") dans l'arbre. */
function libPick(title, fn, kind = "p", songs = null) {
  const choices = [];
  if (kind === "f") choices.push(["(Aucun dossier)", () => fn("")]);
  for (const [n, depth] of libWalk()) {
    if (kind === "f" && n.t !== "f") continue;
    const label = "   ".repeat(depth) + (n.t === "f" ? "📁 " : "♫ ") + n.n;
    if (n.t === kind) choices.push([label, () => fn(n.id)]);
  }
  if (kind === "p") choices.push(["＋ Nouvelle playlist…", () => libCreate("p", "", songs)]);
  libSheet(title, choices);
}

/** Depuis la fiche d'un morceau : « ＋ Playlist ». */
function addCurrentSongToPlaylist() {
  if (!current || !current.k) { toast("Mets à jour l'app du Mac (2.4) pour ranger ce morceau"); return; }
  const songs = [{ k: current.k }];
  libPick("Ajouter à une playlist", (id) => {
    const n = libAddSongs(id, songs);
    toast(n ? `Ajouté à « ${libGet(id).n} »` : `Déjà dans « ${libGet(id).n} »`);
  }, "p", songs);
}

function wireLibrary() {
  $("lib-new").addEventListener("click", () => libSheet("Créer", [
    ["♫ Nouvelle playlist", () => libCreate("p")],
    ["📁 Nouveau dossier", () => libCreate("f")],
  ]));
  $("lib-sheet").addEventListener("click", (e) => { if (e.target.id === "lib-sheet") $("lib-sheet").hidden = true; });
  $("song-add").addEventListener("click", addCurrentSongToPlaylist);
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) { if (libDirty) syncLibrary(); return; }
    if (store.get("tok") && (libDirty || libNow() - store.get("libAt", 0) > 30_000)) syncLibrary();
  });
  window.addEventListener("online", () => { if (libDirty) syncLibrary(); });
}

wireLibrary();
if (store.get("tok")) syncLibrary();
