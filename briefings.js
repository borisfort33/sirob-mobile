/* SIRØB Mobile : briefings Ghost Production, classés par artiste / client, synchronisés avec le Mac.
 *
 * Fichier partagé : OneDrive « SIROB Browser/briefings.json » = {"version", "updated", "briefings": [...]}
 * (même format que flbrowser/briefings.py). Un briefing = {"id", "x" (texte, 1re ligne = titre),
 * "f" (artiste / client), "c", "u", "p" (épinglé), "k" (fait : barré)}. Fusion comme les idées
 * (ideas.js : mergeIdeas, la modification la plus récente gagne), écriture avec If-Match.
 *
 * Liste : un groupe par personne (repliable), rond à gauche pour marquer « fait » d'un appui.
 * Les liens YouTube / SoundCloud / Spotify / Deezer se lisent ici ; « ↗ App » ouvre Spotify.
 */
"use strict";

const BRIEFS_PATH = "SIROB Browser/briefings.json";
const NO_NAME = "Sans nom";

let briefs = store.get("briefs", []);
let briefsETag = store.get("briefsETag", null);
let briefsDirty = store.get("briefsDirty", false);
let currentBrief = null;
let briefsSyncing = null;
let briefSaveTimer;
let briefsError = "";
let briefsClosed = new Set(store.get("briefsClosed", []));   // personnes repliées

const personOf = (it) => (it.f || "").trim();
const personKey = (name) => (name || "").trim().toLocaleLowerCase("fr");
const liveBriefs = () => briefs.filter((it) => !it.d);
const getBrief = (id) => briefs.find((it) => it.id === id && !it.d);
// À faire d'abord (épinglés en tête), faits en bas, puis les plus récents.
const briefOrder = (a, b) => (a.k ? 1 : 0) - (b.k ? 1 : 0) || (b.p ? 1 : 0) - (a.p ? 1 : 0) || (b.u || 0) - (a.u || 0);

function briefActive() { return store.get("tab", "songs") === "brief"; }

function saveBriefsLocal() {
  store.set("briefs", briefs);
  store.set("briefsDirty", briefsDirty);
}

function touchBrief(id, fields) {
  const it = getBrief(id);
  if (!it) return null;
  if (Object.entries(fields).every(([k, v]) => it[k] === v)) return it;
  Object.assign(it, fields, { u: Math.max(nowMs(), (it.u || 0) + 1) });
  briefsDirty = true;
  saveBriefsLocal();
  scheduleBriefSync();
  return it;
}

// --- Synchro OneDrive -----------------------------------------------------------------

async function fetchBriefs() {
  try {
    const item = await graph(`/me/drive/root:/${encPath(BRIEFS_PATH)}`);
    const url = item["@microsoft.graph.downloadUrl"];
    if (!url) throw new Error("OneDrive n'a pas donné le fichier des briefings");
    const resp = await fetch(url, { cache: "no-store" });
    if (!resp.ok) throw new Error(`Téléchargement des briefings refusé (${resp.status})`);
    const data = await resp.json().catch(() => ({}));
    return { eTag: item.eTag, items: Array.isArray(data) ? data : (data.briefings || []) };
  } catch (e) {
    if (e.status === 404) return { eTag: null, items: [] };
    throw e;
  }
}

async function putBriefs(items, eTag) {
  const body = JSON.stringify({ version: 1, updated: nowMs(), briefings: items }, null, 1);
  const headers = eTag ? { "If-Match": eTag } : {};
  const path = `/me/drive/root:/${encPath(BRIEFS_PATH)}:/content` + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
  const item = await graph(path, { method: "PUT", body, headers });
  return item.eTag;
}

function scheduleBriefSync(delay = 1500) {
  clearTimeout(briefSaveTimer);
  briefSaveTimer = setTimeout(syncBriefs, delay);
}

async function syncBriefs() {
  if (!store.get("tok")) return;
  if (briefsSyncing) { await briefsSyncing; if (!briefsDirty) return; }
  briefsSyncing = (async () => {
    for (let attempt = 0; attempt < 4; attempt++) {
      let remote;
      try { remote = await fetchBriefs(); } catch (e) {
        if (e instanceof NeedLogin) return;
        briefsError = `Synchro impossible : ${e.message}${briefsDirty ? " · modifications gardées sur l'iPhone" : ""}`;
        setBriefsStatus();
        return;
      }
      const merged = mergeIdeas(briefs, remote.items);
      const changedHere = !sameIdeas(merged, briefs);
      briefs = merged;
      if (!sameIdeas(merged, remote.items)) {
        try {
          briefsETag = await putBriefs(merged, remote.eTag);
        } catch (e) {
          if (e.status === 412 || e.status === 409) continue;   // le Mac vient d'écrire : on refusionne
          briefsError = `Synchro impossible (${e.message}) · modifications gardées sur l'iPhone`;
          saveBriefsLocal();
          setBriefsStatus();
          return;
        }
      } else {
        briefsETag = remote.eTag;
      }
      briefsDirty = false;
      briefsError = "";
      saveBriefsLocal();
      store.set("briefsETag", briefsETag);
      store.set("briefsAt", nowMs());
      if (changedHere) refreshBriefsUi();
      setBriefsStatus();
      return;
    }
  })();
  try { await briefsSyncing; } finally { briefsSyncing = null; }
}

// --- Liste (onglet Briefs) ----------------------------------------------------------------

function setBriefsStatus() {
  if (!briefActive()) return;
  const all = liveBriefs();
  const todo = all.filter((it) => !it.k).length;
  const base = `${todo} briefing${todo > 1 ? "s" : ""} à faire sur ${all.length}`;
  $("status").textContent = briefsError ? `${base} · ${briefsError}`
    : `${base} · synchronisés avec le Mac${briefsDirty ? " · envoi…" : ""}`;
}

function briefGroups(list) {
  const groups = new Map();
  for (const it of list) {
    const key = personKey(personOf(it));
    if (!groups.has(key)) groups.set(key, { name: personOf(it), items: [] });
    groups.get(key).items.push(it);
  }
  return [...groups.entries()]
    .sort(([a], [b]) => (a === "") - (b === "") || a.localeCompare(b, "fr"))
    .map(([key, g]) => ({ key, ...g, items: g.items.sort(briefOrder) }));
}

function renderBriefs() {
  const box = $("brief-list");
  box.replaceChildren();
  const q = norm($("q").value.trim());
  $("clear").hidden = !$("q").value;
  const list = liveBriefs().filter((it) => !q || norm(`${it.x || ""} ${it.f || ""}`).includes(q));
  for (const g of briefGroups(list)) {
    const todo = g.items.filter((it) => !it.k).length;
    const closed = !q && briefsClosed.has(g.key);
    const head = el("div", "brief-person" + (closed ? " closed" : ""));
    const name = el("button", "brief-name");
    name.append(el("span", "brief-chev", closed ? "›" : "⌄"), el("span", "", g.name || NO_NAME));
    name.append(el("span", "brief-count", todo ? `${todo} à faire` : "tout fait"));
    name.addEventListener("click", () => {
      if (briefsClosed.has(g.key)) briefsClosed.delete(g.key); else briefsClosed.add(g.key);
      store.set("briefsClosed", [...briefsClosed]);
      renderBriefs();
    });
    const add = el("button", "brief-add", "＋");
    add.setAttribute("aria-label", `Nouveau briefing pour ${g.name || NO_NAME}`);
    add.addEventListener("click", () => newBrief(g.name));
    head.append(name, add);
    box.append(head);
    if (closed) continue;
    const card = el("ul", "notes-card");
    for (const it of g.items) {
      const li = el("li", "note brief" + (it.k ? " done" : ""));
      const check = el("button", "brief-check" + (it.k ? " on" : ""), it.k ? "✓" : "");
      check.setAttribute("aria-label", it.k ? "Remettre à faire" : "Marquer comme fait");
      check.addEventListener("click", (ev) => {
        ev.stopPropagation();
        touchBrief(it.id, { k: it.k ? 0 : 1 });
        renderBriefs();
      });
      const body = el("div", "brief-body");
      body.append(el("div", "note-title", (it.p ? "📌 " : "") + ideaTitle(it).replace(/^Nouvelle idée$/, "Nouveau briefing")));
      const sub = el("div", "note-sub");
      sub.append(el("span", "note-date", shortDate(it.u || it.c || 0)));
      const links = ideaLinks(it);
      sub.append(document.createTextNode(" " + (ideaPreview(it) || (links.length ? linkKind(links[0])[1] : "Aucun texte supplémentaire"))));
      body.append(sub);
      li.append(check, body);
      li.addEventListener("click", () => openBrief(it.id));
      card.append(li);
    }
    box.append(card);
  }
  if (!list.length) {
    box.append(el("p", "empty center", q ? "Aucun briefing ne correspond."
      : "Aucun briefing pour l'instant.\nAppuie sur ＋ pour noter le brief d'un artiste ou d'un client."));
  }
  setBriefsStatus();
}

function refreshBriefsUi() {
  if (briefActive()) renderBriefs();
  if (currentBrief && !$("brief").hidden) {
    const it = getBrief(currentBrief);
    if (!it) { history.back(); toast("Briefing supprimé sur le Mac"); return; }
    if (document.activeElement !== $("brief-text") && $("brief-text").value !== (it.x || "")) $("brief-text").value = it.x || "";
    if (document.activeElement !== $("brief-person") && $("brief-person").value !== (it.f || "")) $("brief-person").value = it.f || "";
    fillBriefMeta(it);
  }
}

// --- Briefing ouvert ----------------------------------------------------------------------

function openBrief(id, focus = false) {
  const it = getBrief(id);
  if (!it) return;
  currentBrief = id;
  $("brief-text").value = it.x || "";
  $("brief-person").value = it.f || "";
  const names = new Set(liveBriefs().map(personOf).filter(Boolean));
  $("brief-people").replaceChildren(...[...names].sort().map((n) => Object.assign(document.createElement("option"), { value: n })));
  closeEmbed();
  fillBriefMeta(it);
  show("brief");
  $("brief").scrollTop = 0;
  history.pushState({ view: "brief" }, "");
  briefGrow();
  if (focus) setTimeout(() => $(it.f ? "brief-text" : "brief-person").focus(), 50);
}

function fillBriefMeta(it) {
  $("brief-date").textContent = longDate(it.u || it.c || nowMs());
  $("brief-pin").classList.toggle("on", !!it.p);
  $("brief-done").classList.toggle("on", !!it.k);
  $("brief-done").textContent = it.k ? "✓ Fait" : "Fait ?";
  $("brief-text").classList.toggle("done", !!it.k);
  const box = $("brief-links");
  fillLinks(box, ideaLinks(it), () => $("brief-text").value);
  $("brief-links-title").hidden = !box.childElementCount;
}

function flushBrief() {
  clearTimeout(flushBrief.t);
  if (!currentBrief) return;
  touchBrief(currentBrief, { x: $("brief-text").value, f: $("brief-person").value.trim() });
}

function newBrief(person = "") {
  const t = nowMs();
  const it = { id: newId(), x: "", f: person || "", c: t, u: t, p: 0, k: 0 };
  briefs.push(it);
  briefsDirty = true;
  saveBriefsLocal();
  openBrief(it.id, true);
}

function closeBrief() {
  if (!currentBrief) return;
  flushBrief();
  const it = getBrief(currentBrief);
  // Briefing laissé vide : il disparaît (comme Notes).
  if (it && !(it.x || "").trim()) deleteBrief(it.id, true);
  currentBrief = null;
  closeEmbed();
  if (briefActive()) renderBriefs();
  syncBriefs();
}

function deleteBrief(id, silent = false) {
  const i = briefs.findIndex((x) => x.id === id);
  if (i < 0) return;
  briefs[i] = { id, u: Math.max(nowMs(), (briefs[i].u || 0) + 1), d: 1 };
  briefsDirty = true;
  saveBriefsLocal();
  if (!silent) scheduleBriefSync(10);
}

function briefGrow() {
  const t = $("brief-text");
  t.style.height = "auto";
  t.style.height = Math.max(240, t.scrollHeight + 4) + "px";
}

// --- Branchements ------------------------------------------------------------------------

function wireBriefs() {
  $("brief-new").addEventListener("click", () => newBrief(""));
  $("brief-back").addEventListener("click", () => history.back());
  $("brief-text").addEventListener("input", () => { briefGrow(); clearTimeout(flushBrief.t); flushBrief.t = setTimeout(flushBrief, 800); });
  $("brief-text").addEventListener("blur", () => { flushBrief(); const it = getBrief(currentBrief); if (it) fillBriefMeta(it); });
  $("brief-person").addEventListener("input", () => { clearTimeout(flushBrief.t); flushBrief.t = setTimeout(flushBrief, 800); });
  $("brief-person").addEventListener("change", flushBrief);
  $("brief-pin").addEventListener("click", () => { flushBrief(); const it = getBrief(currentBrief); if (it) fillBriefMeta(touchBrief(it.id, { p: it.p ? 0 : 1 })); });
  $("brief-done").addEventListener("click", () => {
    flushBrief();
    const it = getBrief(currentBrief);
    if (it) { fillBriefMeta(touchBrief(it.id, { k: it.k ? 0 : 1 })); toast(it.k ? "Briefing fait : barré" : "Briefing remis à faire"); }
  });
  $("brief-delete").addEventListener("click", () => {
    const it = getBrief(currentBrief);
    if (!it || ((it.x || "").trim() && !confirm(`Supprimer « ${ideaTitle(it)} » ?\nIl disparaît aussi du Mac.`))) return;
    deleteBrief(it.id);
    currentBrief = null;
    history.back();
  });
  $("brief-paste").addEventListener("click", async () => {
    let text = "";
    try { text = await navigator.clipboard.readText(); } catch { /* refusé */ }
    if (!text) { toast("Presse-papiers vide ou refusé : appui long dans le texte › Coller"); return; }
    const t = $("brief-text");
    t.value = (t.value.trimEnd() ? t.value.trimEnd() + "\n" : "") + text.trim();
    briefGrow();
    flushBrief();
    fillBriefMeta(getBrief(currentBrief));
  });
  window.addEventListener("popstate", () => { if (currentBrief && $("brief").hidden) closeBrief(); });
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) { flushBrief(); if (briefsDirty) syncBriefs(); return; }
    if (store.get("tok") && (briefsDirty || nowMs() - store.get("briefsAt", 0) > 30_000)) syncBriefs();
  });
  window.addEventListener("online", () => { if (briefsDirty) syncBriefs(); });
}

wireBriefs();
if (store.get("tab", "songs") !== "songs") setTab(store.get("tab"));
if (store.get("tok")) syncBriefs();
