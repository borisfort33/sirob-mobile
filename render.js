/* SIRØB Mobile — « ⏺ Rendre sur le Mac » : demander au Mac allumé de rendre un .flp avec FL Studio.
 *
 * L'iPhone écrit OneDrive « SIROB Browser/rendus-demandes.json » (lui seul l'écrit) ;
 * l'app du Mac, si elle est ouverte, lance le rendu (MP3 320 ou WAV 24 bits) et écrit l'avancement
 * dans « SIROB Browser/rendus-etat.json » (lui seul l'écrit). Voir flbrowser/render_requests.py.
 * Le Mac ne ferme jamais FL Studio pour une demande iPhone : si FL est ouvert, elle attend.
 */

const REQ_PATH = "SIROB Browser/rendus-demandes.json";
const STATE_PATH = "SIROB Browser/rendus-etat.json";
const REQ_KEEP_MS = 7 * 24 * 3600 * 1000;
const FINAL = new Set(["fait", "erreur", "annulé", "expirée"]);
let renderState = { jobs: {} };
let renderPoll = null;
let renderFmt = store.get("renderFmt", "mp3");

async function readJson(path) {
  try {
    const item = await graph(`/me/drive/root:/${encPath(path)}`);
    const resp = await fetch(item["@microsoft.graph.downloadUrl"], { cache: "no-store" });
    if (!resp.ok) throw new Error(`Téléchargement refusé (${resp.status})`);
    return { eTag: item.eTag, data: await resp.json().catch(() => ({})) };
  } catch (e) {
    if (e.status === 404) return { eTag: null, data: {} };
    throw e;
  }
}

async function sendRenderRequest(req) {
  for (let attempt = 0; attempt < 4; attempt++) {
    const { eTag, data } = await readJson(REQ_PATH);
    const now = Date.now();
    const reqs = (data.requests || []).filter((r) => now - (r.c || 0) < REQ_KEEP_MS);
    reqs.push(req);
    const body = JSON.stringify({ version: 1, requests: reqs }, null, 1);
    const path = `/me/drive/root:/${encPath(REQ_PATH)}:/content` + (eTag ? "" : "?@microsoft.graph.conflictBehavior=fail");
    try {
      await graph(path, { method: "PUT", body, headers: eTag ? { "If-Match": eTag } : {} });
      store.set("renderMine", [...store.get("renderMine", []).filter((r) => now - r.c < REQ_KEEP_MS), req]);
      return;
    } catch (e) {
      if (e.status === 412 || e.status === 409) continue;
      throw e;
    }
  }
  throw new Error("OneDrive occupé, réessaie dans un instant");
}

function macSeen() {
  if (!renderState.mac) return "Le Mac n'a encore jamais répondu (app 3.3 ou plus récente nécessaire).";
  const min = (Date.now() - renderState.mac) / 60000;
  return min < 8 ? "Mac allumé, app ouverte." : `Mac vu pour la dernière fois ${ago(renderState.mac / 1000)}.`;
}

async function refreshRenderState() {
  try {
    renderState = (await readJson(STATE_PATH)).data || {};
    renderState.jobs = renderState.jobs || {};
  } catch (e) {
    if (e instanceof NeedLogin) return;
  }
  fillRenders();
}

const ST_TEXT = { "attente": "⏳ En attente du Mac", "en cours": "⏺ Rendu en cours sur le Mac",
                  "fait": "✓ Rendu terminé", "erreur": "⚠️ Erreur", "annulé": "Annulé", "expirée": "Expirée" };

function fillRenders() {
  const list = $("renders");
  if (!list) return;
  const k = current && current.k;
  const mine = store.get("renderMine", []).filter((r) => r.s === k).sort((a, b) => b.c - a.c).slice(0, 5);
  list.replaceChildren();
  let open = false;
  for (const r of mine) {
    const job = renderState.jobs[r.id] || { st: "attente", msg: "Envoyé : le Mac le lancera s'il est allumé avec l'app ouverte" };
    if (!FINAL.has(job.st)) open = true;
    if (job.st === "fait" && job.out) {
      const f = fileObj([job.out, Math.round((job.u || Date.now()) / 1000)]);
      list.append(fileRow(f, current.n, [f]));
      continue;
    }
    const li = el("li", "row");
    const main = el("div", "row-main");
    main.append(el("div", "row-title", `${ST_TEXT[job.st] || job.st} · ${stripExt(baseName(r.k))} (${r.fmt === "wav" ? "WAV 24 bits" : "MP3 320"})`),
                el("div", "row-sub", job.msg || ""));
    li.append(el("span", "play", job.st === "erreur" ? "!" : "⏺"), main);
    list.append(li);
  }
  $("renders-title").hidden = !mine.length;
  clearInterval(renderPoll);
  renderPoll = open && !$("song").hidden ? setInterval(refreshRenderState, 10_000) : null;
  if (!open && mine.some((r) => (renderState.jobs[r.id] || {}).st === "fait")) loadCatalogue();
}

function openRender() {
  if (!current) return;
  const choices = current.r || [];
  if (!choices.length) { toast("Mets à jour l'app du Mac (3.3) pour rendre depuis l'iPhone"); return; }
  const box = $("lib-sheet-box");
  box.replaceChildren(el("p", "share-title", "Rendre sur le Mac"), el("p", "dim small", macSeen()));
  const select = el("select", "render-select");
  for (const [label, key] of choices) {
    const o = el("option", "", label);
    o.value = key;
    select.append(o);
  }
  const chips = el("div", "chips");
  for (const [k, label] of [["mp3", "MP3 320 kbps"], ["wav", "WAV 24 bits"]]) {
    const b = el("button", k === renderFmt ? "on" : "", label);
    b.addEventListener("click", () => {
      renderFmt = k;
      store.set("renderFmt", k);
      chips.querySelectorAll("button").forEach((x) => x.classList.toggle("on", x === b));
    });
    chips.append(b);
  }
  const go = el("button", "primary big", "⏺ Lancer le rendu");
  go.addEventListener("click", async () => {
    go.disabled = true;
    const req = { id: (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random()),
                  k: select.value, s: current.k, n: current.n, fmt: renderFmt, c: Date.now() };
    try {
      await sendRenderRequest(req);
      $("lib-sheet").hidden = true;
      toast("Demande envoyée au Mac");
      fillRenders();
    } catch (e) {
      if (e instanceof NeedLogin) return showLogin();
      toast("Envoi impossible : " + e.message);
      go.disabled = false;
    }
  });
  const cancel = el("button", "ghost", "Annuler");
  cancel.addEventListener("click", () => { $("lib-sheet").hidden = true; });
  box.append(el("p", "small", "Projet :"), select, el("p", "small", "Format :"), chips,
             el("p", "dim small", "Rangé dans Export (Mastering pour une session de mastering). Un fichier du même nom est remplacé."),
             go, cancel);
  $("lib-sheet").hidden = false;
}

function wireRender() {
  $("song-render").addEventListener("click", openRender);
  // La fiche d'un morceau s'ouvre : on montre ses rendus demandés et leur état.
  window.onSongOpened = () => {
    $("song-render").hidden = !(current && current.r && current.r.length);
    fillRenders();
    if (store.get("renderMine", []).some((r) => r.s === (current && current.k))) refreshRenderState();
  };
  new MutationObserver(() => {
    if ($("song").hidden) { clearInterval(renderPoll); renderPoll = null; }
  }).observe($("song"), { attributes: true, attributeFilter: ["hidden"] });
}

wireRender();
if (store.get("tok")) refreshRenderState();
