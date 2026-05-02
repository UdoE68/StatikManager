import "./styles/main.css";

import type {
  BrowseEntry,
  BrowseResponse,
  ErrorResponse,
  FileKind,
  FileMetaResponse,
  ProjectsResponse,
  SessionResponse,
} from "./types/api";

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) throw new Error("#app fehlt");

let browsePfadAktuell = "";
let ausgewaehlteDateiRel: string | null = null;

app.innerHTML = `
  <main class="shell app-shell">
    <header class="app-header">
      <div class="app-header-brand">
        <h1 class="app-title">StatikManager Web</h1>
        <span class="app-header-tag muted">Metadaten &amp; Vorschau</span>
      </div>
      <section class="app-header-project" aria-labelledby="projekt-label">
        <h2 id="projekt-label" class="h2 app-header-projekt-label">Projekt</h2>
        <div class="app-header-row">
          <label class="label label--inline" for="pfad-input">Pfad</label>
          <input id="pfad-input" class="input input--compact" type="text" autocomplete="off"
            placeholder="z. B. C:\\Projekte\\MeinStatikProjekt" spellcheck="false" />
          <button id="btn-oeffnen" type="button" class="btn btn--compact">Öffnen</button>
          <button id="btn-ordner" type="button" class="btn btn-secondary btn--compact">Ordner …</button>
        </div>
        <div class="app-header-row app-header-row-projekte">
          <label class="label label--inline" for="projekt-select">Projekte</label>
          <select id="projekt-select" class="select select--compact" aria-label="Gespeicherte Projekte">
            <option value="">— Projekt wählen —</option>
          </select>
          <button id="btn-projekt-speichern" type="button" class="btn btn-secondary btn--compact" title="Aktuelles Projekt in die Liste speichern">Speichern</button>
          <button id="btn-projekt-entfernen" type="button" class="btn btn-secondary btn--compact" title="Ausgewähltes Projekt aus der Liste entfernen">Entfernen</button>
        </div>
        <p id="ordner-hinweis" class="hinweis hinweis--compact" role="status" hidden></p>
        <p id="fehler" class="fehler fehler--compact" role="alert" hidden></p>
        <div class="app-root-row">
          <span class="aktuell-label aktuell-label--inline">Root</span>
          <p id="root-anzeige" class="root-anzeige root-anzeige--compact">—</p>
        </div>
      </section>
    </header>

    <div class="layout-haupt app-workspace">
      <aside class="layout-spalte-links app-pane-tree">
        <section class="panel panel-inner panel--chrome" id="browse-panel" aria-labelledby="browse-label">
          <h2 id="browse-label" class="h2 h2--pane">Ordner</h2>
          <div class="browse-toolbar browse-toolbar--compact">
            <button id="btn-hoch" type="button" class="btn btn-secondary btn--compact" disabled>Nach oben</button>
            <span id="browse-pfad-anzeige" class="browse-pfad"></span>
          </div>
          <p id="browse-fehler" class="fehler fehler--compact" role="alert" hidden></p>
          <ul id="browse-liste" class="browse-liste browse-tree-root browse-tree--compact" role="tree" aria-label="Projektordner"></ul>
        </section>
      </aside>

      <div class="layout-spalte-rechts layout-spalte-rechts-stapel app-pane-detail">
        <section class="panel panel-inner panel--chrome panel--meta" aria-labelledby="meta-label">
          <h2 id="meta-label" class="h2 h2--pane">Datei-Info</h2>
          <p id="meta-fehler" class="fehler fehler--compact" role="alert" hidden></p>
          <div id="meta-inhalt" class="meta-inhalt">
            <p class="meta-platzhalter">Keine Datei ausgewählt.</p>
          </div>
        </section>

        <section class="panel panel-inner panel--chrome panel--preview" aria-labelledby="vorschau-label">
          <h2 id="vorschau-label" class="h2 h2--pane">Vorschau</h2>
          <p id="vorschau-fehler" class="fehler fehler--compact" role="alert" hidden></p>
          <div id="vorschau-inhalt" class="vorschau-inhalt">
            <p class="vorschau-platzhalter">Keine Datei ausgewählt.</p>
          </div>
        </section>
      </div>
    </div>

    <footer class="app-statusbar">
      <p id="health-status" class="status status--bar">API …</p>
    </footer>
  </main>
`;

function el<K extends keyof HTMLElementTagNameMap>(
  sel: string,
  tag: K
): HTMLElementTagNameMap[K] {
  const node = document.querySelector(sel);
  if (!node || node.tagName.toLowerCase() !== tag)
    throw new Error(`Element fehlt oder falscher Typ: ${sel}`);
  return node as HTMLElementTagNameMap[K];
}

function kindLabel(kind: FileKind): string {
  const map: Record<FileKind, string> = {
    pdf: "PDF",
    image: "Bild",
    html: "HTML",
    json: "JSON",
    text: "Text",
    other: "Sonstiges",
  };
  return map[kind] ?? String(kind);
}

async function checkHealth(): Promise<void> {
  const statusEl = el("#health-status", "p");
  try {
    const res = await fetch("/api/health");
    const data = (await res.json()) as { ok?: boolean };
    if (res.ok && data.ok === true) {
      statusEl.textContent = "API: /api/health → { ok: true }";
      statusEl.classList.add("ok");
      return;
    }
    statusEl.textContent = `Health: unerwartet (${res.status})`;
    statusEl.classList.add("err");
  } catch {
    statusEl.textContent =
      "API nicht erreichbar. Backend starten (dotnet run) oder Vite-Proxy prüfen.";
    statusEl.classList.add("err");
  }
}

function setFehler(text: string | null): void {
  const fehler = el("#fehler", "p");
  if (text === null || text === "") {
    fehler.hidden = true;
    fehler.textContent = "";
    return;
  }
  fehler.hidden = false;
  fehler.textContent = text;
}

function setBrowseFehler(text: string | null): void {
  const fehler = el("#browse-fehler", "p");
  if (text === null || text === "") {
    fehler.hidden = true;
    fehler.textContent = "";
    return;
  }
  fehler.hidden = false;
  fehler.textContent = text;
}

function setMetaFehler(text: string | null): void {
  const fehler = el("#meta-fehler", "p");
  if (text === null || text === "") {
    fehler.hidden = true;
    fehler.textContent = "";
    return;
  }
  fehler.hidden = false;
  fehler.textContent = text;
}

function setVorschauFehler(text: string | null): void {
  const fehler = el("#vorschau-fehler", "p");
  if (text === null || text === "") {
    fehler.hidden = true;
    fehler.textContent = "";
    return;
  }
  fehler.hidden = false;
  fehler.textContent = text;
}

function vorschauLeeren(): void {
  setVorschauFehler(null);
  const w = el("#vorschau-inhalt", "div");
  w.innerHTML = '<p class="vorschau-platzhalter">Keine Datei ausgewählt.</p>';
}

function previewStreamUrl(relPath: string): string {
  return `/api/preview/stream?path=${encodeURIComponent(relPath)}`;
}

async function leseFehlerAusAntwort(res: Response): Promise<string> {
  try {
    const j = (await res.json()) as Partial<ErrorResponse>;
    if (typeof j.error === "string") return j.error;
  } catch {
    /* Antwort war kein JSON */
  }
  return `Fehler (${res.status})`;
}

async function zeigeVorschau(meta: FileMetaResponse): Promise<void> {
  const wrap = el("#vorschau-inhalt", "div");
  setVorschauFehler(null);
  wrap.innerHTML = "";

  if (meta.kind === "other") {
    const p = document.createElement("p");
    p.className = "vorschau-keine muted-small";
    p.textContent =
      "Für diesen Dateityp gibt es keine Vorschau — es werden nur die Metadaten angezeigt.";
    wrap.appendChild(p);
    return;
  }

  const url = previewStreamUrl(meta.relativePath);

  try {
    switch (meta.kind) {
      case "pdf": {
        const iframe = document.createElement("iframe");
        iframe.className = "vorschau-iframe vorschau-pdf";
        iframe.title = `PDF: ${meta.name}`;
        iframe.src = url;
        wrap.appendChild(iframe);
        break;
      }
      case "image": {
        const img = document.createElement("img");
        img.className = "vorschau-img";
        img.alt = meta.name;
        img.src = url;
        img.addEventListener("error", () => {
          setVorschauFehler("Bild konnte nicht geladen werden.");
        });
        wrap.appendChild(img);
        break;
      }
      case "html": {
        const iframe = document.createElement("iframe");
        iframe.className = "vorschau-iframe vorschau-html";
        iframe.title = `HTML: ${meta.name}`;
        iframe.setAttribute("sandbox", "allow-same-origin");
        iframe.src = url;
        wrap.appendChild(iframe);
        break;
      }
      case "json":
      case "text": {
        const res = await fetch(url);
        if (!res.ok) {
          setVorschauFehler(await leseFehlerAusAntwort(res));
          return;
        }
        const rawText = await res.text();
        if (meta.kind === "json") {
          try {
            const parsed: unknown = JSON.parse(rawText);
            const pre = document.createElement("pre");
            pre.className = "vorschau-pre";
            pre.textContent = JSON.stringify(parsed, null, 2);
            wrap.appendChild(pre);
          } catch {
            setVorschauFehler("Inhalt ist kein gültiges JSON.");
          }
        } else {
          const pre = document.createElement("pre");
          pre.className = "vorschau-pre";
          pre.textContent = rawText;
          wrap.appendChild(pre);
        }
        break;
      }
      default:
        break;
    }
  } catch {
    setVorschauFehler("Vorschau konnte nicht geladen werden.");
  }
}

function setRootAnzeige(rootPath: string | null): void {
  const rootAnzeige = el("#root-anzeige", "p");
  rootAnzeige.textContent =
    rootPath === null || rootPath === "" ? "— (nicht gesetzt)" : rootPath;
}

function metaLeeren(): void {
  ausgewaehlteDateiRel = null;
  setMetaFehler(null);
  const inh = el("#meta-inhalt", "div");
  inh.innerHTML = '<p class="meta-platzhalter">Keine Datei ausgewählt.</p>';
  vorschauLeeren();
}

function parentRelativePath(rel: string): string {
  const t = rel.replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
  if (!t) return "";
  const parts = t.split("/").filter(Boolean);
  parts.pop();
  return parts.join("/");
}

function formatBytes(n: number | null): string {
  if (n === null || n === undefined) return "—";
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

function formatZeit(iso: string): string {
  try {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return iso;
    return d.toLocaleString(undefined, {
      dateStyle: "short",
      timeStyle: "short",
    });
  } catch {
    return iso;
  }
}

function aktualisiereBrowseToolbar(): void {
  const btn = el("#btn-hoch", "button");
  const pfadAnzeige = el("#browse-pfad-anzeige", "span");
  btn.disabled = browsePfadAktuell === "";
  pfadAnzeige.textContent =
    browsePfadAktuell === ""
      ? "(Projekt-Root)"
      : browsePfadAktuell.replace(/\//g, " \\ ");
}

function normalisiereRelPfad(rel: string): string {
  return rel.replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
}

function findeDirektesOrdnerKind(
  ul: HTMLUListElement,
  relPath: string
): HTMLLIElement | null {
  for (const child of ul.children) {
    if (child.tagName !== "LI") continue;
    const li = child as HTMLLIElement;
    if (!li.classList.contains("browse-dir")) continue;
    if (li.dataset.relativePath === relPath) return li;
  }
  return null;
}

function erzeugeDateiKnoten(e: BrowseEntry): HTMLLIElement {
  const li = document.createElement("li");
  li.className =
    "browse-tree-node browse-zeile browse-file browse-tree-item";
  li.dataset.relativePath = e.relativePath;
  li.dataset.isDirectory = "false";
  li.tabIndex = 0;
  li.setAttribute("role", "treeitem");

  const row = document.createElement("div");
  row.className = "browse-tree-row browse-tree-row-leaf";
  const spacer = document.createElement("span");
  spacer.className = "browse-tree-leaf-spacer";
  spacer.setAttribute("aria-hidden", "true");
  const name = document.createElement("span");
  name.className = "browse-name";
  name.textContent = e.name;
  const meta = document.createElement("span");
  meta.className = "browse-meta";
  meta.textContent = `${formatZeit(e.modifiedUtc)} · ${formatBytes(e.sizeBytes)}`;
  row.append(spacer, name, meta);
  li.appendChild(row);
  return li;
}

function erzeugeOrdnerKnoten(e: BrowseEntry): HTMLLIElement {
  const li = document.createElement("li");
  li.className =
    "browse-tree-node browse-zeile browse-dir browse-tree-item";
  li.dataset.relativePath = e.relativePath;
  li.dataset.isDirectory = "true";
  li.dataset.loaded = "false";
  li.tabIndex = 0;
  li.setAttribute("role", "treeitem");
  li.setAttribute("aria-expanded", "false");

  const row = document.createElement("div");
  row.className = "browse-tree-row";

  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "browse-tree-toggle";
  toggle.setAttribute("aria-expanded", "false");
  toggle.setAttribute(
    "aria-label",
    `Ordner „${e.name}“ ein- oder ausklappen`
  );
  toggle.textContent = "\u25B8";

  const name = document.createElement("span");
  name.className = "browse-dir-label browse-name";
  name.textContent = e.name;

  const meta = document.createElement("span");
  meta.className = "browse-meta";
  meta.textContent = `${formatZeit(e.modifiedUtc)} · Ordner`;

  row.append(toggle, name, meta);

  const kindUl = document.createElement("ul");
  kindUl.className = "browse-tree-children";
  kindUl.hidden = true;
  kindUl.setAttribute("role", "group");

  li.append(row, kindUl);
  return li;
}

function fuelleOrdnerKinder(ul: HTMLUListElement, entries: BrowseEntry[]): void {
  ul.innerHTML = "";
  if (entries.length === 0) {
    const leer = document.createElement("li");
    leer.className = "browse-tree-leer browse-leer";
    leer.textContent = "(Ordner ist leer)";
    ul.appendChild(leer);
    return;
  }
  for (const e of entries) {
    ul.appendChild(
      e.isDirectory ? erzeugeOrdnerKnoten(e) : erzeugeDateiKnoten(e)
    );
  }
}

async function holeBrowseEntries(relPath: string): Promise<BrowseEntry[]> {
  const query =
    relPath === "" ? "" : `?path=${encodeURIComponent(relPath)}`;
  const res = await fetch(`/api/browse${query}`);
  const raw: unknown = await res.json();
  if (!res.ok) {
    const err = raw as Partial<ErrorResponse>;
    throw new Error(
      typeof err.error === "string" ? err.error : `Fehler (${res.status})`
    );
  }
  const data = raw as BrowseResponse;
  return data.entries ?? [];
}

async function ladeKinderInOrdner(li: HTMLLIElement): Promise<void> {
  if (li.dataset.loaded === "true") return;
  const relPath = li.dataset.relativePath ?? "";
  const entries = await holeBrowseEntries(relPath);
  const kindUl = li.querySelector(
    ":scope > ul.browse-tree-children"
  ) as HTMLUListElement;
  fuelleOrdnerKinder(kindUl, entries);
  li.dataset.loaded = "true";
}

function setToggleExpanded(li: HTMLLIElement, expanded: boolean): void {
  const btn = li.querySelector(".browse-tree-toggle");
  const kindUl = li.querySelector(
    ":scope > ul.browse-tree-children"
  ) as HTMLUListElement | null;
  li.classList.toggle("browse-tree-expanded", expanded);
  btn?.setAttribute("aria-expanded", expanded ? "true" : "false");
  li.setAttribute("aria-expanded", expanded ? "true" : "false");
  if (kindUl) kindUl.hidden = !expanded;
}

async function toggleOrdnerExpand(li: HTMLLIElement): Promise<void> {
  const wirdGeoeffnet = !li.classList.contains("browse-tree-expanded");
  if (wirdGeoeffnet) {
    try {
      setBrowseFehler(null);
      await ladeKinderInOrdner(li);
    } catch (e) {
      const msg =
        e instanceof Error ? e.message : "Ordner konnte nicht geladen werden.";
      setBrowseFehler(msg);
      return;
    }
    setToggleExpanded(li, true);
  } else {
    setToggleExpanded(li, false);
  }
}

async function expandOrdnerOeffnen(li: HTMLLIElement): Promise<void> {
  try {
    setBrowseFehler(null);
    await ladeKinderInOrdner(li);
  } catch (e) {
    const msg =
      e instanceof Error ? e.message : "Ordner konnte nicht geladen werden.";
    setBrowseFehler(msg);
    throw e;
  }
  setToggleExpanded(li, true);
}

async function ensureExpandedPath(relPath: string): Promise<void> {
  const norm = normalisiereRelPfad(relPath);
  if (!norm) return;

  const segments = norm.split("/").filter(Boolean);
  let pathSoFar = "";
  let parentUl = el("#browse-liste", "ul");

  for (const seg of segments) {
    pathSoFar = pathSoFar ? `${pathSoFar}/${seg}` : seg;
    const knoten = findeDirektesOrdnerKind(parentUl, pathSoFar);
    if (!knoten) break;
    await expandOrdnerOeffnen(knoten);
    const nextUl = knoten.querySelector(
      ":scope > ul.browse-tree-children"
    ) as HTMLUListElement | null;
    if (!nextUl) break;
    parentUl = nextUl;
  }
}

function aktualisiereBaumHighlight(): void {
  const ul = el("#browse-liste", "ul");
  ul.querySelectorAll("li.browse-tree-node").forEach((node) => {
    const li = node as HTMLLIElement;
    const path = li.dataset.relativePath ?? "";
    if (li.classList.contains("browse-dir")) {
      const cur = browsePfadAktuell === path;
      li.classList.toggle("browse-tree-current-dir", cur);
    }
    if (li.classList.contains("browse-file")) {
      const sel =
        ausgewaehlteDateiRel !== null && ausgewaehlteDateiRel === path;
      li.classList.toggle("browse-selected", sel);
      if (sel) li.setAttribute("aria-current", "true");
      else li.removeAttribute("aria-current");
    }
  });
}

function scrollAktuellenOrdnerInsSichtfeld(): void {
  const ul = el("#browse-liste", "ul");
  const hit = ul.querySelector(
    "li.browse-dir.browse-tree-current-dir"
  ) as HTMLElement | null;
  hit?.scrollIntoView({ block: "nearest", behavior: "smooth" });
}

function baumeRootAusEntries(entries: BrowseEntry[]): void {
  const ul = el("#browse-liste", "ul");
  ul.innerHTML = "";
  fuelleOrdnerKinder(ul, entries);
  aktualisiereBaumHighlight();
}

async function navigiereZuOrdner(relPath: string): Promise<void> {
  browsePfadAktuell = normalisiereRelPfad(relPath);
  ausgewaehlteDateiRel = null;
  metaLeeren();
  aktualisiereBrowseToolbar();
  try {
    await ensureExpandedPath(browsePfadAktuell);
  } catch {
    /* Fehler bereits via expandOrdnerOeffnen / setBrowseFehler */
  }
  aktualisiereBaumHighlight();
  scrollAktuellenOrdnerInsSichtfeld();
}

async function initialisiereBrowseBaum(): Promise<void> {
  browsePfadAktuell = "";
  ausgewaehlteDateiRel = null;
  aktualisiereBrowseToolbar();
  setBrowseFehler(null);
  metaLeeren();

  const ul = el("#browse-liste", "ul");
  ul.innerHTML = "";

  try {
    const entries = await holeBrowseEntries("");
    baumeRootAusEntries(entries);
  } catch (e) {
    const msg =
      e instanceof Error
        ? e.message
        : "Ordnerliste konnte nicht geladen werden (Netzwerk).";
    setBrowseFehler(msg);
    baumeRootAusEntries([]);
  }
}

function rendereMeta(meta: FileMetaResponse): void {
  const inh = el("#meta-inhalt", "div");
  const dl = document.createElement("dl");
  dl.className = "meta-dl";

  function zeile(dt: string, dd: string): void {
    const dEl = document.createElement("dt");
    dEl.textContent = dt;
    const ddEl = document.createElement("dd");
    ddEl.textContent = dd;
    dl.append(dEl, ddEl);
  }

  zeile("Name", meta.name);
  zeile("Relativer Pfad", meta.relativePath);
  zeile("Art", kindLabel(meta.kind));
  zeile("MIME-Typ", meta.mimeType);
  zeile("Größe", formatBytes(meta.sizeBytes));
  zeile("Zuletzt geändert", formatZeit(meta.modifiedUtc));

  inh.innerHTML = "";
  inh.appendChild(dl);
}

async function ladeDateiMeta(relPath: string): Promise<void> {
  ausgewaehlteDateiRel = relPath;
  setMetaFehler(null);
  setVorschauFehler(null);
  el("#vorschau-inhalt", "div").innerHTML =
    '<p class="vorschau-platzhalter">Lade Vorschau …</p>';

  const q = `?path=${encodeURIComponent(relPath)}`;

  try {
    const res = await fetch(`/api/file/meta${q}`);
    const raw: unknown = await res.json();

    if (!res.ok) {
      const err = raw as Partial<ErrorResponse>;
      const msg =
        typeof err.error === "string" ? err.error : `Fehler (${res.status})`;
      setMetaFehler(msg);
      el("#meta-inhalt", "div").innerHTML = "";
      vorschauLeeren();
      aktualisiereBaumHighlight();
      return;
    }

    const meta = raw as FileMetaResponse;
    rendereMeta(meta);
    aktualisiereBaumHighlight();
    await zeigeVorschau(meta);
  } catch {
    setMetaFehler("Metadaten konnten nicht geladen werden (Netzwerk).");
    el("#meta-inhalt", "div").innerHTML = "";
    vorschauLeeren();
  }
}

function logProjectsApiError(
  operation: string,
  res: Response,
  bodyText: string
): void {
  const snippet =
    bodyText.length > 500 ? `${bodyText.slice(0, 500)}…` : bodyText;
  console.error(
    `[StatikManager /api/projects] ${operation} → ${res.status} ${res.statusText}`,
    snippet || "(leer)"
  );
}

async function ladeProjektliste(aktuellesRoot: string | null): Promise<void> {
  const sel = el("#projekt-select", "select");
  try {
    const res = await fetch("/api/projects");
    const text = await res.text();
    let raw: unknown;
    try {
      raw = text.trim() ? JSON.parse(text) : {};
    } catch {
      logProjectsApiError("GET", res, text);
      return;
    }
    if (!res.ok) {
      logProjectsApiError("GET", res, text);
      return;
    }
    const data = raw as ProjectsResponse;
    const projects = data.projects ?? [];
    sel.innerHTML = '<option value="">— Projekt wählen —</option>';
    for (const p of projects) {
      const opt = document.createElement("option");
      opt.value = p.path;
      opt.textContent =
        p.name && p.name.trim() !== "" ? `${p.name} — ${p.path}` : p.path;
      sel.appendChild(opt);
    }
    if (aktuellesRoot) {
      const found = projects.some((p) => p.path === aktuellesRoot);
      sel.value = found ? aktuellesRoot : "";
    } else {
      sel.value = "";
    }
  } catch (e) {
    console.error("[StatikManager /api/projects] GET (Netzwerk)", e);
  }
}

async function aktuellesProjektInListeSpeichern(): Promise<void> {
  setFehler(null);
  try {
    const res = await fetch("/api/session/root");
    const data = (await res.json()) as SessionResponse;
    if (!res.ok || !data.rootPath) {
      setFehler("Kein geöffnetes Projekt — bitte zuerst einen Ordner öffnen.");
      return;
    }
    const res2 = await fetch("/api/projects", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: data.rootPath, name: null }),
    });
    const body2 = await res2.text();
    let raw2: unknown;
    try {
      raw2 = body2.trim() ? JSON.parse(body2) : {};
    } catch {
      logProjectsApiError("POST", res2, body2);
      setFehler(
        `Unerwartete Antwort (${res2.status}). Details in der Browser-Konsole.`
      );
      return;
    }
    if (!res2.ok) {
      logProjectsApiError("POST", res2, body2);
      const err = raw2 as Partial<ErrorResponse>;
      setFehler(
        typeof err.error === "string" ? err.error : `Fehler (${res2.status})`
      );
      return;
    }
    await ladeProjektliste(data.rootPath);
  } catch (e) {
    console.error("[StatikManager /api/projects] POST (Netzwerk)", e);
    setFehler("Projektliste konnte nicht aktualisiert werden (Netzwerk).");
  }
}

async function projektAusListeEntfernen(): Promise<void> {
  const sel = el("#projekt-select", "select");
  const path = sel.value;
  if (!path) {
    setFehler("Bitte ein Projekt in der Liste auswählen.");
    return;
  }
  setFehler(null);
  try {
    const res = await fetch(
      `/api/projects?path=${encodeURIComponent(path)}`,
      { method: "DELETE" }
    );
    const textDel = await res.text();
    let raw: unknown;
    try {
      raw = textDel.trim() ? JSON.parse(textDel) : {};
    } catch {
      logProjectsApiError("DELETE", res, textDel);
      setFehler(
        `Unerwartete Antwort (${res.status}). Details in der Browser-Konsole.`
      );
      return;
    }
    if (!res.ok) {
      logProjectsApiError("DELETE", res, textDel);
      const err = raw as Partial<ErrorResponse>;
      setFehler(
        typeof err.error === "string" ? err.error : `Fehler (${res.status})`
      );
      return;
    }
    const resSess = await fetch("/api/session/root");
    const session = (await resSess.json()) as SessionResponse;
    await ladeProjektliste(session.rootPath ?? null);
  } catch {
    setFehler("Projekt konnte nicht entfernt werden.");
  }
}

function aufProjektSelectChange(): void {
  const sel = el("#projekt-select", "select");
  const path = sel.value;
  if (!path) return;
  void projektRootSetzen(path);
}

async function ladeSession(): Promise<void> {
  let rootFuerSelect: string | null = null;
  try {
    const res = await fetch("/api/session/root");
    const data = (await res.json()) as SessionResponse;
    if (!res.ok) {
      setFehler("Session konnte nicht geladen werden.");
      await ladeProjektliste(null);
      return;
    }
    setRootAnzeige(data.rootPath ?? null);
    setFehler(null);
    rootFuerSelect = data.rootPath ?? null;
    if (data.rootPath) {
      el("#pfad-input", "input").value = data.rootPath;
      browsePfadAktuell = "";
      metaLeeren();
      ordnerHinweisAusblenden();
      await initialisiereBrowseBaum();
    } else {
      el("#pfad-input", "input").value = "";
      el("#browse-liste", "ul").innerHTML = "";
      aktualisiereBrowseToolbar();
      setBrowseFehler(null);
      metaLeeren();
    }
  } catch {
    setFehler("Verbindung zur API fehlgeschlagen.");
    rootFuerSelect = null;
  }
  await ladeProjektliste(rootFuerSelect);
}

async function projektRootSetzen(pfad: string): Promise<void> {
  setFehler(null);

  try {
    const res = await fetch("/api/session/root", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rootPath: pfad }),
    });

    const raw: unknown = await res.json();

    if (!res.ok) {
      const err = raw as Partial<ErrorResponse>;
      setFehler(
        typeof err.error === "string" ? err.error : `Fehler (${res.status})`
      );
      return;
    }

    const session = raw as SessionResponse;
    setRootAnzeige(session.rootPath ?? null);
    setFehler(null);
    ordnerHinweisAusblenden();
    browsePfadAktuell = "";
    metaLeeren();
    await initialisiereBrowseBaum();
    el("#pfad-input", "input").value = session.rootPath ?? "";
    await ladeProjektliste(session.rootPath ?? null);
  } catch {
    setFehler("Projekt konnte nicht gesetzt werden (Netzwerk).");
  }
}

async function projektOeffnen(): Promise<void> {
  const input = el("#pfad-input", "input");
  await projektRootSetzen(input.value);
}

function ordnerHinweisAusblenden(): void {
  const h = el("#ordner-hinweis", "p");
  h.hidden = true;
  h.textContent = "";
}

function ordnerHinweisAnzeigen(): void {
  setFehler(null);
  const h = el("#ordner-hinweis", "p");
  h.hidden = false;
  h.textContent =
    "Bitte Projektpfad oben einfügen. Der native Ordnerdialog wird später über Desktop-Wrapper/Tauri umgesetzt.";
}

function aufBrowseListeKlick(ev: MouseEvent): void {
  const t = ev.target as HTMLElement | null;
  if (t?.closest?.("button.browse-tree-toggle")) {
    const btn = t.closest("button.browse-tree-toggle");
    const li = btn?.closest("li.browse-dir") as HTMLLIElement | null;
    if (li) {
      ev.preventDefault();
      void toggleOrdnerExpand(li);
    }
    return;
  }

  const fileLi = t?.closest?.("li.browse-file") as HTMLLIElement | null;
  if (fileLi?.dataset.relativePath) {
    void ladeDateiMeta(fileLi.dataset.relativePath);
    return;
  }

  const dirLi = t?.closest?.("li.browse-dir") as HTMLLIElement | null;
  if (dirLi?.dataset.relativePath) {
    void navigiereZuOrdner(dirLi.dataset.relativePath);
  }
}

function aufBrowseListeKey(ev: KeyboardEvent): void {
  if (ev.key !== "Enter" && ev.key !== " ") return;
  const t = ev.target as HTMLElement | null;

  const toggle = t?.closest?.("button.browse-tree-toggle");
  if (toggle) {
    ev.preventDefault();
    const li = toggle.closest("li.browse-dir") as HTMLLIElement | null;
    if (li) void toggleOrdnerExpand(li);
    return;
  }

  const li = t?.closest?.("li.browse-dir, li.browse-file") as
    | HTMLLIElement
    | null;
  if (!li?.dataset.relativePath) return;
  ev.preventDefault();

  if (li.classList.contains("browse-dir")) {
    void navigiereZuOrdner(li.dataset.relativePath);
    return;
  }
  void ladeDateiMeta(li.dataset.relativePath);
}

void checkHealth();
void ladeSession();

const btn = el("#btn-oeffnen", "button");
btn.addEventListener("click", () => void projektOeffnen());

const btnOrdner = el("#btn-ordner", "button");
btnOrdner.addEventListener("click", () => ordnerHinweisAnzeigen());

const input = el("#pfad-input", "input");
input.addEventListener("keydown", (ev) => {
  if (ev.key === "Enter") void projektOeffnen();
});

const btnHoch = el("#btn-hoch", "button");
btnHoch.addEventListener("click", () => {
  const neu = parentRelativePath(browsePfadAktuell);
  void navigiereZuOrdner(neu);
});

const browseListe = el("#browse-liste", "ul");
browseListe.addEventListener("click", aufBrowseListeKlick);
browseListe.addEventListener("keydown", aufBrowseListeKey);

const projektSelect = el("#projekt-select", "select");
projektSelect.addEventListener("change", aufProjektSelectChange);

el("#btn-projekt-speichern", "button").addEventListener("click", () =>
  void aktuellesProjektInListeSpeichern()
);
el("#btn-projekt-entfernen", "button").addEventListener("click", () =>
  void projektAusListeEntfernen()
);
