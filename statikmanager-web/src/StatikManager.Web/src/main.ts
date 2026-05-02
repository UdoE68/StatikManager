import "./styles/main.css";

import type {
  BrowseEntry,
  BrowseResponse,
  ErrorResponse,
  FileKind,
  FileMetaResponse,
  SessionResponse,
} from "./types/api";

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) throw new Error("#app fehlt");

let browsePfadAktuell = "";
let ausgewaehlteDateiRel: string | null = null;
let letzteBrowseEntries: BrowseEntry[] = [];

app.innerHTML = `
  <main class="shell">
    <h1>StatikManager Web</h1>
    <p class="muted">Milestone 5 – Metadaten &amp; Vorschau</p>

    <section class="panel" aria-labelledby="projekt-label">
      <h2 id="projekt-label" class="h2">Projekt</h2>
      <label class="label" for="pfad-input">Projektpfad</label>
      <div class="row">
        <input id="pfad-input" class="input" type="text" autocomplete="off"
          placeholder="z. B. C:\\Projekte\\MeinStatikProjekt" spellcheck="false" />
        <button id="btn-oeffnen" type="button" class="btn">Projekt öffnen</button>
        <button id="btn-ordner" type="button" class="btn btn-secondary">Ordner wählen</button>
      </div>
      <p id="fehler" class="fehler" role="alert" hidden></p>
      <p class="aktuell-label">Aktuelles Projekt-Root</p>
      <p id="root-anzeige" class="root-anzeige">—</p>
    </section>

    <div class="layout-haupt">
      <div class="layout-spalte-links">
        <section class="panel panel-inner" id="browse-panel" aria-labelledby="browse-label">
          <h2 id="browse-label" class="h2">Ordnerinhalt</h2>
          <div class="browse-toolbar">
            <button id="btn-hoch" type="button" class="btn btn-secondary" disabled>Nach oben</button>
            <span id="browse-pfad-anzeige" class="browse-pfad"></span>
          </div>
          <p id="browse-fehler" class="fehler" role="alert" hidden></p>
          <ul id="browse-liste" class="browse-liste" aria-live="polite"></ul>
        </section>
      </div>

      <div class="layout-spalte-rechts layout-spalte-rechts-stapel">
        <section class="panel panel-inner" aria-labelledby="meta-label">
          <h2 id="meta-label" class="h2">Datei-Info</h2>
          <p id="meta-fehler" class="fehler" role="alert" hidden></p>
          <div id="meta-inhalt" class="meta-inhalt">
            <p class="meta-platzhalter">Keine Datei ausgewählt.</p>
          </div>
        </section>

        <section class="panel panel-inner" aria-labelledby="vorschau-label">
          <h2 id="vorschau-label" class="h2">Vorschau</h2>
          <p id="vorschau-fehler" class="fehler" role="alert" hidden></p>
          <div id="vorschau-inhalt" class="vorschau-inhalt">
            <p class="vorschau-platzhalter">Keine Datei ausgewählt.</p>
          </div>
        </section>
      </div>
    </div>

    <section class="panel muted-small">
      <p id="health-status" class="status">API …</p>
    </section>
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

function rendereBrowseListe(entries: BrowseEntry[]): void {
  const ul = el("#browse-liste", "ul");
  ul.innerHTML = "";
  if (entries.length === 0) {
    const li = document.createElement("li");
    li.className = "browse-leer";
    li.textContent = "(Ordner ist leer)";
    ul.appendChild(li);
    return;
  }

  for (const e of entries) {
    const li = document.createElement("li");
    li.className = "browse-zeile";
    if (e.isDirectory) {
      li.classList.add("browse-dir");
      li.dataset.relativePath = e.relativePath;
      li.tabIndex = 0;
      li.setAttribute("role", "button");
      const name = document.createElement("span");
      name.className = "browse-name";
      name.textContent = e.name;
      const meta = document.createElement("span");
      meta.className = "browse-meta";
      meta.textContent = `${formatZeit(e.modifiedUtc)} · Ordner`;
      li.append(name, meta);
    } else {
      li.classList.add("browse-file");
      li.dataset.relativePath = e.relativePath;
      li.tabIndex = 0;
      li.setAttribute("role", "button");
      if (ausgewaehlteDateiRel === e.relativePath) {
        li.classList.add("browse-selected");
        li.setAttribute("aria-current", "true");
      }
      const name = document.createElement("span");
      name.className = "browse-name";
      name.textContent = e.name;
      const meta = document.createElement("span");
      meta.className = "browse-meta";
      meta.textContent = `${formatZeit(e.modifiedUtc)} · ${formatBytes(e.sizeBytes)}`;
      li.append(name, meta);
    }
    ul.appendChild(li);
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
      rendereBrowseListe(letzteBrowseEntries);
      return;
    }

    const meta = raw as FileMetaResponse;
    rendereMeta(meta);
    rendereBrowseListe(letzteBrowseEntries);
    await zeigeVorschau(meta);
  } catch {
    setMetaFehler("Metadaten konnten nicht geladen werden (Netzwerk).");
    el("#meta-inhalt", "div").innerHTML = "";
    vorschauLeeren();
  }
}

async function ladeBrowse(relPath: string): Promise<void> {
  browsePfadAktuell = relPath;
  ausgewaehlteDateiRel = null;
  aktualisiereBrowseToolbar();
  setBrowseFehler(null);
  metaLeeren();

  const query =
    relPath === "" ? "" : `?path=${encodeURIComponent(relPath)}`;

  try {
    const res = await fetch(`/api/browse${query}`);
    const raw: unknown = await res.json();

    if (!res.ok) {
      const err = raw as Partial<ErrorResponse>;
      setBrowseFehler(
        typeof err.error === "string" ? err.error : `Fehler (${res.status})`
      );
      letzteBrowseEntries = [];
      rendereBrowseListe([]);
      return;
    }

    const data = raw as BrowseResponse;
    letzteBrowseEntries = data.entries ?? [];
    rendereBrowseListe(letzteBrowseEntries);
  } catch {
    setBrowseFehler("Ordnerliste konnte nicht geladen werden (Netzwerk).");
    letzteBrowseEntries = [];
    rendereBrowseListe([]);
  }
}

async function ladeSession(): Promise<void> {
  try {
    const res = await fetch("/api/session/root");
    const data = (await res.json()) as SessionResponse;
    if (!res.ok) {
      setFehler("Session konnte nicht geladen werden.");
      return;
    }
    setRootAnzeige(data.rootPath ?? null);
    setFehler(null);
    if (data.rootPath) {
      browsePfadAktuell = "";
      metaLeeren();
      await ladeBrowse("");
    } else {
      letzteBrowseEntries = [];
      rendereBrowseListe([]);
      aktualisiereBrowseToolbar();
      setBrowseFehler(null);
      metaLeeren();
    }
  } catch {
    setFehler("Verbindung zur API fehlgeschlagen.");
  }
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
    browsePfadAktuell = "";
    metaLeeren();
    await ladeBrowse("");
  } catch {
    setFehler("Projekt konnte nicht gesetzt werden (Netzwerk).");
  }
}

async function projektOeffnen(): Promise<void> {
  const input = el("#pfad-input", "input");
  await projektRootSetzen(input.value);
}

async function ordnerWaehlen(): Promise<void> {
  setFehler(null);

  try {
    const res = await fetch("/api/session/pick-root", {
      method: "POST",
    });

    if (res.status === 204) {
      return;
    }

    const raw: unknown = await res.json();

    if (!res.ok) {
      const err = raw as Partial<ErrorResponse>;
      setFehler(
        typeof err.error === "string"
          ? err.error
          : `Ordnerdialog: Fehler (${res.status})`
      );
      return;
    }

    const session = raw as SessionResponse;
    const input = el("#pfad-input", "input");
    input.value = session.rootPath ?? "";
    setRootAnzeige(session.rootPath ?? null);
    setFehler(null);
    browsePfadAktuell = "";
    metaLeeren();
    await ladeBrowse("");
  } catch {
    setFehler("Ordnerdialog nicht erreichbar (Netzwerk oder Server).");
  }
}

function aufBrowseListeKlick(ev: MouseEvent): void {
  const t = ev.target as HTMLElement | null;
  const li = t?.closest?.("li.browse-dir, li.browse-file") as HTMLLIElement | null;
  if (!li) return;
  const p = li.dataset.relativePath;
  if (p === undefined) return;

  if (li.classList.contains("browse-dir")) {
    void ladeBrowse(p);
    return;
  }

  if (li.classList.contains("browse-file")) {
    void ladeDateiMeta(p);
  }
}

function aufBrowseListeKey(ev: KeyboardEvent): void {
  if (ev.key !== "Enter" && ev.key !== " ") return;
  const li = ev.target as HTMLElement | null;
  if (!li?.matches("li.browse-dir, li.browse-file")) return;
  ev.preventDefault();
  const p = li.dataset.relativePath;
  if (p === undefined) return;

  if (li.classList.contains("browse-dir")) {
    void ladeBrowse(p);
    return;
  }

  void ladeDateiMeta(p);
}

void checkHealth();
void ladeSession();

const btn = el("#btn-oeffnen", "button");
btn.addEventListener("click", () => void projektOeffnen());

const btnOrdner = el("#btn-ordner", "button");
btnOrdner.addEventListener("click", () => void ordnerWaehlen());

const input = el("#pfad-input", "input");
input.addEventListener("keydown", (ev) => {
  if (ev.key === "Enter") void projektOeffnen();
});

const btnHoch = el("#btn-hoch", "button");
btnHoch.addEventListener("click", () => {
  const neu = parentRelativePath(browsePfadAktuell);
  void ladeBrowse(neu);
});

const browseListe = el("#browse-liste", "ul");
browseListe.addEventListener("click", aufBrowseListeKlick);
browseListe.addEventListener("keydown", aufBrowseListeKey);
