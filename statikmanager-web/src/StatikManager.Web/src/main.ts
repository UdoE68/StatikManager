import "./styles/main.css";

import type {
  BrowseEntry,
  BrowseResponse,
  ErrorResponse,
  SessionResponse,
} from "./types/api";

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) throw new Error("#app fehlt");

let browsePfadAktuell = "";

app.innerHTML = `
  <main class="shell">
    <h1>StatikManager Web</h1>
    <p class="muted">Milestone 3 – Projekt &amp; Ordnerliste</p>

    <section class="panel" aria-labelledby="projekt-label">
      <h2 id="projekt-label" class="h2">Projekt</h2>
      <label class="label" for="pfad-input">Projektpfad</label>
      <div class="row">
        <input id="pfad-input" class="input" type="text" autocomplete="off"
          placeholder="z. B. C:\\Projekte\\MeinStatikProjekt" spellcheck="false" />
        <button id="btn-oeffnen" type="button" class="btn">Projekt öffnen</button>
      </div>
      <p id="fehler" class="fehler" role="alert" hidden></p>
      <p class="aktuell-label">Aktuelles Projekt-Root</p>
      <p id="root-anzeige" class="root-anzeige">—</p>
    </section>

    <section class="panel" id="browse-panel" aria-labelledby="browse-label">
      <h2 id="browse-label" class="h2">Ordnerinhalt</h2>
      <div class="browse-toolbar">
        <button id="btn-hoch" type="button" class="btn btn-secondary" disabled>Nach oben</button>
        <span id="browse-pfad-anzeige" class="browse-pfad"></span>
      </div>
      <p id="browse-fehler" class="fehler" role="alert" hidden></p>
      <ul id="browse-liste" class="browse-liste" aria-live="polite"></ul>
    </section>

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

function setRootAnzeige(rootPath: string | null): void {
  const rootAnzeige = el("#root-anzeige", "p");
  rootAnzeige.textContent =
    rootPath === null || rootPath === "" ? "— (nicht gesetzt)" : rootPath;
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

async function ladeBrowse(relPath: string): Promise<void> {
  browsePfadAktuell = relPath;
  aktualisiereBrowseToolbar();
  setBrowseFehler(null);

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
      rendereBrowseListe([]);
      return;
    }

    const data = raw as BrowseResponse;
    rendereBrowseListe(data.entries ?? []);
  } catch {
    setBrowseFehler("Ordnerliste konnte nicht geladen werden (Netzwerk).");
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
      await ladeBrowse("");
    } else {
      rendereBrowseListe([]);
      aktualisiereBrowseToolbar();
      setBrowseFehler(null);
    }
  } catch {
    setFehler("Verbindung zur API fehlgeschlagen.");
  }
}

async function projektOeffnen(): Promise<void> {
  const input = el("#pfad-input", "input");
  const pfad = input.value;

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
    await ladeBrowse("");
  } catch {
    setFehler("Projekt konnte nicht gesetzt werden (Netzwerk).");
  }
}

function aufBrowseListeKlick(ev: MouseEvent): void {
  const t = ev.target as HTMLElement | null;
  const li = t?.closest?.("li.browse-dir") as HTMLLIElement | null;
  if (!li?.dataset.relativePath) return;
  void ladeBrowse(li.dataset.relativePath);
}

function aufBrowseListeKey(ev: KeyboardEvent): void {
  if (ev.key !== "Enter" && ev.key !== " ") return;
  const li = ev.target as HTMLElement | null;
  if (!li?.classList.contains("browse-dir")) return;
  ev.preventDefault();
  const p = li.dataset.relativePath;
  if (p !== undefined) void ladeBrowse(p);
}

void checkHealth();
void ladeSession();

const btn = el("#btn-oeffnen", "button");
btn.addEventListener("click", () => void projektOeffnen());

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
