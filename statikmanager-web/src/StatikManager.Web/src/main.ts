import "./styles/main.css";

import type { ErrorResponse, SessionResponse } from "./types/api";

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) throw new Error("#app fehlt");

app.innerHTML = `
  <main class="shell">
    <h1>StatikManager Web</h1>
    <p class="muted">Milestone 2 – Projekt-Root</p>

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

function setRootAnzeige(rootPath: string | null): void {
  const rootAnzeige = el("#root-anzeige", "p");
  rootAnzeige.textContent =
    rootPath === null || rootPath === "" ? "— (nicht gesetzt)" : rootPath;
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
  } catch {
    setFehler("Projekt konnte nicht gesetzt werden (Netzwerk).");
  }
}

void checkHealth();
void ladeSession();

const btn = el("#btn-oeffnen", "button");
btn.addEventListener("click", () => void projektOeffnen());

const input = el("#pfad-input", "input");
input.addEventListener("keydown", (ev) => {
  if (ev.key === "Enter") void projektOeffnen();
});
