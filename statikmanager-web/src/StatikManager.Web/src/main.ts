import "./styles/main.css";

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) throw new Error("#app fehlt");

app.innerHTML = `
  <main class="shell">
    <h1>StatikManager Web</h1>
    <p class="muted">Milestone 1 – Health-Check</p>
    <p id="health-status" class="status">Lade …</p>
  </main>
`;

function getStatusEl(): HTMLParagraphElement {
  const el = document.querySelector<HTMLParagraphElement>("#health-status");
  if (!el) throw new Error("#health-status fehlt");
  return el;
}

async function checkHealth(): Promise<void> {
  const statusEl = getStatusEl();
  try {
    const res = await fetch("/api/health");
    const data = (await res.json()) as { ok?: boolean };
    if (res.ok && data.ok === true) {
      statusEl.textContent = "API: /api/health → { ok: true }";
      statusEl.classList.add("ok");
      return;
    }
    statusEl.textContent = `Unerwartete Antwort (${res.status})`;
    statusEl.classList.add("err");
  } catch {
    statusEl.textContent =
      "API nicht erreichbar. Backend starten (dotnet run) oder Vite-Proxy prüfen.";
    statusEl.classList.add("err");
  }
}

void checkHealth();
