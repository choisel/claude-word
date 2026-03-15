/* globals Office, Word */

const SERVER = "https://localhost:5000";

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("ask-btn").addEventListener("click", onAsk);
    document.getElementById("refresh-btn").addEventListener("click", refreshSelection);
    document.getElementById("question").addEventListener("keydown", (e) => {
      // Ctrl+Enter or Cmd+Enter to submit
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        onAsk();
      }
    });
    refreshSelection();
    checkServer();
  }
});

// ---------------------------------------------------------------------------
// Selection
// ---------------------------------------------------------------------------
async function refreshSelection() {
  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      const text = sel.text.trim();
      const preview = document.getElementById("selection-preview");
      preview.textContent = text || "(aucun texte sélectionné)";
    });
  } catch (err) {
    console.error("refreshSelection error:", err);
  }
}

// ---------------------------------------------------------------------------
// Server health check
// ---------------------------------------------------------------------------
async function checkServer() {
  const indicator = document.getElementById("status-indicator");
  try {
    const resp = await fetch(`${SERVER}/health`, { method: "GET" });
    if (resp.ok) {
      setIndicator("idle");
    } else {
      setIndicator("error");
    }
  } catch {
    setIndicator("error");
    showError("Serveur local non joignable. Lancez start.sh et réessayez.");
  }
}

// ---------------------------------------------------------------------------
// Ask Claude
// ---------------------------------------------------------------------------
async function onAsk() {
  const questionEl = document.getElementById("question");
  const question = questionEl.value.trim();
  if (!question) return;

  const selectedText = document.getElementById("selection-preview").textContent;
  const isPlaceholder =
    selectedText === "(aucun texte sélectionné)" ||
    selectedText === "(sélectionnez du texte dans le document)";

  hideError();
  setIndicator("busy");
  setLoading(true);

  // Show user message in chat
  appendMessage("user", question);
  questionEl.value = "";

  // Placeholder while waiting
  const loadingEl = appendMessage("loading", "Claude réfléchit…");

  try {
    const resp = await fetch(`${SERVER}/ask`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        selected_text: isPlaceholder ? "" : selectedText,
        question,
      }),
    });

    loadingEl.remove();

    if (!resp.ok) {
      let detail = `Erreur HTTP ${resp.status}`;
      try {
        const err = await resp.json();
        detail = err.detail || detail;
      } catch {}
      showError(detail);
      setIndicator("error");
      return;
    }

    const data = await resp.json();
    appendMessage("claude", data.answer, data.duration_ms);
    setIndicator("idle");

  } catch (err) {
    loadingEl.remove();
    showError("Impossible de joindre le serveur. Est-ce que start.sh tourne ?");
    setIndicator("error");
    console.error("fetch error:", err);
  } finally {
    setLoading(false);
  }
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------
function appendMessage(role, text, duration_ms) {
  const chat = document.getElementById("chat-history");
  const el = document.createElement("div");
  el.className = `chat-message ${role}`;
  el.textContent = text;

  if (role === "claude" && duration_ms !== undefined) {
    const meta = document.createElement("div");
    meta.className = "chat-meta";
    meta.textContent = `${(duration_ms / 1000).toFixed(1)}s`;
    el.appendChild(meta);
  }

  chat.appendChild(el);
  el.scrollIntoView({ behavior: "smooth", block: "end" });
  return el;
}

function setLoading(isLoading) {
  const btn = document.getElementById("ask-btn");
  btn.disabled = isLoading;
  btn.textContent = isLoading ? "…" : "Envoyer";
}

function setIndicator(state) {
  const el = document.getElementById("status-indicator");
  el.className = `status-${state}`;
}

function showError(msg) {
  const bar = document.getElementById("error-bar");
  bar.textContent = msg;
  bar.classList.remove("hidden");
}

function hideError() {
  document.getElementById("error-bar").classList.add("hidden");
}
