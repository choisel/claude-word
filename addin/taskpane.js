/* globals Office, Word */

const SERVER = "https://localhost:5000";

let sessionId = null;
let docMode = null;      // "full" | "summarized" | null
let sectionCount = 0;

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("ask-btn").addEventListener("click", onAsk);
    document.getElementById("refresh-btn").addEventListener("click", refreshSelection);
    document.getElementById("load-doc-btn").addEventListener("click", loadDocument);
    document.getElementById("clear-section-btn").addEventListener("click", () => {
      document.getElementById("section-number").value = "";
    });
    document.getElementById("question").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        onAsk();
      }
    });

    refreshSelection();
    checkServer().then(() => loadDocument());
  }
});

// ---------------------------------------------------------------------------
// Server health
// ---------------------------------------------------------------------------
async function checkServer() {
  try {
    const resp = await fetch(`${SERVER}/health`);
    if (resp.ok) {
      setIndicator("idle");
      return true;
    }
  } catch {}
  setIndicator("error");
  showError("Serveur local non joignable. Lancez start.sh et réessayez.");
  return false;
}

// ---------------------------------------------------------------------------
// Document loading
// ---------------------------------------------------------------------------
async function loadDocument() {
  setInitLabel("Lecture du document…", "loading");
  setIndicator("busy");
  hideError();

  let fullText = "";
  try {
    fullText = await readFullDocument();
  } catch (err) {
    setInitLabel("Erreur de lecture du document", "");
    setIndicator("error");
    showError("Impossible de lire le document Word.");
    console.error("readFullDocument error:", err);
    return;
  }

  if (!fullText.trim()) {
    setInitLabel("Document vide", "");
    setIndicator("idle");
    return;
  }

  try {
    const resp = await fetch(`${SERVER}/init`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: fullText, session_id: sessionId }),
    });

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      throw new Error(err.detail || `HTTP ${resp.status}`);
    }

    const data = await resp.json();
    sessionId = data.session_id;
    docMode = data.mode;
    sectionCount = data.section_count;

    const modeLabel = data.mode === "summarized" ? "résumé" : "complet";
    const label = `${data.section_count} sections · ${data.page_count} pages · mode ${modeLabel}`;
    setInitLabel(`Document chargé — ${label}`, "ready");
    setDocStatus(`${data.page_count}p`);
    setIndicator("idle");

    // Clear previous chat on reload
    document.getElementById("chat-history").innerHTML = "";

  } catch (err) {
    setInitLabel("Échec du chargement", "");
    setIndicator("error");
    showError(`Erreur initialisation : ${err.message}`);
    console.error("loadDocument error:", err);
  }
}

async function readFullDocument() {
  return new Promise((resolve, reject) => {
    Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      resolve(body.text);
    }).catch(reject);
  });
}

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
      document.getElementById("selection-preview").textContent =
        text || "(aucun texte sélectionné)";
    });
  } catch (err) {
    console.error("refreshSelection error:", err);
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
  const isPlaceholder = selectedText.startsWith("(");
  const sectionNumber = document.getElementById("section-number").value.trim() || null;

  hideError();
  setIndicator("busy");
  setLoading(true);

  appendMessage("user", question, sectionNumber);
  questionEl.value = "";

  // Create the Claude response bubble immediately (empty, will fill via streaming)
  const claudeEl = createStreamingBubble();

  try {
    const resp = await fetch(`${SERVER}/ask-stream`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        question,
        selected_text: isPlaceholder ? "" : selectedText,
        section_number: sectionNumber,
        session_id: sessionId,
      }),
    });

    if (!resp.ok) {
      claudeEl.remove();
      const err = await resp.json().catch(() => ({}));
      showError(err.detail || `Erreur HTTP ${resp.status}`);
      setIndicator("error");
      return;
    }

    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop(); // keep incomplete last line

      for (const line of lines) {
        if (!line.startsWith("data: ")) continue;
        try {
          const event = JSON.parse(line.slice(6));

          if (event.type === "token") {
            appendTokenToStream(claudeEl, event.text);
          } else if (event.type === "done") {
            finalizeStreamBubble(claudeEl, event.duration_ms);
            setIndicator("idle");
          } else if (event.type === "error") {
            claudeEl.remove();
            showError(event.detail);
            setIndicator("error");
          }
        } catch {}
      }
    }

  } catch (err) {
    claudeEl.remove();
    showError("Impossible de joindre le serveur. Est-ce que start.sh tourne ?");
    setIndicator("error");
    console.error("fetch /ask-stream error:", err);
  } finally {
    setLoading(false);
  }
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------
function createStreamingBubble() {
  const chat = document.getElementById("chat-history");
  const el = document.createElement("div");
  el.className = "chat-message claude streaming";

  const body = document.createElement("span");
  body.className = "stream-body";
  el.appendChild(body);

  // Blinking cursor
  const cursor = document.createElement("span");
  cursor.className = "stream-cursor";
  cursor.textContent = "▋";
  el.appendChild(cursor);

  chat.appendChild(el);
  el.scrollIntoView({ behavior: "smooth", block: "end" });
  return el;
}

function appendTokenToStream(el, text) {
  el.querySelector(".stream-body").textContent += text;
  el.scrollIntoView({ behavior: "smooth", block: "end" });
}

function finalizeStreamBubble(el, duration_ms) {
  el.classList.remove("streaming");
  const cursor = el.querySelector(".stream-cursor");
  if (cursor) cursor.remove();

  const meta = document.createElement("div");
  meta.className = "chat-meta";
  meta.textContent = `${(duration_ms / 1000).toFixed(1)}s`;
  el.appendChild(meta);
}

function appendMessage(role, text, sectionNumber = null, duration_ms = undefined) {
  const chat = document.getElementById("chat-history");
  const el = document.createElement("div");
  el.className = `chat-message ${role}`;

  if (sectionNumber && role === "user") {
    const tag = document.createElement("div");
    tag.className = "chat-section-tag";
    tag.textContent = `Section ${sectionNumber}`;
    el.appendChild(tag);
  }

  const body = document.createElement("span");
  body.textContent = text;
  el.appendChild(body);

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
  document.getElementById("status-indicator").className = `status-${state}`;
}

function setInitLabel(text, cls) {
  const el = document.getElementById("init-label");
  el.textContent = text;
  el.className = cls || "";
}

function setDocStatus(text) {
  document.getElementById("doc-status").textContent = text;
}

function showError(msg) {
  const bar = document.getElementById("error-bar");
  bar.textContent = msg;
  bar.classList.remove("hidden");
}

function hideError() {
  document.getElementById("error-bar").classList.add("hidden");
}
