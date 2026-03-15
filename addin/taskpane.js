/* globals Office, Word */

const SERVER = "https://localhost:5000";

let sessionId = null;
let docMode = null;
let sectionCount = 0;
let knownSections = [];
let sectionDetectionInterval = null;
let documentLoaded = false;
let pendingSelection = "";

// ---------------------------------------------------------------------------
// i18n — labels keyed by Office locale (2-letter prefix, fallback to "en")
// ---------------------------------------------------------------------------
const STRINGS = {
  fr: {
    noText:          "(aucun texte sélectionné)",
    docNotLoaded:    "Document non chargé",
    loadBtn:         "Charger le document",
    reading:         "Lecture du document…",
    readError:       "Erreur de lecture du document",
    serverError:     "Serveur local non joignable. Lancez start.sh et réessayez.",
    emptyDoc:        "Document vide",
    loadedDoc:       (sections, pages, mode) =>
      `Document chargé — ${sections} sections · ${pages} pages · mode ${mode === "summarized" ? "résumé" : "complet"}`,
    loadFailed:      "Échec du chargement",
    initError:       (msg) => `Erreur initialisation : ${msg}`,
    thinking:        "Claude réfléchit…",
    send:            "Envoyer",
    sending:         "…",
    serverUnreach:   "Impossible de joindre le serveur. Est-ce que start.sh tourne ?",
    httpError:       (status) => `Erreur HTTP ${status}`,
    sectionLabel:    "Section",
    sectionTarget:   "Section ciblée",
    optional:        "(optionnel)",
    placeholder:     "Posez votre question à Claude…",
    sectionAll:      "— document entier —",
    selectionChip:   "Texte sélectionné :",
  },
  en: {
    noText:          "(no text selected)",
    docNotLoaded:    "Document not loaded",
    loadBtn:         "Load document",
    reading:         "Reading document…",
    readError:       "Error reading document",
    serverError:     "Local server unreachable. Run start.sh and try again.",
    emptyDoc:        "Empty document",
    loadedDoc:       (sections, pages, mode) =>
      `Document loaded — ${sections} sections · ${pages} pages · ${mode} mode`,
    loadFailed:      "Load failed",
    initError:       (msg) => `Init error: ${msg}`,
    thinking:        "Claude is thinking…",
    send:            "Send",
    sending:         "…",
    serverUnreach:   "Cannot reach server. Is start.sh running?",
    httpError:       (status) => `HTTP error ${status}`,
    sectionLabel:    "Section",
    sectionTarget:   "Target section",
    optional:        "(optional)",
    placeholder:     "Ask Claude a question…",
    sectionAll:      "— full document —",
    selectionChip:   "Selected text:",
  },
  de: {
    noText:          "(kein Text ausgewählt)",
    docNotLoaded:    "Dokument nicht geladen",
    loadBtn:         "Dokument laden",
    reading:         "Dokument wird gelesen…",
    readError:       "Fehler beim Lesen des Dokuments",
    serverError:     "Lokaler Server nicht erreichbar. Starten Sie start.sh.",
    emptyDoc:        "Leeres Dokument",
    loadedDoc:       (sections, pages, mode) =>
      `Dokument geladen — ${sections} Abschnitte · ${pages} Seiten · Modus ${mode}`,
    loadFailed:      "Laden fehlgeschlagen",
    initError:       (msg) => `Initialisierungsfehler: ${msg}`,
    thinking:        "Claude denkt…",
    send:            "Senden",
    sending:         "…",
    serverUnreach:   "Server nicht erreichbar. Läuft start.sh?",
    httpError:       (status) => `HTTP-Fehler ${status}`,
    sectionLabel:    "Abschnitt",
    sectionTarget:   "Zielabschnitt",
    optional:        "(optional)",
    placeholder:     "Stellen Sie Claude eine Frage…",
    sectionAll:      "— gesamtes Dokument —",
    selectionChip:   "Ausgewählter Text:",
  },
  es: {
    noText:          "(ningún texto seleccionado)",
    docNotLoaded:    "Documento no cargado",
    loadBtn:         "Cargar documento",
    reading:         "Leyendo documento…",
    readError:       "Error al leer el documento",
    serverError:     "Servidor local no disponible. Ejecute start.sh.",
    emptyDoc:        "Documento vacío",
    loadedDoc:       (sections, pages, mode) =>
      `Documento cargado — ${sections} secciones · ${pages} páginas · modo ${mode}`,
    loadFailed:      "Error al cargar",
    initError:       (msg) => `Error de inicialización: ${msg}`,
    thinking:        "Claude está pensando…",
    send:            "Enviar",
    sending:         "…",
    serverUnreach:   "No se puede conectar al servidor. ¿Está ejecutándose start.sh?",
    httpError:       (status) => `Error HTTP ${status}`,
    sectionLabel:    "Sección",
    sectionTarget:   "Sección objetivo",
    optional:        "(opcional)",
    placeholder:     "Haga una pregunta a Claude…",
    sectionAll:      "— documento completo —",
    selectionChip:   "Texto seleccionado:",
  },
};

let t = STRINGS.en; // active locale strings, set in initLocale()

function initLocale() {
  const locale = (Office.context.displayLanguage || "en-US").substring(0, 2).toLowerCase();
  t = STRINGS[locale] || STRINGS.en;
  document.getElementById("load-doc-btn").textContent = t.loadBtn;
  document.getElementById("init-label").textContent   = t.docNotLoaded;
  document.getElementById("question").placeholder     = t.placeholder;
  document.getElementById("ask-btn").textContent      = t.send;
  const labels = document.querySelectorAll("[data-i18n]");
  labels.forEach((el) => {
    const key = el.dataset.i18n;
    if (t[key]) el.textContent = t[key];
  });
}

// ---------------------------------------------------------------------------
// Office theme — use isDarkTheme as reliable signal, apply our own palette
// ---------------------------------------------------------------------------
function applyTheme(isDark) {
  document.documentElement.setAttribute("data-theme", isDark ? "dark" : "light");
}

function initTheme() {
  try {
    const theme = Office.context.officeTheme;
    if (theme && typeof theme.isDarkTheme !== "undefined") {
      applyTheme(theme.isDarkTheme);
      Office.context.officeTheme.addHandlerAsync(
        Office.EventType.OfficeThemeChanged,
        (e) => applyTheme(e.officeTheme && e.officeTheme.isDarkTheme)
      );
      return;
    }
    // isDarkTheme not available — try to infer from bodyBackgroundColor luminance
    if (theme && theme.bodyBackgroundColor) {
      const hex = theme.bodyBackgroundColor.replace("#", "");
      const r = parseInt(hex.substring(0, 2), 16);
      const g = parseInt(hex.substring(2, 4), 16);
      const b = parseInt(hex.substring(4, 6), 16);
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      applyTheme(luminance < 0.5);
    }
  } catch {
    // Fallback: CSS prefers-color-scheme media query handles it automatically
  }
}

// ---------------------------------------------------------------------------
// Office.onReady
// ---------------------------------------------------------------------------
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initTheme();
    initLocale();

    document.getElementById("ask-btn").addEventListener("click", onAsk);
    document.getElementById("load-doc-btn").addEventListener("click", loadDocument);
    document.getElementById("selection-chip-clear").addEventListener("click", clearSelectionChip);
    document.getElementById("clear-section-btn").addEventListener("click", () => {
      const sel = document.getElementById("section-number");
      sel.value = "";
      delete sel.dataset.userEdited;
    });
    document.getElementById("section-number").addEventListener("change", (e) => {
      if (e.target.value) {
        e.target.dataset.userEdited = "1";
      } else {
        delete e.target.dataset.userEdited;
      }
    });
    document.getElementById("question").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        onAsk();
      }
    });
    document.getElementById("question").addEventListener("input", updateSendButton);

    updateSendButton();
    checkServer().then(() => loadDocument());
    // If the taskpane was opened via the context menu, capture the selection
    askClaudeFromSelection();
  }
});

// ---------------------------------------------------------------------------
// Server health
// ---------------------------------------------------------------------------
async function checkServer() {
  try {
    const resp = await fetch(`${SERVER}/health`);
    if (resp.ok) { setIndicator("idle"); return true; }
  } catch {}
  setIndicator("error");
  showError(t.serverError);
  return false;
}

// ---------------------------------------------------------------------------
// Document loading
// ---------------------------------------------------------------------------
async function loadDocument() {
  documentLoaded = false;
  updateSendButton();
  setInitLabel(t.reading, "loading");
  setIndicator("busy");
  hideError();

  let fullText = "";
  try {
    fullText = await readFullDocument();
  } catch (err) {
    setInitLabel(t.readError, "");
    setIndicator("error");
    showError(t.readError);
    console.error("readFullDocument error:", err);
    return;
  }

  if (!fullText.trim()) {
    setInitLabel(t.emptyDoc, "");
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
      throw new Error(err.detail || t.httpError(resp.status));
    }

    const data = await resp.json();
    sessionId    = data.session_id;
    docMode      = data.mode;
    sectionCount = data.section_count;
    knownSections = data.structure || [];
    documentLoaded = true;
    updateSendButton();

    setInitLabel(t.loadedDoc(data.section_count, data.page_count, data.mode), "ready");
    setDocStatus(`${data.page_count}p`);
    setIndicator("idle");
    document.getElementById("chat-history").innerHTML = "";
    populateSectionDropdown(knownSections);
    startSectionDetection();

  } catch (err) {
    setInitLabel(t.loadFailed, "");
    setIndicator("error");
    showError(t.initError(err.message));
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
// Context menu action — called when user clicks "Demander à Claude"
// ---------------------------------------------------------------------------
async function askClaudeFromSelection() {
  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      const text = sel.text.trim();
      if (text) {
        pendingSelection = text;
        showSelectionChip(text);
      }
    });
  } catch (err) {
    console.error("askClaudeFromSelection error:", err);
  }
}

function showSelectionChip(text) {
  const chip = document.getElementById("selection-chip");
  document.getElementById("selection-chip-text").textContent = text;
  chip.classList.remove("hidden");
}

function clearSelectionChip() {
  pendingSelection = "";
  document.getElementById("selection-chip").classList.add("hidden");
  document.getElementById("selection-chip-text").textContent = "";
}

// ---------------------------------------------------------------------------
// Ask Claude
// ---------------------------------------------------------------------------
async function onAsk() {
  const questionEl = document.getElementById("question");
  const question = questionEl.value.trim();
  if (!question) return;

  const selectedText = pendingSelection;
  const sectionNumber = document.getElementById("section-number").value.trim() || null;

  hideError();
  setIndicator("busy");
  setLoading(true);

  appendMessage("user", question, sectionNumber);
  questionEl.value = "";

  const claudeEl = createStreamingBubble();

  try {
    const resp = await fetch(`${SERVER}/ask-stream`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        question,
        selected_text: selectedText,
        section_number: sectionNumber,
        session_id: sessionId,
      }),
    });

    if (!resp.ok) {
      claudeEl.remove();
      const err = await resp.json().catch(() => ({}));
      showError(err.detail || t.httpError(resp.status));
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
      buffer = lines.pop();

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
    showError(t.serverUnreach);
    setIndicator("error");
    console.error("fetch /ask-stream error:", err);
  } finally {
    setLoading(false);
    clearSelectionChip();
  }
}

// ---------------------------------------------------------------------------
// Section auto-detection
// ---------------------------------------------------------------------------
function startSectionDetection() {
  if (sectionDetectionInterval) clearInterval(sectionDetectionInterval);
  sectionDetectionInterval = setInterval(detectCurrentSection, 2000);
}

async function detectCurrentSection() {
  if (!knownSections.length) return;
  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const para = sel.paragraphs.getFirst();
      para.load("text");
      await context.sync();
      const detectedNumber = findSectionNumberInText(para.text);
      if (detectedNumber) {
        const field = document.getElementById("section-number");
        if (!field.dataset.userEdited) field.value = detectedNumber;
      }
    });
  } catch {}
}

function findSectionNumberInText(text) {
  if (!text || !knownSections.length) return null;
  const trimmed = text.trim();
  for (const s of knownSections) {
    if (s.title && trimmed.toLowerCase().includes(s.title.toLowerCase()) && s.title.length > 3) {
      return s.number;
    }
  }
  const m = trimmed.match(
    /^(?:(?:article|section|chapitre|chapter|partie|part|abschnitt|sección)\s+(\d+(?:\.\d+)*)|((\d+(?:\.\d+)*)[\.\s\-–]))/i
  );
  if (m) {
    const candidate = m[1] || m[3];
    const found = knownSections.find((s) => s.number === candidate);
    if (found) return candidate;
  }
  return null;
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

function appendMessage(role, text, sectionNumber = null) {
  const chat = document.getElementById("chat-history");
  const el = document.createElement("div");
  el.className = `chat-message ${role}`;

  if (sectionNumber && role === "user") {
    const tag = document.createElement("div");
    tag.className = "chat-section-tag";
    tag.textContent = `${t.sectionLabel} ${sectionNumber}`;
    el.appendChild(tag);
  }

  const body = document.createElement("span");
  body.textContent = text;
  el.appendChild(body);

  chat.appendChild(el);
  el.scrollIntoView({ behavior: "smooth", block: "end" });
  return el;
}

function setLoading(isLoading) {
  const btn = document.getElementById("ask-btn");
  btn.textContent = isLoading ? t.sending : t.send;
  if (isLoading) {
    btn.disabled = true;
  } else {
    updateSendButton();
  }
}

function updateSendButton() {
  const btn = document.getElementById("ask-btn");
  const question = document.getElementById("question").value.trim();
  btn.disabled = !documentLoaded || !question;
}

function formatSectionNumber(number) {
  // "1.1.2" → "1 - 1 - 2"
  return number.split(".").join(" - ");
}

function populateSectionDropdown(sections) {
  const sel = document.getElementById("section-number");
  sel.innerHTML = "";
  const defaultOpt = document.createElement("option");
  defaultOpt.value = "";
  defaultOpt.textContent = t.sectionAll || "— document entier —";
  sel.appendChild(defaultOpt);
  for (const s of sections) {
    const opt = document.createElement("option");
    opt.value = s.number;
    const label = formatSectionNumber(s.number);
    opt.textContent = s.title ? `${label} — ${s.title}` : label;
    sel.appendChild(opt);
  }
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
