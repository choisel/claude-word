import asyncio
import json
import logging
import logging.handlers
import os
import time
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
load_dotenv()

CLAUDE_CLI = os.getenv("CLAUDE_CLI_PATH", "/opt/homebrew/bin/claude")
TIMEOUT = int(os.getenv("CLAUDE_TIMEOUT_SECONDS", "120"))
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / "server.log"

fmt = logging.Formatter("%(asctime)s %(levelname)-8s %(name)s - %(message)s")

file_handler = logging.handlers.RotatingFileHandler(
    LOG_FILE, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8"
)
file_handler.setFormatter(fmt)

console_handler = logging.StreamHandler()
console_handler.setFormatter(fmt)

root_logger = logging.getLogger()
root_logger.setLevel(LOG_LEVEL)
root_logger.addHandler(file_handler)
root_logger.addHandler(console_handler)

logger = logging.getLogger("claude-bridge")

# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
app = FastAPI(title="Claude Word Bridge", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
class AskRequest(BaseModel):
    selected_text: str = ""
    question: str


class AskResponse(BaseModel):
    answer: str
    duration_ms: int


# ---------------------------------------------------------------------------
# Claude call
# ---------------------------------------------------------------------------
def build_prompt(selected_text: str, question: str) -> str:
    if selected_text and selected_text.strip():
        return (
            f"Selected text:\n---\n{selected_text}\n---\n\n{question}"
        )
    return question


async def call_claude(prompt: str) -> tuple[str, int]:
    logger.debug("Spawning claude CLI: %s", CLAUDE_CLI)
    t0 = time.monotonic()

    try:
        proc = await asyncio.create_subprocess_exec(
            CLAUDE_CLI,
            "--print",
            "--output-format", "json",
            "--permission-mode", "bypassPermissions",
            stdin=asyncio.subprocess.PIPE,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
    except FileNotFoundError:
        logger.error("Claude CLI not found at: %s", CLAUDE_CLI)
        raise HTTPException(
            status_code=500,
            detail=f"Claude CLI not found at '{CLAUDE_CLI}'. Check CLAUDE_CLI_PATH in .env"
        )

    try:
        stdout, stderr = await asyncio.wait_for(
            proc.communicate(input=prompt.encode("utf-8")),
            timeout=TIMEOUT,
        )
    except asyncio.TimeoutError:
        proc.kill()
        logger.error("Claude CLI timed out after %ds", TIMEOUT)
        raise HTTPException(status_code=504, detail=f"Claude timed out after {TIMEOUT}s")

    duration_ms = int((time.monotonic() - t0) * 1000)

    if stderr:
        logger.warning("Claude CLI stderr: %s", stderr.decode("utf-8", errors="replace").strip())

    if proc.returncode != 0:
        raw_err = stderr.decode("utf-8", errors="replace").strip()
        logger.error("Claude CLI exited with code %d. stderr: %s", proc.returncode, raw_err)
        raise HTTPException(
            status_code=500,
            detail=f"Claude CLI error (exit {proc.returncode}): {raw_err}"
        )

    raw = stdout.decode("utf-8", errors="replace").strip()

    # Try to parse JSON output; fall back to raw text
    try:
        data = json.loads(raw)
        # The JSON format has a "result" field
        answer = data.get("result") or data.get("response") or raw
    except json.JSONDecodeError:
        logger.debug("Claude output is not JSON, using raw text")
        answer = raw

    return answer, duration_ms


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------
@app.get("/health")
async def health():
    return {"status": "ok", "claude_cli": CLAUDE_CLI}


@app.post("/ask", response_model=AskResponse)
async def ask(req: AskRequest, request: Request):
    client = request.client.host if request.client else "unknown"
    logger.info(
        "[REQUEST] client=%s question_length=%d selected_text_length=%d | question=%r",
        client,
        len(req.question),
        len(req.selected_text),
        req.question[:120],
    )

    if not req.question.strip():
        raise HTTPException(status_code=400, detail="question cannot be empty")

    prompt = build_prompt(req.selected_text, req.question)
    logger.debug("Prompt sent to Claude (%d chars):\n%s", len(prompt), prompt[:500])

    answer, duration_ms = await call_claude(prompt)

    logger.info(
        "[RESPONSE] duration_ms=%d answer_length=%d | answer_preview=%r",
        duration_ms,
        len(answer),
        answer[:120],
    )

    return AskResponse(answer=answer, duration_ms=duration_ms)


# ---------------------------------------------------------------------------
# Static files (add-in)
# ---------------------------------------------------------------------------
ADDIN_DIR = Path(__file__).parent.parent / "addin"
if ADDIN_DIR.exists():
    app.mount("/", StaticFiles(directory=str(ADDIN_DIR), html=True), name="static")
    logger.info("Serving add-in static files from: %s", ADDIN_DIR)
else:
    logger.warning("addin/ directory not found at %s — static files not served", ADDIN_DIR)
