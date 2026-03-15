import asyncio
import json
import logging
import logging.handlers
import os
import tempfile
import time
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

import document as doc_module
import session as session_module

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
app = FastAPI(title="Claude Word Bridge", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
class InitRequest(BaseModel):
    text: str
    session_id: Optional[str] = None


class InitResponse(BaseModel):
    session_id: str
    page_count: int
    section_count: int
    mode: str          # "full" | "summarized"
    structure: list[dict]


class AskRequest(BaseModel):
    question: str
    selected_text: str = ""
    section_number: Optional[str] = None
    session_id: Optional[str] = None


class AskResponse(BaseModel):
    answer: str
    duration_ms: int
    session_id: str


# ---------------------------------------------------------------------------
# Claude call
# ---------------------------------------------------------------------------
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
            cwd=tempfile.gettempdir(),
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
        stderr_text = stderr.decode("utf-8", errors="replace").strip()
        if stderr_text:
            logger.warning("Claude CLI stderr: %s", stderr_text)

    if proc.returncode != 0:
        raw_err = stderr.decode("utf-8", errors="replace").strip()
        logger.error("Claude CLI exited with code %d. stderr: %s", proc.returncode, raw_err)
        raise HTTPException(
            status_code=500,
            detail=f"Claude CLI error (exit {proc.returncode}): {raw_err}"
        )

    raw = stdout.decode("utf-8", errors="replace").strip()

    try:
        data = json.loads(raw)
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


@app.post("/init", response_model=InitResponse)
async def init_document(req: InitRequest, request: Request):
    client = request.client.host if request.client else "unknown"
    page_count = doc_module.estimate_pages(req.text)
    sections = doc_module.extract_sections(req.text)

    logger.info(
        "[INIT] client=%s pages=%d sections=%d text_length=%d",
        client, page_count, len(sections), len(req.text)
    )

    session = session_module.get_or_create_session(req.session_id)
    session.page_count = page_count
    session.section_count = len(sections)
    session.sections = sections
    session.history = []  # reset history on document reload

    if page_count < doc_module.SHORT_DOC_PAGE_THRESHOLD:
        session.mode = "full"
        session.full_text = req.text
        session.summary = ""
        session.structure_text = doc_module.build_structure_text(sections)
        logger.info("[INIT] mode=full, no summarization needed")
    else:
        session.mode = "summarized"
        session.full_text = ""
        session.structure_text = doc_module.build_structure_text(sections)
        logger.info("[INIT] mode=summarized, calling Claude for summary...")
        init_prompt = doc_module.build_init_prompt(req.text)
        summary, duration_ms = await call_claude(init_prompt)
        session.summary = summary
        logger.info("[INIT] summary generated in %dms (%d chars)", duration_ms, len(summary))

    structure_list = [
        {"number": s.number, "title": s.title, "sort_key": s.sort_key}
        for s in sections
    ]

    return InitResponse(
        session_id=session.session_id,
        page_count=page_count,
        section_count=len(sections),
        mode=session.mode,
        structure=structure_list,
    )


@app.post("/ask", response_model=AskResponse)
async def ask(req: AskRequest, request: Request):
    client = request.client.host if request.client else "unknown"
    logger.info(
        "[REQUEST] client=%s question_length=%d section=%s session=%s | question=%r",
        client,
        len(req.question),
        req.section_number or "none",
        req.session_id or "none",
        req.question[:120],
    )

    if not req.question.strip():
        raise HTTPException(status_code=400, detail="question cannot be empty")

    session = session_module.get_session(req.session_id) if req.session_id else None

    # Determine section sort key for history filtering
    section_sort_key: Optional[float] = None
    if session and req.section_number:
        target = doc_module.find_section(session.sections, req.section_number)
        if target:
            section_sort_key = target.sort_key

    # Build prompt
    if session:
        history = session_module.get_relevant_history(session, section_sort_key)
        prompt = doc_module.build_ask_prompt(session, req.question, req.section_number, history)
    else:
        # No session — fall back to simple prompt with selected text
        if req.selected_text.strip():
            prompt = f"Selected text:\n---\n{req.selected_text}\n---\n\nQuestion: {req.question}"
        else:
            prompt = req.question

    logger.debug("Prompt sent to Claude (%d chars):\n%s", len(prompt), prompt[:800])

    answer, duration_ms = await call_claude(prompt)

    # Save to history
    if session:
        session_module.add_exchange(session, "user", req.question, section_sort_key)
        session_module.add_exchange(session, "claude", answer, section_sort_key)

    logger.info(
        "[RESPONSE] duration_ms=%d answer_length=%d | answer_preview=%r",
        duration_ms,
        len(answer),
        answer[:120],
    )

    return AskResponse(
        answer=answer,
        duration_ms=duration_ms,
        session_id=session.session_id if session else "",
    )


# ---------------------------------------------------------------------------
# Streaming endpoint
# ---------------------------------------------------------------------------
async def stream_claude(prompt: str, session: object, section_sort_key: Optional[float], question: str):
    """
    Generator that spawns Claude with stream-json output and yields SSE-formatted lines.
    Emits: data: {"type":"token","text":"..."}
           data: {"type":"done","duration_ms":N}
           data: {"type":"error","detail":"..."}
    """
    t0 = time.monotonic()
    full_answer = []

    try:
        proc = await asyncio.create_subprocess_exec(
            CLAUDE_CLI,
            "--print",
            "--output-format", "stream-json",
            "--verbose",
            "--permission-mode", "bypassPermissions",
            stdin=asyncio.subprocess.PIPE,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            cwd=tempfile.gettempdir(),
        )
    except FileNotFoundError:
        yield f"data: {json.dumps({'type': 'error', 'detail': f'Claude CLI not found at {CLAUDE_CLI}'})}\n\n"
        return

    # Write prompt to stdin and close it
    proc.stdin.write(prompt.encode("utf-8"))
    await proc.stdin.drain()
    proc.stdin.close()

    try:
        async def read_lines():
            async for raw_line in proc.stdout:
                line = raw_line.decode("utf-8", errors="replace").strip()
                if not line:
                    continue
                try:
                    chunk = json.loads(line)
                except json.JSONDecodeError:
                    continue

                chunk_type = chunk.get("type", "")

                # stream-json emits assistant text in content_block_delta events
                if chunk_type == "assistant":
                    for block in chunk.get("message", {}).get("content", []):
                        if block.get("type") == "text":
                            text = block["text"]
                            full_answer.append(text)
                            yield f"data: {json.dumps({'type': 'token', 'text': text})}\n\n"

                elif chunk_type == "content_block_delta":
                    delta = chunk.get("delta", {})
                    if delta.get("type") == "text_delta":
                        text = delta.get("text", "")
                        if text:
                            full_answer.append(text)
                            yield f"data: {json.dumps({'type': 'token', 'text': text})}\n\n"

                elif chunk_type == "result":
                    # Final result block — may contain the full text too
                    result_text = chunk.get("result", "")
                    if result_text and not full_answer:
                        full_answer.append(result_text)
                        yield f"data: {json.dumps({'type': 'token', 'text': result_text})}\n\n"

        async for item in read_lines():
            yield item

        await asyncio.wait_for(proc.wait(), timeout=TIMEOUT)

    except asyncio.TimeoutError:
        proc.kill()
        yield f"data: {json.dumps({'type': 'error', 'detail': f'Claude timed out after {TIMEOUT}s'})}\n\n"
        return

    duration_ms = int((time.monotonic() - t0) * 1000)
    final_answer = "".join(full_answer)

    # Save to session history
    if session and final_answer:
        session_module.add_exchange(session, "user", question, section_sort_key)
        session_module.add_exchange(session, "claude", final_answer, section_sort_key)

    logger.info(
        "[STREAM RESPONSE] duration_ms=%d answer_length=%d | preview=%r",
        duration_ms, len(final_answer), final_answer[:120],
    )

    yield f"data: {json.dumps({'type': 'done', 'duration_ms': duration_ms})}\n\n"


@app.post("/ask-stream")
async def ask_stream(req: AskRequest, request: Request):
    client = request.client.host if request.client else "unknown"
    logger.info(
        "[STREAM REQUEST] client=%s question_length=%d section=%s session=%s | question=%r",
        client, len(req.question),
        req.section_number or "none",
        req.session_id or "none",
        req.question[:120],
    )

    if not req.question.strip():
        raise HTTPException(status_code=400, detail="question cannot be empty")

    session = session_module.get_session(req.session_id) if req.session_id else None

    section_sort_key: Optional[float] = None
    if session and req.section_number:
        target = doc_module.find_section(session.sections, req.section_number)
        if target:
            section_sort_key = target.sort_key

    if session:
        history = session_module.get_relevant_history(session, section_sort_key)
        prompt = doc_module.build_ask_prompt(session, req.question, req.section_number, history)
    else:
        if req.selected_text.strip():
            prompt = f"Selected text:\n---\n{req.selected_text}\n---\n\nQuestion: {req.question}"
        else:
            prompt = req.question

    logger.debug("Stream prompt (%d chars):\n%s", len(prompt), prompt[:800])

    return StreamingResponse(
        stream_claude(prompt, session, section_sort_key, req.question),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


# ---------------------------------------------------------------------------
# Static files (add-in)
# ---------------------------------------------------------------------------
ADDIN_DIR = Path(__file__).parent.parent / "addin"
if ADDIN_DIR.exists():
    app.mount("/", StaticFiles(directory=str(ADDIN_DIR), html=True), name="static")
    logger.info("Serving add-in static files from: %s", ADDIN_DIR)
else:
    logger.warning("addin/ directory not found at %s — static files not served", ADDIN_DIR)
