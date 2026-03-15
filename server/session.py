"""
In-memory session store.
Each session holds the analysed document state and conversation history.
"""

import time
import uuid
from dataclasses import dataclass, field
from typing import Optional

# Sessions older than 4 hours are purged automatically
SESSION_TTL_SECONDS = 4 * 60 * 60


@dataclass
class Exchange:
    role: str          # "user" | "claude"
    content: str
    section_number: Optional[float] = None   # None = no specific section


@dataclass
class DocumentSection:
    number: str        # raw label, e.g. "1", "2.3", "Article 5"
    sort_key: float    # numeric sort key for proximity calculation
    title: str         # heading text following the number
    content: str       # body text of the section


@dataclass
class Session:
    session_id: str
    mode: str                              # "full" | "summarized"
    full_text: str = ""                    # used in "full" mode
    summary: str = ""                      # used in "summarized" mode
    structure_text: str = ""               # human-readable structure overview
    sections: list[DocumentSection] = field(default_factory=list)
    history: list[Exchange] = field(default_factory=list)
    page_count: int = 0
    section_count: int = 0
    created_at: float = field(default_factory=time.time)
    updated_at: float = field(default_factory=time.time)


# ---------------------------------------------------------------------------
# Store
# ---------------------------------------------------------------------------
_sessions: dict[str, Session] = {}


def create_session() -> Session:
    sid = str(uuid.uuid4())
    session = Session(session_id=sid, mode="full")
    _sessions[sid] = session
    return session


def get_session(session_id: str) -> Optional[Session]:
    _purge_expired()
    return _sessions.get(session_id)


def get_or_create_session(session_id: Optional[str]) -> Session:
    if session_id:
        session = get_session(session_id)
        if session:
            return session
    return create_session()


def add_exchange(session: Session, role: str, content: str, section_number: Optional[float] = None):
    session.history.append(Exchange(role=role, content=content, section_number=section_number))
    session.updated_at = time.time()


def get_relevant_history(session: Session, section_number: Optional[float], proximity: int = 5) -> list[Exchange]:
    """
    Return history exchanges relevant to the current section.
    - If no section_number: return last 3 exchanges.
    - If section_number given: return exchanges whose section is within
      `proximity` of the current section, plus the last exchange always.
    """
    if not session.history:
        return []

    if section_number is None:
        return session.history[-3:]

    relevant = []
    for ex in session.history:
        if ex.section_number is None:
            # General exchange — always include if recent (last 2)
            continue
        if abs(ex.section_number - section_number) <= proximity:
            relevant.append(ex)

    # Always include the last exchange for conversational continuity
    if session.history and session.history[-1] not in relevant:
        relevant.append(session.history[-1])

    return relevant


def _purge_expired():
    now = time.time()
    expired = [sid for sid, s in _sessions.items() if now - s.updated_at > SESSION_TTL_SECONDS]
    for sid in expired:
        del _sessions[sid]
