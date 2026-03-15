"""
Document parsing and context building.
Extracts numbered sections, estimates page count, builds prompts.
"""

import re
import logging
from typing import Optional
from session import DocumentSection, Session

logger = logging.getLogger("claude-bridge.document")

# ~1800 characters per page (A4, standard font)
CHARS_PER_PAGE = 1800
SHORT_DOC_PAGE_THRESHOLD = 8

# Matches numbered sections at the start of a line:
# "1.", "1.1", "1.1.1", "Article 1", "Section 2", "ARTICLE 3"
SECTION_RE = re.compile(
    r'^(?:'
    r'(?:article|section|chapitre|chapter|partie|part)\s+(\d+(?:\.\d+)*)'  # keyword + number
    r'|(\d+(?:\.\d+)*)\.'                                                    # "1." or "1.1."
    r'|(\d+(?:\.\d+)*)\s*[-–—]'                                             # "1 -" or "1.1 —"
    r')',
    re.IGNORECASE | re.MULTILINE
)


def estimate_pages(text: str) -> int:
    return max(1, round(len(text) / CHARS_PER_PAGE))


def extract_sections(text: str) -> list[DocumentSection]:
    """
    Parse the document text and extract numbered sections.
    Returns a list of DocumentSection ordered by appearance.
    """
    lines = text.splitlines()
    sections: list[DocumentSection] = []
    current_section: Optional[dict] = None

    for line in lines:
        m = SECTION_RE.match(line.strip())
        if m:
            # Save previous section
            if current_section is not None:
                sections.append(_make_section(current_section))

            raw_number = m.group(1) or m.group(2) or m.group(3)
            # Title = rest of the line after the matched prefix
            title = line.strip()[m.end():].strip(" :-–—")
            current_section = {
                "number": raw_number,
                "title": title,
                "lines": [],
            }
        elif current_section is not None:
            current_section["lines"].append(line)

    if current_section is not None:
        sections.append(_make_section(current_section))

    logger.info("Extracted %d sections from document", len(sections))
    return sections


def _make_section(data: dict) -> DocumentSection:
    raw = data["number"]
    try:
        # Convert "1.2.3" → 1.0203 for numeric proximity comparisons
        parts = [int(p) for p in raw.split(".")]
        sort_key = parts[0] + sum(p / (100 ** (i)) for i, p in enumerate(parts[1:], 1))
    except ValueError:
        sort_key = 0.0

    return DocumentSection(
        number=raw,
        sort_key=sort_key,
        title=data["title"],
        content="\n".join(data["lines"]).strip(),
    )


def build_structure_text(sections: list[DocumentSection]) -> str:
    """Human-readable outline of the document structure."""
    lines = ["Document structure:"]
    for s in sections:
        indent = "  " * (s.number.count("."))
        title_part = f" — {s.title}" if s.title else ""
        lines.append(f"{indent}[{s.number}]{title_part}")
    return "\n".join(lines)


def find_section(sections: list[DocumentSection], number: str) -> Optional[DocumentSection]:
    """Find a section by its number label (exact or prefix match)."""
    number = number.strip()
    for s in sections:
        if s.number == number:
            return s
    # Fallback: prefix match (e.g. "1" matches "1.0")
    for s in sections:
        if s.number.startswith(number + ".") or s.number == number:
            return s
    return None


def get_neighboring_sections(
    sections: list[DocumentSection],
    target: DocumentSection,
    radius: int = 2,
) -> list[DocumentSection]:
    """Return target section plus `radius` sections before and after."""
    try:
        idx = sections.index(target)
    except ValueError:
        return [target]
    start = max(0, idx - radius)
    end = min(len(sections), idx + radius + 1)
    return sections[start:end]


# ---------------------------------------------------------------------------
# Prompt builders
# ---------------------------------------------------------------------------

def build_init_prompt(text: str) -> str:
    return (
        "You are analysing a document. "
        "Provide a concise summary (3-5 sentences) of the document's purpose and main themes, "
        "followed by a structured outline of its numbered sections.\n\n"
        f"Document:\n---\n{text}\n---"
    )


def build_ask_prompt(
    session: Session,
    question: str,
    section_number: Optional[str],
    history_exchanges: list,
) -> str:
    parts: list[str] = []

    if session.mode == "full":
        parts.append(f"Document:\n---\n{session.full_text}\n---\n")
    else:
        parts.append(f"Document summary:\n{session.summary}\n")
        parts.append(f"{session.structure_text}\n")

        if section_number:
            target = find_section(session.sections, section_number)
            if target:
                neighbors = get_neighboring_sections(session.sections, target)
                section_block = "\n\n".join(
                    f"[Section {s.number}]{' — ' + s.title if s.title else ''}\n{s.content}"
                    for s in neighbors
                )
                parts.append(f"Relevant sections:\n---\n{section_block}\n---\n")
            else:
                parts.append(f"(Section {section_number} not found in document)\n")

    # Conversation history
    if history_exchanges:
        history_lines = []
        for ex in history_exchanges:
            prefix = "User" if ex.role == "user" else "Assistant"
            history_lines.append(f"{prefix}: {ex.content}")
        parts.append("Previous exchanges:\n" + "\n".join(history_lines) + "\n")

    parts.append(f"Question: {question}")
    return "\n".join(parts)
