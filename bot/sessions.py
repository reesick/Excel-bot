"""
Session Manager — In-memory conversation state tracker.
Tracks where each Telegram user is in the conversation flow.
"""

import time

# ─── State Constants ───────────────────────────────────────────────────────────

IDLE = "idle"
EXCEL_ANALYZING = "excel_analyzing"
EXCEL_AWAITING_LEVEL = "excel_awaiting_level"
EXCEL_AWAITING_ANSWERS = "excel_awaiting_answers"
EXCEL_GENERATING = "excel_generating"
PPT_AWAITING_ANSWERS = "ppt_awaiting_answers"
PPT_GENERATING = "ppt_generating"

# ─── Session Store ─────────────────────────────────────────────────────────────

_sessions: dict[int, dict] = {}

SESSION_TTL_SECONDS = 2 * 60 * 60  # 2 hours


def _new_session() -> dict:
    """Create a fresh session with default values."""
    return {
        "state": IDLE,
        "file_path": None,
        "analysis": None,
        "level": None,
        "excel_output_path": None,
        "chart_paths": [],
        "last_activity": time.time(),
    }


def get_session(chat_id: int) -> dict:
    """Return the session for this chat_id, creating one if it doesn't exist."""
    if chat_id not in _sessions:
        _sessions[chat_id] = _new_session()
    _sessions[chat_id]["last_activity"] = time.time()
    return _sessions[chat_id]


def update_session(chat_id: int, **kwargs) -> None:
    """Update specific fields in an existing session."""
    session = get_session(chat_id)
    session.update(kwargs)
    session["last_activity"] = time.time()


def clear_session(chat_id: int) -> None:
    """Reset session to IDLE, clearing all stored data but keeping the dict entry."""
    _sessions[chat_id] = _new_session()


def cleanup_old_sessions() -> int:
    """
    Remove sessions inactive for more than SESSION_TTL_SECONDS.
    Returns the number of sessions removed.
    """
    now = time.time()
    stale = [
        cid for cid, s in _sessions.items()
        if now - s["last_activity"] > SESSION_TTL_SECONDS
    ]
    for cid in stale:
        del _sessions[cid]
    return len(stale)
