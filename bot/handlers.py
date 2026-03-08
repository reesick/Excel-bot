"""
Handlers — The brain of the Telegram bot.
Orchestrates Claude Analyst → Excel Skill → PPT Skill pipeline.
No questions asked to user — Claude figures everything out.
"""

import asyncio
import logging
import os
import re
import time
import traceback

import httpx

from bot import sessions
from bot.sessions import (
    IDLE,
    EXCEL_ANALYZING,
    EXCEL_GENERATING,
    PPT_GENERATING,
)

# Skill imports — wrapped in try/except for clear startup errors
try:
    from skills.analyst import analyze_and_plan, ask_question
except ImportError as e:
    logging.error("Failed to import analyst skill: %s", e)
    analyze_and_plan = ask_question = None

try:
    from skills.excel_skill import generate_excel
except ImportError as e:
    logging.error("Failed to import excel_skill: %s", e)
    generate_excel = None

try:
    from skills.ppt_skill import generate_ppt
except ImportError as e:
    logging.error("Failed to import ppt_skill: %s", e)
    generate_ppt = None


logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"


# ─────────────────────────────────────────────────────────────────────────────
#  Telegram API Helpers
# ─────────────────────────────────────────────────────────────────────────────

async def send_message(chat_id: int, text: str, parse_mode: str = "Markdown") -> None:
    """Send a text message via Telegram Bot API with retry on 429."""
    payload = {"chat_id": chat_id, "text": text, "parse_mode": parse_mode}
    for attempt in range(3):
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.post(f"{TELEGRAM_API}/sendMessage", json=payload)
                if resp.status_code == 429:
                    retry_after = resp.json().get("parameters", {}).get("retry_after", 5)
                    logger.warning("Rate limited, retrying after %s seconds", retry_after)
                    await asyncio.sleep(retry_after)
                    continue
                if resp.status_code != 200:
                    logger.error("sendMessage failed: %s %s", resp.status_code, resp.text)
                return
        except Exception:
            logger.error("sendMessage error (attempt %d): %s", attempt + 1, traceback.format_exc())
            if attempt < 2:
                await asyncio.sleep(1)


async def send_document(chat_id: int, file_path: str, caption: str = "") -> None:
    """Send a file via Telegram Bot API sendDocument with retry on 429."""
    for attempt in range(3):
        try:
            async with httpx.AsyncClient(timeout=120) as client:
                with open(file_path, "rb") as f:
                    files = {"document": (os.path.basename(file_path), f)}
                    data = {"chat_id": str(chat_id)}
                    if caption:
                        data["caption"] = caption
                        data["parse_mode"] = "Markdown"
                    resp = await client.post(
                        f"{TELEGRAM_API}/sendDocument", data=data, files=files
                    )
                    if resp.status_code == 429:
                        retry_after = resp.json().get("parameters", {}).get("retry_after", 5)
                        await asyncio.sleep(retry_after)
                        continue
                    if resp.status_code != 200:
                        logger.error("sendDocument failed: %s %s", resp.status_code, resp.text)
                    return
        except Exception:
            logger.error("sendDocument error (attempt %d): %s", attempt + 1, traceback.format_exc())
            if attempt < 2:
                await asyncio.sleep(1)


async def send_typing(chat_id: int, action: str = "typing") -> None:
    """Send a chat action indicator."""
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            await client.post(
                f"{TELEGRAM_API}/sendChatAction",
                json={"chat_id": chat_id, "action": action},
            )
    except Exception:
        pass


async def _keep_typing(chat_id: int, stop_event: asyncio.Event) -> None:
    """Background task: send typing indicator every 5 seconds until stopped."""
    while not stop_event.is_set():
        await send_typing(chat_id)
        try:
            await asyncio.wait_for(stop_event.wait(), timeout=5)
        except asyncio.TimeoutError:
            pass


# ─────────────────────────────────────────────────────────────────────────────
#  File Download Helper
# ─────────────────────────────────────────────────────────────────────────────

async def _download_telegram_file(file_id: str, save_path: str) -> str:
    """Download a file from Telegram servers to a local path."""
    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.get(f"{TELEGRAM_API}/getFile", params={"file_id": file_id})
        resp.raise_for_status()
        file_path = resp.json()["result"]["file_path"]
        file_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
        resp = await client.get(file_url)
        resp.raise_for_status()
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        with open(save_path, "wb") as f:
            f.write(resp.content)
    return save_path


# ─────────────────────────────────────────────────────────────────────────────
#  Main Dispatch
# ─────────────────────────────────────────────────────────────────────────────

async def handle_update(update: dict) -> None:
    """Route an incoming Telegram Update to the appropriate handler."""
    message = update.get("message")
    if not message:
        return

    chat_id = message["chat"]["id"]
    text = message.get("text", "").strip()
    document = message.get("document")
    session = sessions.get_session(chat_id)

    logger.info("Update from chat_id=%s | type=%s | state=%s",
                chat_id, "document" if document else "text", session["state"])

    try:
        # /start or /help
        if text.lower() in ("/start", "/help"):
            await handle_start(chat_id)
            return

        # /cancel or /reset
        if text.lower() in ("/cancel", "/reset"):
            await handle_cancel(chat_id)
            return

        # File upload
        if document:
            await handle_file_upload(chat_id, document)
            return

        # PPT request keywords
        if text and re.search(r"\b(ppt|presentation|slides)\b", text, re.IGNORECASE):
            await handle_ppt_request(chat_id, text)
            return

        # Data question (if a file was previously analyzed)
        if text and session.get("analysis"):
            await handle_data_question(chat_id, text)
            return

        # Fallback
        await send_message(
            chat_id,
            "👋 I'm your AI Office Assistant!\n\n"
            "Here's what I can do:\n"
            "📊 *Upload any Excel or CSV file* — I'll analyze it, build a dashboard AND a presentation\n"
            "🎨 *Type 'make a ppt'* — Create a presentation from your data\n"
            "❓ *Ask me anything* about your data after uploading\n"
            "❌ */cancel* — Reset and start over\n\n"
            "_Send me a file to get started!_",
        )

    except Exception:
        logger.error("Unhandled error for chat_id=%s:\n%s", chat_id, traceback.format_exc())
        await send_message(chat_id, "❌ Something unexpected went wrong. Please try again or type /cancel to reset.")
        sessions.clear_session(chat_id)


# ─────────────────────────────────────────────────────────────────────────────
#  Individual Handlers
# ─────────────────────────────────────────────────────────────────────────────

async def handle_start(chat_id: int) -> None:
    """Send welcome message."""
    await send_message(
        chat_id,
        "👋 *Welcome to Office Bot!*\n\n"
        "I'm your AI-powered business analyst. Send me *any* business file "
        "and I'll give you:\n\n"
        "📊 A *professional Excel dashboard* with charts, KPIs, pivots, and insights\n"
        "🎨 A *polished PowerPoint presentation* ready for your next meeting\n\n"
        "I handle CSV, XLS, and XLSX files — clean or messy, any industry.\n\n"
        "Just *upload a file* and I'll take care of everything. No questions asked.\n\n"
        "Type */cancel* at any time to start over.",
    )


async def handle_file_upload(chat_id: int, document: dict) -> None:
    """Process an uploaded file: analyze → generate dashboard + PPT → send both."""
    file_name = document.get("file_name", "unknown")
    file_size = document.get("file_size", 0)
    file_id = document["file_id"]

    # Validate extension
    ext = os.path.splitext(file_name)[1].lower()
    if ext not in (".xlsx", ".xls", ".csv"):
        await send_message(
            chat_id,
            "⚠️ I can process *Excel* (`.xlsx`, `.xls`) and *CSV* (`.csv`) files. "
            "Please upload a valid business data file.",
        )
        return

    # Validate size (15 MB cap)
    if file_size > 15 * 1024 * 1024:
        await send_message(chat_id, "⚠️ That file is too large (max 15 MB). Please upload a smaller file.")
        return

    # Download the file
    timestamp = int(time.time())
    save_path = f"/tmp/uploads/{chat_id}_{timestamp}_{file_name}"
    try:
        await _download_telegram_file(file_id, save_path)
    except Exception:
        logger.error("File download failed:\n%s", traceback.format_exc())
        await send_message(chat_id, "❌ Failed to download the file. Please try uploading again.")
        return

    session_id = f"{chat_id}_{timestamp}"
    sessions.update_session(chat_id, file_path=save_path, state=EXCEL_ANALYZING)

    await send_message(chat_id, "📊 Got it! My analyst is reviewing your file...")

    # Start typing indicator background task
    stop_typing = asyncio.Event()
    typing_task = asyncio.ensure_future(_keep_typing(chat_id, stop_typing))

    # ── Step 1: Claude Analysis ──────────────────────────────────────────
    try:
        loop = asyncio.get_event_loop()
        analysis = await loop.run_in_executor(
            None, analyze_and_plan, save_path, session_id
        )
    except Exception:
        stop_typing.set()
        await typing_task
        logger.error("analyze_and_plan failed:\n%s", traceback.format_exc())
        await send_message(
            chat_id,
            "❌ I couldn't read this file. Please make sure it's a valid "
            "Excel or CSV file and try again.",
        )
        sessions.clear_session(chat_id)
        return

    # Store analysis in session for follow-up questions
    sessions.update_session(chat_id, analysis=analysis)

    # Send Claude's summary
    summary = analysis.get("analysis_summary", "Analysis complete.")
    await send_message(chat_id, f"✅ *Analysis Complete*\n\n{summary}")
    await send_message(chat_id, "⚙️ Generating your dashboard and presentation now...")

    # ── Step 2: Generate Excel Dashboard ─────────────────────────────────
    sessions.update_session(chat_id, state=EXCEL_GENERATING)

    excel_output = f"/tmp/outputs/{session_id}_dashboard.xlsx"
    analysis["output_path"] = excel_output

    excel_path = None
    try:
        excel_path = await loop.run_in_executor(None, generate_excel, analysis)
    except Exception:
        logger.error("generate_excel failed:\n%s", traceback.format_exc())
        # Non-fatal — continue with PPT

    # ── Step 3: Generate PPT ─────────────────────────────────────────────
    sessions.update_session(chat_id, state=PPT_GENERATING)

    ppt_output = f"/tmp/outputs/{session_id}_presentation.pptx"
    analysis["output_path"] = ppt_output

    ppt_path = None
    try:
        ppt_path = await loop.run_in_executor(None, generate_ppt, analysis)
    except Exception:
        logger.error("generate_ppt failed:\n%s", traceback.format_exc())
        # Non-fatal — send what we have

    # Stop typing indicator
    stop_typing.set()
    await typing_task

    # ── Step 4: Send results ─────────────────────────────────────────────
    files_sent = 0

    if excel_path and os.path.exists(excel_path):
        await send_document(chat_id, excel_path, caption="📊 Your Excel Dashboard")
        sessions.update_session(chat_id, excel_output_path=excel_path)
        files_sent += 1

    if ppt_path and os.path.exists(ppt_path):
        await send_document(chat_id, ppt_path, caption="🎨 Your PowerPoint Presentation")
        files_sent += 1

    if files_sent == 0:
        await send_message(
            chat_id,
            "❌ Something went wrong while generating your reports. Please try again.",
        )
        sessions.clear_session(chat_id)
        return

    # Send key insights as follow-up message
    insights = analysis.get("insights", [])
    if insights:
        insight_text = "\n".join(f"• {ins}" for ins in insights[:4])
        await send_message(
            chat_id,
            f"💡 *Key Findings:*\n\n{insight_text}\n\n"
            "_You can ask me any question about your data, or upload another file._",
        )
    else:
        await send_message(
            chat_id,
            "✅ All done! You can ask me questions about your data, "
            "or upload another file.",
        )

    sessions.update_session(chat_id, state=IDLE)


async def handle_ppt_request(chat_id: int, text: str) -> None:
    """Handle PPT regeneration request."""
    session = sessions.get_session(chat_id)

    if not session.get("analysis"):
        await send_message(
            chat_id,
            "📎 Please upload a file first — I'll analyze it and create "
            "both a dashboard and presentation for you.",
        )
        return

    analysis = session["analysis"]

    await send_message(chat_id, "🎨 Creating your presentation...")
    sessions.update_session(chat_id, state=PPT_GENERATING)

    stop_typing = asyncio.Event()
    typing_task = asyncio.ensure_future(_keep_typing(chat_id, stop_typing))

    timestamp = int(time.time())
    ppt_output = f"/tmp/outputs/{chat_id}_{timestamp}_presentation.pptx"
    analysis["output_path"] = ppt_output

    try:
        loop = asyncio.get_event_loop()
        ppt_path = await loop.run_in_executor(None, generate_ppt, analysis)
    except Exception:
        stop_typing.set()
        await typing_task
        logger.error("generate_ppt failed:\n%s", traceback.format_exc())
        await send_message(chat_id, "❌ Something went wrong. Please try again.")
        sessions.update_session(chat_id, state=IDLE)
        return

    stop_typing.set()
    await typing_task

    await send_document(chat_id, ppt_path, caption="🎨 Your PowerPoint Presentation")
    sessions.update_session(chat_id, state=IDLE)


async def handle_data_question(chat_id: int, question: str) -> None:
    """Answer a question about previously analyzed data using Claude."""
    session = sessions.get_session(chat_id)
    analysis = session.get("analysis", {})
    csv_path = analysis.get("file_path")

    if not csv_path or not os.path.exists(csv_path):
        await send_message(
            chat_id,
            "📎 I don't have your data loaded. Please upload a file first.",
        )
        return

    await send_typing(chat_id)
    session_id = f"{chat_id}_{int(time.time())}"

    try:
        loop = asyncio.get_event_loop()
        answer = await loop.run_in_executor(
            None, ask_question, question, csv_path, session_id
        )
    except Exception:
        logger.error("ask_question failed:\n%s", traceback.format_exc())
        await send_message(
            chat_id,
            "❌ I couldn't answer that question. Please try rephrasing it.",
        )
        return

    await send_message(chat_id, answer)


async def handle_cancel(chat_id: int) -> None:
    """Cancel current operation and clean up."""
    session = sessions.get_session(chat_id)

    # Clean up temp files
    for key in ("file_path", "excel_output_path"):
        fpath = session.get(key)
        if fpath and os.path.exists(fpath):
            try:
                os.remove(fpath)
            except Exception:
                pass

    # Clean up analyst output dir
    output_dir = session.get("analysis", {}).get("output_dir")
    if output_dir and os.path.exists(output_dir):
        import shutil
        try:
            shutil.rmtree(output_dir)
        except Exception:
            pass

    sessions.clear_session(chat_id)
    await send_message(
        chat_id,
        "✅ Cancelled. Upload a new file whenever you're ready.",
    )


# ─────────────────────────────────────────────────────────────────────────────
#  Periodic Cleanup Task
# ─────────────────────────────────────────────────────────────────────────────

def _cleanup_temp_files() -> int:
    """Delete files in /tmp/uploads and /tmp/outputs older than 2 hours."""
    count = 0
    cutoff = time.time() - (2 * 60 * 60)
    for folder in ("/tmp/uploads", "/tmp/outputs", "/tmp/analyst_output"):
        if not os.path.exists(folder):
            continue
        for fname in os.listdir(folder):
            fpath = os.path.join(folder, fname)
            try:
                if os.path.isfile(fpath) and os.path.getmtime(fpath) < cutoff:
                    os.remove(fpath)
                    count += 1
                elif os.path.isdir(fpath) and os.path.getmtime(fpath) < cutoff:
                    import shutil
                    shutil.rmtree(fpath)
                    count += 1
            except Exception:
                pass
    return count


async def _cleanup_loop() -> None:
    """Run cleanup every 30 minutes."""
    while True:
        await asyncio.sleep(30 * 60)
        try:
            stale_sessions = sessions.cleanup_old_sessions()
            stale_files = _cleanup_temp_files()
            logger.info("Cleanup: removed %d sessions, %d temp files",
                        stale_sessions, stale_files)
        except Exception:
            logger.error("Cleanup error:\n%s", traceback.format_exc())


def start_cleanup_task() -> None:
    """Register the periodic cleanup as an asyncio background task."""
    asyncio.ensure_future(_cleanup_loop())
