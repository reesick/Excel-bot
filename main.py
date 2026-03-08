"""
Office Bot — Entry Point
Supports two modes based on LOCAL_MODE env variable:
  LOCAL_MODE=true  → Telegram polling (for local development)
  LOCAL_MODE=false → FastAPI webhook server (for Render deployment)
"""

import asyncio
import logging
import os
import sys
import traceback

from dotenv import load_dotenv

# Load .env file (only used locally — Render uses its own env var dashboard)
load_dotenv()

# ─── Mode Detection ──────────────────────────────────────────────────────────

LOCAL_MODE = os.environ.get("LOCAL_MODE", "false").strip().lower() in ("true", "1", "yes")

# ─── Validate Environment Variables ───────────────────────────────────────────

# In local (polling) mode, WEBHOOK_SECRET and APP_URL are not needed
if LOCAL_MODE:
    REQUIRED_VARS = ["TELEGRAM_BOT_TOKEN", "ANTHROPIC_API_KEY"]
else:
    REQUIRED_VARS = ["TELEGRAM_BOT_TOKEN", "ANTHROPIC_API_KEY", "WEBHOOK_SECRET", "APP_URL"]

missing = [v for v in REQUIRED_VARS if not os.environ.get(v)]
if missing:
    print(f"❌ FATAL: Missing required environment variables: {', '.join(missing)}", file=sys.stderr)
    print("   Set them in .env (local) or in the Render dashboard (production).", file=sys.stderr)
    sys.exit(1)

TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "")
APP_URL = os.environ.get("APP_URL", "").rstrip("/")

# ─── Logging ──────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
#  POLLING MODE  (LOCAL_MODE=true)
# ═══════════════════════════════════════════════════════════════════════════════

async def run_polling():
    """Delete any existing webhook, then poll getUpdates in a loop."""
    import httpx
    from bot.handlers import handle_update, start_cleanup_task

    api = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"

    # Create temp directories
    os.makedirs("/tmp/uploads", exist_ok=True)
    os.makedirs("/tmp/outputs", exist_ok=True)

    # Step 1: Delete any existing webhook
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(f"{api}/deleteWebhook", json={"drop_pending_updates": False})
        result = resp.json()
        if result.get("ok"):
            logger.info("✅ Webhook deleted — polling mode active")
        else:
            logger.error("❌ Failed to delete webhook: %s", result)

    # Start periodic cleanup
    start_cleanup_task()

    # Step 2: Poll loop
    offset = 0
    logger.info("🔄 Polling for updates... (Ctrl+C to stop)")

    async with httpx.AsyncClient(timeout=60) as client:
        while True:
            try:
                resp = await client.get(
                    f"{api}/getUpdates",
                    params={"offset": offset, "timeout": 30},
                )
                data = resp.json()

                if not data.get("ok"):
                    logger.error("getUpdates error: %s", data)
                    await asyncio.sleep(5)
                    continue

                for update in data.get("result", []):
                    offset = update["update_id"] + 1
                    try:
                        await handle_update(update)
                    except Exception:
                        logger.error("Error handling update:\n%s", traceback.format_exc())

            except httpx.ReadTimeout:
                # Normal — long-poll timeout with no updates
                continue
            except asyncio.CancelledError:
                logger.info("Polling stopped.")
                return
            except Exception:
                logger.error("Polling error:\n%s", traceback.format_exc())
                await asyncio.sleep(5)


# ═══════════════════════════════════════════════════════════════════════════════
#  WEBHOOK MODE  (default / Render)
# ═══════════════════════════════════════════════════════════════════════════════

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

app = FastAPI(title="Office Bot", version="1.0.0")


@app.on_event("startup")
async def on_startup():
    """Set Telegram webhook, create temp dirs, start cleanup task."""
    import httpx
    from bot.handlers import start_cleanup_task

    # Create temp directories
    os.makedirs("/tmp/uploads", exist_ok=True)
    os.makedirs("/tmp/outputs", exist_ok=True)
    logger.info("Created /tmp/uploads and /tmp/outputs directories")

    # Register Telegram webhook
    webhook_url = f"{APP_URL}/webhook/{WEBHOOK_SECRET}"
    api_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/setWebhook"

    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.post(api_url, json={"url": webhook_url})
            result = resp.json()
            if result.get("ok"):
                logger.info("✅ Telegram webhook set to: %s", webhook_url)
            else:
                logger.error("❌ Failed to set webhook: %s", result)
    except Exception:
        logger.error("❌ Webhook registration failed:\n%s", traceback.format_exc())

    # Start periodic cleanup
    start_cleanup_task()
    logger.info("✅ Periodic cleanup task started (every 30 min)")


@app.get("/health")
async def health_check():
    """Health check endpoint for Render."""
    return {"status": "ok"}


@app.post("/webhook/{secret}")
async def webhook(secret: str, request: Request):
    """Telegram webhook endpoint."""
    if secret != WEBHOOK_SECRET:
        logger.warning("Webhook called with invalid secret")
        return JSONResponse(status_code=403, content={"error": "forbidden"})

    # Always return 200 to Telegram — even on errors
    try:
        update = await request.json()
        chat_id = update.get("message", {}).get("chat", {}).get("id", "unknown")
        msg_type = "document" if update.get("message", {}).get("document") else "text"
        logger.info("Incoming update: chat_id=%s type=%s", chat_id, msg_type)

        from bot.handlers import handle_update
        await handle_update(update)

    except Exception:
        logger.error("Webhook processing error:\n%s", traceback.format_exc())

    return JSONResponse(status_code=200, content={})


# ═══════════════════════════════════════════════════════════════════════════════
#  Entry point — branch on LOCAL_MODE
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if LOCAL_MODE:
        logger.info("🏠 Starting in LOCAL polling mode")
        try:
            asyncio.run(run_polling())
        except KeyboardInterrupt:
            logger.info("👋 Bot stopped by user")
    else:
        import uvicorn
        logger.info("🌐 Starting in WEBHOOK mode (FastAPI)")
        uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
