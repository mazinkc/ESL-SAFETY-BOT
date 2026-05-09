"""
ESL Safety Observation Bot — Firebase Edition
==============================================
Officers post in WhatsApp group → Mazin forwards to this bot
→ Auto-stored in Firestore + image URL from Telegram
→ /report generates PPTX + Excel instantly
"""

import os
import io
import json
import re
import logging
import requests
from datetime import datetime

import firebase_admin
from firebase_admin import credentials, firestore

from telegram import Update, InputFile
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes
)

from report_excel import generate_excel
from report_pptx  import generate_pptx

# ─── LOGGING ────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO
)
log = logging.getLogger(__name__)

# ─── FIREBASE CONFIG ────────────────────────────────────────────────────────
_firebase_creds = json.loads(os.environ["FIREBASE_CREDS_JSON"])
_cred = credentials.Certificate(_firebase_creds)
firebase_admin.initialize_app(_cred)

db = firestore.client()

BOT_TOKEN = os.environ["BOT_TOKEN"]

# ─── FIELD PARSER ───────────────────────────────────────────────────────────
_PATTERNS = {
    "observer_name":  r"(?:name|observer)[:\-\s]+(.+)",
    "datetime":       r"(?:date\s*[/\-]?\s*time|date)[:\-\s]+(.+)",
    "observation":    r"(?:observation|obs)[:\-\s]+(.+)",
    "area":           r"area[:\-\s]+(.+)",
    "responsible":    r"(?:responsible|responsibility)[:\-\s]+(.+)",
    "status":         r"status[:\-\s]+(.+)",
    "target_date":    r"(?:target\s*date|target)[:\-\s]+(.+)",
}

def parse_observation(text: str) -> dict | None:
    obs = {}
    for field, pattern in _PATTERNS.items():
        m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if m:
            obs[field] = m.group(1).strip()
    if "observer_name" in obs and "observation" in obs:
        return obs
    return None

# ─── FIREBASE HELPERS ────────────────────────────────────────────────────────
def _next_ref() -> str:
    counter_ref = db.collection("meta").document("counter")
    counter_doc = counter_ref.get()
    n = (counter_doc.to_dict().get("count", 0) + 1) if counter_doc.exists else 1
    counter_ref.set({"count": n})
    year = datetime.now().strftime("%y")
    return f"OBS-{year}-{n:04d}"


def save_observation(obs: dict) -> str:
    ref_no = _next_ref()
    obs["ref_no"]    = ref_no
    obs["logged_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    obs.setdefault("status",    "Open")
    obs.setdefault("image_url", "")
    db.collection("observations").document(ref_no).set(obs)
    log.info(f"Saved {ref_no}")
    return ref_no


def load_all() -> list[dict]:
    docs = db.collection("observations").order_by("logged_at").stream()
    return [d.to_dict() for d in docs]

# ─── BOT HANDLERS ────────────────────────────────────────────────────────────
HELP = (
    "👷 *ESL Safety Observation Bot*\n"
    "━━━━━━━━━━━━━━━━━━━━━\n\n"
    "*Forward observations in this format:*\n\n"
    "```\n"
    "Name: Rahul Kumar\n"
    "Date/Time: 09/05/2026 10:30\n"
    "Observation: Loose scaffold near BF3\n"
    "Area: Sinter Plant - K2 Conveyor\n"
    "Responsible: Suresh Singh\n"
    "Status: Open\n"
    "Target Date: 15/05/2026\n"
    "```\n"
    "📷 Attach photo with the message\n\n"
    "*Commands:*\n"
    "/report — Generate PPTX + Excel\n"
    "/stats  — Live summary\n"
    "/start  — Show this help"
)


async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP, parse_mode="Markdown")


async def cmd_stats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        all_obs  = load_all()
        total    = len(all_obs)
        open_c   = sum(1 for o in all_obs if o.get("status","").lower() == "open")
        prog_c   = sum(1 for o in all_obs if "progress" in o.get("status","").lower())
        closed_c = sum(1 for o in all_obs if o.get("status","").lower() == "closed")
        rate     = f"{closed_c/total*100:.1f}%" if total else "N/A"

        await update.message.reply_text(
            f"📊 *Live Stats*\n"
            f"━━━━━━━━━━━━━━\n"
            f"Total:       *{total}*\n"
            f"🔴 Open:        *{open_c}*\n"
            f"🟡 In Progress: *{prog_c}*\n"
            f"🟢 Closed:      *{closed_c}*\n"
            f"✅ Closure Rate: *{rate}*",
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error(e)
        await update.message.reply_text("❌ Could not fetch stats.")


async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = await update.message.reply_text("⏳ Generating reports…")
    try:
        all_obs = load_all()
        if not all_obs:
            await msg.edit_text("⚠️ No observations found yet.")
            return

        date_str = datetime.now().strftime("%d%m%Y")

        xl = generate_excel(all_obs)
        pp = generate_pptx(all_obs)

        await msg.delete()

        await update.message.reply_document(
            InputFile(io.BytesIO(xl), filename=f"ESL_Observations_{date_str}.xlsx"),
            caption=f"📊 *Excel* — {len(all_obs)} observations",
            parse_mode="Markdown"
        )
        await update.message.reply_document(
            InputFile(io.BytesIO(pp), filename=f"ESL_Observations_{date_str}.pptx"),
            caption=f"📑 *PowerPoint* — {len(all_obs)} observations",
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error(e, exc_info=True)
        await msg.edit_text(f"❌ Failed: `{e}`", parse_mode="Markdown")


async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    message = update.message
    if not message:
        return

    text = (message.text or message.caption or "").strip()
    if not text:
        return

    obs = parse_observation(text)
    if not obs:
        await message.reply_text(
            "⚠️ Could not parse. Message must have at least:\n"
            "`Name:` and `Observation:` fields.",
            parse_mode="Markdown"
        )
        return

    # ── Get Telegram image URL (no Firebase Storage needed) ───────────────
    obs["image_url"] = ""
    obs["file_id"]   = ""

    if message.photo:
        try:
            photo   = message.photo[-1]          # largest size
            tg_file = await ctx.bot.get_file(photo.file_id)
            # file_path is the full HTTPS download URL from Telegram servers
            obs["image_url"] = tg_file.file_path
            obs["file_id"]   = photo.file_id
            log.info(f"Image URL captured from Telegram")
        except Exception as e:
            log.warning(f"Image capture failed: {e}")

    # ── Save to Firestore ─────────────────────────────────────────────────
    try:
        ref_no = save_observation(obs)

        img_note = "📷 Photo saved ✅" if obs.get("image_url") else "📷 No photo"

        await message.reply_text(
            f"✅ *Observation Saved!*\n"
            f"━━━━━━━━━━━━━━━━━━\n"
            f"Ref:         `{ref_no}`\n"
            f"Observer:  {obs.get('observer_name','—')}\n"
            f"Area:        {obs.get('area','—')}\n"
            f"Status:     {obs.get('status','—')}\n"
            f"Target:     {obs.get('target_date','—')}\n"
            f"{img_note}",
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error(e, exc_info=True)
        await message.reply_text("❌ Save failed. Try again.")


# ─── MAIN ────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start",  cmd_start))
    app.add_handler(CommandHandler("stats",  cmd_stats))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(MessageHandler(
        filters.TEXT | filters.CAPTION | filters.PHOTO,
        handle_message
    ))
    log.info("Bot started.")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
