"""
ESL Safety Observation Bot — Firebase Edition
==============================================
Officers forward observations → Bot parses + stores in Firestore
→ /report  /today  /areas  /shift  /stats
"""

import os
import io
import json
import re
import logging
import requests
from datetime import datetime, timedelta

import firebase_admin
from firebase_admin import credentials, firestore

from telegram import Update, InputFile
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes
)

from report_excel import generate_excel
from report_pptx  import generate_pptx, generate_pptx_per_area

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

# ─── DEPARTMENTS ────────────────────────────────────────────────────────────
AREAS = [
    "RMHS", "SINTER", "COKE OVEN", "POWERPLANT",
    "BLAST FURNACE 2", "BLAST FURNACE 3", "PCI", "PCM", "SECONDARY OPERATIONS"
]

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
    # Only Observation field is mandatory (Name is optional)
    if "observation" in obs:
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
    obs["log_date"]  = datetime.now().strftime("%d/%m/%Y")
    obs["log_ts"]    = datetime.now().isoformat()
    obs.setdefault("status",        "Open")
    obs.setdefault("image_url",     "")
    obs.setdefault("observer_name", "—")
    db.collection("observations").document(ref_no).set(obs)
    log.info(f"Saved {ref_no}")
    return ref_no


def load_all() -> list[dict]:
    docs = db.collection("observations").order_by("logged_at").stream()
    return [d.to_dict() for d in docs]


def load_today() -> list[dict]:
    today = datetime.now().strftime("%d/%m/%Y")
    return [o for o in load_all() if o.get("log_date") == today]


def load_shift(hours: int = 8) -> list[dict]:
    cutoff = (datetime.now() - timedelta(hours=hours)).isoformat()
    return [o for o in load_all() if o.get("log_ts", "") >= cutoff]

# ─── HELP TEXT ───────────────────────────────────────────────────────────────
HELP = (
    "👷 *ESL Safety Observation Bot*\n"
    "━━━━━━━━━━━━━━━━━━━━━\n\n"
    "*Forward observations in this format:*\n\n"
    "```\n"
    "Date/Time: 09/05/2026 10:30\n"
    "Observation: Loose scaffold near BF3\n"
    "Area: SINTER\n"
    "Responsible: Suresh Singh\n"
    "Status: Open\n"
    "Target Date: 15/05/2026\n"
    "```\n"
    "📷 Attach photo with the message\n\n"
    "*Commands:*\n"
    "/report  — Full report (all observations)\n"
    "/today   — Today's report only\n"
    "/shift   — Last 8 hours report\n"
    "/areas   — Separate PPTX per department\n"
    "/stats   — Live summary\n"
    "/start   — Show this help\n\n"
    "*Departments:*\n"
    "RMHS | SINTER | COKE OVEN | POWERPLANT\n"
    "BLAST FURNACE 2 | BLAST FURNACE 3\n"
    "PCI | PCM | SECONDARY OPERATIONS"
)

# ─── BOT HANDLERS ────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP, parse_mode="Markdown")


async def cmd_stats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        all_obs   = load_all()
        today_obs = load_today()
        total     = len(all_obs)
        open_c    = sum(1 for o in all_obs if "open" in o.get("status","").lower() and "progress" not in o.get("status","").lower())
        prog_c    = sum(1 for o in all_obs if "progress" in o.get("status","").lower())
        closed_c  = sum(1 for o in all_obs if "closed" in o.get("status","").lower())
        rate      = f"{closed_c/total*100:.1f}%" if total else "N/A"

        await update.message.reply_text(
            f"📊 *Live Stats*\n"
            f"━━━━━━━━━━━━━━\n"
            f"📅 Today:        *{len(today_obs)}* observations\n\n"
            f"📋 Total:        *{total}*\n"
            f"🔴 Open:         *{open_c}*\n"
            f"🟡 In Progress:  *{prog_c}*\n"
            f"🟢 Closed:       *{closed_c}*\n"
            f"✅ Closure Rate: *{rate}*",
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error(e)
        await update.message.reply_text("❌ Could not fetch stats.")


async def _send_reports(update, obs_list, label, date_label=""):
    msg = await update.message.reply_text(f"⏳ Generating {label} report…")
    try:
        if not obs_list:
            await msg.edit_text(f"⚠️ No observations found for: {label}")
            return
        date_str = datetime.now().strftime("%d%m%Y")
        xl = generate_excel(obs_list)
        pp = generate_pptx(obs_list, date_label=date_label or label)
        await msg.delete()
        await update.message.reply_document(
            InputFile(io.BytesIO(xl), filename=f"ESL_{label}_{date_str}.xlsx"),
            caption=f"📊 *Excel* — {len(obs_list)} observations | {label}",
            parse_mode="Markdown"
        )
        await update.message.reply_document(
            InputFile(io.BytesIO(pp), filename=f"ESL_{label}_{date_str}.pptx"),
            caption=f"📑 *PowerPoint* — {len(obs_list)} observations | {label}",
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error(e, exc_info=True)
        await msg.edit_text(f"❌ Failed: `{e}`", parse_mode="Markdown")


async def cmd_report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await _send_reports(update, load_all(), "All_Observations",
                        date_label=f"Full Report — {datetime.now().strftime('%d/%m/%Y')}")


async def cmd_today(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    today = datetime.now().strftime("%d/%m/%Y")
    await _send_reports(update, load_today(), "Today",
                        date_label=f"Daily Report — {today}")


async def cmd_shift(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    obs   = load_shift(hours=8)
    label = f"Shift_{datetime.now().strftime('%d%m%Y_%H%M')}"
    await _send_reports(update, obs, label,
                        date_label=f"Shift Report — Last 8 hrs — {datetime.now().strftime('%d/%m/%Y %H:%M')}")


async def cmd_areas(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = await update.message.reply_text("⏳ Generating per-area reports…")
    try:
        all_obs = load_all()
        if not all_obs:
            await msg.edit_text("⚠️ No observations found.")
            return
        today        = datetime.now().strftime("%d/%m/%Y")
        area_reports = generate_pptx_per_area(all_obs, date_label=f"Area Report — {today}")
        if not area_reports:
            await msg.edit_text("⚠️ No area-tagged observations found.")
            return
        await msg.delete()
        date_str = datetime.now().strftime("%d%m%Y")
        for area_name, pptx_bytes in area_reports.items():
            safe_name = area_name.replace(" ", "_")
            area_obs  = [o for o in all_obs if area_name.lower() in o.get("area","").lower()]
            await update.message.reply_document(
                InputFile(io.BytesIO(pptx_bytes),
                          filename=f"ESL_{safe_name}_{date_str}.pptx"),
                caption=f"📑 *{area_name}* — {len(area_obs)} observations",
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
            "`Observation:` field.",
            parse_mode="Markdown"
        )
        return

    obs["image_url"] = ""
    obs["file_id"]   = ""
    if message.photo:
        try:
            photo         = message.photo[-1]
            tg_file       = await ctx.bot.get_file(photo.file_id)
            obs["image_url"] = tg_file.file_path
            obs["file_id"]   = photo.file_id
            log.info("Image URL captured from Telegram")
        except Exception as e:
            log.warning(f"Image capture failed: {e}")

    try:
        ref_no   = save_observation(obs)
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
    app.add_handler(CommandHandler("today",  cmd_today))
    app.add_handler(CommandHandler("shift",  cmd_shift))
    app.add_handler(CommandHandler("areas",  cmd_areas))
    app.add_handler(MessageHandler(
        filters.TEXT | filters.CAPTION | filters.PHOTO,
        handle_message
    ))
    log.info("Bot started.")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
