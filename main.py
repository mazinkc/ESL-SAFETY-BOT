"""
ESL Safety Observation Bot — Firebase Edition
==============================================
Fixes:
  ✅ Flask keep-alive (fixes Render port timeout)
  ✅ delete_webhook on startup (fixes Conflict error)

New Features:
  ✅ /close OBS-26-0014 [photo] — close with proof
  ✅ /stats — department-wise breakdown
  ✅ Auto area classification from observation text
  ✅ /mystats — per-officer stats
"""

import os
import io
import json
import re
import logging
import threading
import requests
from datetime import datetime, timedelta

import firebase_admin
from firebase_admin import credentials, firestore

from flask import Flask, request, Response

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

# ─── FLASK KEEP-ALIVE (fixes Render port timeout) ───────────────────────────
flask_app = Flask(__name__)

@flask_app.route("/")
def home():
    return "ESL Safety Bot is running! ✅", 200

@flask_app.route("/health")
def health():
    return {"status": "ok", "bot": "ESL-SAFETY-BOT"}, 200


@flask_app.route("/photo-proxy")
def photo_proxy():
    """
    Proxy Telegram photo URLs for the dashboard PPTX export.
    The browser can't fetch api.telegram.org directly (CORS + bot token embedded
    in the URL path), so the dashboard calls this endpoint instead.

    Usage:  GET /photo-proxy?url=https://api.telegram.org/file/bot.../photo.jpg
    Returns the image bytes with Access-Control-Allow-Origin: * so the browser
    can convert it to a base64 data URL and embed it in PptxGenJS.
    """
    url = request.args.get("url", "").strip()

    # Only proxy Telegram file URLs — reject anything else
    if not url or "api.telegram.org/file/" not in url:
        return "Only Telegram file URLs are accepted.", 400

    try:
        r = requests.get(url, timeout=10, stream=False)
        r.raise_for_status()
        content_type = r.headers.get("Content-Type", "image/jpeg")
        return Response(
            r.content,
            status=200,
            headers={
                "Content-Type":                content_type,
                "Access-Control-Allow-Origin": "*",
                "Cache-Control":               "public, max-age=3600",
            }
        )
    except Exception as e:
        log.warning(f"[photo-proxy] failed for {url!r}: {e}")
        return "Proxy fetch failed", 502

def run_flask():
    port = int(os.environ.get("PORT", 8080))
    flask_app.run(host="0.0.0.0", port=port)

threading.Thread(target=run_flask, daemon=True).start()
log.info("Flask keep-alive server started.")

# ─── FIREBASE CONFIG ─────────────────────────────────────────────────────────
_firebase_creds = json.loads(os.environ["FIREBASE_CREDS_JSON"])
_cred = credentials.Certificate(_firebase_creds)
firebase_admin.initialize_app(_cred)
db = firestore.client()

BOT_TOKEN = os.environ["BOT_TOKEN"]

# ─── DEPARTMENTS ─────────────────────────────────────────────────────────────
AREAS = [
    "RMHS", "SINTER", "COKE OVEN", "POWERPLANT",
    "BLAST FURNACE 2", "BLAST FURNACE 3", "PCI", "PCM", "SECONDARY OPERATIONS"
]

# ─── AUTO AREA CLASSIFIER ────────────────────────────────────────────────────
AREA_KEYWORDS = {
    "RMHS":                 ["rmhs", "raw material", "conveyor", "stockyard", "stacker", "reclaimer"],
    "SINTER":               ["sinter", "sintering", "sinter plant"],
    "COKE OVEN":            ["coke", "oven", "coke oven", "battery", "coke battery"],
    "POWERPLANT":           ["power", "turbine", "boiler", "generator", "power plant", "powerplant", "dg"],
    "BLAST FURNACE 2":      ["bf2", "blast furnace 2", "furnace 2", "bf-2", "bf 2"],
    "BLAST FURNACE 3":      ["bf3", "blast furnace 3", "furnace 3", "bf-3", "bf 3"],
    "PCI":                  ["pci", "pulverized coal", "pulverised coal", "coal injection"],
    "PCM":                  ["pcm", "process control", "casting", "metal"],
    "SECONDARY OPERATIONS": ["secondary", "rolling", "SMS", "steel melt", "ladle", "torpedo"],
}

def classify_area(text: str) -> str | None:
    """Auto-detect area from observation text. Returns area name or None."""
    text_lower = text.lower()
    for area, keywords in AREA_KEYWORDS.items():
        if any(kw.lower() in text_lower for kw in keywords):
            return area
    return None

# ─── FIELD PARSER ────────────────────────────────────────────────────────────
_PATTERNS = {
    "observer_name": r"(?:name|observer)[:\-\s]+(.+)",
    "datetime":      r"(?:date\s*[/\-]?\s*time|date)[:\-\s]+(.+)",
    "observation":   r"(?:observation|obs)[:\-\s]+(.+)",
    "area":          r"area[:\-\s]+(.+)",
    "responsible":   r"(?:responsible|responsibility)[:\-\s]+(.+)",
    "status":        r"status[:\-\s]+(.+)",
    "target_date":   r"(?:target\s*date|target)[:\-\s]+(.+)",
}

def parse_observation(text: str) -> dict | None:
    obs = {}
    for field, pattern in _PATTERNS.items():
        m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if m:
            obs[field] = m.group(1).strip()
    if "observation" not in obs:
        return None

    # Auto-classify area if not provided
    if "area" not in obs or not obs["area"].strip():
        detected = classify_area(obs.get("observation", "") + " " + text)
        if detected:
            obs["area"]          = detected
            obs["area_auto"]     = True   # flag so we can tell user it was auto-detected
        else:
            obs["area"]          = "UNKNOWN"
            obs["area_auto"]     = False
    else:
        obs["area_auto"] = False

    return obs

# ─── FIREBASE HELPERS ─────────────────────────────────────────────────────────
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


def get_observation(ref_no: str) -> dict | None:
    doc = db.collection("observations").document(ref_no.upper()).get()
    return doc.to_dict() if doc.exists else None


def update_observation(ref_no: str, updates: dict):
    db.collection("observations").document(ref_no.upper()).update(updates)
    log.info(f"Updated {ref_no}: {list(updates.keys())}")


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
    "*Submit an observation in this format:*\n\n"
    "```\n"
    "Date/Time: 09/05/2026 10:30\n"
    "Observation: Loose scaffold near BF3\n"
    "Area: SINTER\n"
    "Responsible: Suresh Singh\n"
    "Status: Open\n"
    "Target Date: 15/05/2026\n"
    "```\n"
    "📷 Attach a photo with the message\n"
    "📍 Area is *auto-detected* if not provided\n\n"
    "*Commands:*\n"
    "/report   — Full report (all observations)\n"
    "/today    — Today's report\n"
    "/shift    — Last 8 hours report\n"
    "/areas    — Separate PPTX per department\n"
    "/stats    — Live summary + dept breakdown\n"
    "/close    — Close an observation\n"
    "/mystats  — Your personal submission stats\n"
    "/start    — Show this help\n\n"
    "*Close an observation:*\n"
    "`/close OBS-26-0014` (attach closure photo)\n\n"
    "*Departments:*\n"
    "RMHS | SINTER | COKE OVEN | POWERPLANT\n"
    "BLAST FURNACE 2 | BLAST FURNACE 3\n"
    "PCI | PCM | SECONDARY OPERATIONS"
)

# ─── HANDLERS ────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP, parse_mode="Markdown")


async def cmd_stats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Live stats with department-wise breakdown."""
    try:
        all_obs   = load_all()
        today_obs = load_today()
        total     = len(all_obs)

        if total == 0:
            await update.message.reply_text("⚠️ No observations recorded yet.")
            return

        open_c    = sum(1 for o in all_obs if "open"     in o.get("status","").lower() and "progress" not in o.get("status","").lower())
        prog_c    = sum(1 for o in all_obs if "progress" in o.get("status","").lower())
        closed_c  = sum(1 for o in all_obs if "closed"   in o.get("status","").lower())
        rate      = f"{closed_c / total * 100:.1f}%" if total else "N/A"

        # Department-wise breakdown
        dept_lines = []
        for area in AREAS:
            area_obs    = [o for o in all_obs if area.lower() in o.get("area", "").lower()]
            if not area_obs:
                continue
            a_total  = len(area_obs)
            a_open   = sum(1 for o in area_obs if "open"     in o.get("status","").lower() and "progress" not in o.get("status","").lower())
            a_prog   = sum(1 for o in area_obs if "progress" in o.get("status","").lower())
            a_closed = sum(1 for o in area_obs if "closed"   in o.get("status","").lower())
            dept_lines.append(
                f"  *{area}*\n"
                f"    Total: {a_total}  🔴{a_open}  🟡{a_prog}  🟢{a_closed}"
            )

        dept_text = "\n".join(dept_lines) if dept_lines else "  No department data."

        msg = (
            f"📊 *ESL Safety — Live Stats*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"📅 Today:        *{len(today_obs)}*\n"
            f"📋 Total:        *{total}*\n"
            f"🔴 Open:         *{open_c}*\n"
            f"🟡 In Progress:  *{prog_c}*\n"
            f"🟢 Closed:       *{closed_c}*\n"
            f"✅ Closure Rate: *{rate}*\n\n"
            f"🏭 *Department Breakdown:*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"{dept_text}\n\n"
            f"🕐 Updated: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        )
        await update.message.reply_text(msg, parse_mode="Markdown")

    except Exception as e:
        log.error(e, exc_info=True)
        await update.message.reply_text("❌ Could not fetch stats.")


async def cmd_mystats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Show personal submission stats for the requesting user."""
    try:
        user     = update.effective_user
        name     = user.full_name or user.username or str(user.id)
        all_obs  = load_all()
        my_obs   = [o for o in all_obs
                    if name.lower() in o.get("observer_name", "").lower()
                    or (user.username or "").lower() in o.get("observer_name", "").lower()]

        if not my_obs:
            await update.message.reply_text(
                f"📋 No observations found for *{name}*.\n"
                f"Make sure your name appears in the `Name:` field.",
                parse_mode="Markdown"
            )
            return

        total    = len(my_obs)
        open_c   = sum(1 for o in my_obs if "open"     in o.get("status","").lower() and "progress" not in o.get("status","").lower())
        prog_c   = sum(1 for o in my_obs if "progress" in o.get("status","").lower())
        closed_c = sum(1 for o in my_obs if "closed"   in o.get("status","").lower())

        await update.message.reply_text(
            f"👤 *Your Stats — {name}*\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"📋 Submitted:    *{total}*\n"
            f"🔴 Open:         *{open_c}*\n"
            f"🟡 In Progress:  *{prog_c}*\n"
            f"🟢 Closed:       *{closed_c}*",
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error(e, exc_info=True)
        await update.message.reply_text("❌ Could not fetch your stats.")


async def cmd_close(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    Close an observation.
    Usage: /close OBS-26-0014
    Attach a photo for closure proof (optional but recommended).
    """
    message = update.message
    user    = update.effective_user

    # Parse ref number from command args
    args = ctx.args
    if not args:
        await message.reply_text(
            "❌ Usage: `/close OBS-26-0014`\n"
            "Optionally attach a closure photo.",
            parse_mode="Markdown"
        )
        return

    ref_no = args[0].strip().upper()
    if not re.match(r"OBS-\d{2}-\d{4}", ref_no):
        await message.reply_text(
            f"❌ Invalid ref number: `{ref_no}`\n"
            f"Format should be: `OBS-26-0014`",
            parse_mode="Markdown"
        )
        return

    obs = get_observation(ref_no)
    if not obs:
        await message.reply_text(
            f"❌ Observation `{ref_no}` not found.",
            parse_mode="Markdown"
        )
        return

    if "closed" in obs.get("status", "").lower():
        await message.reply_text(
            f"⚠️ `{ref_no}` is already *Closed*.\n"
            f"Closed on: {obs.get('closed_at', '—')}",
            parse_mode="Markdown"
        )
        return

    # Save closure photo if attached
    closure_image_url = ""
    if message.photo:
        try:
            photo             = message.photo[-1]
            tg_file           = await ctx.bot.get_file(photo.file_id)
            closure_image_url = tg_file.file_path
            log.info(f"Closure photo saved for {ref_no}")
        except Exception as e:
            log.warning(f"Closure photo capture failed: {e}")

    closed_by = user.full_name or user.username or str(user.id)
    closed_at = datetime.now().strftime("%d/%m/%Y %H:%M")

    update_observation(ref_no, {
        "status":              "Closed",
        "closed_by":           closed_by,
        "closed_at":           closed_at,
        "closure_image_url":   closure_image_url,
    })

    photo_note = "📷 Closure photo saved ✅" if closure_image_url else "📷 No closure photo"

    await message.reply_text(
        f"✅ *Observation Closed!*\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"Ref:         `{ref_no}`\n"
        f"Area:        {obs.get('area', '—')}\n"
        f"Closed by:   {closed_by}\n"
        f"Closed at:   {closed_at}\n"
        f"{photo_note}",
        parse_mode="Markdown"
    )


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
            area_obs  = [o for o in all_obs if area_name.lower() in o.get("area", "").lower()]
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
    """Handle incoming observation messages (text or photo with caption)."""
    message = update.message
    if not message:
        return

    # Ignore /close with photo (handled by cmd_close)
    text = (message.text or message.caption or "").strip()
    if text.lower().startswith("/close"):
        return
    if not text:
        return

    obs = parse_observation(text)
    if not obs:
        await message.reply_text(
            "⚠️ Could not parse. Message must include:\n"
            "`Observation:` field at minimum.",
            parse_mode="Markdown"
        )
        return

    # Capture photo
    obs["image_url"] = ""
    obs["file_id"]   = ""
    if message.photo:
        try:
            photo            = message.photo[-1]
            tg_file          = await ctx.bot.get_file(photo.file_id)
            obs["image_url"] = tg_file.file_path
            obs["file_id"]   = photo.file_id
            log.info("Image URL captured from Telegram")
        except Exception as e:
            log.warning(f"Image capture failed: {e}")

    try:
        ref_no   = save_observation(obs)
        img_note = "📷 Photo saved ✅" if obs.get("image_url") else "📷 No photo"

        # Note if area was auto-detected
        area_note = ""
        if obs.get("area_auto"):
            area_note = f"\n🤖 Area auto-detected from text"

        await message.reply_text(
            f"✅ *Observation Saved!*\n"
            f"━━━━━━━━━━━━━━━━━━\n"
            f"Ref:        `{ref_no}`\n"
            f"Observer:   {obs.get('observer_name', '—')}\n"
            f"Area:       {obs.get('area', '—')}{area_note}\n"
            f"Status:     {obs.get('status', '—')}\n"
            f"Target:     {obs.get('target_date', '—')}\n"
            f"{img_note}\n\n"
            f"To close: `/close {ref_no}`",
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error(e, exc_info=True)
        await message.reply_text("❌ Save failed. Try again.")


# ─── MAIN ────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start",    cmd_start))
    app.add_handler(CommandHandler("stats",    cmd_stats))
    app.add_handler(CommandHandler("mystats",  cmd_mystats))
    app.add_handler(CommandHandler("report",   cmd_report))
    app.add_handler(CommandHandler("today",    cmd_today))
    app.add_handler(CommandHandler("shift",    cmd_shift))
    app.add_handler(CommandHandler("areas",    cmd_areas))
    app.add_handler(CommandHandler("close",    cmd_close))
    app.add_handler(MessageHandler(
        filters.TEXT | filters.CAPTION | filters.PHOTO,
        handle_message
    ))

    log.info("Bot started. Polling...")
    app.run_polling(drop_pending_updates=True)   # drop_pending clears old conflict


if __name__ == "__main__":
    main()
