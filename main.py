"""
ESL Safety Observation Bot — Firebase Edition
==============================================
Fixes:
  ✅ Flask keep-alive (fixes Render port timeout)
  ✅ delete_webhook on startup (fixes Conflict error)
  ✅ Crash-loop fixed — let Render auto-restart container, not Python while-loop
  ✅ Event loop closure issue fixed (close_loop=False)

New Features:
  ✅ /close OBS-26-0014 [photo] — close with proof
  ✅ /stats — department-wise breakdown
  ✅ Auto area classification from observation text
  ✅ /mystats — per-officer stats

Photo Storage:
  ✅ Photos compressed with Pillow (max 800px, JPEG q=60)
  ✅ Stored as base64 data URL in Firestore field `image_b64`
  ✅ No Firebase Storage / no external URL needed
"""

import os
import io
import json
import re
import base64
import logging
import threading
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import firebase_admin
from firebase_admin import credentials, firestore

from PIL import Image

from flask import Flask

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
    level=logging.INFO,
    force=True,
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

def run_flask():
    port = int(os.environ.get("PORT", 8080))
    flask_app.run(host="0.0.0.0", port=port)

def start_flask_keepalive():
    thread = threading.Thread(target=run_flask, daemon=True, name="flask-keepalive")
    thread.start()
    log.info("Flask keep-alive server started.")

# ─── FIREBASE CONFIG ─────────────────────────────────────────────────────────
db = None
BOT_TOKEN = None

try:
    LOCAL_TZ = ZoneInfo(os.environ.get("LOCAL_TZ", "Asia/Kolkata"))
except Exception:
    LOCAL_TZ = timezone(timedelta(hours=5, minutes=30))


def require_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required Render environment variable: {name}")
    return value


def init_runtime_config():
    global db, BOT_TOKEN

    if db is None:
        raw_creds = require_env("FIREBASE_CREDS_JSON")
        firebase_creds = json.loads(raw_creds)

        if not firebase_admin._apps:
            cred = credentials.Certificate(firebase_creds)
            firebase_admin.initialize_app(cred)          # no storageBucket needed
        db = firestore.client()
        log.info("Firebase initialized.")

    if BOT_TOKEN is None:
        BOT_TOKEN = require_env("BOT_TOKEN")
        log.info("Telegram bot token loaded.")


def local_now() -> datetime:
    return datetime.now(LOCAL_TZ)


def to_local_dt(value: datetime | None = None) -> datetime:
    if value is None:
        return local_now()
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(LOCAL_TZ)


def telegram_message_time(message) -> tuple[datetime, str]:
    origin = getattr(message, "forward_origin", None)
    origin_date = getattr(origin, "date", None)
    if origin_date:
        return to_local_dt(origin_date), "telegram_forward_origin"
    return to_local_dt(getattr(message, "date", None)), "telegram_message"


def timestamp_fields(value: datetime | None, source: str) -> dict:
    local_dt = to_local_dt(value)
    human = local_dt.strftime("%d/%m/%Y %H:%M")
    iso = local_dt.isoformat(timespec="seconds")
    return {
        "logged_at": human,
        "log_date": local_dt.strftime("%d/%m/%Y"),
        "log_ts": iso,
        "telegram_at": human,
        "telegram_ts": iso,
        "date_source": source,
        "timezone": str(LOCAL_TZ),
    }

# ─── IMAGE COMPRESSION → BASE64 ──────────────────────────────────────────────
MAX_PX   = 800    # longest side in pixels
QUALITY  = 60     # JPEG quality (0-100); 60 ≈ 80-150 KB for a typical site photo

def compress_to_b64(image_bytes: bytes) -> str:
    """
    Resize image to MAX_PX on longest side, re-encode as JPEG at QUALITY,
    and return a base64 data-URL string suitable for storing in Firestore.

    Returns "" on any failure so callers can degrade gracefully.
    Typical output: 80–200 KB base64 string (~107–267 KB stored).
    Firestore 1 MB document limit is safe for all normal site photos.
    """
    try:
        img = Image.open(io.BytesIO(image_bytes))
        img = img.convert("RGB")                       # drop alpha / palette
        img.thumbnail((MAX_PX, MAX_PX), Image.LANCZOS) # resize in place, keeps ratio
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=QUALITY, optimize=True)
        encoded = base64.b64encode(buf.getvalue()).decode("utf-8")
        data_url = f"data:image/jpeg;base64,{encoded}"
        kb = len(buf.getvalue()) / 1024
        log.info(f"[compress] {kb:.0f} KB JPEG → base64 ({len(data_url)//1024} KB string)")
        return data_url
    except Exception as e:
        log.warning(f"[compress] failed: {e}")
        return ""

# ─── DEPARTMENTS ─────────────────────────────────────────────────────────────
AREAS = [
    "RMHS", "SINTER", "COKE OVEN", "POWERPLANT",
    "BLAST FURNACE 2", "BLAST FURNACE 3", "PCI", "PCM", "SECONDARY OPERATIONS"
]

# ─── AUTO AREA CLASSIFIER ────────────────────────────────────────────────────
AREA_KEYWORDS = {
    "RMHS":                 ["rmhs", "raw material", "stockyard", "stacker", "reclaimer", "wagon", "tippler"],
    "SINTER":               ["sinter", "sintering", "sinter plant", "sinter machine"],
    "COKE OVEN":            ["coke oven", "coke", "oven", "battery", "coke battery", "vco", "cpcs", "pusher", "charging ram", "crusher"],
    "POWERPLANT":           ["power", "turbine", "boiler", "generator", "power plant", "powerplant", "dg"],
    "BLAST FURNACE 2":      ["bf#2", "bf2", "blast furnace 2", "furnace 2", "bf-2", "bf 2"],
    "BLAST FURNACE 3":      ["bf#3", "bf3", "blast furnace 3", "furnace 3", "bf-3", "bf 3"],
    "PCI":                  ["pci", "pulverized coal", "pulverised coal", "coal injection"],
    "PCM":                  ["pcm", "process control"],
    "SECONDARY OPERATIONS": ["secondary operations", "secondary"],
}

def classify_area(text: str) -> str | None:
    """Auto-detect area from observation text. Returns area name or None."""
    text_lower = text.lower()
    best_area = None
    best_score = 0
    for area, keywords in AREA_KEYWORDS.items():
        score = 0
        for kw in keywords:
            kw_l = kw.lower()
            if kw_l in text_lower:
                score += 2 + len(kw_l.split())
        if score > best_score:
            best_area = area
            best_score = score
    return best_area


def clean_forwarded_message(text: str) -> str:
    """Remove WhatsApp/Telegram forwarding noise while preserving the observation."""
    lines = []
    for raw in text.splitlines():
        line = raw.strip().strip("*")
        if not line:
            continue
        low = line.lower()
        if low in {"forwarded", "this message was deleted"}:
            continue
        if "(file attached)" in low:
            continue
        lines.append(line)
    return "\n".join(lines).strip()


def split_location_and_observation(text: str, detected_area: str | None) -> tuple[str, str]:
    """
    Supports WhatsApp captions like:
      BF#2, HBS
      Pencil grinder and grinding machine were found...
    """
    cleaned = clean_forwarded_message(text)
    lines = [l.strip().strip("*") for l in cleaned.splitlines() if l.strip()]
    if not lines:
        return "", ""

    first = lines[0]
    rest = "\n".join(lines[1:]).strip()

    first_has_area = bool(classify_area(first))
    short_location = len(first) <= 60 and (
        "," in first or "#" in first or bool(re.search(r"\b(bf|hbs|vco|cpcs|rmhs|sinter|pci|pcm)\b", first, re.I))
    )

    if rest and short_location:
        return first, rest

    return "", cleaned

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

    detected = classify_area(text)
    location_line = ""

    if "observation" not in obs:
        location_line, flexible_observation = split_location_and_observation(text, detected)
        if len(flexible_observation) < 10:
            return None
        obs["observation"] = flexible_observation
        obs["source_format"] = "whatsapp_forward"
    else:
        obs["source_format"] = "field_format"

    if "area" not in obs or not obs["area"].strip():
        if detected:
            obs["area"]      = detected
            obs["area_auto"] = True
        else:
            obs["area"]      = "UNKNOWN"
            obs["area_auto"] = False
    else:
        obs["area_auto"] = False

    if location_line:
        obs["location"] = location_line
    obs.setdefault("responsible", "-")
    obs.setdefault("target_date", "-")
    obs.setdefault("status", "Open")

    return obs

# ─── FIREBASE HELPERS ─────────────────────────────────────────────────────────
def _next_ref() -> str:
    counter_ref = db.collection("meta").document("counter")
    counter_doc = counter_ref.get()
    n = (counter_doc.to_dict().get("count", 0) + 1) if counter_doc.exists else 1
    counter_ref.set({"count": n})
    year = local_now().strftime("%y")
    return f"OBS-{year}-{n:04d}"


def save_observation(obs: dict, shared_at: datetime | None = None, date_source: str = "server_time") -> str:
    ref_no = _next_ref()
    obs["ref_no"] = ref_no
    obs.update(timestamp_fields(shared_at, date_source))
    if obs.get("datetime"):
        obs["message_datetime"] = obs["datetime"]
    obs["datetime"] = obs["telegram_at"]
    obs.setdefault("status",        "Open")
    obs.setdefault("image_b64",     "")     # base64 data-URL or ""
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
    docs = db.collection("observations").stream()
    return sorted(
        [d.to_dict() for d in docs],
        key=lambda o: o.get("telegram_ts") or o.get("log_ts") or o.get("logged_at") or "",
    )


def load_today() -> list[dict]:
    today = local_now().strftime("%d/%m/%Y")
    return [o for o in load_all() if o.get("log_date") == today]


def load_shift(hours: int = 8) -> list[dict]:
    cutoff = (local_now() - timedelta(hours=hours)).isoformat(timespec="seconds")
    return [
        o for o in load_all()
        if (o.get("telegram_ts") or o.get("log_ts") or "") >= cutoff
    ]

# ─── HELP TEXT ───────────────────────────────────────────────────────────────
HELP = (
    "👷 *ESL Safety Observation Bot*\n"
    "━━━━━━━━━━━━━━━━━━━━━\n\n"
    "*Submit an observation in this format:*\n\n"
    "```\n"
    "Observation: Loose scaffold near BF3\n"
    "Area: SINTER\n"
    "Responsible: Suresh Singh\n"
    "Status: Open\n"
    "Target Date: 15/05/2026\n"
    "```\n"
    "Date/time is captured automatically from Telegram.\n"
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

        open_c   = sum(1 for o in all_obs if "open"     in o.get("status","").lower() and "progress" not in o.get("status","").lower())
        prog_c   = sum(1 for o in all_obs if "progress" in o.get("status","").lower())
        closed_c = sum(1 for o in all_obs if "closed"   in o.get("status","").lower())
        rate     = f"{closed_c / total * 100:.1f}%" if total else "N/A"

        dept_lines = []
        for area in AREAS:
            area_obs = [o for o in all_obs if area.lower() in o.get("area", "").lower()]
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
            f"🕐 Updated: {local_now().strftime('%d/%m/%Y %H:%M')}"
        )
        await update.message.reply_text(msg, parse_mode="Markdown")

    except Exception as e:
        log.error(e, exc_info=True)
        await update.message.reply_text("❌ Could not fetch stats.")


async def cmd_mystats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Show personal submission stats for the requesting user."""
    try:
        user    = update.effective_user
        name    = user.full_name or user.username or str(user.id)
        all_obs = load_all()
        my_obs  = [o for o in all_obs
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

    # Compress closure photo → base64
    closure_b64 = ""
    if message.photo:
        try:
            photo     = message.photo[-1]
            tg_file   = await ctx.bot.get_file(photo.file_id)
            img_bytes = await tg_file.download_as_bytearray()
            closure_b64 = compress_to_b64(bytes(img_bytes))
            log.info(f"Closure photo compressed for {ref_no}")
        except Exception as e:
            log.warning(f"Closure photo compression failed: {e}")

    closed_by = user.full_name or user.username or str(user.id)
    closed_at = local_now().strftime("%d/%m/%Y %H:%M")

    update_observation(ref_no, {
        "status":           "Closed",
        "closed_by":        closed_by,
        "closed_at":        closed_at,
        "closure_b64":      closure_b64,   # base64 closure photo (or "")
    })

    photo_note = "📷 Closure photo saved ✅" if closure_b64 else "📷 No closure photo"

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
        date_str = local_now().strftime("%d%m%Y")
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
                        date_label=f"Full Report — {local_now().strftime('%d/%m/%Y')}")


async def cmd_today(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    today = local_now().strftime("%d/%m/%Y")
    await _send_reports(update, load_today(), "Today",
                        date_label=f"Daily Report — {today}")


async def cmd_shift(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    obs   = load_shift(hours=8)
    label = f"Shift_{local_now().strftime('%d%m%Y_%H%M')}"
    await _send_reports(update, obs, label,
                        date_label=f"Shift Report — Last 8 hrs — {local_now().strftime('%d/%m/%Y %H:%M')}")


async def cmd_areas(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = await update.message.reply_text("⏳ Generating per-area reports…")
    try:
        all_obs = load_all()
        if not all_obs:
            await msg.edit_text("⚠️ No observations found.")
            return
        today        = local_now().strftime("%d/%m/%Y")
        area_reports = generate_pptx_per_area(all_obs, date_label=f"Area Report — {today}")
        if not area_reports:
            await msg.edit_text("⚠️ No area-tagged observations found.")
            return
        await msg.delete()
        date_str = local_now().strftime("%d%m%Y")
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

    # Compress photo → base64, store inline in Firestore
    shared_at, date_source = telegram_message_time(message)

    obs["image_b64"] = ""
    if message.photo:
        try:
            photo     = message.photo[-1]          # highest resolution
            tg_file   = await ctx.bot.get_file(photo.file_id)
            img_bytes = await tg_file.download_as_bytearray()
            obs["image_b64"] = compress_to_b64(bytes(img_bytes))
        except Exception as e:
            log.warning(f"Image capture/compression failed: {e}")

    try:
        ref_no   = save_observation(obs, shared_at=shared_at, date_source=date_source)
        img_note = "📷 Photo saved ✅" if obs.get("image_b64") else "📷 No photo"

        area_note = ""
        if obs.get("area_auto"):
            area_note = f"\n🤖 Area auto-detected from text"

        await message.reply_text(
            f"✅ *Observation Saved!*\n"
            f"━━━━━━━━━━━━━━━━━━\n"
            f"Ref:        `{ref_no}`\n"
            f"Date/Time:  {obs.get('telegram_at', obs.get('logged_at', ''))}\n"
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
async def bot_post_init(app):
    await app.bot.delete_webhook(drop_pending_updates=True)
    log.info("Telegram webhook cleared. Polling can start.")


def build_bot_application():
    init_runtime_config()
    app = Application.builder().token(BOT_TOKEN).post_init(bot_post_init).build()

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
    return app


def main():
    """
    Run the bot ONCE per process.

    Why no while-True loop:
      • Application.run_polling() closes the asyncio event loop on exit.
      • Restarting it in the same process triggers:
            RuntimeError: Event loop is closed
      • The correct pattern is: exit the process, let Render restart the container.
      • Exit code != 0 tells Render to auto-restart with a fresh process.
    """
    start_flask_keepalive()

    try:
        app = build_bot_application()
        log.info("Bot started. Polling...")
        app.run_polling(
            drop_pending_updates=True,
            close_loop=False,        # keep event loop healthy during shutdown
            stop_signals=None,       # let Render handle SIGTERM cleanly
        )
        # If run_polling returns normally (rare, e.g. clean shutdown),
        # exit normally so Render doesn't keep restarting us.
        log.info("Polling stopped cleanly. Exiting.")

    except KeyboardInterrupt:
        log.info("KeyboardInterrupt — shutting down.")

    except Exception:
        log.exception("Bot crashed. Exiting so Render restarts the container.")
        # Non-zero exit → Render auto-restart kicks in with a fresh process.
        raise SystemExit(1)


if __name__ == "__main__":
    main()
