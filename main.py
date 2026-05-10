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
import time
import requests
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
        try:
            firebase_creds = json.loads(raw_creds)
        except json.JSONDecodeError as e:
