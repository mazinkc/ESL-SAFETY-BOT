"""
PPTX report generator — ESL Safety Observations
Layout: 2 observations per slide side by side with image + details
Brand: ESL Steel / Vedanta (dark theme: #0D1117, orange #F97316, blue #3B82F6)
"""

import io
import requests
import logging
from datetime import datetime
from itertools import zip_longest

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

log = logging.getLogger(__name__)

# ─── BRAND COLORS ────────────────────────────────────────────────────────────
DARK    = RGBColor(0x0D, 0x11, 0x17)
SURF    = RGBColor(0x11, 0x18, 0x27)
SURF2   = RGBColor(0x1F, 0x29, 0x37)
BORDER  = RGBColor(0x37, 0x41, 0x51)
ORANGE  = RGBColor(0xF9, 0x73, 0x16)
BLUE    = RGBColor(0x3B, 0x82, 0xF6)
GREEN   = RGBColor(0x22, 0xC5, 0x5E)
YELLOW  = RGBColor(0xEA, 0xB3, 0x08)
RED     = RGBColor(0xEF, 0x44, 0x44)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
MUTED   = RGBColor(0x9C, 0xA3, 0xAF)
TEXT    = RGBColor(0xE2, 0xE8, 0xF0)

STATUS_COLOR = {
    "open":             RED,
    "partially closed": YELLOW,
    "in progress":      YELLOW,
    "closed":           GREEN,
}

# ─── SLIDE DIMENSIONS (widescreen 16:9) ──────────────────────────────────────
W = Inches(13.33)
H = Inches(7.5)

# Column layout: left panel x=0.15, right panel x=6.74
PANEL_W  = Inches(6.44)
PANEL_H  = Inches(5.85)
PANEL_Y  = Inches(1.45)
LEFT_X   = Inches(0.15)
RIGHT_X  = Inches(6.74)
GAP      = Inches(0.15)


# ─── HELPERS ─────────────────────────────────────────────────────────────────
def rect(slide, x, y, w, h, fill: RGBColor, line: RGBColor = None, line_w=0.75):
    s = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        s.line.width     = Pt(line_w)
    else:
        s.line.fill.background()
    return s


def txbox(slide, text, x, y, w, h,
          size=11, bold=False, color=None, align=PP_ALIGN.LEFT, wrap=True):
    color = color or TEXT
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text           = str(text)
    run.font.size      = Pt(size)
    run.font.bold      = bold
    run.font.color.rgb = color
    run.font.name      = "Segoe UI"
    return tb


def fetch_image(url: str) -> bytes | None:
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return r.content
    except Exception as e:
        log.warning(f"Image fetch failed: {e}")
        return None


# ─── SLIDE BUILDERS ──────────────────────────────────────────────────────────
def _slide_title(prs, total, open_c, closed_c):
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # Full dark background
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = DARK

    # Orange hazard stripe top
    rect(sl, 0, 0, W, Inches(0.18), ORANGE)

    # Orange hazard stripe bottom
    rect(sl, 0, H - Inches(0.18), W, Inches(0.18), ORANGE)

    # Center glow panel
    rect(sl, Inches(1.2), Inches(1.6), Inches(10.93), Inches(4.6), SURF)

    # Orange left accent bar
    rect(sl, Inches(1.2), Inches(1.6), Inches(0.12), Inches(4.6), ORANGE)

    # ESL label
    txbox(sl, "VEDANTA GROUP  |  ESL STEEL LIMITED  |  BOKARO",
          Inches(1.5), Inches(1.9), Inches(10.4), Inches(0.6),
          size=12, color=MUTED, align=PP_ALIGN.CENTER)

    # Main title
    txbox(sl, "Safety Observations Report",
          Inches(1.5), Inches(2.55), Inches(10.4), Inches(1.4),
          size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Divider
    rect(sl, Inches(3.5), Inches(3.95), Inches(6.33), Inches(0.04), ORANGE)

    # Stats row
    stats = [
        (str(total),    "TOTAL",   BLUE),
        (str(open_c),   "OPEN",    RED),
        (str(closed_c), "CLOSED",  GREEN),
    ]
    for i, (val, label, clr) in enumerate(stats):
        cx = Inches(3.7 + i * 2.1)
        txbox(sl, val,   cx, Inches(4.1), Inches(2.0), Inches(0.75),
              size=28, bold=True, color=clr, align=PP_ALIGN.CENTER)
        txbox(sl, label, cx, Inches(4.85), Inches(2.0), Inches(0.45),
              size=10, color=MUTED, align=PP_ALIGN.CENTER)

    # Date
    txbox(sl, datetime.now().strftime("%d %B %Y"),
          Inches(1.5), Inches(5.6), Inches(10.4), Inches(0.5),
          size=12, color=MUTED, align=PP_ALIGN.CENTER)


def _slide_summary(prs, total, open_c, partial_c, closed_c):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = DARK

    # Header
    rect(sl, 0, 0, W, Inches(1.2), SURF)
    rect(sl, 0, 0, Inches(0.12), Inches(1.2), ORANGE)
    txbox(sl, "Observations Summary",
          Inches(0.35), Inches(0.2), Inches(9), Inches(0.8),
          size=26, bold=True, color=WHITE)
    txbox(sl, datetime.now().strftime("%d/%m/%Y"),
          Inches(10), Inches(0.35), Inches(3.0), Inches(0.5),
          size=12, color=MUTED, align=PP_ALIGN.RIGHT)

    rate  = f"{closed_c / total * 100:.1f}%" if total else "N/A"
    cards = [
        ("Total",            str(total),     BLUE,   SURF2),
        ("Open",             str(open_c),    RED,    SURF2),
        ("Partially Closed", str(partial_c), YELLOW, SURF2),
        ("Closed",           str(closed_c),  GREEN,  SURF2),
    ]

    card_w = Inches(2.9)
    gap    = Inches(0.28)
    sx     = (W - (card_w * 4 + gap * 3)) / 2

    for i, (label, val, clr, bg) in enumerate(cards):
        cx = sx + i * (card_w + gap)
        cy = Inches(1.45)

        # Card
        rect(sl, cx, cy, card_w, Inches(3.5), SURF2)
        # Top accent line
        rect(sl, cx, cy, card_w, Inches(0.1), clr)

        # Big number
        txbox(sl, val,
              cx, cy + Inches(0.4), card_w, Inches(1.6),
              size=52, bold=True, color=clr, align=PP_ALIGN.CENTER)

        # Label
        txbox(sl, label,
              cx, cy + Inches(2.1), card_w, Inches(0.6),
              size=12, color=MUTED, align=PP_ALIGN.CENTER)

    # Closure rate bar
    rect(sl, Inches(3.2), Inches(5.4), Inches(6.93), Inches(0.95), SURF2)
    rect(sl, Inches(3.2), Inches(5.4), Inches(0.12), Inches(0.95), ORANGE)
    txbox(sl, f"Closure Rate:  {rate}",
          Inches(3.5), Inches(5.5), Inches(6.63), Inches(0.75),
          size=22, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)


def _draw_obs_panel(sl, obs, panel_x):
    """Draw one observation panel (left or right side of slide)."""
    status     = obs.get("status", "Open").strip()
    status_key = status.lower()
    s_color    = STATUS_COLOR.get(status_key, MUTED)
    image_url  = obs.get("image_url", "").strip()

    # ── Panel background ──────────────────────────────────────────────────────
    rect(sl, panel_x, PANEL_Y, PANEL_W, PANEL_H, SURF2)

    # ── Top accent bar (status color) ─────────────────────────────────────────
    rect(sl, panel_x, PANEL_Y, PANEL_W, Inches(0.1), s_color)

    # ── Ref No + Status badge row ─────────────────────────────────────────────
    txbox(sl, obs.get("ref_no", ""),
          panel_x + Inches(0.15), PANEL_Y + Inches(0.15),
          Inches(3.5), Inches(0.45),
          size=13, bold=True, color=WHITE)

    # Status badge
    badge_w = Inches(1.8)
    badge_x = panel_x + PANEL_W - badge_w - Inches(0.12)
    rect(sl, badge_x, PANEL_Y + Inches(0.14), badge_w, Inches(0.42), s_color)
    txbox(sl, status.upper(),
          badge_x, PANEL_Y + Inches(0.14), badge_w, Inches(0.42),
          size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # ── Image section (top half of panel) ─────────────────────────────────────
    img_y = PANEL_Y + Inches(0.65)
    img_h = Inches(2.5)
    img_w = PANEL_W - Inches(0.3)
    img_x = panel_x + Inches(0.15)

    placed = False
    if image_url:
        img_bytes = fetch_image(image_url)
        if img_bytes:
            try:
                sl.shapes.add_picture(io.BytesIO(img_bytes), img_x, img_y, img_w, img_h)
                placed = True
            except Exception as e:
                log.warning(f"Picture insert failed: {e}")

    if not placed:
        rect(sl, img_x, img_y, img_w, img_h, SURF)
        txbox(sl, "📷  No Image",
              img_x, img_y + Inches(1.0), img_w, Inches(0.6),
              size=13, color=BORDER, align=PP_ALIGN.CENTER)

    # ── Details section (bottom half) ─────────────────────────────────────────
    details_y  = img_y + img_h + Inches(0.12)
    detail_h   = Inches(0.42)
    detail_gap = Inches(0.44)

    fields = [
        ("👤", obs.get("observer_name", "—")),
        ("📅", obs.get("datetime",      "—")),
        ("📍", obs.get("area",          "—")),
        ("👷", obs.get("responsible",   "—")),
        ("🎯", obs.get("target_date",   "—")),
    ]

    for i, (icon, val) in enumerate(fields):
        fy = details_y + i * detail_gap
        # Field label chip
        rect(sl, panel_x + Inches(0.15), fy,
             Inches(0.32), detail_h, SURF)
        txbox(sl, icon,
              panel_x + Inches(0.15), fy,
              Inches(0.32), detail_h,
              size=11, align=PP_ALIGN.CENTER)
        # Value
        txbox(sl, val,
              panel_x + Inches(0.52), fy,
              PANEL_W - Inches(0.65), detail_h,
              size=10, color=TEXT)

    # ── Observation text at bottom ────────────────────────────────────────────
    obs_text = obs.get("observation", "")
    obs_y    = PANEL_Y + PANEL_H - Inches(0.72)
    rect(sl, panel_x, obs_y, PANEL_W, Inches(0.72), SURF)
    rect(sl, panel_x, obs_y, Inches(0.08), Inches(0.72), ORANGE)
    txbox(sl, f"  {obs_text}",
          panel_x + Inches(0.12), obs_y + Inches(0.06),
          PANEL_W - Inches(0.18), Inches(0.62),
          size=10, color=TEXT, wrap=True)


def _slide_pair(prs, obs1, obs2):
    """One slide with 2 observations side by side."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = DARK

    # Top header bar
    rect(sl, 0, 0, W, Inches(1.3), SURF)
    rect(sl, 0, 0, Inches(0.12), Inches(1.3), ORANGE)
    txbox(sl, "ESL STEEL  |  Safety Observations",
          Inches(0.3), Inches(0.22), Inches(9), Inches(0.55),
          size=18, bold=True, color=WHITE)
    txbox(sl, datetime.now().strftime("%d/%m/%Y"),
          Inches(10), Inches(0.3), Inches(3.1), Inches(0.5),
          size=12, color=MUTED, align=PP_ALIGN.RIGHT)

    # Divider between panels
    rect(sl, LEFT_X + PANEL_W + Inches(0.08), PANEL_Y,
         Inches(0.02), PANEL_H, BORDER)

    # Draw observation panels
    _draw_obs_panel(sl, obs1, LEFT_X)

    if obs2:
        _draw_obs_panel(sl, obs2, RIGHT_X)
    else:
        # Empty right panel placeholder
        rect(sl, RIGHT_X, PANEL_Y, PANEL_W, PANEL_H, SURF2)
        txbox(sl, "—",
              RIGHT_X, PANEL_Y + Inches(2.5), PANEL_W, Inches(0.7),
              size=20, color=BORDER, align=PP_ALIGN.CENTER)


# ─── MAIN ENTRY ──────────────────────────────────────────────────────────────
def generate_pptx(observations: list[dict]) -> bytes:
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    total     = len(observations)
    open_c    = sum(1 for o in observations if o.get("status","").lower() == "open")
    partial_c = sum(1 for o in observations if "partial" in o.get("status","").lower()
                                            or "progress" in o.get("status","").lower())
    closed_c  = sum(1 for o in observations if o.get("status","").lower() == "closed")

    # Title + Summary slides
    _slide_title(prs, total, open_c, closed_c)
    _slide_summary(prs, total, open_c, partial_c, closed_c)

    # Pair up observations — 2 per slide
    pairs = list(zip_longest(observations[::2], observations[1::2]))
    for obs1, obs2 in pairs:
        _slide_pair(prs, obs1, obs2)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()
