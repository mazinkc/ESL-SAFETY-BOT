"""
ESL Safety Observation — PPTX Report Generator
White background | Vedanta/ESL Logo | 2 observations per slide
"""

import io
import os
import base64
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ─── BRAND COLORS ─────────────────────────────────────────────────────────────
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DARK_BLUE  = RGBColor(0x00, 0x33, 0x6B)
ORANGE     = RGBColor(0xE8, 0x5A, 0x0E)
LIGHT_GRAY = RGBColor(0xF2, 0xF4, 0xF7)
MID_GRAY   = RGBColor(0xCC, 0xCC, 0xCC)
TEXT_DARK  = RGBColor(0x1A, 0x1A, 0x2E)

STATUS_COLORS = {
    "open":        RGBColor(0xDC, 0x26, 0x26),
    "in progress": RGBColor(0xD9, 0x77, 0x06),
    "closed":      RGBColor(0x16, 0xA3, 0x4A),
}

AREAS = [
    "RMHS", "SINTER", "COKE OVEN", "POWERPLANT",
    "BLAST FURNACE 2", "BLAST FURNACE 3", "PCI", "PCM", "SECONDARY OPERATIONS"
]

W = Inches(13.33)
H = Inches(7.5)

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")


def _logo_stream():
    """Load logo from file. Returns None if not found."""
    try:
        return open(LOGO_PATH, "rb")
    except Exception:
        return None


def _add_rect(slide, left, top, width, height, color: RGBColor):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _add_text(slide, text, left, top, width, height,
              size, bold=False, color=TEXT_DARK, align=PP_ALIGN.LEFT, wrap=True):
    tx = slide.shapes.add_textbox(left, top, width, height)
    tf = tx.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def _status_color(status: str) -> RGBColor:
    s = (status or "open").lower()
    for key, col in STATUS_COLORS.items():
        if key in s:
            return col
    return STATUS_COLORS["open"]


def _fetch_image(obs: dict):
    """
    Decode the base64 photo stored in `image_b64`.
    Returns a BytesIO stream, or None if no photo is present.
    """
    data_url = obs.get("image_b64", "")
    if not data_url:
        return None
    try:
        # Strip the data-URL prefix: "data:image/jpeg;base64,<data>"
        if "," in data_url:
            data_url = data_url.split(",", 1)[1]
        return io.BytesIO(base64.b64decode(data_url))
    except Exception:
        return None


def _add_logo(slide, left, top, width, height):
    logo = _logo_stream()
    if logo:
        try:
            slide.shapes.add_picture(logo, left, top, width, height)
            logo.close()
        except Exception:
            pass


# ─── TITLE SLIDE ──────────────────────────────────────────────────────────────
def _add_title_slide(prs, title_line1, title_line2, obs_list):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = WHITE

    # Top navy bar
    _add_rect(slide, 0, 0, W, Inches(1.3), DARK_BLUE)

    # Logo
    _add_logo(slide, Inches(0.2), Inches(0.1), Inches(2.5), Inches(1.1))

    # Header text
    _add_text(slide, "ESL STEEL LIMITED | Safety Observations",
              Inches(3), Inches(0.15), Inches(9.8), Inches(1.0),
              size=18, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)

    # Orange accent
    _add_rect(slide, 0, Inches(1.3), W, Inches(0.08), ORANGE)

    # Main title
    _add_text(slide, title_line1,
              Inches(1.5), Inches(2.0), Inches(10), Inches(1.2),
              size=36, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)

    # Subtitle
    _add_text(slide, title_line2,
              Inches(1.5), Inches(3.3), Inches(10), Inches(0.8),
              size=22, color=ORANGE, align=PP_ALIGN.CENTER)

    # Stats cards
    total    = len(obs_list)
    open_c   = sum(1 for o in obs_list if "open" in o.get("status","").lower() and "progress" not in o.get("status","").lower())
    prog_c   = sum(1 for o in obs_list if "progress" in o.get("status","").lower())
    closed_c = sum(1 for o in obs_list if "closed" in o.get("status","").lower())

    stats = [
        ("Total",       str(total),    DARK_BLUE),
        ("Open",        str(open_c),   STATUS_COLORS["open"]),
        ("In Progress", str(prog_c),   STATUS_COLORS["in progress"]),
        ("Closed",      str(closed_c), STATUS_COLORS["closed"]),
    ]
    box_w   = Inches(2.5)
    start_x = Inches(1.9)
    for i, (label, val, col) in enumerate(stats):
        bx = start_x + i * (box_w + Inches(0.3))
        _add_rect(slide, bx, Inches(4.4), box_w, Inches(1.8), col)
        _add_text(slide, val,  bx, Inches(4.4), box_w, Inches(1.0),
                  size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text(slide, label, bx, Inches(5.3), box_w, Inches(0.5),
                  size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Footer
    _add_rect(slide, 0, Inches(6.9), W, Inches(0.6), LIGHT_GRAY)
    _add_text(slide,
              f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Vedanta Group — ESL Steel Limited",
              Inches(0.3), Inches(6.95), Inches(12.5), Inches(0.4),
              size=9, color=MID_GRAY, align=PP_ALIGN.CENTER)


# ─── OBSERVATION CARD ─────────────────────────────────────────────────────────
def _add_obs_card(slide, obs: dict, left, top, width, height):
    status = obs.get("status", "Open")
    sc     = _status_color(status)

    # Card background
    card = slide.shapes.add_shape(1, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_GRAY
    card.line.color.rgb = MID_GRAY
    card.line.width = Pt(0.5)

    # Status strip
    _add_rect(slide, left, top, width, Inches(0.18), sc)

    # Ref No
    ref_no = obs.get("ref_no", "—")
    _add_text(slide, ref_no, left + Inches(0.12), top + Inches(0.22),
              Inches(3), Inches(0.3), size=9, bold=True, color=DARK_BLUE)

    # Status badge
    badge_w = Inches(1.4)
    _add_rect(slide, left + width - badge_w - Inches(0.1), top + Inches(0.22),
              badge_w, Inches(0.28), sc)
    _add_text(slide, status.upper(),
              left + width - badge_w - Inches(0.1), top + Inches(0.22),
              badge_w, Inches(0.28), size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Photo
    img_top = top + Inches(0.58)
    img_h   = Inches(2.1)
    img_data = _fetch_image(obs)
    if img_data:
        try:
            slide.shapes.add_picture(img_data, left + Inches(0.12), img_top,
                                     width - Inches(0.24), img_h)
        except Exception:
            _add_rect(slide, left + Inches(0.12), img_top, width - Inches(0.24), img_h, MID_GRAY)
    else:
        _add_rect(slide, left + Inches(0.12), img_top, width - Inches(0.24), img_h, MID_GRAY)
        _add_text(slide, "No Photo",
                  left + Inches(0.12), img_top + Inches(0.85),
                  width - Inches(0.24), Inches(0.4),
                  size=9, color=WHITE, align=PP_ALIGN.CENTER)

    # Details
    detail_top = img_top + img_h + Inches(0.1)
    obs_text   = obs.get("observation", "—")
    _add_text(slide, obs_text,
              left + Inches(0.12), detail_top,
              width - Inches(0.24), Inches(0.65),
              size=8, bold=True, color=TEXT_DARK, wrap=True)
    detail_top += Inches(0.68)

    for icon, val in [
        ("📅", obs.get("datetime",     "—")),
        ("📍", obs.get("area",         "—")),
        ("👷", obs.get("responsible",  "—")),
        ("🎯", obs.get("target_date",  "—")),
    ]:
        _add_text(slide, f"{icon} {val}",
                  left + Inches(0.12), detail_top,
                  width - Inches(0.24), Inches(0.28),
                  size=8, color=TEXT_DARK, wrap=True)
        detail_top += Inches(0.3)


# ─── CONTENT SLIDE (2 cards) ──────────────────────────────────────────────────
def _add_content_slide(prs, obs_pair: list, area_label: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = WHITE

    # Header
    _add_rect(slide, 0, 0, W, Inches(0.65), DARK_BLUE)
    _add_logo(slide, Inches(0.1), Inches(0.05), Inches(1.6), Inches(0.56))
    _add_text(slide, f"ESL STEEL | {area_label} | Safety Observations",
              Inches(1.9), Inches(0.1), Inches(11.1), Inches(0.5),
              size=11, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)
    _add_rect(slide, 0, Inches(0.65), W, Inches(0.05), ORANGE)

    # Two cards
    pad    = Inches(0.18)
    card_w = (W - pad * 3) / 2
    card_h = H - Inches(0.7) - pad * 2

    for i, obs in enumerate(obs_pair):
        _add_obs_card(slide, obs, pad + i * (card_w + pad), Inches(0.78), card_w, card_h)


# ─── PUBLIC API ───────────────────────────────────────────────────────────────
def generate_pptx(obs_list: list, area: str = "ALL", date_label: str = "") -> bytes:
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    area_display = area if area != "ALL" else "All Areas"
    subtitle     = date_label or datetime.now().strftime("%d/%m/%Y")

    _add_title_slide(prs, f"Safety Observations — {area_display}", subtitle, obs_list)

    for i in range(0, len(obs_list), 2):
        _add_content_slide(prs, obs_list[i:i+2], area_display)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def generate_pptx_per_area(obs_list: list, date_label: str = "") -> dict:
    result = {}
    for area in AREAS:
        filtered = [o for o in obs_list if area.lower() in o.get("area", "").lower()]
        if filtered:
            result[area] = generate_pptx(filtered, area=area, date_label=date_label)
    return result
