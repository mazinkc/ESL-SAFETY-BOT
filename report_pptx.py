"""
ESL Safety Observation — PPTX Report Generator
White background | Vedanta/ESL Logo | 2 observations per slide
Supports: full report, daily report, per-area report
"""

import io
import base64
import requests
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ─── LOGO (Vedanta / ESL Steel — base64 embedded) ────────────────────────────
LOGO_B64 = "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"

# ─── BRAND COLORS ─────────────────────────────────────────────────────────────
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DARK_BLUE  = RGBColor(0x00, 0x33, 0x6B)   # Vedanta navy
ORANGE     = RGBColor(0xE8, 0x5A, 0x0E)   # Vedanta orange
LIGHT_GRAY = RGBColor(0xF2, 0xF4, 0xF7)
MID_GRAY   = RGBColor(0xCC, 0xCC, 0xCC)
TEXT_DARK  = RGBColor(0x1A, 0x1A, 0x2E)

STATUS_COLORS = {
    "open":        RGBColor(0xDC, 0x26, 0x26),  # red
    "in progress": RGBColor(0xD9, 0x77, 0x06),  # amber
    "closed":      RGBColor(0x16, 0xA3, 0x4A),  # green
}

AREAS = [
    "RMHS", "SINTER", "COKE OVEN", "POWERPLANT",
    "BLAST FURNACE 2", "BLAST FURNACE 3", "PCI", "PCM", "SECONDARY OPERATIONS"
]

# ─── SLIDE DIMENSIONS (Widescreen 13.33" × 7.5") ─────────────────────────────
W  = Inches(13.33)
H  = Inches(7.5)


def _logo_stream() -> io.BytesIO:
    # Fix base64 padding if needed
    padded = LOGO_B64 + "=" * (4 - len(LOGO_B64) % 4) if len(LOGO_B64) % 4 else LOGO_B64
    return io.BytesIO(base64.b64decode(padded))


def _fill_white(shape):
    shape.fill.solid()
    shape.fill.fore_color.rgb = WHITE


def _add_rect(slide, left, top, width, height, color: RGBColor):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _tf(shape, text, size, bold=False, color=TEXT_DARK, align=PP_ALIGN.LEFT, wrap=True):
    tf = shape.text_frame
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


def _fetch_image(url: str) -> io.BytesIO | None:
    if not url:
        return None
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return io.BytesIO(r.content)
    except Exception:
        return None


# ─── TITLE SLIDE ─────────────────────────────────────────────────────────────
def _add_title_slide(prs, title_line1, title_line2, obs_list):
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)

    # White background
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = WHITE

    # Top navy bar
    bar = _add_rect(slide, 0, 0, W, Inches(1.3), DARK_BLUE)

    # Logo in top-left of bar
    logo = slide.shapes.add_picture(_logo_stream(), Inches(0.2), Inches(0.1), Inches(2.5), Inches(1.1))

    # Title text in bar
    tx = slide.shapes.add_textbox(Inches(3), Inches(0.15), Inches(9.5), Inches(1.0))
    tf = tx.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    run.text = "ESL STEEL LIMITED | Safety Observations"
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = WHITE

    # Orange accent bar
    _add_rect(slide, 0, Inches(1.3), W, Inches(0.08), ORANGE)

    # Main title
    tx2 = slide.shapes.add_textbox(Inches(1.5), Inches(2.0), Inches(10), Inches(1.2))
    tf2 = tx2.text_frame
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = title_line1
    r2.font.size = Pt(36)
    r2.font.bold = True
    r2.font.color.rgb = DARK_BLUE

    # Subtitle
    tx3 = slide.shapes.add_textbox(Inches(1.5), Inches(3.3), Inches(10), Inches(0.8))
    tf3 = tx3.text_frame
    p3 = tf3.paragraphs[0]
    p3.alignment = PP_ALIGN.CENTER
    r3 = p3.add_run()
    r3.text = title_line2
    r3.font.size = Pt(22)
    r3.font.color.rgb = ORANGE

    # Stats row
    total    = len(obs_list)
    open_c   = sum(1 for o in obs_list if "open" in (o.get("status","")).lower() and "progress" not in (o.get("status","")).lower())
    prog_c   = sum(1 for o in obs_list if "progress" in (o.get("status","")).lower())
    closed_c = sum(1 for o in obs_list if "closed" in (o.get("status","")).lower())

    stats = [
        ("Total", str(total), DARK_BLUE),
        ("Open", str(open_c), STATUS_COLORS["open"]),
        ("In Progress", str(prog_c), STATUS_COLORS["in progress"]),
        ("Closed", str(closed_c), STATUS_COLORS["closed"]),
    ]
    box_w = Inches(2.5)
    start_x = Inches(1.9)
    for i, (label, val, col) in enumerate(stats):
        bx = start_x + i * (box_w + Inches(0.3))
        card = _add_rect(slide, bx, Inches(4.4), box_w, Inches(1.8), col)
        tx4 = slide.shapes.add_textbox(bx, Inches(4.4), box_w, Inches(1.0))
        tf4 = tx4.text_frame
        p4 = tf4.paragraphs[0]
        p4.alignment = PP_ALIGN.CENTER
        r4 = p4.add_run()
        r4.text = val
        r4.font.size = Pt(36)
        r4.font.bold = True
        r4.font.color.rgb = WHITE

        tx5 = slide.shapes.add_textbox(bx, Inches(5.3), box_w, Inches(0.5))
        tf5 = tx5.text_frame
        p5 = tf5.paragraphs[0]
        p5.alignment = PP_ALIGN.CENTER
        r5 = p5.add_run()
        r5.text = label
        r5.font.size = Pt(13)
        r5.font.bold = True
        r5.font.color.rgb = WHITE

    # Bottom footer
    _add_rect(slide, 0, Inches(6.9), W, Inches(0.6), LIGHT_GRAY)
    tx6 = slide.shapes.add_textbox(Inches(0.3), Inches(6.95), Inches(12), Inches(0.4))
    tf6 = tx6.text_frame
    p6 = tf6.paragraphs[0]
    p6.alignment = PP_ALIGN.CENTER
    r6 = p6.add_run()
    r6.text = f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Vedanta Group — ESL Steel Limited"
    r6.font.size = Pt(9)
    r6.font.color.rgb = MID_GRAY

    return slide


# ─── OBSERVATION CARD (half-slide) ───────────────────────────────────────────
def _add_obs_card(slide, obs: dict, left: Emu, top: Emu, width: Emu, height: Emu):
    status = obs.get("status", "Open")
    sc     = _status_color(status)

    # Card background
    card = _add_rect(slide, left, top, width, height, LIGHT_GRAY)
    card.line.color.rgb = MID_GRAY
    card.line.width = Pt(0.5)

    # Status colour strip on top
    _add_rect(slide, left, top, width, Inches(0.18), sc)

    # Ref + Status badge
    ref_no = obs.get("ref_no", "—")
    badge_w = Inches(1.4)
    badge = _add_rect(slide, left + width - badge_w - Inches(0.1), top + Inches(0.22),
                      badge_w, Inches(0.28), sc)
    tx_b = slide.shapes.add_textbox(left + width - badge_w - Inches(0.1), top + Inches(0.22),
                                    badge_w, Inches(0.28))
    tf_b = tx_b.text_frame
    p_b  = tf_b.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    r_b  = p_b.add_run()
    r_b.text = status.upper()
    r_b.font.size = Pt(7)
    r_b.font.bold = True
    r_b.font.color.rgb = WHITE

    tx_ref = slide.shapes.add_textbox(left + Inches(0.12), top + Inches(0.22),
                                      Inches(3.5), Inches(0.3))
    _tf(tx_ref, ref_no, 9, bold=True, color=DARK_BLUE)

    # Photo section
    img_top  = top + Inches(0.58)
    img_h    = Inches(2.1)
    img_data = _fetch_image(obs.get("image_url", ""))
    if img_data:
        try:
            slide.shapes.add_picture(img_data, left + Inches(0.12), img_top,
                                     width - Inches(0.24), img_h)
        except Exception:
            _add_rect(slide, left + Inches(0.12), img_top,
                      width - Inches(0.24), img_h, MID_GRAY)
    else:
        ph = _add_rect(slide, left + Inches(0.12), img_top,
                       width - Inches(0.24), img_h, MID_GRAY)
        tx_ph = slide.shapes.add_textbox(left + Inches(0.12), img_top + Inches(0.85),
                                         width - Inches(0.24), Inches(0.4))
        tf_ph = tx_ph.text_frame
        p_ph  = tf_ph.paragraphs[0]
        p_ph.alignment = PP_ALIGN.CENTER
        r_ph  = p_ph.add_run()
        r_ph.text = "No Photo"
        r_ph.font.size = Pt(9)
        r_ph.font.color.rgb = WHITE

    # Details section
    detail_top = img_top + img_h + Inches(0.1)
    fields = [
        ("📅", obs.get("datetime", "—")),
        ("📍", obs.get("area", "—")),
        ("👷", obs.get("responsible", "—")),
        ("🎯", obs.get("target_date", "—")),
    ]

    obs_text = obs.get("observation", "—")
    tx_obs = slide.shapes.add_textbox(left + Inches(0.12), detail_top,
                                      width - Inches(0.24), Inches(0.65))
    tf_obs = tx_obs.text_frame
    tf_obs.word_wrap = True
    p_obs  = tf_obs.paragraphs[0]
    r_obs  = p_obs.add_run()
    r_obs.text = obs_text
    r_obs.font.size = Pt(8)
    r_obs.font.bold = True
    r_obs.font.color.rgb = TEXT_DARK

    detail_top += Inches(0.68)
    for icon, val in fields:
        tx_f = slide.shapes.add_textbox(left + Inches(0.12), detail_top,
                                        width - Inches(0.24), Inches(0.28))
        tf_f = tx_f.text_frame
        tf_f.word_wrap = True
        p_f  = tf_f.paragraphs[0]
        r_f  = p_f.add_run()
        r_f.text = f"{icon} {val}"
        r_f.font.size = Pt(8)
        r_f.font.color.rgb = TEXT_DARK
        detail_top += Inches(0.3)


# ─── CONTENT SLIDE (2 observations) ──────────────────────────────────────────
def _add_content_slide(prs, obs_pair: list, area_label: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = WHITE

    # Header bar
    _add_rect(slide, 0, 0, W, Inches(0.65), DARK_BLUE)
    logo = slide.shapes.add_picture(_logo_stream(), Inches(0.1), Inches(0.05), Inches(1.6), Inches(0.56))

    tx_h = slide.shapes.add_textbox(Inches(1.9), Inches(0.1), Inches(9), Inches(0.5))
    tf_h = tx_h.text_frame
    p_h  = tf_h.paragraphs[0]
    p_h.alignment = PP_ALIGN.RIGHT
    r_h  = p_h.add_run()
    r_h.text = f"ESL STEEL | {area_label} | Safety Observations"
    r_h.font.size = Pt(11)
    r_h.font.bold = True
    r_h.font.color.rgb = WHITE

    _add_rect(slide, 0, Inches(0.65), W, Inches(0.05), ORANGE)

    # Two cards side by side
    padding = Inches(0.18)
    card_w  = (W - padding * 3) / 2
    card_h  = H - Inches(0.65) - Inches(0.05) - padding * 2

    for i, obs in enumerate(obs_pair):
        left = padding + i * (card_w + padding)
        _add_obs_card(slide, obs, left, Inches(0.78), card_w, card_h)

    # If only 1 obs on slide, fill right half with blank card style
    if len(obs_pair) == 1:
        left2 = padding + card_w + padding
        ph = _add_rect(slide, left2, Inches(0.78), card_w, card_h, LIGHT_GRAY)


# ─── PUBLIC API ───────────────────────────────────────────────────────────────
def generate_pptx(obs_list: list, area: str = "ALL", date_label: str = "") -> bytes:
    """
    Generate PPTX for the given observations.
    area: "ALL" or area name
    date_label: descriptive string shown in title slide subtitle
    """
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    if not obs_list:
        # Return minimal PPTX with just title
        _add_title_slide(prs, "No Observations Found", date_label or "", [])
        buf = io.BytesIO()
        prs.save(buf)
        return buf.getvalue()

    area_display = area if area != "ALL" else "All Areas"
    subtitle = date_label if date_label else datetime.now().strftime("%d/%m/%Y")

    _add_title_slide(prs, f"Safety Observations — {area_display}", subtitle, obs_list)

    # Content slides: 2 per slide
    for i in range(0, len(obs_list), 2):
        pair = obs_list[i:i+2]
        _add_content_slide(prs, pair, area_display)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def generate_pptx_per_area(obs_list: list, date_label: str = "") -> dict:
    """
    Returns dict: { area_name: pptx_bytes }
    Only areas that have observations are included.
    """
    result = {}
    for area in AREAS:
        filtered = [o for o in obs_list if area.lower() in o.get("area","").lower()]
        if filtered:
            result[area] = generate_pptx(filtered, area=area, date_label=date_label)
    return result
