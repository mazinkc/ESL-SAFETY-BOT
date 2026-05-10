"""
ESL Safety Observation - PPTX Report Generator
White background | Vedanta/ESL Logo | 2 observations per slide
Embeds photos directly in PPTX. No Telegram hyperlinks.
"""

import io
import os
import base64
from datetime import datetime

import requests
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK_BLUE = RGBColor(0x00, 0x33, 0x6B)
ORANGE = RGBColor(0xE8, 0x5A, 0x0E)
LIGHT_GRAY = RGBColor(0xF2, 0xF4, 0xF7)
MID_GRAY = RGBColor(0xCC, 0xCC, 0xCC)
TEXT_DARK = RGBColor(0x1A, 0x1A, 0x2E)

STATUS_COLORS = {
    "open": RGBColor(0xDC, 0x26, 0x26),
    "in progress": RGBColor(0xD9, 0x77, 0x06),
    "closed": RGBColor(0x16, 0xA3, 0x4A),
}

AREAS = [
    "RMHS", "SINTER", "COKE OVEN", "POWERPLANT",
    "BLAST FURNACE 2", "BLAST FURNACE 3", "PCI", "PCM",
    "SECONDARY OPERATIONS"
]

W = Inches(13.33)
H = Inches(7.5)

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")


def _logo_stream():
    try:
        return open(LOGO_PATH, "rb")
    except Exception:
        return None


def _add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def _add_text(slide, text, left, top, width, height,
              size, bold=False, color=TEXT_DARK,
              align=PP_ALIGN.LEFT, wrap=True):
    tx = slide.shapes.add_textbox(left, top, width, height)
    tf = tx.text_frame
    tf.word_wrap = wrap
    tf.margin_left = Pt(4)
    tf.margin_right = Pt(4)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)

    p = tf.paragraphs[0]
    p.alignment = align

    run = p.add_run()
    run.text = str(text or "")
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color

    return tx


def _status_color(status):
    s = (status or "open").lower()
    for key, col in STATUS_COLORS.items():
        if key in s:
            return col
    return STATUS_COLORS["open"]


def _photo_url_from_obs(obs):
    """
    Supports old records where photo was stored as Telegram/file URL.
    """
    for key in [
        "image_url",
        "photo_url",
        "photo",
        "photo_link",
        "image",
        "image_link",
        "telegram_photo_url",
    ]:
        val = obs.get(key)
        if val and str(val).startswith("http"):
            return str(val)
    return ""


def _fetch_image(obs):
    """
    Return a clean JPEG stream for python-pptx.

    Supports:
    1. image_b64 data URL
    2. closure_b64 data URL
    3. old Telegram/photo URL fields

    Important: this embeds the image into PPTX. It does not create hyperlinks.
    """
    try:
        data_url = obs.get("image_b64") or obs.get("closure_b64") or ""

        if data_url:
            if "," in data_url:
                data_url = data_url.split(",", 1)[1]
            raw = base64.b64decode(data_url)
        else:
            photo_url = _photo_url_from_obs(obs)
            if not photo_url:
                return None

            response = requests.get(photo_url, timeout=20)
            response.raise_for_status()
            raw = response.content

        img = Image.open(io.BytesIO(raw))
        img = img.convert("RGB")
        img.thumbnail((1600, 1600), Image.LANCZOS)

        out = io.BytesIO()
        img.save(out, format="JPEG", quality=88, optimize=True)
        out.seek(0)
        return out

    except Exception as e:
        print(f"Photo embed failed: {e}")
        return None


def _add_logo(slide, left, top, width, height):
    logo = _logo_stream()
    if logo:
        try:
            slide.shapes.add_picture(logo, left, top, width, height)
        except Exception:
            pass
        finally:
            logo.close()


def _add_photo_box(slide, obs, left, top, width, height):
    img_data = _fetch_image(obs)

    if img_data:
        try:
            slide.shapes.add_picture(img_data, left, top, width, height)
            return
        except Exception as e:
            print(f"PPT photo add failed: {e}")
            label = "Photo Error"
    else:
        label = "No Photo"

    _add_rect(slide, left, top, width, height, MID_GRAY)
    _add_text(
        slide, label,
        left, top + height / 2 - Inches(0.2),
        width, Inches(0.4),
        size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER
    )


def _add_title_slide(prs, title_line1, title_line2, obs_list):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = WHITE

    _add_rect(slide, 0, 0, W, Inches(1.3), DARK_BLUE)
    _add_logo(slide, Inches(0.2), Inches(0.1), Inches(2.5), Inches(1.1))

    _add_text(
        slide,
        "ESL STEEL LIMITED | Safety Observations",
        Inches(3), Inches(0.15), Inches(9.8), Inches(1.0),
        size=18, bold=True, color=WHITE, align=PP_ALIGN.RIGHT
    )

    _add_rect(slide, 0, Inches(1.3), W, Inches(0.08), ORANGE)

    _add_text(
        slide, title_line1,
        Inches(1.5), Inches(2.0), Inches(10), Inches(1.2),
        size=36, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER
    )

    _add_text(
        slide, title_line2,
        Inches(1.5), Inches(3.3), Inches(10), Inches(0.8),
        size=22, color=ORANGE, align=PP_ALIGN.CENTER
    )

    total = len(obs_list)
    open_c = sum(
        1 for o in obs_list
        if "open" in o.get("status", "").lower()
        and "progress" not in o.get("status", "").lower()
    )
    prog_c = sum(1 for o in obs_list if "progress" in o.get("status", "").lower())
    closed_c = sum(1 for o in obs_list if "closed" in o.get("status", "").lower())

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

        _add_rect(slide, bx, Inches(4.4), box_w, Inches(1.8), col)

        _add_text(
            slide, val,
            bx, Inches(4.4), box_w, Inches(1.0),
            size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER
        )

        _add_text(
            slide, label,
            bx, Inches(5.3), box_w, Inches(0.5),
            size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER
        )

    _add_rect(slide, 0, Inches(6.9), W, Inches(0.6), LIGHT_GRAY)

    _add_text(
        slide,
        f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Vedanta Group - ESL Steel Limited",
        Inches(0.3), Inches(6.95), Inches(12.5), Inches(0.4),
        size=9, color=MID_GRAY, align=PP_ALIGN.CENTER
    )


def _add_obs_card(slide, obs, left, top, width, height):
    status = obs.get("status", "Open")
    sc = _status_color(status)

    card = slide.shapes.add_shape(1, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_GRAY
    card.line.color.rgb = MID_GRAY
    card.line.width = Pt(0.5)

    _add_rect(slide, left, top, width, Inches(0.18), sc)

    ref_no = obs.get("ref_no", "-")
    _add_text(
        slide, ref_no,
        left + Inches(0.12), top + Inches(0.22),
        Inches(3), Inches(0.3),
        size=9, bold=True, color=DARK_BLUE
    )

    badge_w = Inches(1.4)
    badge_left = left + width - badge_w - Inches(0.1)

    _add_rect(
        slide,
        badge_left,
        top + Inches(0.22),
        badge_w,
        Inches(0.28),
        sc
    )

    _add_text(
        slide, status.upper(),
        badge_left,
        top + Inches(0.22),
        badge_w,
        Inches(0.28),
        size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER
    )

    img_left = left + Inches(0.12)
    img_top = top + Inches(0.58)
    img_w = width - Inches(0.24)
    img_h = Inches(2.1)

    _add_photo_box(slide, obs, img_left, img_top, img_w, img_h)

    detail_top = img_top + img_h + Inches(0.1)

    _add_text(
        slide,
        obs.get("observation", "-"),
        left + Inches(0.12), detail_top,
        width - Inches(0.24), Inches(0.65),
        size=8, bold=True, color=TEXT_DARK, wrap=True
    )

    detail_top += Inches(0.68)

    details = [
        ("Date", obs.get("datetime", "-")),
        ("Area", obs.get("area", "-")),
        ("Responsible", obs.get("responsible", "-")),
        ("Target", obs.get("target_date", "-")),
    ]

    for label, val in details:
        _add_text(
            slide,
            f"{label}: {val}",
            left + Inches(0.12), detail_top,
            width - Inches(0.24), Inches(0.28),
            size=8, color=TEXT_DARK, wrap=True
        )
        detail_top += Inches(0.3)


def _add_content_slide(prs, obs_pair, area_label):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = WHITE

    _add_rect(slide, 0, 0, W, Inches(0.65), DARK_BLUE)
    _add_logo(slide, Inches(0.1), Inches(0.05), Inches(1.6), Inches(0.56))

    _add_text(
        slide,
        f"ESL STEEL | {area_label} | Safety Observations",
        Inches(1.9), Inches(0.1), Inches(11.1), Inches(0.5),
        size=11, bold=True, color=WHITE, align=PP_ALIGN.RIGHT
    )

    _add_rect(slide, 0, Inches(0.65), W, Inches(0.05), ORANGE)

    pad = Inches(0.18)
    card_w = (W - pad * 3) / 2
    card_h = H - Inches(0.7) - pad * 2

    for i, obs in enumerate(obs_pair):
        _add_obs_card(
            slide,
            obs,
            pad + i * (card_w + pad),
            Inches(0.78),
            card_w,
            card_h
        )


def generate_pptx(obs_list, area="ALL", date_label=""):
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    area_display = area if area != "ALL" else "All Areas"
    subtitle = date_label or datetime.now().strftime("%d/%m/%Y")

    _add_title_slide(
        prs,
        f"Safety Observations - {area_display}",
        subtitle,
        obs_list
    )

    for i in range(0, len(obs_list), 2):
        _add_content_slide(prs, obs_list[i:i + 2], area_display)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def generate_pptx_per_area(obs_list, date_label=""):
    result = {}

    for area in AREAS:
        filtered = [
            o for o in obs_list
            if area.lower() in o.get("area", "").lower()
        ]

        if filtered:
            result[area] = generate_pptx(
                filtered,
                area=area,
                date_label=date_label
            )

    return result
