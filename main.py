# main.py
# Sapience ODCs — Render + FastAPI — Excel only (XlsxWriter)
# v1.1.0 (provider block + date in banner + bill-to compressed + advance as line item)
#
# Changes vs v1.0.6:
# - Date moved to top-right inside banner (no longer in left meta rows)
# - Left block becomes "DATOS DEL PROVEEDOR" with fields: NOMBRE, RFC, E-MAIL, SERVICIO, PROYECTO
# - Bill-to compressed into ONE merged cell (name + RFC + address)
# - Advance/Anticipo becomes a row inside the Concepts table (negative amount)
# - Summary becomes ONLY "TOTAL" (since anticipo is already in table)
# - Keeps Terms page (Page 2) as-is

from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from datetime import datetime
from io import BytesIO
import math
import textwrap
import requests
import struct

import xlsxwriter
from xlsxwriter.utility import xl_range

from terms import TERMS_TITLE, TERMS_LEFT, TERMS_RIGHT

app = FastAPI(title="Sapience ODCs (Excel)", version="1.1.0")


# -----------------------------
# Models
# -----------------------------
class ODCItem(BaseModel):
    concept: str = ""
    unit_cost: float = 0
    units: float = 0
    subtotal: Optional[float] = None  # if omitted, we compute = unit_cost * units


class ODCPayload(BaseModel):
    # Header
    odc_number: str
    date_str: str

    # Provider block
    provider: str  # provider name (kept for backwards compat)
    provider_rfc: Optional[str] = ""
    provider_email: Optional[str] = ""

    service: str
    project: str

    # Bill-to
    bill_to_title: str = "FACTURAR A:"
    bill_to_name: str
    bill_to_rfc: str
    bill_to_address_1: str
    bill_to_address_2: str

    # Items (concept rows)
    items: List[ODCItem] = Field(default_factory=list)

    # Advance (now becomes a line item if provided)
    advance_amount: Optional[float] = None

    currency_symbol: str = "$"

    # Logo
    logo_url: Optional[str] = "https://i.postimg.cc/8CxrRbft/logo-sapience-blanco-alargado.png"


# -----------------------------
# Helpers
# -----------------------------
def _png_size(img_bytes: bytes):
    """Return (w,h) for PNG bytes or None."""
    if len(img_bytes) < 24 or img_bytes[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    try:
        w = struct.unpack(">I", img_bytes[16:20])[0]
        h = struct.unpack(">I", img_bytes[20:24])[0]
        return int(w), int(h)
    except Exception:
        return None


def range_a1(row1: int, col1: int, row2: int, col2: int) -> str:
    """1-based row/col -> Excel range like A1:Z20"""
    return xl_range(row1 - 1, col1 - 1, row2 - 1, col2 - 1)


def safe_float(x) -> float:
    try:
        return float(x)
    except Exception:
        return 0.0


def fill_range(ws, row1: int, col1: int, row2: int, col2: int, fmt):
    """Paint a rectangle explicitly with blank cells (avoid transparent background)."""
    for rr in range(row1, row2 + 1):
        for cc in range(col1, col2 + 1):
            ws.write_blank(rr - 1, cc - 1, "", fmt)


def row_height_for_wrapped_text(
    text: str,
    wrap_width_chars: int,
    base_line_height: float = 11.0,
    extra_lines: float = 0.6,
) -> float:
    """Estimate wrapped lines and return a row height."""
    if not text:
        return base_line_height * (1 + extra_lines)

    paragraphs = str(text).splitlines() or [""]
    total_lines = 0
    for p in paragraphs:
        wrapped = textwrap.wrap(
            p,
            width=max(1, wrap_width_chars),
            break_long_words=True,
            replace_whitespace=False,
        ) or [""]
        total_lines += len(wrapped)

    return base_line_height * (total_lines + extra_lines)


def add_terms_page(
    ws,
    wb,
    *,
    start_row: int,
    white_bg,
    FONT: str,
    TEAL_2: str,
    BLACK: str,
    GRID_LIGHT: str,
):
    """
    Adds Page 2: TERMS_TITLE + 1 column of clauses (TERMS_LEFT + TERMS_RIGHT).
    Returns last row used (1-based).
    """
    ws.set_h_pagebreaks([start_row - 1])

    ALL_TERMS = (TERMS_LEFT or []) + (TERMS_RIGHT or [])

    title_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK, "bg_color": "#FFFFFF",
    })
    section_hdr = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "align": "left", "valign": "top",
        "font_color": BLACK, "bg_color": "#FFFFFF",
    })
    bullet_fmt = wb.add_format({
        "font_name": FONT, "font_size": 7,
        "align": "left", "valign": "top",
        "font_color": BLACK, "bg_color": "#FFFFFF",
        "text_wrap": True,
    })

    COL1, COL2 = 2, 26

    rr = start_row
    ws.set_row(rr - 1, 22)
    fill_range(ws, rr, 2, rr, 26, white_bg)
    ws.merge_range(rr - 1, COL1 - 1, rr - 1, COL2 - 1, TERMS_TITLE, title_fmt)
    rr += 2

    fill_range(ws, rr, 2, rr + 260, 26, white_bg)

    WRAP_WIDTH = 150

    for (hdr, bullets) in ALL_TERMS:
        ws.set_row(rr - 1, 16)
        ws.merge_range(rr - 1, COL1 - 1, rr - 1, COL2 - 1, hdr, section_hdr)
        rr += 1

        for b in bullets:
            text = f"– {b}"
            h = row_height_for_wrapped_text(
                text,
                wrap_width_chars=WRAP_WIDTH,
                base_line_height=9.0,
                extra_lines=0.9,
            )
            ws.set_row(rr - 1, int(max(14, math.ceil(h))))
            ws.merge_range(rr - 1, COL1 - 1, rr - 1, COL2 - 1, text, bullet_fmt)
            rr += 1

        ws.set_row(rr - 1, 8)
        rr += 1

    return rr


# -----------------------------
# Routes
# -----------------------------
@app.get("/health")
def health():
    return {"ok": True, "time": datetime.now().isoformat()}


@app.post("/generate-odc-excel")
def generate_odc_excel(payload: ODCPayload):
    try:
        xlsx_bytes = build_odc_excel(payload)
        filename = f"ODC_{payload.odc_number}_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        return StreamingResponse(
            BytesIO(xlsx_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


# -----------------------------
# Excel Builder (XlsxWriter)
# -----------------------------
def build_odc_excel(payload: ODCPayload) -> bytes:
    out = BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("ODC")

    # -------- Palette / Font --------
    FONT = "Montserrat"
    TEAL = "#0F3D4C"
    TEAL_2 = "#0E4A5A"
    WHITE = "#FFFFFF"
    LIGHT_GRAY = "#EFEFEF"
    GRID = "#7C7C7C"
    GRID_LIGHT = "#C9C9C9"
    RED = "#E10600"
    BLACK = "#111111"
    GRAY_TEXT = "#666666"

    # -------- Grid columns (uniform 2.5) --------
    for c in range(0, 26):
        ws.set_column(c, c, 2.5)

    # -------- Explicit white canvas --------
    white_bg = wb.add_format({"bg_color": WHITE})
    fill_range(ws, 1, 1, 420, 50, white_bg)

    banner_fill = wb.add_format({"bg_color": TEAL})
    gray_fill = wb.add_format({"bg_color": LIGHT_GRAY})

    # -------- Formats --------
    odc_box_lbl = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
        "border": 1, "border_color": GRID_LIGHT,
    })
    odc_box_val = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": RED, "bg_color": WHITE,
        "border": 1, "border_color": GRID_LIGHT,
    })

    date_in_banner = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "align": "right", "valign": "vcenter",
        "font_color": WHITE, "bg_color": TEAL,
    })

    section_title = wb.add_format({
        "font_name": FONT, "font_size": 16, "bold": True,
        "align": "left", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })

    label_gray = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": LIGHT_GRAY
    })
    label_white = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })

    value_gray_div = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": LIGHT_GRAY,
        "text_wrap": True,
        "left": 1, "left_color": GRID_LIGHT,
    })
    value_white_div = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE,
        "text_wrap": True,
        "left": 1, "left_color": GRID_LIGHT,
    })

    bill_title_fmt = wb.add_format({
        "font_name": FONT, "font_size": 16, "bold": True,
        "align": "left", "valign": "top",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    bill_block_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": False,
        "align": "left", "valign": "top",
        "font_color": BLACK, "bg_color": WHITE,
        "text_wrap": True,
    })
    bill_block_bold = wb.add_format({
        "font_name": FONT, "font_size": 12, "bold": True,
        "align": "left", "valign": "top",
        "font_color": BLACK, "bg_color": WHITE,
        "text_wrap": True,
    })

    th_fmt = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": WHITE, "bg_color": TEAL_2,
        "border": 1, "border_color": GRID,
    })

    concept_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK,
        "text_wrap": True,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    concept_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK,
        "text_wrap": True,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
    })

    money_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })
    money_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })
    money_w_red = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": RED,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })
    money_g_red = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": RED,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })

    units_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    units_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
    })

    total_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 18, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    total_value_fmt = wb.add_format({
        "font_name": FONT, "font_size": 18, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })

    # -------- Banner (rows 1..3) --------
    ws.set_row(0, 16)
    ws.set_row(1, 16)
    ws.set_row(2, 16)
    fill_range(ws, 1, 2, 3, 26, banner_fill)  # B1:Z3

    # Date in banner (top-right) — spans X..Z on row 1
    ws.merge_range(0, 23, 0, 25, payload.date_str, date_in_banner)

    # ODC box top-right within banner (rows 2..3 area visually)
    ws.merge_range(0, 19, 1, 22, "ODC #:", odc_box_lbl)           # T..W
    ws.merge_range(0, 23, 1, 25, payload.odc_number, odc_box_val) # X..Z (date is on row 1 only)

    # Insert logo
    if payload.logo_url:
        try:
            resp = requests.get(payload.logo_url, timeout=15)
            resp.raise_for_status()
            img = resp.content
            wh = _png_size(img)

            banner_h_px = 64
            safe_h_px = 52
            safe_w_px = 190

            x_scale = y_scale = 0.09
            x_off = 10
            y_off = 6

            if wh:
                w_px, h_px = wh
                scale = min(safe_w_px / max(1, w_px), safe_h_px / max(1, h_px))
                scale = max(0.07, min(scale, 0.095))
                x_scale = y_scale = scale

                scaled_h = h_px * y_scale
                y_off = max(0, int((banner_h_px - scaled_h) / 2))

            ws.insert_image(
                0, 2, "logo.png",
                {
                    "image_data": BytesIO(img),
                    "x_scale": x_scale,
                    "y_scale": y_scale,
                    "x_offset": x_off,
                    "y_offset": y_off,
                    "object_position": 1,
                },
            )
        except Exception:
            pass

    # -------- Provider section title --------
    # Row 4 (Excel): big "DATOS DEL PROVEEDOR" on left
    ws.set_row(3, 22)
    ws.merge_range(3, 1, 3, 13, "DATOS DEL PROVEEDOR", section_title)

    # -------- Provider block rows 5..9 (Excel) --------
    provider_rows = [
        ("NOMBRE:", payload.provider),
        ("RFC:", payload.provider_rfc or ""),
        ("E-MAIL:", payload.provider_email or ""),
        ("SERVICIO:", payload.service),
        ("PROYECTO:", payload.project),
    ]

    for i, (lab, val) in enumerate(provider_rows):
        rr = 5 + i
        ws.set_row(rr - 1, 20)
        is_gray = (i % 2 == 0)  # starts with gray stripe (looks like template)
        fill_range(ws, rr, 2, rr, 14, gray_fill if is_gray else white_bg)

        ws.merge_range(rr - 1, 1, rr - 1, 4, lab, label_gray if is_gray else label_white)
        ws.merge_range(rr - 1, 5, rr - 1, 13, val, value_gray_div if is_gray else value_white_div)

    # -------- Bill-to block (compressed) --------
    # Title on row 4 (same row as provider title area but right side)
    fill_range(ws, 4, 15, 9, 26, white_bg)
    ws.merge_range(3, 14, 3, 25, payload.bill_to_title, bill_title_fmt)

    bill_lines = [
        payload.bill_to_name.strip(),
        f"RFC: {payload.bill_to_rfc}".strip(),
        payload.bill_to_address_1.strip(),
        payload.bill_to_address_2.strip(),
    ]
    bill_text = "\n".join([x for x in bill_lines if x])

    # Merge rows 5..9 into one cell O..Z
    ws.merge_range(4, 14, 8, 25, bill_text, bill_block_fmt)

    # Spacer row 10
    ws.set_row(9, 10)

    # -------- Table header row 11 (Excel) --------
    header_row = 11
    ws.set_row(header_row - 1, 26)

    ws.merge_range(header_row - 1, 1, header_row - 1, 13, "Concepto", th_fmt)
    ws.merge_range(header_row - 1, 14, header_row - 1, 18, "Costo unitario", th_fmt)
    ws.merge_range(header_row - 1, 19, header_row - 1, 21, "Unidades", th_fmt)
    ws.merge_range(header_row - 1, 22, header_row - 1, 25, "Subtotal", th_fmt)

    # -------- Items start row 12 (Excel) --------
    start_items = 12
    base_items = payload.items or []

    # If advance exists: append as line item with negative amount (units=1)
    adv = safe_float(payload.advance_amount) if payload.advance_amount is not None else 0.0
    items: List[ODCItem] = list(base_items)

    if adv != 0.0:
        # Make it negative in the table
        adv_val = -abs(adv)
        items.append(ODCItem(
            concept=f"Anticipo ODC {payload.odc_number}",
            unit_cost=adv_val,
            units=1,
            subtotal=adv_val,
        ))

    if not items:
        items = [ODCItem(concept="", unit_cost=0, units=0)]

    max_items = min(len(items), 18)
    wrap_chars = 58
    min_row_h = 26

    last_item_row = start_items - 1
    computed_total = 0.0

    for idx in range(max_items):
        rr = start_items + idx
        it = items[idx]

        zebra = (idx % 2 == 1)
        row_fill = gray_fill if zebra else white_bg
        fill_range(ws, rr, 2, rr, 26, row_fill)

        unit_cost = safe_float(it.unit_cost)
        units = safe_float(it.units)
        subtotal = safe_float(it.subtotal) if it.subtotal is not None else (unit_cost * units)

        computed_total += subtotal

        needed = row_height_for_wrapped_text(it.concept, wrap_chars, base_line_height=11.0, extra_lines=0.7)
        ws.set_row(rr - 1, int(max(min_row_h, math.ceil(needed))))

        ws.merge_range(rr - 1, 1, rr - 1, 13, it.concept, concept_g if zebra else concept_w)

        # Red formatting for negative (anticipo)
        is_negative = subtotal < 0 or unit_cost < 0
        mfmt = (money_g_red if zebra else money_w_red) if is_negative else (money_g if zebra else money_w)

        ws.merge_range(rr - 1, 14, rr - 1, 18, unit_cost, mfmt)
        ws.merge_range(rr - 1, 19, rr - 1, 21, units, units_g if zebra else units_w)
        ws.merge_range(rr - 1, 22, rr - 1, 25, subtotal, mfmt)

        last_item_row = rr

    # -------- TOTAL only (bottom right) --------
    total_row = last_item_row + 3
    ws.set_row(total_row - 1, 28)
    fill_range(ws, total_row, 2, total_row, 26, white_bg)

    ws.merge_range(total_row - 1, 19, total_row - 1, 21, "TOTAL:", total_label_fmt)
    ws.merge_range(total_row - 1, 22, total_row - 1, 25, computed_total, total_value_fmt)

    # ✅ Terms on Page 2
    terms_start_row = total_row + 6
    terms_end_row = add_terms_page(
        ws,
        wb,
        start_row=terms_start_row,
        white_bg=white_bg,
        FONT=FONT,
        TEAL_2=TEAL_2,
        BLACK=BLACK,
        GRID_LIGHT=GRID_LIGHT,
    )

    # -------- Print settings --------
    ws.set_portrait()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.fit_to_pages(1, 0)

    ws.print_area(range_a1(1, 1, terms_end_row + 2, 26))

    wb.close()
    return out.getvalue()
