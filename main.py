# main.py
# Sapience ODCs — Render + FastAPI — Excel only (XlsxWriter)
# v1.0.6 (patched-final) + Terms page
# - Adds: Notes/Terms on Page 2 (NOW 1 column), forced page break
# - Does NOT change existing layout/typography, only appends

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

from terms import TERMS_TITLE, TERMS_LEFT, TERMS_RIGHT  # ✅ NEW

app = FastAPI(title="Sapience ODCs (Excel)", version="1.0.6")


# -----------------------------
# Models
# -----------------------------
class ODCItem(BaseModel):
    concept: str = ""
    unit_cost: float = 0
    units: float = 0
    subtotal: Optional[float] = None  # if omitted, we compute = unit_cost * units


class ODCPayload(BaseModel):
    # Header / left meta
    odc_number: str
    date_str: str
    provider: str
    service: str
    project: str

    # Bill-to
    bill_to_title: str = "FACTURAR A:"
    bill_to_name: str
    bill_to_rfc: str
    bill_to_address_1: str
    bill_to_address_2: str

    # Items
    items: List[ODCItem] = Field(default_factory=list)

    # Summary amounts (bottom right)
    sum_amount: Optional[float] = None
    advance_amount: Optional[float] = None
    total_due: Optional[float] = None

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


def _wrap_lines(text: str, width: int) -> List[str]:
    """Wrap a paragraph into multiple lines; keeps words, breaks long words."""
    if not text:
        return [""]
    lines = []
    for p in str(text).splitlines():
        wrapped = textwrap.wrap(
            p,
            width=max(1, width),
            break_long_words=True,
            replace_whitespace=False,
        ) or [""]
        lines.extend(wrapped)
    return lines


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
    Minimal: uses existing palette and Montserrat.
    Returns last row used (1-based).
    """
    # Force a page break so this starts on Page 2 in PDF export
    ws.set_h_pagebreaks([start_row - 1])

    # ✅ Combine both columns into one stream (leave terms.py unchanged)
    ALL_TERMS = (TERMS_LEFT or []) + (TERMS_RIGHT or [])

    # Formats (smaller body text to avoid cutting)
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
        "font_name": FONT, "font_size": 7,  # ✅ smaller
        "align": "left", "valign": "top",
        "font_color": BLACK, "bg_color": "#FFFFFF",
        "text_wrap": True,
    })

    # Single column range (B..Z)
    COL1, COL2 = 2, 26

    # Title row
    rr = start_row
    ws.set_row(rr - 1, 22)
    fill_range(ws, rr, 2, rr, 26, white_bg)
    ws.merge_range(rr - 1, COL1 - 1, rr - 1, COL2 - 1, TERMS_TITLE, title_fmt)
    rr += 2

    # Paint a clean white area to avoid leftovers
    fill_range(ws, rr, 2, rr + 260, 26, white_bg)

    # Wider wrap since it's one big merged column
    WRAP_WIDTH = 150

    for (hdr, bullets) in ALL_TERMS:
        # Header
        ws.set_row(rr - 1, 16)
        ws.merge_range(rr - 1, COL1 - 1, rr - 1, COL2 - 1, hdr, section_hdr)
        rr += 1

        # Bullets
        for b in bullets:
            # ✅ NO metemos saltos de línea manuales; dejamos que Excel haga el wrap
            text = f"– {b}"

            # Height tuned for smaller font (estima wrap sin \n)
            h = row_height_for_wrapped_text(
                text,
                wrap_width_chars=WRAP_WIDTH,
                base_line_height=9.0,
                extra_lines=0.9,
            )
            ws.set_row(rr - 1, int(max(14, math.ceil(h))))
            ws.merge_range(rr - 1, COL1 - 1, rr - 1, COL2 - 1, text, bullet_fmt)
            rr += 1

        # Spacer
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

    # -------- Explicit white canvas (no transparency) --------
    white_bg = wb.add_format({"bg_color": WHITE})
    fill_range(ws, 1, 1, 400, 50, white_bg)  # ✅ slightly larger canvas for page 2

    # Base fills
    banner_fill = wb.add_format({"bg_color": TEAL})
    gray_fill = wb.add_format({"bg_color": LIGHT_GRAY})

    # Thin underline for summary rows (horizontal separators)
    hline_fmt = wb.add_format({"bg_color": WHITE, "bottom": 1, "bottom_color": GRID_LIGHT})

    # ODC box top-right
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

    # Left meta labels/values (alternating stripes)
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

    # ✅ Values with LEFT BORDER (this is the divider line) — no empty divider column
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

    # Bill-to block (painted white)
    bill_title_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "left", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    bill_bold_fmt = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE
    })
    bill_norm_fmt = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE
    })

    # Table header (9 pt)
    th_fmt = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": WHITE, "bg_color": TEAL_2,
        "border": 1, "border_color": GRID,
    })

    # Table cells (8 pt + vertical center)
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

    # Summary formats
    sum_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    adv_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    total_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })

    sum_value_fmt = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "align": "center", "valign": "vcenter",
        "font_color": GRAY_TEXT, "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })
    adv_value_fmt = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": RED, "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })
    total_value_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })

    # -------- Banner: 3 rows x 16 pts (rows 1..3 in Excel) --------
    ws.set_row(0, 16)
    ws.set_row(1, 16)
    ws.set_row(2, 16)

    fill_range(ws, 1, 2, 3, 26, banner_fill)  # B1:Z3

    # ODC box top-right within banner (rows 1..2)
    ws.merge_range(0, 19, 1, 22, "ODC #:", odc_box_lbl)           # T..W
    ws.merge_range(0, 23, 1, 25, payload.odc_number, odc_box_val) # X..Z

    # Insert logo (smaller + vertically centered inside banner)
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

    # -------- Left meta block rows 4..8 (Excel) --------
    meta_rows = [
        ("ODC #", payload.odc_number),
        ("FECHA:", payload.date_str),
        ("PROVEEDOR:", payload.provider),
        ("SERVICIO:", payload.service),
        ("PROYECTO:", payload.project),
    ]

    for i, (lab, val) in enumerate(meta_rows):
        rr = 4 + i
        ws.set_row(rr - 1, 20)

        # Blanco/Gris/Blanco alternado empezando en BLANCO
        is_gray = (i % 2 == 1)
        fill_range(ws, rr, 2, rr, 14, gray_fill if is_gray else white_bg)

        # label B..E (0-based 1..4)
        ws.merge_range(rr - 1, 1, rr - 1, 4, lab, label_gray if is_gray else label_white)

        # ✅ value F..N (0-based 5..13) with LEFT BORDER (divider)
        ws.merge_range(rr - 1, 5, rr - 1, 13, val, value_gray_div if is_gray else value_white_div)

    # -------- Bill-to block right side rows 4..8 and cols O..Z --------
    fill_range(ws, 4, 15, 8, 26, white_bg)
    ws.merge_range(3, 14, 3, 25, payload.bill_to_title, bill_title_fmt)  # row 4
    ws.merge_range(4, 14, 4, 25, payload.bill_to_name, bill_bold_fmt)    # row 5
    ws.merge_range(5, 14, 5, 25, f"RFC: {payload.bill_to_rfc}", bill_bold_fmt)
    ws.merge_range(6, 14, 6, 25, payload.bill_to_address_1, bill_norm_fmt)
    ws.merge_range(7, 14, 7, 25, payload.bill_to_address_2, bill_norm_fmt)

    # Spacer row 9
    ws.set_row(8, 10)

    # -------- Table header row 10 (Excel) --------
    header_row = 10
    ws.set_row(header_row - 1, 26)

    ws.merge_range(header_row - 1, 1, header_row - 1, 13, "Concepto", th_fmt)
    ws.merge_range(header_row - 1, 14, header_row - 1, 18, "Costo unitario", th_fmt)
    ws.merge_range(header_row - 1, 19, header_row - 1, 21, "Unidades", th_fmt)
    ws.merge_range(header_row - 1, 22, header_row - 1, 25, "Subtotal", th_fmt)

    # -------- Items start row 11 (Excel) --------
    start_items = 11
    items = payload.items or [ODCItem(concept="", unit_cost=0, units=0)]
    max_items = min(len(items), 18)
    wrap_chars = 58
    min_row_h = 26

    last_item_row = start_items - 1
    computed_sum = 0.0

    for idx in range(max_items):
        rr = start_items + idx
        it = items[idx]

        zebra = (idx % 2 == 1)
        row_fill = gray_fill if zebra else white_bg
        fill_range(ws, rr, 2, rr, 26, row_fill)

        unit_cost = safe_float(it.unit_cost)
        units = safe_float(it.units)
        subtotal = safe_float(it.subtotal) if it.subtotal is not None else (unit_cost * units)

        computed_sum += subtotal

        needed = row_height_for_wrapped_text(it.concept, wrap_chars, base_line_height=11.0, extra_lines=0.7)
        ws.set_row(rr - 1, int(max(min_row_h, math.ceil(needed))))

        ws.merge_range(rr - 1, 1, rr - 1, 13, it.concept, concept_g if zebra else concept_w)
        ws.merge_range(rr - 1, 14, rr - 1, 18, unit_cost, money_g if zebra else money_w)
        ws.merge_range(rr - 1, 19, rr - 1, 21, units, units_g if zebra else units_w)
        ws.merge_range(rr - 1, 22, rr - 1, 25, subtotal, money_g if zebra else money_w)

        last_item_row = rr

    # -------- Summary block (SUMA / ANTICIPO / TOTAL) --------
    sum_amount = safe_float(payload.sum_amount) if payload.sum_amount is not None else computed_sum
    advance_amount = safe_float(payload.advance_amount) if payload.advance_amount is not None else 0.0
    total_due = safe_float(payload.total_due) if payload.total_due is not None else (sum_amount - advance_amount)

    summary_start = last_item_row + 2
    ws.set_row(summary_start - 1, 10)

    sum_row = summary_start + 1
    adv_row = summary_start + 2
    tot_row = summary_start + 3

    for r in [sum_row, adv_row, tot_row]:
        ws.set_row(r - 1, 22)
        fill_range(ws, r, 2, r, 26, white_bg)

    # Underlines across T..Z
    fill_range(ws, sum_row, 20, sum_row, 26, hline_fmt)  # T..Z
    fill_range(ws, adv_row, 20, adv_row, 26, hline_fmt)  # T..Z

    ws.merge_range(sum_row - 1, 19, sum_row - 1, 21, "SUMA:", sum_label_fmt)
    ws.merge_range(adv_row - 1, 19, adv_row - 1, 21, "ANTICIPO:", adv_label_fmt)
    ws.merge_range(tot_row - 1, 19, tot_row - 1, 21, "TOTAL:", total_label_fmt)

    ws.merge_range(sum_row - 1, 22, sum_row - 1, 25, sum_amount, sum_value_fmt)
    ws.merge_range(adv_row - 1, 22, adv_row - 1, 25, advance_amount, adv_value_fmt)
    ws.merge_range(tot_row - 1, 22, tot_row - 1, 25, total_due, total_value_fmt)

    # ✅ Terms on Page 2 (start after some spacing)
    terms_start_row = tot_row + 6
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

    # ✅ Extend print area to include page 2
    ws.print_area(range_a1(1, 1, terms_end_row + 2, 26))  # A1:Z...

    wb.close()
    return out.getvalue()
