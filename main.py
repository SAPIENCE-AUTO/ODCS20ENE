# main.py
# Sapience ODCs — Render + FastAPI — Excel only (XlsxWriter)
# v2.0.0 — New layout (Provider block + date in banner + bill-to compressed + anticipo as line item)
#
# Changes:
# - Banner: date on top-right (white text)
# - Provider block: "DATOS DEL PROVEEDOR" + NOMBRE/RFC/E-MAIL/SERVICIO/PROYECTO
# - Bill-to block: compressed into 1 merged cell with line breaks
# - Summary: removes SUMA/ANTICIPO, keeps only TOTAL (items already include anticipo negative line)
# - Keeps terms page (Page 2) as single column

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

app = FastAPI(title="Sapience ODCs (Excel)", version="2.0.0")


# -----------------------------
# Models
# -----------------------------
class ODCItem(BaseModel):
    concept: str = ""
    unit_cost: float = 0
    units: float = 0
    subtotal: Optional[float] = None  # if omitted, we compute = unit_cost * units


class ODCPayload(BaseModel):
    odc_number: str
    date_str: str  # shown in banner top-right (e.g. "17 - nov - 2025")

    # Provider block (NEW)
    provider_name: Optional[str] = None
    provider_rfc: Optional[str] = None
    provider_email: Optional[str] = None

    # Backward-compat (old)
    provider: Optional[str] = None

    # Service / Project
    service: str
    project: str

    # Bill-to
    bill_to_title: str = "FACTURAR A:"
    bill_to_name: str
    bill_to_rfc: str
    bill_to_address_1: str
    bill_to_address_2: str

    # Items (includes anticipo as negative line item if applicable)
    items: List[ODCItem] = Field(default_factory=list)

    # Total override (optional)
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
    fill_range(ws, 1, 1, 520, 50, white_bg)

    banner_fill = wb.add_format({"bg_color": TEAL})
    gray_fill = wb.add_format({"bg_color": LIGHT_GRAY})

    # -------- Banner formats --------
    date_banner_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10,
        "align": "right", "valign": "top",
        "font_color": WHITE, "bg_color": TEAL,
    })

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

    # -------- Provider section title --------
    provider_title_fmt = wb.add_format({
        "font_name": FONT, "font_size": 16, "bold": True,
        "align": "left", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
    })

    # Provider labels/values (striped)
    label_gray = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": LIGHT_GRAY
    })
    label_white = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    value_gray_div = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": LIGHT_GRAY,
        "text_wrap": True,
        "left": 1, "left_color": GRID_LIGHT,
    })
    value_white_div = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE,
        "text_wrap": True,
        "left": 1, "left_color": GRID_LIGHT,
    })
    email_value_gray_div = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "align": "left", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": LIGHT_GRAY,
        "underline": 1,
        "text_wrap": True,
        "left": 1, "left_color": GRID_LIGHT,
    })
    email_value_white_div = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "align": "left", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
        "underline": 1,
        "text_wrap": True,
        "left": 1, "left_color": GRID_LIGHT,
    })

    # -------- Bill-to block (single merged cell) --------
    bill_block_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10,
        "align": "left", "valign": "top",
        "font_color": BLACK, "bg_color": WHITE,
        "text_wrap": True
    })
    bill_title_bold = wb.add_format({
        "font_name": FONT, "font_size": 14, "bold": True,
        "font_color": TEAL_2,
    })
    bill_name_bold = wb.add_format({
        "font_name": FONT, "font_size": 12, "bold": True,
        "font_color": BLACK,
    })
    bill_rfc_label = wb.add_format({
        "font_name": FONT, "font_size": 12, "bold": True,
        "font_color": TEAL_2,
    })
    bill_rfc_val = wb.add_format({
        "font_name": FONT, "font_size": 12, "bold": True,
        "font_color": BLACK,
    })
    bill_addr = wb.add_format({
        "font_name": FONT, "font_size": 11,
        "font_color": BLACK,
    })

    # -------- Table header --------
    th_fmt = wb.add_format({
        "font_name": FONT, "font_size": 11, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": WHITE, "bg_color": TEAL_2,
        "border": 1, "border_color": GRID,
    })

    # -------- Table cells --------
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
    money_red_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "font_color": RED,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })
    money_red_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
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

    # -------- Total formats --------
    total_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 16, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    total_value_fmt = wb.add_format({
        "font_name": FONT, "font_size": 12, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })

    # -----------------------------
    # BANNER
    # -----------------------------
    ws.set_row(0, 16)
    ws.set_row(1, 16)
    ws.set_row(2, 16)
    fill_range(ws, 1, 2, 3, 26, banner_fill)  # B1:Z3

    # Date on banner (top-right but not overlapping ODC box)
    ws.merge_range(0, 14, 0, 18, payload.date_str, date_banner_fmt)  # O1:S1

    # ODC box top-right (rows 1..2)
    ws.merge_range(0, 19, 1, 22, "ODC #:", odc_box_lbl)            # T..W
    ws.merge_range(0, 23, 1, 25, payload.odc_number, odc_box_val)  # X..Z

    # Logo
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

    # -----------------------------
    # PROVIDER TITLE + PROVIDER ROWS (left)
    # -----------------------------
    # Row 4: section title
    ws.set_row(3, 28)  # Excel row 4
    ws.merge_range(3, 1, 3, 13, "DATOS DEL PROVEEDOR", provider_title_fmt)  # B..N

    # Resolve provider fields (fallbacks)
    provider_name = (payload.provider_name or payload.provider or "").strip()
    provider_rfc = (payload.provider_rfc or "").strip()
    provider_email = (payload.provider_email or "").strip()

    provider_rows = [
        ("NOMBRE:", provider_name),
        ("RFC:", provider_rfc),
        ("E-MAIL:", provider_email),
        ("SERVICIO:", payload.service),
        ("PROYECTO:", payload.project),
    ]

    # Rows 5..9 in Excel
    for i, (lab, val) in enumerate(provider_rows):
        rr = 5 + i
        ws.set_row(rr - 1, 22)

        # stripes start gray on first row? in your example: first looks gray
        is_gray = (i % 2 == 0)
        fill_range(ws, rr, 2, rr, 14, gray_fill if is_gray else white_bg)

        ws.merge_range(rr - 1, 1, rr - 1, 4, lab, label_gray if is_gray else label_white)

        # email in teal + underline
        if lab.startswith("E-MAIL"):
            ws.merge_range(rr - 1, 5, rr - 1, 13, val, email_value_gray_div if is_gray else email_value_white_div)
        else:
            ws.merge_range(rr - 1, 5, rr - 1, 13, val, value_gray_div if is_gray else value_white_div)

    # -----------------------------
    # BILL-TO BLOCK (right) — single merged cell with rich text
    # -----------------------------
    # Merge O4:Z9 (rows 4..9) i.e. Excel rows 4-9 => 0-based rows 3-8
    fill_range(ws, 4, 15, 9, 26, white_bg)
    ws.merge_range(3, 14, 8, 25, "", bill_block_fmt)

    # Write rich text on the top-left cell of that merged area (row=3 col=14)
    ws.write_rich_string(
        3, 14,
        bill_title_bold, payload.bill_to_title + "\n",
        bill_name_bold, payload.bill_to_name + "\n",
        bill_rfc_label, "RFC: ",
        bill_rfc_val, payload.bill_to_rfc + "\n",
        bill_addr, payload.bill_to_address_1 + "\n",
        bill_addr, payload.bill_to_address_2,
        bill_block_fmt
    )

    # Spacer row 10
    ws.set_row(9, 12)

    # -----------------------------
    # TABLE HEADER (row 11)
    # -----------------------------
    header_row = 11
    ws.set_row(header_row - 1, 30)

    ws.merge_range(header_row - 1, 1, header_row - 1, 13, "Concepto", th_fmt)
    ws.merge_range(header_row - 1, 14, header_row - 1, 18, "Costo unitario", th_fmt)
    ws.merge_range(header_row - 1, 19, header_row - 1, 21, "Unidades", th_fmt)
    ws.merge_range(header_row - 1, 22, header_row - 1, 25, "Subtotal", th_fmt)

    # -----------------------------
    # ITEMS (start row 12)
    # -----------------------------
    start_items = 12
    items = payload.items or [ODCItem(concept="", unit_cost=0, units=0)]
    max_items = min(len(items), 18)
    wrap_chars = 58
    min_row_h = 28

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

        needed = row_height_for_wrapped_text(it.concept, wrap_chars, base_line_height=12.0, extra_lines=0.7)
        ws.set_row(rr - 1, int(max(min_row_h, math.ceil(needed))))

        # negative values in red (like your example)
        money_fmt = (money_red_g if zebra else money_red_w) if subtotal < 0 or unit_cost < 0 else (money_g if zebra else money_w)

        ws.merge_range(rr - 1, 1, rr - 1, 13, it.concept, concept_g if zebra else concept_w)
        ws.merge_range(rr - 1, 14, rr - 1, 18, unit_cost, money_fmt)
        ws.merge_range(rr - 1, 19, rr - 1, 21, units, units_g if zebra else units_w)
        ws.merge_range(rr - 1, 22, rr - 1, 25, subtotal, money_fmt)

        last_item_row = rr

    # -----------------------------
    # TOTAL (only)
    # -----------------------------
    total_due = safe_float(payload.total_due) if payload.total_due is not None else computed_total

    # leave some air like the example
    tot_row = last_item_row + 3
    ws.set_row(tot_row - 1, 30)
    fill_range(ws, tot_row, 2, tot_row, 26, white_bg)

    ws.merge_range(tot_row - 1, 18, tot_row - 1, 21, "TOTAL:", total_label_fmt)     # S..V
    ws.merge_range(tot_row - 1, 22, tot_row - 1, 25, total_due, total_value_fmt)   # W..Z

    # -----------------------------
    # TERMS PAGE (Page 2)
    # -----------------------------
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

    # -----------------------------
    # PRINT SETTINGS
    # -----------------------------
    ws.set_portrait()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.fit_to_pages(1, 0)

    ws.print_area(range_a1(1, 1, terms_end_row + 2, 26))  # A1:Z...

    wb.close()
    return out.getvalue()
