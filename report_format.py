#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
The shared design system for the feedback reports.

Every report in this repo is laid out from the master sheet ``summative`` in
``Summary AiFormative_and_Summative_Feedback_Report 2025-2026.xlsx``.  This
module reads that design straight out of the workbook (fonts, sizes, border
style/colour, row heights, column-width rule, logo image and logo offsets) and
re-applies it to any other sheet.

Nothing here touches the *content* of a sheet - callers pass the rows in and
they are written back verbatim.
"""

import os
import re
import shutil
import tempfile
import zipfile
from html import unescape

import openpyxl
from PIL import ImageFont
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import (
    OneCellAnchor, AnchorMarker, XDRPositiveSize2D)
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MASTER_SHEET = "summative"
COLLEGE = "KIT's College of Engineering,Kolhapur"
DEPARTMENT = "Department of Electronics and Telecommunication Engineering"

# logo placement, copied from the master sheet (EMU)
LOGO_FROM_COL, LOGO_FROM_ROW = 0, 0
LOGO_COL_OFF, LOGO_ROW_OFF = 130630, 64355
LOGO_EXT_CX, LOGO_EXT_CY = 1415142, 577902

# Calibri-vs-DejaVu advance-width ratio and padding, calibrated against the
# master sheet's own column widths (reproduces 5 of its 7 widths to within
# ~0.1 of a character, and lands ~2 wide on the other two).
GLYPH_FACTOR = 0.88
CELL_PADDING = 5
PX_PER_CHAR = 7


# --------------------------------------------------------------------------- #
# reading the design out of the master
# --------------------------------------------------------------------------- #
def copy_font(font):
    """Rebuild a real Font (cell.font is a read-only StyleProxy)."""
    if font.color is None:
        color = None
    elif font.color.type == "theme":
        color = Color(theme=font.color.theme, tint=font.color.tint)
    elif font.color.type == "indexed":
        color = Color(indexed=font.color.indexed)
    else:
        color = Color(rgb=font.color.rgb)
    return Font(name=font.name, size=font.size, bold=font.bold, italic=font.italic,
                underline=font.underline, strike=font.strike, color=color)


def copy_side(side):
    """Rebuild a real Side (cell.border.left is a read-only StyleProxy)."""
    color = None
    if side.color is not None:
        if side.color.type == "indexed":
            color = Color(indexed=side.color.indexed)
        elif side.color.type == "theme":
            color = Color(theme=side.color.theme, tint=side.color.tint)
        else:
            color = Color(rgb=side.color.rgb)
    return Side(style=side.style, color=color)


def read_design(master_wb, sheet=MASTER_SHEET):
    """Pull the whole design system out of the master sheet."""
    ws = master_wb[sheet]
    return {
        "title_font": copy_font(ws["A1"].font),
        "header_font": copy_font(ws["A4"].font),
        "data_font": copy_font(ws["A5"].font),
        "thin_side": copy_side(ws["A5"].border.left),
        "title_row_height": ws.row_dimensions[1].height,
        "data_row_height": ws.row_dimensions[4].height,
        "text_align": (ws["A5"].alignment.horizontal, ws["A5"].alignment.vertical),
        "num_align": (ws["E5"].alignment.horizontal, ws["E5"].alignment.vertical),
        "zoom": ws.sheet_view.zoomScale,
    }


# --------------------------------------------------------------------------- #
# the logo
# --------------------------------------------------------------------------- #
_TMP_DIR = None


def extract_logo(xlsx_path):
    """Grab the first picture out of an xlsx package."""
    with zipfile.ZipFile(xlsx_path) as z:
        for name in sorted(z.namelist()):
            if name.startswith("xl/media/") and name.lower().endswith((".png", ".jpg", ".jpeg")):
                return z.read(name)
    return None


def _logo_path(png_bytes):
    global _TMP_DIR
    if _TMP_DIR is None:
        _TMP_DIR = tempfile.mkdtemp(prefix="xlsxlogo_")
    path = os.path.join(_TMP_DIR, "logo.png" if png_bytes[:8] == b"\x89PNG\r\n\x1a\n" else "logo.jpg")
    with open(path, "wb") as fh:
        fh.write(png_bytes)
    return path


def reseat_images(wb, png_bytes):
    """
    openpyxl hands loaded images a BytesIO that Pillow closes on GC; point every
    existing picture at a real file so it survives the save.
    """
    if png_bytes is None:
        return
    path = _logo_path(png_bytes)
    for ws in wb.worksheets:
        for img in ws._images:
            img.ref = path


def make_logo(png_bytes):
    """A copy of the master logo, anchored and sized exactly like the master's."""
    img = Image(_logo_path(png_bytes))
    anchor = OneCellAnchor()
    anchor._from = AnchorMarker(col=LOGO_FROM_COL, colOff=LOGO_COL_OFF,
                                row=LOGO_FROM_ROW, rowOff=LOGO_ROW_OFF)
    anchor.ext = XDRPositiveSize2D(LOGO_EXT_CX, LOGO_EXT_CY)
    img.anchor = anchor
    return img


def cleanup():
    global _TMP_DIR
    if _TMP_DIR:
        shutil.rmtree(_TMP_DIR, ignore_errors=True)
        _TMP_DIR = None


# --------------------------------------------------------------------------- #
# column widths
# --------------------------------------------------------------------------- #
_fonts = {}


def metric_font(size_pt, bold=False):
    key = (size_pt, bold)
    if key not in _fonts:
        path = ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
                else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
        _fonts[key] = ImageFont.truetype(path, int(round(size_pt * 96 / 72)))
    return _fonts[key]


def text_px(text, size_pt=12, bold=False):
    return metric_font(size_pt, bold).getlength(str(text))


def fit_widths(rows, ncols, minimum=10, maximum=80):
    """
    The designer's rule, read off the master sheet: a column is exactly as wide
    as its longest entry (header in bold included), plus Excel's cell padding.
    """
    widths = []
    for c in range(ncols):
        longest = 0.0
        for i, row in enumerate(rows):
            if c < len(row) and row[c] is not None:
                longest = max(longest, text_px(row[c], 12, bold=(i == 0)))
        w = (longest * GLYPH_FACTOR + CELL_PADDING) / PX_PER_CHAR
        widths.append(round(min(max(w, minimum), maximum), 2))
    return widths


def apply_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# --------------------------------------------------------------------------- #
# building a sheet
# --------------------------------------------------------------------------- #
def thin(side):
    return Border(left=side, right=side, top=side, bottom=side)


def border_for(side, left=False, right=False, top=False, bottom=False):
    return Border(left=side if left else Side(),
                  right=side if right else Side(),
                  top=side if top else Side(),
                  bottom=side if bottom else Side())


def fresh_sheet(wb, old_ws):
    """
    Swap a worksheet for a brand-new empty one with the same name, in the same
    tab position.  Rebuilding in place leaves blank styled cells past the table
    and inflates the sheet's used range.
    """
    index = wb.worksheets.index(old_ws)
    title = old_ws.title
    wb.remove(old_ws)
    return wb.create_sheet(title, index)


def build_sheet(ws, design, title, headers, data, numeric_cols=(), logo=None,
                extra_cells=(), keep_fills=None):
    """
    Lay a sheet out with the master's design system.

    design        tokens from ``read_design``
    title         row-3 title (the report's own AY_..._... label)
    headers       list of column headings
    data          list of rows - written back verbatim
    numeric_cols  0-based column indexes that hold numbers
    logo          a copy of the master logo
    extra_cells   (row_index_in_data, col_index_0based, value, is_number)
                  cells outside the table's column span (e.g. stray formulas)
    keep_fills    {(row_index_in_data, col_index_0based): "RRGGBB"} to retain
    """
    ncols = len(headers)
    nrows = len(data)
    side = design["thin_side"]

    # ---------------------------- title block (rows 1-3) -------------------
    for row, text in ((1, COLLEGE), (2, DEPARTMENT), (3, title)):
        ws.cell(row=row, column=1, value=text)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        ws.row_dimensions[row].height = design["title_row_height"]

    for col in range(1, ncols + 1):
        for row in (1, 2, 3):
            cell = ws.cell(row=row, column=col)
            cell.border = border_for(side,
                                     left=(col == 1), right=(col == ncols),
                                     top=(row == 1), bottom=(row == 3))
            cell.font = design["title_font"]
            if col == 1:
                cell.alignment = Alignment(horizontal="center")

    # ------------------------------ header row -----------------------------
    hrow = 4
    ws.row_dimensions[hrow].height = design["data_row_height"]
    for c, text in enumerate(headers, start=1):
        cell = ws.cell(row=hrow, column=c, value=text)
        cell.font = design["header_font"]
        cell.border = thin(side)

    # ------------------------------- data rows -----------------------------
    for r, values in enumerate(data):
        excel_row = hrow + 1 + r
        ws.row_dimensions[excel_row].height = design["data_row_height"]
        for c, value in enumerate(values):
            cell = ws.cell(row=excel_row, column=c + 1)
            if value is not None:
                cell.value = value
            cell.font = design["data_font"]
            cell.border = thin(side)
            if c in numeric_cols:
                cell.alignment = Alignment(horizontal=design["num_align"][0],
                                           vertical=design["num_align"][1])
            else:
                cell.alignment = Alignment(horizontal=design["text_align"][0],
                                           vertical=design["text_align"][1])
            if keep_fills and (r, c) in keep_fills:
                cell.fill = PatternFill(fill_type="solid", fgColor=keep_fills[(r, c)])

    # ---------------------------- cells outside the table ------------------
    for (r, c, value, is_num) in extra_cells:
        cell = ws.cell(row=hrow + 1 + r, column=c + 1)
        cell.value = value
        cell.font = design["data_font"]
        cell.border = thin(side)
        cell.alignment = Alignment(horizontal=design["num_align"][0] if is_num
                                   else design["text_align"][0],
                                   vertical=design["num_align"][1])

    # ------------------------------- dimensions ----------------------------
    rows_for_width = [headers] + [[str(v) if v is not None else "" for v in row] for row in data]
    apply_widths(ws, fit_widths(rows_for_width, ncols))

    if logo is not None:
        ws.add_image(logo)

    ws.sheet_view.zoomScale = design["zoom"]
    ws.sheet_view.showGridLines = True
    ws.freeze_panes = None
    return ws


# --------------------------------------------------------------------------- #
# package tidy-up
# --------------------------------------------------------------------------- #
def find_sheet_part(zf, sheet_name):
    """Resolve a sheet name -> its worksheet part inside an xlsx package."""
    wbxml = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'Id="([^"]+)"', tag)
        tgt = re.search(r'Target="([^"]+)"', tag)
        typ = re.search(r'Type="([^"]+)"', tag)
        if rid and tgt and typ and typ.group(1).endswith("/worksheet"):
            targets[rid.group(1)] = tgt.group(1)
    for name, rid in re.findall(r'<sheet name="([^"]*)"[^>]*r:id="([^"]+)"', wbxml):
        if name == sheet_name:
            t = targets[rid].lstrip("/")
            return t if t.startswith("xl/") else "xl/" + t
    raise RuntimeError("sheet %r not found" % sheet_name)


def tidy_package(out_path, src_path=None, preserve_dimension_of=MASTER_SHEET):
    """
    openpyxl stamps a stray ``style="N"`` on every <col> element when it
    round-trips a workbook; the attribute points into cellStyleXfs and can
    trigger Excel's repair prompt, so strip it.  When a master sheet is carried
    through untouched, put its original <dimension> hint back as well.
    """
    src_dim = None
    out_part = None
    if src_path and preserve_dimension_of:
        try:
            with zipfile.ZipFile(src_path) as zsrc:
                p = find_sheet_part(zsrc, preserve_dimension_of)
                m = re.search(r'<dimension ref="([^"]+)" ?/>', zsrc.read(p).decode("utf-8"))
                src_dim = m.group(1) if m else None
            with zipfile.ZipFile(out_path) as zout:
                out_part = find_sheet_part(zout, preserve_dimension_of)
        except (RuntimeError, KeyError):
            src_dim = out_part = None

    tmp = out_path + ".tmp"
    n = 0
    with zipfile.ZipFile(out_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                n += len(re.findall(r'\s+style="\d+"', text))
                text = re.sub(r"<col\b[^>]*/>", lambda m: re.sub(r'\s+style="\d+"', "", m.group(0)), text)
                if item.filename == out_part and src_dim:
                    text = re.sub(r'<dimension ref="[^"]+" ?/>',
                                  '<dimension ref="%s"/>' % src_dim, text)
                data = text.encode("utf-8")
            zo.writestr(item, data)
    shutil.move(tmp, out_path)
    return n


# --------------------------------------------------------------------------- #
# helpers used by the per-workbook scripts
# --------------------------------------------------------------------------- #
HONORIFICS = {"DR", "MR", "MRS", "MS", "PROF", "SHRI", "SMT", "ER"}


def _name_tokens(name):
    """
    Uppercase word list with honorifics dropped.

    Punctuation becomes whitespace first, so ``Dr.A.M.Pol`` -> [A, M, POL] and
    lands on the same key as ``Dr. Ajay Pol``.
    """
    s = re.sub(r"[^A-Za-z]+", " ", str(name)).upper()
    return [t for t in s.split() if t and t not in HONORIFICS]


def squash(name):
    """Uppercase key for a person's name: honorifics and punctuation removed."""
    return "".join(_name_tokens(name))


def person_key(name):
    """(surname, first-initial) - groups 'Dr.M.D.Sontakke' with 'MANDAR SONTAKKE'."""
    toks = _name_tokens(name)
    if not toks:
        return None
    return (toks[-1], toks[0][0])


def shift_formula(formula, from_col_index, by):
    """Move column references at or after ``from_col_index`` (1-based) by ``by``."""
    def repl(m):
        letters, digits = m.group(1), m.group(2)
        idx = 0
        for ch in letters:
            idx = idx * 26 + (ord(ch) - 64)
        if idx >= from_col_index:
            idx += by
            out = ""
            while idx:
                idx, rem = divmod(idx - 1, 26)
                out = chr(65 + rem) + out
            letters = out
        return letters + digits
    return re.sub(r"\$?([A-Z]+)\$?(\d+)", repl, formula)
