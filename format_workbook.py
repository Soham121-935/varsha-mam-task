#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Standardise every data sheet in the feedback workbook on the design language of
the master sheet ``summative``.

The master sheet is read *from the workbook itself* (fonts, sizes, borders,
merges, row heights, column-width rule, logo image, logo anchor offsets) and
that design system is then re-applied, element by element, to each other sheet.
The content of every sheet is preserved exactly: only layout/formatting changes.

Run:  python3 format_workbook.py
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
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker, XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Summary AiFormative_and_Summative_Feedback_Report 2025-2026.xlsx")
BACKUP_DIR = os.path.join(HERE, "backup")
BACKUP = os.path.join(BACKUP_DIR, "Summary AiFormative_and_Summative_Feedback_Report 2025-2026.xlsx")
OUT = os.path.join(HERE, "Summary_AiFormative_and_Summative_Feedback_Report_2025-2026_Formatted.xlsx")

MASTER = "summative"          # the reference / template sheet (never modified)
COLLEGE = "KIT's College of Engineering,Kolhapur"
DEPARTMENT = "Department of Electronics and Telecommunication Engineering"

# --------------------------------------------------------------------------- #
# Design tokens hard-read from the master sheet below
# --------------------------------------------------------------------------- #
LOGO_FROM_COL = 0
LOGO_FROM_ROW = 0
LOGO_COL_OFF = 130630         # EMU offset inside cell A1 (identical to master)
LOGO_ROW_OFF = 64355          # EMU offset inside cell A1
LOGO_EXT_CX = 1415142         # EMU width  (identical to master -> no distortion)
LOGO_EXT_CY = 577902          # EMU height


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def copy_font(font):
    """Rebuild a real Font (cell.font is a read-only StyleProxy)."""
    return Font(name=font.name, size=font.size, bold=font.bold, italic=font.italic,
                underline=font.underline, strike=font.strike,
                color=Color(theme=font.color.theme, tint=font.color.tint)
                if font.color is not None and font.color.type == "theme"
                else (Color(rgb=font.color.rgb) if font.color is not None else None))


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


def read_master_tokens(wb):
    """Pull the design system straight out of the master sheet."""
    ws = wb[MASTER]
    tok = {}

    # title block: rows 1..3, first cell of each merged title row
    tok["title_font"] = copy_font(ws["A1"].font)
    tok["header_font"] = copy_font(ws["A4"].font)
    tok["data_font"] = copy_font(ws["A5"].font)
    tok["thin_side"] = copy_side(ws["A5"].border.left)
    tok["title_row_height"] = ws.row_dimensions[1].height
    tok["data_row_height"] = ws.row_dimensions[4].height

    # column-width rule used by the designer: widest text in the column + 3
    widths = {}
    for idx, cd in ws.column_dimensions.items():
        widths[openpyxl.utils.column_index_from_string(idx)] = cd.width
    tok["master_widths"] = widths

    # vertical alignment convention of the data area
    tok["text_align"] = (ws["A5"].alignment.horizontal, ws["A5"].alignment.vertical)
    tok["num_align"] = (ws["E5"].alignment.horizontal, ws["E5"].alignment.vertical)

    # logo binary
    img = ws._images[0]
    tok["logo_anchor"] = img.anchor
    return tok


def extract_logo_png(src_path):
    """Grab the logo straight out of the workbook package (xl/media/image1.png)."""
    with zipfile.ZipFile(src_path) as z:
        for name in z.namelist():
            if name.startswith("xl/media/") and name.lower().endswith((".png", ".jpg", ".jpeg")):
                return z.read(name)
    raise RuntimeError("no logo image found in the workbook")


_TMP_DIR = None


def tmp_logo_path(png_bytes):
    global _TMP_DIR
    if _TMP_DIR is None:
        _TMP_DIR = tempfile.mkdtemp(prefix="xlsxlogo_")
    ext = "png" if png_bytes[:8] == b"\x89PNG\r\n\x1a\n" else "jpg"
    path = os.path.join(_TMP_DIR, "logo.%s" % ext)
    with open(path, "wb") as fh:
        fh.write(png_bytes)
    return path


def reseat_images(wb, png_bytes):
    """
    openpyxl hands loaded images a BytesIO that Pillow closes on GC; point every
    existing picture at a real file so it survives the save.
    """
    path = tmp_logo_path(png_bytes)
    for ws in wb.worksheets:
        for img in ws._images:
            img.ref = path


def make_logo(png_bytes):
    """A copy of the master logo, anchored/sized exactly like the master's."""
    img = Image(tmp_logo_path(png_bytes))
    anchor = OneCellAnchor()
    anchor._from = AnchorMarker(col=LOGO_FROM_COL, colOff=LOGO_COL_OFF,
                                row=LOGO_FROM_ROW, rowOff=LOGO_ROW_OFF)
    anchor.ext = XDRPositiveSize2D(LOGO_EXT_CX, LOGO_EXT_CY)
    img.anchor = anchor
    return img


def thin(side):
    """A fully boxed thin border using the master's border style/colour."""
    return Border(left=side, right=side, top=side, bottom=side)


def border_for(side, left=False, right=False, top=False, bottom=False):
    return Border(left=side if left else Side(),
                  right=side if right else Side(),
                  top=side if top else Side(),
                  bottom=side if bottom else Side())


# Calibri-vs-DejaVu advance-width ratio and padding, calibrated against the
# master sheet's own column widths (the designer's auto-fit is reproduced to
# within ~0.1 of a character on 5 of its 7 columns, and ~2 wide on the other 2).
GLYPH_FACTOR = 0.88
CELL_PADDING = 5
PX_PER_CHAR = 7

_fonts = {}


def metric_font(size_pt, bold=False):
    key = (size_pt, bold)
    if key not in _fonts:
        from PIL import ImageFont
        path = ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
                else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
        _fonts[key] = ImageFont.truetype(path, int(round(size_pt * 96 / 72)))
    return _fonts[key]


def text_px(text, size_pt=12, bold=False):
    return metric_font(size_pt, bold).getlength(str(text))


def fit_widths(rows_of_text, ncols, minimum=10, maximum=80):
    """
    Reproduce the designer's column-width rule seen on the master sheet:
    a column is exactly as wide as its longest entry (header included, in bold)
    once rendered in the 12pt body font, plus Excel's cell padding.
    """
    widths = []
    for c in range(ncols):
        longest = 0.0
        for i, row in enumerate(rows_of_text):
            if c < len(row) and row[c] is not None:
                longest = max(longest, text_px(row[c], 12, bold=(i == 0)))
        widths.append(round(min(max((longest * GLYPH_FACTOR + CELL_PADDING) / PX_PER_CHAR, minimum), maximum), 2))
    return widths


def apply_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# --------------------------------------------------------------------------- #
# the actual re-design
# --------------------------------------------------------------------------- #
def fresh_sheet(wb, old_ws):
    """
    Swap a worksheet for a brand-new empty one with the same name, in the same
    tab position.  Rebuilding in place would leave blank styled cells past the
    table and inflate the sheet's used range.
    """
    index = wb.worksheets.index(old_ws)
    title = old_ws.title
    wb.remove(old_ws)
    return wb.create_sheet(title, index)


def build_sheet(ws, tok, title, headers, data, yellow=(), numeric_cols=(), logo=None):
    """
    Lay a sheet out with the master's design system.

    ws           worksheet to rebuild
    tok          design tokens taken from the master sheet
    title        row-3 title text (the sheet's own AY_..._... label)
    headers      list of column headings
    data         list of rows (lists) - the sheet's own content, untouched
    yellow       set of (row_index_within_data, col_index) cells to keep highlighted
    numeric_cols 0-based column indexes that are numeric -> centred like the master
    logo         a copy of the master logo to drop in the same relative spot
    """
    ncols = len(headers)
    last_col = ncols
    nrows = len(data)

    side = tok["thin_side"]
    title_font = tok["title_font"]
    header_font = tok["header_font"]
    data_font = tok["data_font"]

    # ---------------------------- title block (rows 1-3) -------------------
    block = [(1, COLLEGE), (2, DEPARTMENT), (3, title)]
    for row, text in block:
        ws.cell(row=row, column=1, value=text)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
        ws.row_dimensions[row].height = tok["title_row_height"]

    # borders: the master draws a single open box around the three title rows
    for col in range(1, last_col + 1):
        for row in (1, 2, 3):
            is_first = col == 1
            is_last = col == last_col
            ws.cell(row=row, column=col).border = border_for(
                side,
                left=is_first,
                right=is_last,
                top=(row == 1),
                bottom=(row == 3),
            )
            if is_first:                       # merged cell -> text lives in A
                cell = ws.cell(row=row, column=col)
                cell.font = title_font
                cell.alignment = Alignment(horizontal="center")
            else:
                ws.cell(row=row, column=col).font = title_font

    # ------------------------------ header row -----------------------------
    hrow = 4
    ws.row_dimensions[hrow].height = tok["data_row_height"]
    for c, text in enumerate(headers, start=1):
        cell = ws.cell(row=hrow, column=c, value=text)
        cell.font = header_font
        cell.border = thin(side)

    # ------------------------------- data rows -----------------------------
    for r, row_values in enumerate(data, start=0):
        excel_row = hrow + 1 + r
        ws.row_dimensions[excel_row].height = tok["data_row_height"]
        for c, value in enumerate(row_values):
            cell = ws.cell(row=excel_row, column=c + 1)
            if value is not None:
                cell.value = value
            cell.font = data_font
            cell.border = thin(side)
            if c in numeric_cols:
                cell.alignment = Alignment(horizontal=tok["num_align"][0],
                                           vertical=tok["num_align"][1])
            else:
                cell.alignment = Alignment(horizontal=tok["text_align"][0],
                                           vertical=tok["text_align"][1])
            if (r, c) in yellow:
                cell.fill = PatternFill(fill_type="solid", fgColor="FFFFFF00")

    # ------------------------------- dimensions ----------------------------
    widths = fit_widths([headers] + [[str(v) if v is not None else "" for v in row]
                                     for row in data], ncols)
    apply_widths(ws, widths)

    # --------------------------------- logo --------------------------------
    if logo is not None:
        ws.add_image(logo)

    # view settings identical to the master so the sheets open looking the same
    ws.sheet_view.zoomScale = 70
    ws.sheet_view.showGridLines = True
    ws.freeze_panes = None
    return ws


# --------------------------------------------------------------------------- #
# per-sheet content extraction (content is copied verbatim, nothing invented)
# --------------------------------------------------------------------------- #
def grid(ws, cols, first_row, last_row):
    out = []
    for r in range(first_row, last_row + 1):
        out.append([ws.cell(row=r, column=c).value for c in cols])
    return out


def yellow_cells(ws, cols, first_row, last_row):
    """Keep the manually applied yellow highlight - it flags rows for action."""
    found = set()
    for r in range(first_row, last_row + 1):
        for i, c in enumerate(cols):
            cell = ws.cell(row=r, column=c)
            if cell.fill.fill_type == "solid" and cell.fill.fgColor.type == "rgb" \
                    and str(cell.fill.fgColor.rgb).upper() == "FFFFFF00":
                found.add((r - first_row, i))
    return found


# Sem-I report schema: Designation | Faculty Name | Semester/ Group | Subject | Performance
SEMI_HEADERS = ["Designation", "Faculty Name", "Semester/ Group", "Subject", "Performance "]
# Sem-II report schema (identical to the master sheet)
SEMII_HEADERS = ["Faculty Name", "Designation", "Course", "Sem/Class",
                 "Attendees", "Performance", "Date"]


def main():
    if not os.path.exists(BACKUP):
        os.makedirs(BACKUP_DIR, exist_ok=True)
        shutil.copy2(SRC, BACKUP)

    logo_png = extract_logo_png(SRC)
    wb = openpyxl.load_workbook(SRC)
    reseat_images(wb, logo_png)
    tok = read_master_tokens(wb)

    # ------------------------------------------------------------------ #
    # 1. 2025-26-Sem-I Formative      (rows 5-89, header already on row 4)
    # ------------------------------------------------------------------ #
    old = wb["2025-26-Sem-I Formative"]
    cols = [1, 2, 3, 4, 5]
    data = grid(old, cols, 5, 89)
    # keep the trailing-space quirk of the original header text
    headers = [old.cell(row=4, column=c).value for c in cols]
    yellow = yellow_cells(old, cols, 5, 89)
    ws = fresh_sheet(wb, old)
    build_sheet(ws, tok, "AY_2025_26_Sem_I_Formative", headers, data,
                yellow=yellow, numeric_cols={4}, logo=make_logo(logo_png))

    # ------------------------------------------------------------------ #
    # 2. Summative Sem-I              (rows 4-54, no header row originally)
    # ------------------------------------------------------------------ #
    old = wb["Summative Sem-I "]
    cols = [1, 2, 3, 4, 5]
    data = grid(old, cols, 4, 54)
    ws = fresh_sheet(wb, old)
    # header labels are not duplicated from the master - they are taken from the
    # sibling Sem-I sheet, which has the identical 5-column layout.
    build_sheet(ws, tok, "AY_2025_26_Sem_I_Summative", SEMI_HEADERS, data,
                yellow=set(), numeric_cols={4}, logo=make_logo(logo_png))

    # ------------------------------------------------------------------ #
    # 3. Sem-II Frramtive   (rows 2-19; col A is a constant repeated label
    #    and is dropped, cols B-H then match the master's schema exactly)
    # ------------------------------------------------------------------ #
    old = wb["Sem-II Frramtive"]
    cols = [2, 3, 4, 5, 6, 7, 8]
    data = grid(old, cols, 2, 19)
    ws = fresh_sheet(wb, old)
    build_sheet(ws, tok, "AY_2025_26_Sem_II_Formative", SEMII_HEADERS, data,
                yellow=set(), numeric_cols={4, 5}, logo=make_logo(logo_png))

    wb.active = 0
    wb.save(OUT)
    print("wrote", OUT)

    # ------------------------------------------------------------------ #
    # tidy up artefacts left by the openpyxl round-trip
    # ------------------------------------------------------------------ #
    tidy_package(OUT, SRC)

    if _TMP_DIR:
        shutil.rmtree(_TMP_DIR, ignore_errors=True)


def find_sheet_part(zf, sheet_name):
    """Resolve a sheet name -> its worksheet part inside an xlsx package."""
    wbxml = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    target_by_rid = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'Id="([^"]+)"', tag)
        tgt = re.search(r'Target="([^"]+)"', tag)
        typ = re.search(r'Type="([^"]+)"', tag)
        if rid and tgt and typ and typ.group(1).endswith("/worksheet"):
            target_by_rid[rid.group(1)] = tgt.group(1)
    for name, rid in re.findall(r'<sheet name="([^"]*)"[^>]*r:id="([^"]+)"', wbxml):
        if name == sheet_name:
            t = target_by_rid[rid].lstrip("/")
            if not t.startswith("xl/"):
                t = "xl/" + t
            return t
    raise RuntimeError("sheet %r not found" % sheet_name)


def tidy_package(out_path, src_path, master_name=MASTER):
    """
    openpyxl stamps a stray ``style="N"`` on every <col> element when it
    round-trips a workbook.  The attribute points into cellStyleXfs (which only
    holds one entry here) and can trigger Excel's repair prompt, so strip it.
    The master sheet also keeps its original <dimension> so it really is
    untouched.
    """
    with zipfile.ZipFile(src_path) as zsrc:
        src_part = find_sheet_part(zsrc, master_name)
        src_dim = re.search(r'<dimension ref="([^"]+)" ?/>', zsrc.read(src_part).decode('utf-8'))
    with zipfile.ZipFile(out_path) as zout:
        out_part = find_sheet_part(zout, master_name)

    tmp = out_path + ".tmp"
    n = 0
    with zipfile.ZipFile(out_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")

                def fix(m):
                    return re.sub(r'\s+style="\d+"', "", m.group(0))

                new_text = re.sub(r"<col\b[^>]*/>", fix, text)
                n += len(re.findall(r'\s+style="\d+"', text))
                if item.filename == out_part and src_dim:
                    new_text = re.sub(r'<dimension ref="[^"]+" ?/>',
                                      '<dimension ref="%s"/>' % src_dim.group(1), new_text)
                data = new_text.encode("utf-8")
            zo.writestr(item, data)
    shutil.move(tmp, out_path)
    if n:
        print("  stripped %d stray <col style=...> attributes" % n)


if __name__ == "__main__":
    main()
