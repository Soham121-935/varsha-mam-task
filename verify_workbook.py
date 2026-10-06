#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Quality-control checks on the formatted workbook.

  1. the master sheet is still identical to the original (values + every style)
  2. no content was lost or altered on any of the rebuilt sheets
  3. every rebuilt sheet uses the master's design system
  4. the logo is present, undistorted and identically placed on every sheet
"""
import os
import sys
import zipfile
import re

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Summary AiFormative_and_Summative_Feedback_Report 2025-2026.xlsx")
OUT = os.path.join(HERE, "Summary_AiFormative_and_Summative_Feedback_Report_2025-2026_Formatted.xlsx")
MASTER = "summative"

failures = []
notes = []


def check(cond, message):
    if cond:
        print("  ok   %s" % message)
    else:
        print("  FAIL %s" % message)
        failures.append(message)


def note(message):
    print("  note %s" % message)
    notes.append(message)


def cellkey(c):
    f, a, b, fl = c.font, c.alignment, c.border, c.fill
    col = lambda x: (x.type, x.theme, x.tint, x.rgb, x.indexed) if x is not None else None
    return (
        c.value,
        (f.name, f.size, bool(f.bold), bool(f.italic), f.underline, col(f.color)),
        (a.horizontal, a.vertical, a.wrap_text, a.indent),
        (b.left.style, col(b.left.color), b.right.style, col(b.right.color),
         b.top.style, col(b.top.color), b.bottom.style, col(b.bottom.color)),
        (fl.fill_type, col(fl.fgColor)),
        c.number_format,
    )


def main():
    src = openpyxl.load_workbook(SRC)
    out = openpyxl.load_workbook(OUT)

    print("\n[1] sheet inventory")
    check([ws.title for ws in src.worksheets] == [ws.title for ws in out.worksheets],
          "sheet names and order unchanged: %s" % [ws.title for ws in out.worksheets])

    # ------------------------------------------------------------------ #
    print("\n[2] master sheet untouched")
    a, b = src[MASTER], out[MASTER]

    def raw_dimension(path):
        with zipfile.ZipFile(path) as z:
            xml = z.read("xl/worksheets/sheet5.xml").decode("utf-8")
        return re.search(r'<dimension ref="([^"]+)" ?/>', xml).group(1)

    check(raw_dimension(SRC) == raw_dimension(OUT) == "A1:G54",
          "declared used range %s (unchanged)" % raw_dimension(OUT))
    check(sorted(str(r) for r in a.merged_cells.ranges) == sorted(str(r) for r in b.merged_cells.ranges),
          "merged ranges %s" % sorted(str(r) for r in b.merged_cells.ranges))
    check({k: v.width for k, v in a.column_dimensions.items()} ==
          {k: v.width for k, v in b.column_dimensions.items()}, "column widths")
    check({k: v.height for k, v in a.row_dimensions.items()} ==
          {k: v.height for k, v in b.row_dimensions.items()}, "row heights")

    diffs = []
    for r in range(1, max(a.max_row, b.max_row) + 1):
        for c in range(1, max(a.max_column, b.max_column) + 1):
            if cellkey(a.cell(r, c)) != cellkey(b.cell(r, c)):
                diffs.append((r, c, cellkey(a.cell(r, c)), cellkey(b.cell(r, c))))
    check(not diffs, "every cell keeps its value and full style (%d diffs)" % len(diffs))
    for d in diffs[:10]:
        print("        ", d)

    sa, sb = a.sheet_view, b.sheet_view
    check(sa.zoomScale == sb.zoomScale and sa.showGridLines == sb.showGridLines,
          "view settings (zoom=%s, gridlines=%s)" % (sb.zoomScale, sb.showGridLines))
    check((a.page_margins.left, a.page_margins.right, a.page_margins.top, a.page_margins.bottom) ==
          (b.page_margins.left, b.page_margins.right, b.page_margins.top, b.page_margins.bottom),
          "page margins")
    ia, ib = a._images, b._images
    check(len(ia) == len(ib) == 1, "logo still present")
    if ia and ib:
        check((ia[0].anchor._from.col, ia[0].anchor._from.row,
               ia[0].anchor._from.colOff, ia[0].anchor._from.rowOff) ==
              (ib[0].anchor._from.col, ib[0].anchor._from.row,
               ib[0].anchor._from.colOff, ib[0].anchor._from.rowOff),
              "logo anchor unchanged")
        check(ia[0].width == ib[0].width and ia[0].height == ib[0].height,
              "logo size unchanged (%.0fx%.0f px)" % (ib[0].width, ib[0].height))

    # ------------------------------------------------------------------ #
    print("\n[3] content preservation on the rebuilt sheets")
    expect = {
        # sheet: (expected rows of data, expected column count, ncols w/ numbers)
        "2025-26-Sem-I Formative": (85, 5),
        "Summative Sem-I ": (51, 5),
        "Sem-II Frramtive": (18, 7),
    }
    src_rows = {
        "2025-26-Sem-I Formative": [5, 89],
        "Summative Sem-I ": [4, 54],
        "Sem-II Frramtive": [2, 19],
    }
    src_cols = {
        "2025-26-Sem-I Formative": [1, 2, 3, 4, 5],
        "Summative Sem-I ": [1, 2, 3, 4, 5],
        "Sem-II Frramtive": [2, 3, 4, 5, 6, 7, 8],   # col A dropped on request
    }
    for name, (nrows, ncols) in expect.items():
        r0, r1 = src_rows[name]
        cols = src_cols[name]
        original = [[src[name].cell(row=r, column=c).value for c in cols]
                    for r in range(r0, r1 + 1)]
        ws = out[name]
        got = [[ws.cell(row=5 + i, column=1 + j).value for j in range(ncols)]
               for i in range(nrows)]
        check(len(original) == len(got) == nrows, "%s: %d data rows kept" % (name, len(got)))
        bad = [(i, j, original[i][j], got[i][j])
               for i in range(min(len(original), len(got)))
               for j in range(ncols) if original[i][j] != got[i][j]]
        check(not bad, "%s: every value identical to the original" % name)
        for x in bad[:5]:
            print("        ", x)
        # nothing left behind below the table
        below = [ws.cell(row=5 + nrows, column=j).value for j in range(1, ncols + 1)]
        check(all(v is None for v in below), "%s: no stray content under the table" % name)

    # yellow highlight kept
    y_src = set()
    ws = src["2025-26-Sem-I Formative"]
    for r in range(5, 90):
        for c in (2, 3, 4):
            cell = ws.cell(row=r, column=c)
            if cell.fill.fill_type == "solid" and str(cell.fill.fgColor.rgb).upper() == "FFFFFF00":
                y_src.add((r - 5, c - 2))
    y_out = set()
    ws = out["2025-26-Sem-I Formative"]
    for r in range(5, 90):
        for c in (2, 3, 4):
            cell = ws.cell(row=r, column=c)
            if cell.fill.fill_type == "solid" and str(cell.fill.fgColor.rgb).upper() == "FFFFFF00":
                y_out.add((r - 5, c - 2))
    check(y_src == y_out and len(y_src) == 72,
          "Sem-I Formative: all %d yellow-highlighted cells preserved" % len(y_src))

    # ------------------------------------------------------------------ #
    print("\n[4] design system of the master reproduced on every rebuilt sheet")
    m = out[MASTER]
    ref_font = (m["A1"].font.name, m["A1"].font.size, bool(m["A1"].font.bold))
    ref_hdr = (m["A4"].font.name, m["A4"].font.size, bool(m["A4"].font.bold))
    ref_data = (m["A5"].font.name, m["A5"].font.size, bool(m["A5"].font.bold))
    ref_side = m["A5"].border.left.style

    for name in expect:
        ws = out[name]
        ncols = expect[name][1]
        check(ws["A1"].value == m["A1"].value, "%s: college name in row 1" % name)
        check(ws["A2"].value == m["A2"].value, "%s: department in row 2" % name)
        check(sorted(str(r) for r in ws.merged_cells.ranges) ==
              ["A1:%s1" % chr(64 + ncols), "A2:%s2" % chr(64 + ncols), "A3:%s3" % chr(64 + ncols)],
              "%s: title rows merged across the table" % name)
        check((ws["A1"].font.name, ws["A1"].font.size, bool(ws["A1"].font.bold)) == ref_font,
              "%s: title font == master (%s %spt bold=%s)" % ((name,) + ref_font))
        check(ws.row_dimensions[1].height == m.row_dimensions[1].height and
              ws.row_dimensions[2].height == m.row_dimensions[2].height and
              ws.row_dimensions[3].height == m.row_dimensions[3].height,
              "%s: title row heights == master (%s)" % (name, m.row_dimensions[1].height))
        hdr = ws[4]
        check(all((c.font.name, c.font.size, bool(c.font.bold)) == ref_hdr for c in hdr[:ncols]),
              "%s: header row font == master (%s %spt bold=%s)" % ((name,) + ref_hdr))
        check(all(c.border.left.style == ref_side and c.border.right.style == ref_side and
                  c.border.top.style == ref_side and c.border.bottom.style == ref_side
                  for c in hdr[:ncols]) and all(c.value for c in hdr[:ncols]),
              "%s: header row fully boxed with the master's %s border" % (name, ref_side))
        nrows = expect[name][0]
        boxed = True
        for r in range(4, 5 + nrows):
            for c in range(1, ncols + 1):
                cell = ws.cell(row=r, column=c)
                if not all(getattr(cell.border, s).style == ref_side for s in
                           ("left", "right", "top", "bottom")):
                    boxed = False
        check(boxed, "%s: every table cell boxed with %s borders" % (name, ref_side))
        fonts_ok = True
        for r in range(5, 5 + nrows):
            for c in range(1, ncols + 1):
                cell = ws.cell(row=r, column=c)
                if (cell.font.name, cell.font.size, bool(cell.font.bold)) != ref_data:
                    fonts_ok = False
        check(fonts_ok, "%s: data font == master (%s %spt bold=%s)" % ((name,) + ref_data))
        heights_ok = all(ws.row_dimensions[r].height == m.row_dimensions[4].height
                         for r in range(4, 5 + nrows))
        check(heights_ok, "%s: table row heights == master (%s)" % (name, m.row_dimensions[4].height))
        check(ws.sheet_view.zoomScale == m.sheet_view.zoomScale,
              "%s: zoom == master (%s)" % (name, m.sheet_view.zoomScale))
        check(len(ws._images) == 1, "%s: logo present" % name)
        if ws._images:
            im = ws._images[0]
            check((im.anchor._from.col, im.anchor._from.row,
                   im.anchor._from.colOff, im.anchor._from.rowOff) == (0, 0, 130630, 64355),
                  "%s: logo anchored in the same spot as the master" % name)
            check((im.anchor.ext.width, im.anchor.ext.height) == (1415142, 577902),
                  "%s: logo same size as the master (no distortion)" % name)

    # ------------------------------------------------------------------ #
    print("\n[5] readability: nothing clipped, no excessive whitespace")
    # same calibrated metric the formatter uses to size columns
    from PIL import ImageFont
    F = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    FB = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    f12 = ImageFont.truetype(F, 16)
    fb12 = ImageFont.truetype(FB, 16)
    GLYPH_FACTOR, PAD, PPC = 0.88, 5, 7

    for name in list(expect) + [MASTER]:
        ws = out[name]
        if name in expect:
            ncols, nrows = expect[name][1], expect[name][0]
        else:
            ncols, nrows = 7, 40
        over = []
        for c in range(1, ncols + 1):
            letter = openpyxl.utils.get_column_letter(c)
            width = ws.column_dimensions[letter].width or 9.0
            avail = width * PPC + 5
            longest = 0
            for r in range(4, 5 + nrows):
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    longest = max(longest, fb12.getlength(str(v)) if r == 4
                                  else f12.getlength(str(v)))
            need = longest * GLYPH_FACTOR + PAD
            if need > avail + 0.5:
                over.append((letter, round(need), round(avail), "clipped"))
            if avail - need > 45:
                over.append((letter, round(need), round(avail), "wasteful"))
        msg = "%s: every column fits its text without big gaps %s" % (name, over)
        if name == MASTER:          # the reference is allowed its own quirks
            note(msg)
        else:
            check(not over, msg)
        check(all(ws.cell(row=r, column=c).alignment.wrap_text in (None, False)
                  for r in range(1, 6 + nrows) for c in range(1, ncols + 1)),
              "%s: no unintended text wrapping" % name)

    # ------------------------------------------------------------------ #
    print("\n[6] package hygiene")
    with zipfile.ZipFile(OUT) as z:
        bad = z.testzip()
        check(bad is None, "zip archive intact")
        bad_col = []
        for n in z.namelist():
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                if re.search(r'<col\b[^>]*\sstyle="\d+"', z.read(n).decode("utf-8")):
                    bad_col.append(n)
        check(not bad_col, "no invalid <col style=...> attributes left")
        sheets = re.findall(r'<sheet name="([^"]+)"', z.read("xl/workbook.xml").decode())
        check(len(sheets) == 5, "workbook declares all 5 sheets: %s" % sheets)

    print("\n" + "=" * 70)
    if failures:
        print("FAILED %d check(s):" % len(failures))
        for f in failures:
            print("  -", f)
        sys.exit(1)
    print("ALL CHECKS PASSED")


if __name__ == "__main__":
    main()
