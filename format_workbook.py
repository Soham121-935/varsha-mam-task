#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rebuild the 2025-26 feedback workbook on the design system of its master sheet
``summative``.

    python3 format_workbook.py

Writes ``Summary_AiFormative_and_Summative_Feedback_Report_2025-2026_Formatted.xlsx``
and leaves the original (and a copy under ``backup/``) untouched.
"""

import os
import shutil

import openpyxl

from report_format import (MASTER_SHEET, build_sheet, cleanup, extract_logo,
                           fresh_sheet, make_logo, read_design, reseat_images,
                           tidy_package)

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Summary AiFormative_and_Summative_Feedback_Report 2025-2026.xlsx")
BACKUP_DIR = os.path.join(HERE, "backup")
OUT = os.path.join(HERE, "Summary_AiFormative_and_Summative_Feedback_Report_2025-2026_Formatted.xlsx")


def grid(ws, cols, first_row, last_row):
    return [[ws.cell(row=r, column=c).value for c in cols] for r in range(first_row, last_row + 1)]


def yellow_cells(ws, cols, first_row, last_row):
    """The row flagging on this sheet is deliberate, so it is carried over."""
    found = {}
    for i, r in enumerate(range(first_row, last_row + 1)):
        for j, c in enumerate(cols):
            cell = ws.cell(row=r, column=c)
            if cell.fill.fill_type == "solid" and str(cell.fill.fgColor.rgb).upper() == "FFFFFF00":
                found[(i, j)] = "FFFFFF00"
    return found


# Sem-I report schema: Designation | Faculty Name | Semester/ Group | Subject | Performance
SEMI_HEADERS = ["Designation", "Faculty Name", "Semester/ Group", "Subject", "Performance "]
# Sem-II report schema (identical to the master sheet)
SEMII_HEADERS = ["Faculty Name", "Designation", "Course", "Sem/Class",
                 "Attendees", "Performance", "Date"]


def main():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    backup = os.path.join(BACKUP_DIR, os.path.basename(SRC))
    if not os.path.exists(backup):
        shutil.copy2(SRC, backup)

    logo_png = extract_logo(SRC)
    wb = openpyxl.load_workbook(SRC)
    reseat_images(wb, logo_png)
    design = read_design(wb)

    # ------------------------------------------------------------------ #
    # 1. 2025-26-Sem-I Formative      (rows 5-89, header already on row 4)
    # ------------------------------------------------------------------ #
    old = wb["2025-26-Sem-I Formative"]
    data = grid(old, [1, 2, 3, 4, 5], 5, 89)
    headers = [old.cell(row=4, column=c).value for c in (1, 2, 3, 4, 5)]  # verbatim
    keeps = yellow_cells(old, [1, 2, 3, 4, 5], 5, 89)
    ws = fresh_sheet(wb, old)
    build_sheet(ws, design, "AY_2025_26_Sem_I_Formative", headers, data,
                numeric_cols={4}, logo=make_logo(logo_png), keep_fills=keeps)

    # ------------------------------------------------------------------ #
    # 2. Summative Sem-I              (rows 4-54, no header row originally)
    # ------------------------------------------------------------------ #
    old = wb["Summative Sem-I "]
    data = grid(old, [1, 2, 3, 4, 5], 4, 54)
    ws = fresh_sheet(wb, old)
    # labels are not invented: the sibling Sem-I sheet has the same 5-column layout
    build_sheet(ws, design, "AY_2025_26_Sem_I_Summative", SEMI_HEADERS, data,
                numeric_cols={4}, logo=make_logo(logo_png))

    # ------------------------------------------------------------------ #
    # 3. Sem-II Frramtive   (col A is a constant repeated label and is dropped;
    #    cols B-H then match the master's schema exactly)
    # ------------------------------------------------------------------ #
    old = wb["Sem-II Frramtive"]
    data = grid(old, [2, 3, 4, 5, 6, 7, 8], 2, 19)
    ws = fresh_sheet(wb, old)
    build_sheet(ws, design, "AY_2025_26_Sem_II_Formative", SEMII_HEADERS, data,
                numeric_cols={4, 5}, logo=make_logo(logo_png))

    wb.active = 0
    wb.save(OUT)
    stripped = tidy_package(OUT, SRC, preserve_dimension_of=MASTER_SHEET)
    cleanup()

    print("wrote %s" % os.path.basename(OUT))
    print("  %d stray <col style=...> attributes removed" % stripped)


if __name__ == "__main__":
    main()
