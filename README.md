# Feedback Report 2025-26 — formatted workbook

| | |
|---|---|
| **Original (untouched backup)** | `Summary AiFormative_and_Summative_Feedback_Report 2025-2026.xlsx` (also copied to `backup/`) |
| **Deliverable** | `Summary_AiFormative_and_Summative_Feedback_Report_2025-2026_Formatted.xlsx` |
| **Build script** | `format_workbook.py` |
| **QC script** | `verify_workbook.py` |

```bash
python3 format_workbook.py    # rebuilds the deliverable from the original
python3 verify_workbook.py    # runs every quality-control check
```

Requires `openpyxl` and `Pillow`.

## What was done

The workbook's `summative` sheet (AY 2025-26, Sem-II summative) was used as the
master formatting template. Its design system is read **out of the workbook
itself** — fonts, sizes, border style/colour, merges, row heights, column-width
rule, logo image and logo anchor offsets — and re-applied to the other sheets.
`summative` itself is left alone.

### The design system taken from `summative`

| Element | Master's treatment |
|---|---|
| Rows 1–3 | one open box around a 3-row title block, merged across the table: college / department / report name — Calibri 14, centred, 18pt row height |
| Row 4 | column headings — Calibri 12 **bold**, thin box border, 15.6pt row height |
| Rows 5+ | data — Calibri 12, thin box border, left+vertically centred text, centred numbers, 15.6pt row height |
| Borders | `thin` with the automatic (indexed 64) colour, boxed on every table cell |
| Columns | exactly as wide as the longest entry in that column, header included |
| Logo | the workbook's own `image1.png`, anchored inside A1 at its master offsets, 1 415 142 × 577 902 EMU (≈149 × 61 px) — same size and aspect ratio on every sheet |
| View | zoom 70, gridlines on, no freeze panes |

Column widths are derived with a text-metric rule calibrated against the
master's own fitted widths (`0.88 × glyph width + 5 px`, ÷ 7 px per character),
which reproduces 5 of the master's 7 column widths to within about a tenth of a
character.

### Sheets rebuilt

| Sheet | Rows | Columns | Notes |
|---|---|---|---|
| `2025-26-Sem-I Formative` | 85 | 5 | header row 4 kept verbatim |
| `Summative Sem-I ` | 51 | 5 | header row added (see below) |
| `Sem-II Frramtive` | 18 | 7 | header row added; constant column A dropped (see below) |
| `Sheet2` | – | – | empty left-over tab, left as-is |
| `summative` | 40 | 7 | **master — not modified** |

## Judgement calls

Three decisions change what is *in* the sheets, so they were made explicitly:

1. **Missing header rows were added.** `Summative Sem-I ` and `Sem-II Frramtive`
   had no column headings at all. Labels were not invented: `Summative Sem-I `
   took the five headings of its sibling Sem-I sheet, which has the identical
   layout (`Designation / Faculty Name / Semester/ Group / Subject / Performance`),
   and `Sem-II Frramtive` took the master's own seven, since its columns line up
   1:1 with them.

2. **`Sem-II Frramtive` column A was dropped.** Every one of its 18 rows repeated
   the constant `AY_2025_26_Sem_II_Formative` (the same text as the sheet title).
   Columns B–H then match the master's schema exactly.

3. **Yellow row highlighting was kept; grey striping was dropped.** The 24
   yellow-highlighted rows in `2025-26-Sem-I Formative` look like deliberate
   flagging, so all 72 highlighted cells were preserved. The grey/white fills on
   the Sem-I sheets were automatic table banding left over from a paste, and the
   master has no banding, so they were removed.

Everything else is preserved byte for byte: values, text, numbers, spelling and
spacing of names and dates. There were no formulas, conditional formats,
validations or hidden rows/columns to preserve.

## Verification

`verify_workbook.py` checks, and all of it passes:

* the master keeps every value and every style on all 378 cells, plus its merges,
  column widths, row heights, view settings, margins, logo and `A1:G54` used range
* every data row on every rebuilt sheet is value-identical to the original
* all 72 yellow-highlighted cells survive
* every rebuilt sheet matches the master's title font, heading font, body font,
  border style, row heights, zoom and logo position/size
* no column clips its text and none is wastefully wide
* the package is a valid zip; every XML part parses and every relationship resolves

Two cosmetic openpyxl round-trip artefacts are cleaned up in `tidy_package()`
(a stray `style="N"` on `<col>` elements, and the master's `<dimension>` hint).
