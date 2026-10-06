#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Faculty -> designation lookup.

Designations only exist in the 2025-26 workbook, so that is the only source.
Names are grouped by (surname, first initial) so that variants such as
``Dr.M.D.Sontakke`` and ``Dr. MANDAR SONTAKKE`` are recognised as one person.

If a person carries two *different* designations in the source - they look like
promotions inside 2025-26 - the lookup returns ``None`` rather than guessing,
and the cell is left blank.  Names absent from the source also return ``None``.
"""

import collections

import openpyxl

from report_format import HONORIFICS, person_key, squash

SRC = "Summary AiFormative_and_Summative_Feedback_Report 2025-2026.xlsx"

# sheets in the source that carry a designation column
SOURCES = [
    # sheet,                     first row, last row, designation col, name col
    ("summative",                5,  44, 2, 1),
    ("2025-26-Sem-I Formative",  5,  89, 1, 2),
    ("Summative Sem-I ",         5,  55, 1, 2),
]

HONORIFICS_SET = HONORIFICS


def collect(path=SRC):
    """-> {person_key: Counter({designation: n})} and who was seen under which name."""
    wb = openpyxl.load_workbook(path)
    seen = collections.defaultdict(collections.Counter)
    names = collections.defaultdict(set)
    for sheet, r0, r1, dcol, ncol in SOURCES:
        ws = wb[sheet]
        for r in range(r0, r1 + 1):
            name, desig = ws.cell(r, ncol).value, ws.cell(r, dcol).value
            key = person_key(name)
            if key and desig and str(desig).strip():
                seen[key][str(desig).strip()] += 1
                names[key].add(str(name).strip())
    return seen, names


def build(path=SRC, verbose=False):
    """
    -> {squashed name: designation or None}

    ``None`` means "do not fill this in" - either the name is missing from the
    source, or the source disagrees with itself about that person.
    """
    seen, names = collect(path)
    lookup = {}
    report = {"resolved": {}, "blank_conflict": {}, "blank_unknown": set()}

    for key, counter in seen.items():
        designations = set(counter)
        if len(designations) == 1:
            value = designations.pop()
            bucket = "resolved"
        else:
            value = None
            bucket = "blank_conflict"
        for raw in names[key]:
            lookup[squash(raw)] = value
            report[bucket][squash(raw)] = (dict(counter) if bucket == "blank_conflict" else value)

    if verbose:
        print("Designations resolved from %s: %d" % (path, len(report["resolved"])))
        for k in sorted(report["resolved"]):
            print("   %-26s %s" % (k, report["resolved"][k]))
        print("Left blank - source disagrees (promotion?): %d" % len(report["blank_conflict"]))
        for k in sorted(report["blank_conflict"]):
            print("   %-26s %s" % (k, report["blank_conflict"][k]))
    return lookup


if __name__ == "__main__":
    build(verbose=True)
