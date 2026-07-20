"""
_openpyxl_compat.py — make openpyxl tolerant of the DCF template's drawings.

`DCF Template 2020.xlsx` (and any workbook seeded by copying it) carries chart /
text-box drawing XML whose font `pitchFamily` attribute exceeds openpyxl's
hard-coded OOXML cap of 52. A strict, non-read-only `load_workbook()` then dies
with `ValueError: Max value is 52` before any sheet is read — which breaks every
pipeline step that full-loads the workbook to discover layout (excel_writer,
generate_*, fill_*).

Importing this module once (idempotent) widens the offending MinMax validator so
those workbooks load. It touches only the validator bound, never cell data, so
round-tripped files are unaffected.
"""
from __future__ import annotations


def apply() -> None:
    try:
        from openpyxl.drawing.text import Font as _DrawingFont
    except Exception:  # openpyxl layout changed — nothing to patch
        return
    desc = getattr(_DrawingFont, "pitchFamily", None)
    if desc is not None and getattr(desc, "max", None) is not None and desc.max < 127:
        # pitchFamily is a 1-byte field in practice; 127 covers all real values.
        desc.max = 127


apply()
