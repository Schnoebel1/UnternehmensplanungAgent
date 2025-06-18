from __future__ import annotations
import re
from typing import Dict
from openpyxl.worksheet.worksheet import Worksheet

PERIOD_RE = re.compile(r"^t-?\d+$")   # t-2 … t3

def find_header_row(ws: Worksheet) -> int | None:
    """liefert Zeilennummer, in der *t0* steht."""
    for row in ws.iter_rows(min_row=1, max_row=40):
        if any(c.value == "t0" for c in row):
            return row[0].row
    return None

def col_map(ws: Worksheet, header_row: int) -> Dict[str, int]:
    """Spaltenname (t-2 … t3) → Column-Index."""
    return {str(c.value): c.column
            for c in ws[header_row]
            if isinstance(c.value, str) and PERIOD_RE.fullmatch(c.value)}
