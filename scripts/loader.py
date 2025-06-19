from __future__ import annotations
import re
from typing import Dict, List
from openpyxl.worksheet.worksheet import Worksheet

PERIOD_RE = re.compile(r"^t-?\d+$")   # t-2 … t3

def find_header_row(ws: Worksheet, aliases: List[str] = ["t0"]) -> int | None:
    """
    Liefert die Zeilennummer, in der eines der Aliase steht (z.B. 't0').
    Sucht in den ersten 40 Zeilen.
    """
    for row in ws.iter_rows(min_row=1, max_row=40):
        for cell in row:
            # nur echte Zell-Objekte mit .value und .row
            if isinstance(cell.value, str) and cell.value.strip() in aliases:
                return cell.row
    return None

def col_map(ws: Worksheet, header_row: int) -> Dict[str, int]:
    """Spaltenname (t-2 … t3) → Column-Index."""
    return {
        str(c.value): c.column
        for c in ws[header_row]
        if isinstance(c.value, str) and PERIOD_RE.fullmatch(c.value)
    }
