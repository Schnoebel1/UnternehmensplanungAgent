"""
writer_bs.py – Forecast-Writer für BS (2)
----------------------------------------

• schreibt nur die in config/sheets.yml -> forecast_accounts aufgelisteten Konten
• lässt alles andere exakt unverändert
• robustes Matching (Groß/klein, Leer- oder Sonderzeichen werden ignoriert)
• Debug-Ausgabe: zeigt im Terminal, welche Zeile geschrieben wurde
"""

from __future__ import annotations
from pathlib import Path
from typing import Dict, List
import re
import yaml

import numpy as np
from openpyxl import load_workbook

from loader import find_header_row, col_map
from forecast import cagr, project
from explanations import explain

# --------------------------------------------------------------------------- #
#  Config laden                                                               #
# --------------------------------------------------------------------------- #

CFG_FILE = Path(__file__).resolve().parent.parent / "config" / "sheets.yml"
with CFG_FILE.open(encoding="utf-8") as f:
    CFG: Dict = yaml.safe_load(f)["sheets"]["BS (2)"]

FORECAST_COLS = CFG["forecast_cols"]           # ["t1","t2","t3"]

# ---------------- Normalisierung ------------------------------------------- #
_norm = re.compile(r"[^\w]+")   # alles außer a–z, 0–9 als Trennzeichen

def normalize(txt: str) -> str:
    """Konto-String in Vergleichsform bringen."""
    return _norm.sub("", txt).lower()

FORECAST_ACCOUNTS = {normalize(a) for a in CFG["forecast_accounts"]}
READONLY_ACCOUNTS  = {normalize(a) for a in CFG["readonly_accounts"]}

# --------------------------------------------------------------------------- #
#  Konstanten                                                                 #
# --------------------------------------------------------------------------- #

SHEET_NAME   = "BS (2)"
PLACEHOLDERS = {None, "", "-", "–", "."}
PERIODS      = ["t-2", "t-1", "t0"] + FORECAST_COLS   # ['t-2', …, 't3']

# --------------------------------------------------------------------------- #
#  Hauptfunktion                                                              #
# --------------------------------------------------------------------------- #

def write_bs_forecast(src: Path, dst: Path) -> None:
    wb = load_workbook(src, data_only=False)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"'{SHEET_NAME}' fehlt im Workbook")

    ws = wb[SHEET_NAME]
    header = find_header_row(ws)
    if header is None:
        raise RuntimeError("Header-Zeile mit 't0' nicht gefunden")

    col = col_map(ws, header)
    if not all(p in col for p in PERIODS):
        missing = [p for p in PERIODS if p not in col]
        raise RuntimeError(f"Fehlende Spalten im Header: {missing}")

    account_col = 2
    span_years  = 2

    writes = 0  # Zähler für Debug

    for r in range(header + 1, ws.max_row + 1):
        raw = ws.cell(r, account_col).value
        if not isinstance(raw, str):
            continue
        acc = raw.strip()
        acc_norm = normalize(acc)

        # Konto weder forecastbar noch readonly → überspringen
        if acc_norm not in FORECAST_ACCOUNTS | READONLY_ACCOUNTS:
            continue

        # Nur forecasten, wenn es in der Forecast-Liste ist
        if acc_norm not in FORECAST_ACCOUNTS:
            continue

        # Historische Werte lesen
        try:
            t_minus_2 = float(ws.cell(r, col["t-2"]).value)
            t0        = float(ws.cell(r, col["t0"]).value)
        except (TypeError, ValueError):
            continue

        growth  = cagr(t_minus_2, t0, span_years)
        fcst: List[float] = project(t0, growth, horizon=3)

        for idx, p in enumerate(FORECAST_COLS):      # t1, t2, t3
            cell = ws.cell(r, col[p])

            if cell.data_type == "f":                                  # Formel
                continue
            if isinstance(cell.value, (int, float)) and cell.value not in PLACEHOLDERS:
                continue                                               # Zahl da
            cell.value = round(fcst[idx], 2)
            writes += 1

        # Erklärung
        expl = ws.cell(r, col["t3"] + 1)
        if expl.value in PLACEHOLDERS and expl.data_type != "f":
            expl.value = explain(acc, [t_minus_2, ws.cell(r, col["t-1"]).value, t0], fcst)

    dst.parent.mkdir(exist_ok=True)
    wb.save(dst)

    print(f"BS (2): {writes} Zellen neu befüllt – Datei gespeichert → {dst}")


# --------------------------------------------------------------------------- #
#  CLI-Test                                                                   #
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    import sys
    src = Path(sys.argv[1])
    dst = src.parent.parent / "outputs" / "UnternehmensplanungForecast.xlsx"
    write_bs_forecast(src, dst)
