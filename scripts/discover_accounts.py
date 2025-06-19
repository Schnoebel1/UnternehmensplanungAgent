"""
discover_accounts.py
====================

Liest das Mapping aus config/sheets.yml und erzeugt pro Sheet
eine accounts-CSV in config/, mit den Spalten [row,text,category].

Category ∈ {forecast, readonly, ""} automatisch aus den Listen in sheets.yml.
"""

import re
import yaml
from pathlib import Path
from csv import DictWriter
from typing import Dict, List

from openpyxl import load_workbook
from loader import find_header_row

# --- Pfade ---
BASE      = Path(__file__).resolve().parent.parent
SRC_XLSX  = BASE / "data" / "UnternehmensplanungExcel.xlsx"
CFG_FILE  = BASE / "config" / "sheets.yml"
OUT_DIR   = BASE / "config"

# --- Hilfsfunktionen ---
_norm = re.compile(r"[^\w]+")
normalize = lambda s: _norm.sub("", s).lower()

def discover_sheet(sheet_name: str, cfg: Dict) -> None:
    """Erzeuge config/<slug>_accounts.csv für ein einzelnes Sheet."""
    # Forecast- und Readonly-Listen normalisiert
    fcast = {normalize(a) for a in cfg.get("forecast_accounts", [])}
    ro    = {normalize(a) for a in cfg.get("readonly_accounts", [])}
    aliases = cfg.get("header_aliases", ["t0"])

    wb = load_workbook(SRC_XLSX, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        print(f"⚠️  Sheet '{sheet_name}' nicht gefunden – übersprungen.")
        return
    ws = wb[sheet_name]

    # Header-Zeile
    header = find_header_row(ws, aliases)
    if header is None:
        print(f"⚠️  Sheet '{sheet_name}': Header-Zeile alias {aliases} nicht gefunden – übersprungen.")
        return

    # Konto-Spalte automatisch ermitteln
    account_col = None
    for col in range(1, cfg.get("max_scan_col", 15) + 1):
        sample = [ws.cell(r, col).value for r in range(header+1, header+8)]
        if any(isinstance(v, str) and v.strip() for v in sample):
            account_col = col
            break
    if account_col is None:
        print(f"⚠️  Sheet '{sheet_name}': Keine Konto-Spalte erkannt – übersprungen.")
        return

    # Zeilen sammeln
    rows: List[Dict[str,str]] = []
    for r in range(header+1, ws.max_row+1):
        txt = ws.cell(r, account_col).value
        if not isinstance(txt, str): continue
        txt = txt.strip()
        if not txt: continue

        key = normalize(txt)
        if key in fcast:
            cat = "forecast"
        elif key in ro:
            cat = "readonly"
        else:
            cat = ""

        rows.append({"row": r, "text": txt, "category": cat})

    # CSV schreiben
    slug = normalize(sheet_name)
    out_csv = OUT_DIR / f"{slug}_accounts.csv"
    OUT_DIR.mkdir(exist_ok=True)
    with out_csv.open("w", newline="", encoding="utf-8") as f:
        writer = DictWriter(f, fieldnames=["row","text","category"])
        writer.writeheader()
        writer.writerows(rows)

    print(f"✅ Sheet '{sheet_name}': {len(rows)} Zeilen → {out_csv.name}")

def main():
    cfg_all = yaml.safe_load(CFG_FILE.read_text(encoding="utf-8"))["sheets"]
    for sheet, cfg in cfg_all.items():
        discover_sheet(sheet, cfg)

if __name__ == "__main__":
    main()
