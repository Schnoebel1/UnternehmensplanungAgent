# UnternehmensplanungAgent
# KiAgent

KiAgent ist ein Python-basiertes Tool zur automatisierten Forecast-Erstellung für Unternehmensplanungs-Exceldateien. Mit Hilfe von OpenPyXL, LangChain und Ollama generiert es Prognosen für verschiedene Bilanz- und GuV-Positionen. Die Anwendung besteht aus mehreren Modulen zur Kontenerkennung, Forecast-Berechnung und Ergebnisaufbereitung.

## Projektstruktur

* **data/**: Enthält die Quelldatei `UnternehmensplanungExcel.xlsx` und Kontextdaten (`cases.csv`).
* **config/**: YAML- und CSV-Dateien für Sheet-Definitionen und Konten-Mapping (`sheets.yml`, `*_accounts.csv`).
* **scripts/**: Kernskripte für Discover, Loader, Forecast und Hauptprogramm (`discover_accounts.py`, `loader.py`, `forecast.py`, `explanations.py`, `main.py`).
* **scripts/writers/**: Module zum Schreiben von Forecasts in einzelne Sheets (`writer_bs.py`, `writer_pnl.py`, `writer_cfr.py`, u.v.m.).
* **app/**: Optionales Streamlit-Frontend (`streamlit_app.py`) für interaktive Vorschau.
* **outputs/**: Debug-Protokolle (`*_debug.txt`) und das finale Forecast-Workbook `UnternehmensplanungForecast.xlsx`.

## Installation

1. Repository klonen und in das Projektverzeichnis wechseln.
2. `python -m venv .venv` und `source .venv/bin/activate` (Linux/macOS) oder `.\.venv\Scripts\activate` (Windows).
3. `pip install -r requirements.txt` ausführen.
4. Sicherstellen, dass Ollama läuft und das Modell (`llama3:8b`) verfügbar ist.
5. Optional die Umgebungsvariablen `OLLAMA_MODEL` und `OLLAMA_URL` anpassen.

## Nutzung

* **Kontenerkennung**: `python scripts/discover_accounts.py` erzeugt `config/*_accounts.csv`. Mit `--debug SHEETNAME` können Einzeltabellen im Detail geprüft werden. Ist noch etwas statisch an die Excel angepasst.
* **Forecast-Lauf**: `python scripts/main.py` kopiert die Quelldatei, führt alle Forecast-Writer aus und speichert das Ergebnis in `outputs/UnternehmensplanungForecast.xlsx`.
* **Streamlit-Frontend**: `streamlit run app/streamlit_app.py` öffnet eine interaktive Oberfläche zur Ansicht.

## Module im Detail

* **loader.py**: Liefert Funktionen zum Finden von Header-Zeilen und Spalten-Mapping.
* **forecast.py**: Basishilfen für CAGR-Berechnung und Projektion.
* **explanations.py**: Kapselt LLM-Aufruf mit JSON-Output und Fallback auf Baseline-Forecast.
* **discover\_accounts.py**: Erstellt CSV-Dateien zur Kategorisierung von Kontenzeilen (Forecast vs. readonly).
* **Writer-Module**: Lesen historische Werte, rufen `explain()` auf und schreiben Prognosewerte (t1–t3) sowie Begründungen ins Workbook.

## Anpassung und Erweiterung

* Neue Sheets und Alias-Namen konfigurieren in `config/sheets.yml`.
* Konten-Mapping direkt in `config/*_accounts.csv` ändern oder neu generieren.
* Weitere Writer-Module können nach dem vorhandenen Muster in `scripts/writers/` hinzugefügt werden.

## Logging & Debugging

* Debug-Logs pro Modul liegen in `outputs/` als `*_debug.txt`.
* Mit `--debug`-Flags lassen sich zusätzliche Konsolenausgaben aktivieren.

Viel Erfolg mit KiAgent!
