from pathlib import Path
import shutil
from openpyxl import load_workbook

from writers.writer_bs       import write_bs_forecast
from writers.writer_pnl      import write_pnl_forecast
from writers.writer_cfr      import write_cfr_forecast
from writers.writer_rev_sbe  import write_rev_sbe_forecast
from writers.writer_cogs     import write_cogs_forecast
from writers.writer_opex     import write_opex_forecast
from writers.writer_capex    import write_capex_forecast
from writers.writer_staff    import write_staff_forecast

# Basis-Pfad (KiAgent/scripts)
BASE     = Path(__file__).resolve().parent.parent
SRC_XLSX = BASE / "data"    / "UnternehmensplanungExcel.xlsx"
DST_XLSX = BASE / "outputs" / "UnternehmensplanungForecast.xlsx"

def main() -> None:
    # 1) Excel ein einziges Mal kopieren
    DST_XLSX.parent.mkdir(exist_ok=True, parents=True)
    shutil.copy(SRC_XLSX, DST_XLSX)

    # 2) Workbook öffnen (write-enabled)
    wb = load_workbook(DST_XLSX, data_only=False)

    # 3) Nacheinander alle Forecast-Writer aufrufen
    write_bs_forecast(wb)
    write_pnl_forecast(wb)
    write_cfr_forecast(wb)
    write_rev_sbe_forecast(wb)
    write_cogs_forecast(wb)
    write_opex_forecast(wb)
    write_capex_forecast(wb)
    write_staff_forecast(wb)

    # 4) Alles in die Forecast-Datei speichern
    wb.save(DST_XLSX)
    print(f"✅ Alle Forecasts geschrieben in: {DST_XLSX}")

if __name__ == "__main__":
    main()
