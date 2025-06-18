from pathlib import Path
import sys

from writer_bs import write_bs_forecast

def main(src: Path) -> None:
    dst = src.parent.parent / "outputs" / "UnternehmensplanungForecast.xlsx"
    write_bs_forecast(src, dst)

if __name__ == "__main__":
    workbook = (Path(sys.argv[1])
                if len(sys.argv) > 1
                else Path(__file__).resolve().parent.parent
                   / "data" / "UnternehmensplanungExcel.xlsx")
    main(workbook)
