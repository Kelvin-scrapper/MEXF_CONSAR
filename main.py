"""
MEXF_CONSAR — Main Pipeline
Outputs: output/MEXF_CONSAR_DATA_YYYYMMDD.xlsx
         output/MEXF_CONSAR_META_YYYYMMDD.xlsx
         output/MEXF_CONSAR_YYYYMMDD.ZIP
Usage  : python main.py
"""

import os
import sys
import zipfile
from datetime import date

import openpyxl
import pandas as pd

import scraper
import mapper

# openpyxl's safe_string uses "%.16g" which truncates 17-digit IEEE 754 doubles.
# Patch it to use repr(float()) so the full source value is preserved in the
# XML <v> element and shows untruncated in Excel's formula bar.
# Note: _save_data forces engine='openpyxl' so this patch covers both the
# initial DataFrame write and the _apply_number_format rewrite.
import openpyxl.cell._writer as _openpyxl_cell_writer
from math import isnan as _isnan, isinf as _isinf
_orig_safe_string = _openpyxl_cell_writer.safe_string

def _full_precision_safe_string(value):
    if isinstance(value, float) and not _isnan(value) and not _isinf(value):
        return repr(float(value))  # float() strips numpy wrapper; repr gives shortest exact string
    return _orig_safe_string(value)

_openpyxl_cell_writer.safe_string = _full_precision_safe_string

OUTPUT_DIR    = "output"
DOWNLOADS_DIR = "downloads"
OUTPUT_PREFIX = "MEXF_CONSAR"


def _datestamp() -> str:
    return date.today().strftime("%Y%m%d")


def _apply_number_format(filepath: str) -> None:
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    for row in ws.iter_rows(min_row=3, min_col=2):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.##"
    wb.save(filepath)


def _save_data(df: pd.DataFrame, datestamp: str) -> str:
    filename = f"{OUTPUT_PREFIX}_DATA_{datestamp}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    df.to_excel(filepath, index=False, header=False, engine="openpyxl")
    _apply_number_format(filepath)
    print(f"[main] DATA saved: {filepath}")
    return filepath


def _save_metadata(datestamp: str) -> str:
    filename = f"{OUTPUT_PREFIX}_META_{datestamp}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    meta_rows = mapper.build_metadata_rows()
    pd.DataFrame(meta_rows).to_excel(filepath, index=False)
    print(f"[main] META saved: {filepath}")
    return filepath


def _create_zip(data_path: str, meta_path: str, datestamp: str) -> str:
    zip_name = f"{OUTPUT_PREFIX}_{datestamp}.ZIP"
    zip_path = os.path.join(OUTPUT_DIR, zip_name)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(data_path, arcname=os.path.basename(data_path))
        zf.write(meta_path, arcname=os.path.basename(meta_path))
    print(f"[main] ZIP created: {zip_path}")
    return zip_path


def _find_existing_data() -> str | None:
    if not os.path.isdir(OUTPUT_DIR):
        return None
    files = sorted(
        f for f in os.listdir(OUTPUT_DIR)
        if f.startswith(f"{OUTPUT_PREFIX}_DATA_") and f.endswith(".xlsx")
    )
    return os.path.join(OUTPUT_DIR, files[-1]) if files else None


def scrape() -> None:
    """Full pipeline: fetch → parse → merge → save DATA + META + ZIP."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)

    datestamp = _datestamp()

    print("[main] Step 1: Fetching data from all CONSAR sources...")
    data_paths = scraper.fetch_data(downloads_dir=DOWNLOADS_DIR)
    print(f"[main] Downloaded {len(data_paths)} sources")

    existing = _find_existing_data()
    if existing:
        print(f"[main] Merging with existing data: {existing}")

    print("[main] Step 2: Parsing and mapping to output format...")
    out_df = mapper.map_to_output(data_paths, existing_path=existing)
    print(f"[main] Output shape: {out_df.shape}  "
          f"(2 header rows + {out_df.shape[0] - 2} data rows)")

    print("[main] Step 3: Saving files...")
    data_path = _save_data(out_df, datestamp)
    meta_path = _save_metadata(datestamp)
    _create_zip(data_path, meta_path, datestamp)

    print("[main] Done.")


def main() -> None:
    scrape()


if __name__ == "__main__":
    main()
