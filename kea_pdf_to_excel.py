import sys
from pathlib import Path
import pandas as pd
from app import extract_options_from_pdf, make_excel

if len(sys.argv) < 2:
    print("Usage: python kea_pdf_to_excel.py input.pdf [output.xlsx]")
    raise SystemExit(1)

pdf_path = Path(sys.argv[1])
out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else pdf_path.with_suffix("_KEA_Option_Reorder_With_Code.xlsx")
with open(pdf_path, "rb") as f:
    df = extract_options_from_pdf(f)
if df.empty:
    print("No option rows found.")
    raise SystemExit(2)
out_path.write_bytes(make_excel(df))
print(f"Created: {out_path}")
