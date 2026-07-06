# KEA PDF to Excel Reorder Tool

This setup converts a KEA option report PDF into an Excel reorder file.

## What it creates

- Extracted option rows
- Editable **New Option No** column
- Duplicate checks
- Delete check (`0` means delete)
- Console lines like `E005EC 1`
- Master console code inside Excel

## Run on Windows

Double-click:

```bat
RUN_APP_WINDOWS.bat
```

Or manually:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Use

1. Upload KEA option report PDF.
2. Edit **New Option No**.
3. Download Excel With Code.
4. Check Duplicate Check column.
5. Copy Console Lines.
6. Paste Master Console Script on live KEA page console.
7. Paste Console Lines when prompted.
8. Verify all numbers manually.
9. Click the official **Update Options** button.
10. Download final KEA option report and verify again.

## Safety note

This tool only prepares data and fills visible fields faster. Final submission must be done manually through the official KEA page after checking.
