# RFH-Receipt-Reader

Scans receipt photos with Gemini, records them in an Excel workbook, and files
the images into per-currency folders with standardized names.

For Canadian receipts it also checks the numbers: the pre-tax amount is worked
out from the total and tax, and a row is highlighted red if the tax rate is higher
than the province allows or the receipt's printed subtotal doesn't match.

## Setup

1. **Install Python 3.9+** and the dependencies:

   ```
   pip install -r requirements.txt
   ```

2. **Get a free Gemini API key** at https://aistudio.google.com/apikey
   and set it as an environment variable. The scripts read it only from
   `GEMINI_API_KEY`, so it never ends up in the code or in git.

   PowerShell, current window only:

   ```
   $env:GEMINI_API_KEY = "your-key"
   ```

   Or saved for future windows (open a new terminal afterwards):

   ```
   setx GEMINI_API_KEY "your-key"
   ```

3. **Check the settings** at the top of `process_receipts.py`:
   - `ROOT`: the working folder (default `C:\RBT\Receipts`, or set `RECEIPTS_ROOT`)
   - `EXCEL_FILENAME`: the workbook name, which must sit in `ROOT`
   - `MODEL`: the Gemini model (or set `GEMINI_MODEL`)

   The workbook needs a sheet for each region folder (`1_CAD`, `2_USD`, ...) with
   the headers in row 3.

4. **Build the folders** (run once):

   ```
   python folder_setup.py
   ```

   This creates `1_CAD`, `2_USD`, `3_MEX`, `4_EUR`, each with
   `OrigScannedReceipts`, `ProcessedReceipts`, and `ProcessedReceiptsRenamed`.

## Daily use

1. Photograph receipts and drop the images into the matching region's
   `OrigScannedReceipts` folder (e.g. a Canadian receipt goes in
   `1_CAD\OrigScannedReceipts`).

2. **Close the workbook in Excel.** The script tells you if it can't write to it.

3. Run:

   ```
   python process_receipts.py
   ```

   Pick the currency number when prompted. Each receipt is scanned, added as a
   row to its sheet, copied into `ProcessedReceipts`, then moved and renamed into
   `ProcessedReceiptsRenamed` as `YYYYMMDD_1234.56_COMPANY_TAX`.

Options:

```
python process_receipts.py --region CAD     # skip the menu (number, code or folder name)
python process_receipts.py --dry-run        # scan and report, write and move nothing
python process_receipts.py --root D:\Receipts --excel Other.xlsx
```

If a receipt can't be read or the API errors out, that file is left where it is
and the run carries on with the next one. Rate limits (429) and overload errors
(503) are retried automatically.

A receipt with the same date and total as an existing row replaces that row
instead of adding a duplicate. Two different receipts with the same date and
exactly the same total would collide, so check the preview if that could happen.

## Other scripts

- `scan_one.py "path\to\receipt.jpg"` scans one image and prints the result,
  touching nothing. Useful for checking a tricky receipt. `--region EUR` switches
  to day-first dates.
- `hyperlink.py` is a one-off fix that turns old plain-text filenames in the
  workbook into clickable links. Not needed for new rows.

Canadian tax rates live in the `CAD_Tax_Rates` sheet, created on the first live
run. Edit rates there.

## Tests

```
pip install -r requirements-dev.txt
pytest
```

The tests cover the amount and flagging logic and a full run against a temporary
workbook, with the Gemini call stubbed out, so they need no key or network.
