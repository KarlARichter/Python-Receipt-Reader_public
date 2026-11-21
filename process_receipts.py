"""
process_receipts.py

Walks a chosen region's OrigScannedReceipts folder and, for every receipt image:
  STEP 0 - you pick which currency to process by its number (1 = CAD, 2 = USD...)
  STEP 1 - scans it (Gemini) to get date/company/subtotal/tax/total/card/province
  STEP 2 - appends one row to the matching sheet in the Excel workbook. The
           pre-tax Amount is DERIVED (Total - Tax, minus Tip if present) rather
           than trusting the receipt's subtotal label (Level 1). Two type-agnostic
           checks then run (Level 2): the effective rate must not exceed the
           province rate (a ceiling, so mixed/exempt general receipts don't
           false-flag), and any printed subtotal must match the derived Amount.
           A row that fails is flagged CHECK with its Amount cell filled red.
           If a row with the same TransactionDate AND TotalAmount already exists,
           that row is OVERWRITTEN in place (a re-scan updates it) instead of
           adding a duplicate.
  STEP 3 - copies the original into ProcessedReceipts (keeps original name)
  STEP 4 - moves + renames it into ProcessedReceiptsRenamed as
           YYYYMMDD_1234.56_COMPANY_TAX.jpg

Region folders are auto-discovered by pattern, so adding "5_JPY" later needs
no code change.

Setup:
  pip install -r requirements.txt
  set the GEMINI_API_KEY environment variable (see README)

Usage:
  python process_receipts.py                 pick a currency from the menu
  python process_receipts.py --region CAD    skip the menu (number or code)
  python process_receipts.py --dry-run       preview only, nothing is written or moved
"""

import argparse
import json
import os
import re
import sys
import time
import shutil
import datetime
from io import BytesIO

from google import genai
from google.genai import types
from google.genai.errors import ClientError, ServerError
from PIL import Image, ImageOps

import openpyxl
from openpyxl.styles import Font, PatternFill


# Working directory. Override with the RECEIPTS_ROOT environment variable or --root.
ROOT = os.environ.get("RECEIPTS_ROOT", r"C:\RBT\Receipts")

# Excel workbook, sitting in the root. Override with --excel.
EXCEL_FILENAME = "RFH_BioTek_Receipts_KAR.xlsx"

# The key is read from the environment only. Never put it in this file: anything
# committed to git stays in the history even after it is deleted.
MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")

# Preview mode. True = show what WOULD happen, touch nothing (no Excel, no files).
# False = actually write Excel + copy + move + rename. The --dry-run flag sets this.
DRY_RUN = False

# Pause between receipts so we stay under the model's requests-per-minute limit.
# Free tier is 15 RPM. A 6s gap (~10/min) leaves margin so a little timing drift
# doesn't tip over the limit. Set to 0 on a paid tier. Skipped during preview.
SECONDS_BETWEEN_SCANS = 6

# On a 429 (rate limit) the API tells us how long to wait; we sleep and retry the
# SAME receipt this many times before giving up on it. Stops a rate limit from
# crashing the whole run.
MAX_RETRIES_ON_RATE_LIMIT = 5

# Only folders matching this are treated as regions: "1_CAD", "2_USD", "10_JPY"...
REGION_FOLDER_PATTERN = r"^\d+_[A-Za-z]{2,4}$"

SUBDIR_ORIG = "OrigScannedReceipts"
SUBDIR_PROCESSED = "ProcessedReceipts"
SUBDIR_RENAMED = "ProcessedReceiptsRenamed"

SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".heic", ".heif", ".pdf"}

MAX_IMAGE_EDGE = 1568

# Headers live in row 3 of every sheet; data starts on row 4.
HEADER_ROW = 3
FIRST_DATA_ROW = 4

# Rate ceiling tolerance: flag only when the effective rate exceeds the province
# rate by more than this (0.01 = 1 percentage point), allowing for receipt rounding.
RATE_TOLERANCE = 0.01

# Cross-check: a receipt's own pre-tax subtotal (when it prints a real one) should
# match the derived Amount (Total - Tax - Tip). If they disagree by more than this
# many dollars, the tax/tip read is inconsistent or it's a multi-transaction slip
# (e.g. gas + car wash) - flag CHECK. Guarded so a "subtotal" that just equals the
# total (pump receipts print "Fuel sales = total", not a real pre-tax line) is ignored.
CROSS_CHECK_TOLERANCE = 0.05

# Seed values for the CAD_Tax_Rates sheet, written ONCE if that sheet is absent.
# After that, the sheet in the workbook is the source of truth - edit it there.
# Rate is combined sales tax as a decimal (0.13 = 13%). Verify against current rates.
DEFAULT_CAD_TAX_RATES = {
    "AB": ("Alberta", 0.05),
    "BC": ("British Columbia", 0.12),
    "MB": ("Manitoba", 0.12),
    "NB": ("New Brunswick", 0.15),
    "NL": ("Newfoundland and Labrador", 0.15),
    "NS": ("Nova Scotia", 0.14),
    "NT": ("Northwest Territories", 0.05),
    "NU": ("Nunavut", 0.05),
    "ON": ("Ontario", 0.13),
    "PE": ("Prince Edward Island", 0.15),
    "QC": ("Quebec", 0.14975),
    "SK": ("Saskatchewan", 0.11),
    "YT": ("Yukon", 0.05),
}
TAX_RATES_SHEET = "CAD_Tax_Rates"

# US state sales tax rates (combined state + average local, as a reference rate).
# Source: standard combined rates; used only to populate the TaxRate column on
# USD receipts - no checking is done against it.
DEFAULT_USD_TAX_RATES = {
    "AL": ("Alabama", 0.0922), "AK": ("Alaska", 0.0176), "AZ": ("Arizona", 0.0840),
    "AR": ("Arkansas", 0.0947), "CA": ("California", 0.0873), "CO": ("Colorado", 0.0777),
    "CT": ("Connecticut", 0.0635), "DE": ("Delaware", 0.0000), "FL": ("Florida", 0.0708),
    "GA": ("Georgia", 0.0737), "HI": ("Hawaii", 0.0444), "ID": ("Idaho", 0.0603),
    "IL": ("Illinois", 0.0882), "IN": ("Indiana", 0.0700), "IA": ("Iowa", 0.0694),
    "KS": ("Kansas", 0.0869), "KY": ("Kentucky", 0.0600), "LA": ("Louisiana", 0.0952),
    "ME": ("Maine", 0.0550), "MD": ("Maryland", 0.0600), "MA": ("Massachusetts", 0.0625),
    "MI": ("Michigan", 0.0600), "MN": ("Minnesota", 0.0749), "MS": ("Mississippi", 0.0707),
    "MO": ("Missouri", 0.0823), "MT": ("Montana", 0.0000), "NE": ("Nebraska", 0.0694),
    "NV": ("Nevada", 0.0823), "NH": ("New Hampshire", 0.0000), "NJ": ("New Jersey", 0.0663),
    "NM": ("New Mexico", 0.0783), "NY": ("New York", 0.0852), "NC": ("North Carolina", 0.0698),
    "ND": ("North Dakota", 0.0696), "OH": ("Ohio", 0.0723), "OK": ("Oklahoma", 0.0898),
    "OR": ("Oregon", 0.0000), "PA": ("Pennsylvania", 0.0634), "RI": ("Rhode Island", 0.0700),
    "SC": ("South Carolina", 0.0746), "SD": ("South Dakota", 0.0640), "TN": ("Tennessee", 0.0955),
    "TX": ("Texas", 0.0825), "UT": ("Utah", 0.0719), "VT": ("Vermont", 0.0624),
    "VA": ("Virginia", 0.0580), "WA": ("Washington", 0.0993), "WV": ("West Virginia", 0.0650),
    "WI": ("Wisconsin", 0.0543), "WY": ("Wyoming", 0.0534), "DC": ("Washington DC", 0.0600),
}

# Tax cells filled with a calculated (not receipt-read) value get this shade,
# so a calculated tax is visibly distinct from one read off the receipt.
CALC_TAX_FILL = PatternFill("solid", fgColor="FFE699")  # light amber

# Amount cells for rows Level 2 flags as "CHECK" (derived rate doesn't match the
# province, or the Amount couldn't be derived) get this shade, so a row that needs
# a human look is visible in the sheet later - not just in the console during the run.
NEEDS_REVIEW_FILL = PatternFill("solid", fgColor="FFC7CE")  # light red
# --------------------------------------------------------------------------


# ============================ STEP 1: scan =================================

def load_image_bytes(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        import fitz
        with fitz.open(path) as doc:
            pix = doc.load_page(0).get_pixmap(dpi=200)
            img = Image.open(BytesIO(pix.tobytes("png")))
            img.load()
    else:
        if ext in (".heic", ".heif"):
            import pillow_heif
            pillow_heif.register_heif_opener()
        img = Image.open(path)
        # Phone photos are often stored sideways with an EXIF rotation tag.
        # convert("RGB") below drops that tag, so apply it first or the model
        # gets a rotated receipt.
        img = ImageOps.exif_transpose(img)

    img = img.convert("RGB")
    w, h = img.size
    longest = max(w, h)
    if longest > MAX_IMAGE_EDGE:
        s = MAX_IMAGE_EDGE / float(longest)
        img = img.resize((int(w * s), int(h * s)), Image.LANCZOS)

    buf = BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return buf.getvalue()


def build_prompt(region):
    return (
        'You are reading a single receipt image. Return ONLY a JSON object.\n\n'
        'Fields:\n'
        '- "transaction_date": purchase date as "YYYY-MM-DD", or null if unreadable.\n'
        '- "company": merchant / store name, or null.\n'
        '- "subtotal": the pre-tax subtotal as a plain number, no symbol, or null. '
        'This is the amount BEFORE tax is added.\n'
        '- "total": grand total actually paid (tax included, and tip if any) as a plain number, no symbol, or null.\n'
        '- "tip": tip / gratuity amount if a tip line is printed, as a plain number, else null. '
        'Common labels: Tip, Gratuity, Pourboire.\n'
        '- "currency": 3-letter code if visible (CAD, USD, MEX, EUR), else null.\n'
        '- "taxes": list of {"type": "...", "amount": number}. type is HST, GST, PST, QST, or TAX. Empty list if none.\n'
        '- "card_last4": the last 4 digits of the credit/debit card if printed on the receipt, as a string, else null. '
        'Do NOT return the full card number, only the last 4 digits.\n'
        '- "province": for CAD receipts, the 2-letter Canadian province code (ON, AB, QC...) '
        'from the address. For USD receipts, the 2-letter US state code (NY, FL, TX...) '
        'from the address. Null for other currencies or if unreadable.\n'
        '- "address": merchant address as printed, or null.\n'
        '- "confidence": "high", "medium", or "low".\n'
        '- "notes": anything odd (damaged, cut off, handwritten, ambiguous).\n\n'
        'Date rules:\n'
        '- This receipt is from region "%s". For CAD/USD/MEX the printed order is usually month-day-year. For EUR it is day-month-year.\n'
        '- Override that when obvious: if one component is 13 or higher it must be the day.\n'
        '- Always output YYYY-MM-DD regardless of how it was printed.\n\n'
        'Number rules:\n'
        '- European style "1.234,56" means 1234.56. Output a plain decimal.\n'
    ) % region


def get_client():
    key = os.environ.get("GEMINI_API_KEY")
    if not key:
        sys.exit("GEMINI_API_KEY is not set. See the Setup section of the README.")
    return genai.Client(api_key=key)


def scan(path, region):
    image_bytes = load_image_bytes(path)
    client = get_client()

    attempt = 0
    while True:
        try:
            response = client.models.generate_content(
                model=MODEL,
                contents=[
                    types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
                    build_prompt(region),
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0,
                ),
            )
            break
        except (ClientError, ServerError) as e:
            # 429 = rate limit, 503 = model overloaded. Both are temporary: wait the
            # delay the API asks for (or a default) and retry the same receipt.
            if getattr(e, "code", None) in (429, 503) and attempt < MAX_RETRIES_ON_RATE_LIMIT:
                attempt += 1
                wait = _retry_delay_seconds(e, default=SECONDS_BETWEEN_SCANS or 5)
                print("      API busy ({}), waiting {:.0f}s then retrying ({}/{})...".format(
                    e.code, wait, attempt, MAX_RETRIES_ON_RATE_LIMIT))
                time.sleep(wait + 1)  # +1s cushion
                continue
            raise  # some other error, or out of retries

    try:
        parsed = json.loads(response.text)
        if isinstance(parsed, dict):
            return parsed
    except (TypeError, ValueError):
        pass
    print("    Could not parse JSON for {}. Raw:\n{}".format(os.path.basename(path), response.text))
    return None


def _retry_delay_seconds(err, default=5):
    """Pull the 'retry in Xs' hint out of a 429 error; fall back to default."""
    try:
        text = str(err)
        m = re.search(r"retry in ([0-9.]+)s", text, re.IGNORECASE)
        if m:
            return float(m.group(1))
        m = re.search(r"retryDelay['\"]?:\s*['\"]?([0-9.]+)s", text)
        if m:
            return float(m.group(1))
    except Exception:
        pass
    return float(default)


# ============================ small helpers ===============================

def as_number(v):
    """Coerce a scanned value to float, or None if it isn't a clean number."""
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", "").replace("$", "").strip())
        except ValueError:
            return None
    return None


def sum_taxes(data):
    """Total tax from the scanned tax lines, or None if none present."""
    taxes = data.get("taxes") or []
    if not taxes:
        return None
    total = 0.0
    for t in taxes:
        amt = as_number(t.get("amount"))
        if amt is not None:
            total += amt
    return round(total, 2)


def calc_tax_from_rate(subtotal, province, rate_table):
    """
    When no tax was read off a CAD receipt, calculate it from the province rate:
    tax = pre-tax subtotal * rate. Returns (calc_tax, rate) or (None, None) if we
    lack a subtotal or a known province rate.
    """
    if subtotal is None or not province:
        return None, None
    entry = rate_table.get(str(province).upper())
    if not entry:
        return None, None
    rate = entry[1]
    return round(subtotal * rate, 2), rate


def to_yyyymmdd(iso_date):
    if not iso_date:
        return "UNKNOWNDATE"
    try:
        return datetime.datetime.strptime(iso_date, "%Y-%m-%d").strftime("%Y%m%d")
    except (ValueError, TypeError):
        return "UNKNOWNDATE"


def split_ymd(iso_date):
    """'2026-08-19' -> (2026, 8, 19); (None,None,None) on failure."""
    try:
        d = datetime.datetime.strptime(iso_date, "%Y-%m-%d")
        return d.year, d.month, d.day
    except (ValueError, TypeError):
        return None, None, None


# ============================ filename building ============================

def clean_company(name):
    """Keep spaces (your format allows them), strip only Windows-illegal chars."""
    if not name:
        return "UNKNOWN"
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name or "UNKNOWN"


def tax_slot(tax_value, is_cad):
    """CAD: tax amount found, formatted; else NA. Non-CAD: always NA."""
    if not is_cad or tax_value is None:
        return "NA"
    return "{:.2f}".format(tax_value) if tax_value else "NA"


def build_new_name(data, tax_value, is_cad, ext):
    date_part = to_yyyymmdd(data.get("transaction_date"))
    total = as_number(data.get("total"))
    total_part = "{:.2f}".format(total) if total is not None else "0.00"
    company_part = clean_company(data.get("company"))
    tax_part = tax_slot(tax_value, is_cad)
    return "{}_{}_{}_{}{}".format(date_part, total_part, company_part, tax_part, ext)


# ============================ STEP 2: Excel ===============================

def open_workbook():
    """Load the workbook from ROOT. Returns (wb, path) or (None, path) if missing."""
    path = os.path.join(ROOT, EXCEL_FILENAME)
    if not os.path.exists(path):
        return None, path
    return openpyxl.load_workbook(path), path


def ensure_tax_rates_sheet(wb):
    """Create + seed CAD_Tax_Rates once if absent. Return {code: (name, rate)}."""
    if TAX_RATES_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(TAX_RATES_SHEET)
        ws["A1"] = "Code"
        ws["B1"] = "Province"
        ws["C1"] = "Rate (decimal, 0.13 = 13%)"
        r = 2
        for code, (pname, rate) in sorted(DEFAULT_CAD_TAX_RATES.items()):
            ws.cell(row=r, column=1, value=code)
            ws.cell(row=r, column=2, value=pname)
            ws.cell(row=r, column=3, value=rate)
            r += 1

    ws = wb[TAX_RATES_SHEET]
    table = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        code = str(row[0]).upper().strip()
        rate = as_number(row[2]) if len(row) > 2 else None
        if code and rate is not None:
            table[code] = (row[1], rate)
    return table


def next_empty_row(ws):
    """First empty row at or after FIRST_DATA_ROW, keyed off column A."""
    r = FIRST_DATA_ROW
    while ws.cell(row=r, column=1).value not in (None, ""):
        r += 1
    return r


def find_matching_row(ws, date_str, total):
    """
    Return the row number of an existing entry whose TransactionDate (col B) AND
    TotalAmount (col G) both match, so a re-scan of the same receipt overwrites
    that row instead of appending a duplicate. Returns None when there's no match
    or the key is incomplete (no usable date, or no total).

    Note: the key is date+total only. Two genuinely different receipts with the
    same date and identical total (e.g. two same-day fill-ups of the exact same
    amount) would collide - the second would overwrite the first. Rare, but real.
    """
    if not date_str or date_str == "UNKNOWNDATE" or total is None:
        return None
    r = FIRST_DATA_ROW
    while ws.cell(row=r, column=1).value not in (None, ""):
        b = ws.cell(row=r, column=2).value
        g = ws.cell(row=r, column=7).value
        if (b is not None and str(b) == str(date_str)
                and isinstance(g, (int, float)) and abs(float(g) - total) < 0.005):
            return r
        r += 1
    return None


def _renamed_from_hyperlink(cell_value):
    """Pull the display filename out of a =HYPERLINK("path","name") formula."""
    if not isinstance(cell_value, str):
        return None
    quoted = re.findall(r'"((?:[^"]|"")*)"', cell_value)
    if len(quoted) >= 2:
        return quoted[1].replace('""', '"')          # display name
    if len(quoted) == 1:
        return os.path.basename(quoted[0].replace('""', '"'))
    return None


def write_row(ws, r, data, region_code, amount, tax_value, tip_value, total,
              renamed_name, orig_name, region_folder, tax_calculated=False,
              needs_review=False):
    """
    Write one receipt's row. Column layout (Tip added at F, everything after it
    shifted right by one):
      A DateProcessed  B TransactionDate  C CompanyName  D Amount  E Tax
      F Tip  G TotalAmount  H CreditCardDigits  I OrigFilename  J Country
      K ProvState  L RenamedFileName(link)  M Year  N Month  O Day  P TaxRate
    D Amount is the Level 1 DERIVED pre-tax amount (Total - Tax [- Tip]), not the
    receipt's own subtotal label. TaxRate (P) is a live percent formula = Tax /
    Amount (E/D), i.e. the derived rate Level 2 checks against the province.
    A row Level 2 flags as CHECK gets its Amount cell filled red (needs_review).
    """
    year, month, day = split_ymd(data.get("transaction_date"))

    ws.cell(row=r, column=1, value=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))  # A DateProcessed
    ws.cell(row=r, column=2, value=to_yyyymmdd(data.get("transaction_date")))               # B TransactionDate
    ws.cell(row=r, column=3, value=data.get("company"))                                     # C CompanyName
    dcell = ws.cell(row=r, column=4, value=amount)                                           # D Amount (derived)
    # Set the fill explicitly every time (not just when flagged) so overwriting a
    # row that used to be CHECK but is now clean clears the old red.
    dcell.fill = NEEDS_REVIEW_FILL if needs_review else PatternFill(fill_type=None)
    ecell = ws.cell(row=r, column=5, value=tax_value)                                        # E Tax
    ecell.fill = CALC_TAX_FILL if tax_calculated else PatternFill(fill_type=None)
    ws.cell(row=r, column=6, value=tip_value)                                               # F Tip
    ws.cell(row=r, column=7, value=total)                                                   # G TotalAmount
    ws.cell(row=r, column=8, value=data.get("card_last4"))                                  # H CreditCardDigits
    ws.cell(row=r, column=9, value=orig_name)                                               # I OrigFilename
    ws.cell(row=r, column=10, value=region_code)                                            # J Country
    ws.cell(row=r, column=11, value=data.get("province"))                                   # K ProvState
    # L RenamedFileName: clickable relative hyperlink to the filed image.
    # Path is relative to the workbook (which sits in ROOT), so the link survives
    # the whole receipts tree being moved/copied elsewhere.
    _rel = "{}\\{}\\{}".format(region_folder, SUBDIR_RENAMED, renamed_name)
    _esc = lambda s: str(s).replace('"', '""')
    _l = ws.cell(row=r, column=12,
                 value='=HYPERLINK("{}","{}")'.format(_esc(_rel), _esc(renamed_name)))
    _l.font = Font(color="0563C1", underline="single")
    ws.cell(row=r, column=13, value=year)                                                   # M Year
    ws.cell(row=r, column=14, value=month)                                                  # N Month
    ws.cell(row=r, column=15, value=day)                                                    # O Day

    # P TaxRate:
    # CAD -> live formula = Tax / Amount (derived), percent-formatted.
    # USD -> static state rate from the lookup table (informational only, no checking).
    # Other -> leave blank.
    p = ws.cell(row=r, column=16)
    if region_code == "CAD":
        p.value = "=IFERROR(E{r}/D{r},\"\")".format(r=r)
        p.number_format = "0.0%"
    elif region_code == "USD":
        state = str(data.get("province") or "").upper()
        entry = DEFAULT_USD_TAX_RATES.get(state)
        if entry:
            p.value = entry[1]
            p.number_format = "0.0%"


def derive_amount(subtotal, tax_value, tip_value, total, tax_calculated):
    """
    LEVEL 1: work out the pre-tax Amount without trusting the receipt's own
    "subtotal" label, which may be a true pre-tax figure, a tax-inclusive total,
    or one line of a multi-transaction slip.

    Instead we lean on the two numbers that read reliably - the grand total and
    the tax amount - and derive:
        no tip:      Amount = Total - Tax
        tip present: Amount = Total - Tax - Tip   (tip removed first, so a
                                                    tipped meal doesn't inflate
                                                    the pre-tax amount)

    Returns (amount, kind) where kind is:
      "derived"   - computed from total/tax as above (the normal, trusted path)
      "estimated" - tax itself was calculated from the province rate (no read
                    tax to derive from); keep the scanned subtotal as-is
      "fallback"  - total or tax missing, so we couldn't derive; used the
                    scanned subtotal and the row is marked for review
    """
    # When tax was calculated from the province rate (not read), there is no
    # reliable read tax to derive an Amount from. Keep the scanned subtotal; the
    # calculated tax is already the amber-marked estimate. Behaviour unchanged.
    if tax_calculated:
        return subtotal, "estimated"

    if total is not None and tax_value is not None:
        tip = tip_value or 0
        return round(total - tax_value - tip, 2), "derived"

    # Couldn't derive: fall back to whatever subtotal we read, flag it downstream.
    return subtotal, "fallback"


def evaluate_row(amount, subtotal, tax_value, tip_value, total,
                 region_code, province, rate_table, kind):
    """
    LEVEL 2: give each row a verdict. Returns (flags, needs_review).

    The old subtotal+tax+tip==total check is gone on purpose: once Amount is
    derived as Total-Tax-Tip it balances by construction and tells us nothing.
    Two signals replace it, and both are receipt-type agnostic (gas, restaurant,
    retail all go through the same logic):
      1. a rate CEILING (CAD): the effective rate Tax/Amount must not EXCEED the
         province's full rate. Above it = tax over-read. Below it is fine, because
         general receipts legitimately tax below the headline rate (exempt items).
      2. a cross-check of the receipt's own printed subtotal against the derived
         Amount - the real correctness check for anything that prints a subtotal
         (most general receipts): a misread tax or tip makes the two disagree.
    Either firing marks the row CHECK.
    """
    flags = []
    needs_review = False

    # A row we couldn't derive is unverified by definition.
    if kind == "fallback":
        flags.append("could not derive Amount (missing total or tax); used scanned subtotal")
        needs_review = True

    # Signal 1 - rate CEILING (CAD). The effective tax rate can't legitimately
    # exceed the province's full rate, so a derived rate ABOVE it means the tax
    # was over-read (or the Amount is wrong) - flag that. A rate BELOW the full
    # rate is NORMAL on general receipts: zero-rated groceries and other mixed or
    # exempt items are taxed at less than the headline rate, so a low effective
    # rate is expected, not an error. We therefore do NOT flag low rates - doing
    # so would red-flag every mixed-basket receipt. (The trade-off: a fuel receipt
    # that under-reads a tax line, e.g. a gas+wash combo dropping one HST line,
    # comes in low and won't rate-flag - that stays a by-hand catch. Fuel receipts
    # don't print a subtotal, so the cross-check below can't cover them either.)
    # If we can't determine the province's rate, skip this check rather than guess.
    if region_code == "CAD" and amount and tax_value is not None and province:
        entry = rate_table.get(str(province).upper())
        if entry:
            expected = entry[1]
            derived = tax_value / amount
            if derived > expected + RATE_TOLERANCE:
                flags.append(
                    "CHECK: rate {:.1%} exceeds {} rate {:.1%} (tax over-read?)".format(
                        derived, str(province).upper(), expected)
                )
                needs_review = True

    # Signal 2 - cross-check the receipt's printed subtotal against the derived
    # Amount. Only meaningful on a "derived" row that also has a real pre-tax
    # subtotal: one that is present AND clearly less than the total. A "subtotal"
    # that equals the total is just the total grabbed off a pump receipt ("Fuel
    # sales"), not a pre-tax line, so we skip it to avoid false flags.
    if kind == "derived" and subtotal is not None and total is not None:
        is_real_subtotal = subtotal < (total - CROSS_CHECK_TOLERANCE)
        if is_real_subtotal and abs(subtotal - amount) > CROSS_CHECK_TOLERANCE:
            flags.append(
                "CHECK: printed subtotal {:.2f} != derived amount {:.2f} "
                "(possible multi-item slip or bad tax/tip read)".format(subtotal, amount)
            )
            needs_review = True

    return flags, needs_review


# ============================ region selection ============================

def discover_regions(root):
    out = []
    for name in sorted(os.listdir(root)):
        if os.path.isdir(os.path.join(root, name)) and re.match(REGION_FOLDER_PATTERN, name):
            out.append(name)
    return out


def match_region(regions, wanted):
    """
    Resolve a --region value to a folder name. Accepts the number ("1"), the code
    ("cad") or the full folder name ("1_CAD"). Returns None if nothing matches.
    """
    wanted = str(wanted).strip().lower()
    for folder in regions:
        num, code = folder.split("_", 1)
        if wanted in (folder.lower(), num, code.lower()):
            return folder
    return None


def choose_region(regions):
    """
    STEP 0. Show the numbered regions and ask the user to type one number.
    Returns the chosen folder name, or None if the user quits.
    """
    by_number = {}
    for folder in regions:
        num_str = folder.split("_", 1)[0]
        try:
            by_number[int(num_str)] = folder
        except ValueError:
            continue

    print("Which currency do you want to process?")
    for num in sorted(by_number):
        code = by_number[num].split("_", 1)[1]
        print("  {} = {}".format(num, code))
    print("  q = quit")

    while True:
        raw = input("\nEnter the number: ").strip().lower()
        if raw in ("q", ""):
            return None
        try:
            choice = int(raw)
        except ValueError:
            print("  '{}' is not a number. Try again.".format(raw))
            continue
        if choice in by_number:
            return by_number[choice]
        print("  {} is not one of the options. Try again.".format(choice))


# ============================ region walk =================================

def unique_dest(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while os.path.exists("{}_({}){}".format(base, n, ext)):
        n += 1
    return "{}_({}){}".format(base, n, ext)


def process_region(root, region_folder, wb, wb_path, rate_table):
    region_code = region_folder.split("_", 1)[1].upper()
    is_cad = region_code == "CAD"

    orig_dir = os.path.join(root, region_folder, SUBDIR_ORIG)
    processed_dir = os.path.join(root, region_folder, SUBDIR_PROCESSED)
    renamed_dir = os.path.join(root, region_folder, SUBDIR_RENAMED)

    if not os.path.isdir(orig_dir):
        print("  {}: no {} folder, skipping.".format(region_folder, SUBDIR_ORIG))
        return 0

    # the sheet for this region must already exist in the workbook (live runs)
    ws = None
    if not DRY_RUN:
        if region_folder not in wb.sheetnames:
            print("  No sheet named '{}' in the workbook, skipping.".format(region_folder))
            return 0
        ws = wb[region_folder]

    files = [f for f in sorted(os.listdir(orig_dir))
             if os.path.splitext(f)[1].lower() in SUPPORTED_EXTS]
    if not files:
        print("  {} ({}): nothing to process.".format(region_folder, region_code))
        return 0

    print("  {} ({}): {} file(s)".format(region_folder, region_code, len(files)))
    count = 0

    for idx, fname in enumerate(files):
        src = os.path.join(orig_dir, fname)
        ext = os.path.splitext(fname)[1].lower()
        print("    scanning: {}".format(fname))

        try:
            data = scan(src, region_code)
        except (ClientError, ServerError, OSError) as e:
            # One bad file or a network blip shouldn't end the whole batch.
            print("      scan error ({}: {}), leaving in place.".format(type(e).__name__, e))
            data = None
        if data is None:
            print("      scan failed, leaving in place.")
            # still pace ourselves before the next scan
            if not DRY_RUN and SECONDS_BETWEEN_SCANS and idx < len(files) - 1:
                time.sleep(SECONDS_BETWEEN_SCANS)
            continue

        subtotal = as_number(data.get("subtotal"))
        total = as_number(data.get("total"))
        tax_value = sum_taxes(data)
        tip_value = as_number(data.get("tip"))

        # If a CAD receipt had no readable tax, calculate it from the province rate
        # (Amount * rate) and mark it as calculated (coloured cell) downstream.
        tax_calculated = False
        if is_cad and tax_value is None:
            calc, _rate = calc_tax_from_rate(subtotal, data.get("province"), rate_table)
            if calc is not None:
                tax_value = calc
                tax_calculated = True
                print("      tax not on receipt; calculated {:.2f} from {} rate".format(
                    calc, str(data.get("province")).upper()))

        new_name = build_new_name(data, tax_value, is_cad, ext)

        if is_cad:
            # CAD: Level 1 derive Amount, Level 2 evaluate and flag.
            amount, kind = derive_amount(subtotal, tax_value, tip_value, total, tax_calculated)
            flags, needs_review = evaluate_row(amount, subtotal, tax_value, tip_value, total,
                                               region_code, data.get("province"), rate_table, kind)
            for f in flags:
                print("      FLAG: {}".format(f))
        else:
            # Non-CAD: just log what the scanner read. No derivation, no flags,
            # no fills. If a field is missing it goes in as None (blank cell).
            amount = subtotal
            kind = "as read"
            needs_review = False

        if DRY_RUN:
            print("      would rename -> {}".format(new_name))
            print("      would write row: amount({})={} tax={} tip={} total={} card={}".format(
                kind, amount, tax_value, tip_value, total, data.get("card_last4")))
            count += 1
            continue

        os.makedirs(renamed_dir, exist_ok=True)

        # Already logged? Same TransactionDate + TotalAmount -> overwrite that row
        # in place (a re-scan updates the entry instead of adding a duplicate).
        match_row = find_matching_row(ws, to_yyyymmdd(data.get("transaction_date")), total)

        if match_row is not None:
            r = match_row
            # Remove the image this row previously pointed at (it's superseded),
            # then place the new one at its natural name, overwriting on collision -
            # no unique_dest _(2), since we're replacing, not adding.
            old_name = _renamed_from_hyperlink(ws.cell(row=r, column=12).value)
            for p in filter(None, [old_name, new_name]):
                fp = os.path.join(renamed_dir, p)
                if os.path.exists(fp):
                    try:
                        os.remove(fp)
                    except OSError:
                        pass
            dest = os.path.join(renamed_dir, new_name)
            final_name = os.path.basename(dest)
            print("      already logged (same date+total) -> overwriting row {}".format(r))
        else:
            r = next_empty_row(ws)
            # unique_dest may add _(2) on a name clash so the hyperlink points at
            # the real filename.
            dest = unique_dest(os.path.join(renamed_dir, new_name))
            final_name = os.path.basename(dest)

        # STEP 2: write the row (append or overwrite), then save as a checkpoint
        write_row(ws, r, data, region_code, amount, tax_value, tip_value, total,
                  final_name, fname, region_folder, tax_calculated, needs_review)
        wb.save(wb_path)

        # STEP 3: copy original (keep its name) into ProcessedReceipts
        os.makedirs(processed_dir, exist_ok=True)
        shutil.copy2(src, os.path.join(processed_dir, fname))

        # STEP 4: move + rename into ProcessedReceiptsRenamed
        shutil.move(src, dest)

        print("      done -> {}".format(os.path.basename(dest)))
        count += 1

        # pace ourselves so we stay under the per-minute limit (not after the last one)
        if SECONDS_BETWEEN_SCANS and idx < len(files) - 1:
            time.sleep(SECONDS_BETWEEN_SCANS)

    return count


def parse_args(argv=None):
    ap = argparse.ArgumentParser(
        description="Scan receipt images with Gemini, log them to Excel and file them by name.")
    ap.add_argument("--region", help="currency to process, as a number (1), code (CAD) or folder (1_CAD); skips the menu")
    ap.add_argument("--dry-run", action="store_true", help="preview only: scan and report, but write and move nothing")
    ap.add_argument("--root", help="receipts folder (default: RECEIPTS_ROOT env var, else %s)" % ROOT)
    ap.add_argument("--excel", help="workbook filename inside the root (default: %s)" % EXCEL_FILENAME)
    return ap.parse_args(argv)


def workbook_is_locked(path):
    """On Windows, a workbook open in Excel can't be opened for writing."""
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True


def main(argv=None):
    global ROOT, EXCEL_FILENAME, DRY_RUN
    args = parse_args(argv)
    if args.root:
        ROOT = args.root
    if args.excel:
        EXCEL_FILENAME = args.excel
    if args.dry_run:
        DRY_RUN = True

    print("Receipt walker")
    print("Root: {}".format(ROOT))
    print("Mode: {}\n".format("PREVIEW (no files touched)" if DRY_RUN else "LIVE"))

    if not os.path.isdir(ROOT):
        print("Root does not exist. Pass --root, set RECEIPTS_ROOT, or run folder_setup.py first.")
        return

    regions = discover_regions(ROOT)
    if not regions:
        print("No region folders found (expected names like 1_CAD, 2_USD).")
        return

    # load workbook (needed for the real run; preview skips it)
    wb, wb_path = (None, os.path.join(ROOT, EXCEL_FILENAME))
    rate_table = dict((k, v) for k, v in DEFAULT_CAD_TAX_RATES.items())
    if not DRY_RUN:
        if workbook_is_locked(wb_path):
            print("Cannot write to {}.\nClose it in Excel and run again.".format(wb_path))
            return
        wb, wb_path = open_workbook()
        if wb is None:
            print("Workbook not found at {}. Put it in ROOT or pass --excel.".format(wb_path))
            return
        rate_table = ensure_tax_rates_sheet(wb)
        wb.save(wb_path)  # persist the seeded rate sheet if it was just created

    # STEP 0: which currency to process
    if args.region:
        chosen = match_region(regions, args.region)
        if chosen is None:
            print("No region matches '{}'. Found: {}".format(args.region, ", ".join(regions)))
            return
    else:
        chosen = choose_region(regions)
    if chosen is None:
        print("Nothing selected. Exiting.")
        return

    code = chosen.split("_", 1)[1]
    print("\nProcessing {} ({}) only.\n".format(chosen, code))

    total = process_region(ROOT, chosen, wb, wb_path, rate_table)

    print("\nFinished. {} receipt(s) {}.".format(
        total, "previewed" if DRY_RUN else "processed"))
    if DRY_RUN:
        print("This was a PREVIEW. Run again without --dry-run to write Excel + copy + rename.")


if __name__ == "__main__":
    main()
