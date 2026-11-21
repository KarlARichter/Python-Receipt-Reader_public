r"""
scan_one.py

Scan a single receipt image and print what the model read, as JSON. Nothing is
written, moved or logged, so it is safe for checking a tricky receipt.

Uses the same scan() and prompt as process_receipts.py, so the output matches
what a real run would get.

Run:
  python scan_one.py "C:\path\to\receipt.jpg"
  python scan_one.py receipt.jpg --region EUR
"""

import argparse
import json
import os
import sys

from process_receipts import scan


def main():
    ap = argparse.ArgumentParser(description="Scan one receipt and print the result.")
    ap.add_argument("path", help="image or PDF to scan")
    ap.add_argument("--region", default="CAD",
                    help="CAD, USD, MEX or EUR. Only affects how dates are read (EUR is day-first). Default: CAD")
    args = ap.parse_args()

    if not os.path.isfile(args.path):
        sys.exit("File not found: {}".format(args.path))

    data = scan(args.path, args.region.upper())
    if data is not None:
        print(json.dumps(data, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
