r"""
folder_setup.py

Run this ONCE to build the receipt folder structure. Safe to run again later;
it only creates what's missing and never touches existing files.

After running, you'll have:

    <ROOT>\
        1_CAD\
            OrigScannedReceipts\        <- you drop phone photos here
            ProcessedReceipts\          <- copies of originals land here
            ProcessedReceiptsRenamed\   <- renamed copies land here
        2_USD\
            (same three)
        3_MEX\  (same three)
        4_EUR\  (same three)
"""

import argparse
import os

# Where to build the tree. Same setting process_receipts.py uses: the
# RECEIPTS_ROOT environment variable, or pass --root.
ROOT = os.environ.get("RECEIPTS_ROOT", r"C:\RBT\Receipts")

# Region folders to create. Trim or add as you like (e.g. add "5_JPY" later).
# The names drive everything downstream.
REGIONS = ["1_CAD", "2_USD", "3_MEX", "4_EUR"]

# The three subfolders every region gets.
SUBFOLDERS = ["OrigScannedReceipts", "ProcessedReceipts", "ProcessedReceiptsRenamed"]


def main():
    global ROOT
    ap = argparse.ArgumentParser(description="Create the receipt folder tree.")
    ap.add_argument("--root", help="where to build it (default: %s)" % ROOT)
    args = ap.parse_args()
    if args.root:
        ROOT = args.root

    print("Building folder structure under:")
    print("  {}\n".format(ROOT))

    created = 0
    existed = 0

    for region in REGIONS:
        for sub in SUBFOLDERS:
            path = os.path.join(ROOT, region, sub)
            if os.path.isdir(path):
                existed += 1
                print("  exists : {}".format(path))
            else:
                os.makedirs(path, exist_ok=True)
                created += 1
                print("  created: {}".format(path))

    print("\nDone. {} folder(s) created, {} already there.".format(created, existed))
    print("Drop your phone photos into each region's OrigScannedReceipts folder.")


if __name__ == "__main__":
    main()