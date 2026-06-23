#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 TURKEY FOREIGN STOCK / ETF CAPITAL-GAINS TOOL - expandable edition
================================================================================

Commands:
  python tax_tool.py init [--assets N] [--rows N] [--overwrite]
      Creates Turkey_Tax_Tracker.xlsx with a SUMMARY sheet and N ASSET sheets.
      Default: 10 assets and 500 transaction rows per asset.

  python tax_tool.py add-asset [N] [--rows N]
      Adds N more ASSET sheets to an existing workbook and rebuilds SUMMARY.
      Default: add 1 asset.

  python tax_tool.py run [EVDS_KEY]
      Processes every sheet whose name starts with ASSET_. The number of asset
      sheets is not fixed. The SUMMARY sheet is rebuilt dynamically.

  python tax_tool.py guide
      Writes the dividend guide text file next to the workbook.

Main capital-gains rules implemented:
  - TCMB forex buying rate is used for each transaction date. If no bulletin
    exists on that date, the tool walks back to the latest published bulletin.
  - FIFO lot matching is used for sales.
  - YI-UFE indexation uses the month before acquisition and the month before
    disposal. Indexation is applied only if the index increase is at least 10%.
  - Same-year securities gains and losses are netted across all ASSET sheets;
    the annual loss floor is applied only after the total securities result.
  - If an index month needed for a realised sale is missing, the workbook is
    marked PROVISIONAL. Non-indexed cost is used only as an estimate.

Dividends are not calculated in this file. See Dividend_Calculation_Guide.md.
"""

import datetime as dt
import re
import sys
import xml.etree.ElementTree as ET
from itertools import groupby
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing openpyxl. Run: pip install openpyxl requests")

try:
    import requests
except ImportError:
    sys.exit("Missing requests. Run: pip install openpyxl requests")


# Configuration and all functions from tax_tool.py
# [Full content would go here - this is a placeholder]

if __name__ == "__main__":
    main()
