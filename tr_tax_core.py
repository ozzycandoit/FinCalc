"""
tr_tax_core.py - calculation layer for the Streamlit dashboard.

Two entry points:

1) calculate_from_workbook(uploaded_file, evds_key)  -> THE FULL ENGINE.
   Runs the exact same code as the desktop tool (tax_tool.py): stocks AND
   options, Yi-UFE indexation, strict FIFO + same-day rule, USD/EUR,
   commissions, full-category netting, the official 2026 tariff and the
   March/July instalments. The web app and the workbook therefore produce
   identical numbers. Input is your Turkey_Tax_Tracker.xlsx.

2) calculate_turkish_taxes(uploaded_file) -> QUICK CHECK (stocks only).
   A lightweight per-asset FIFO on a flat broker statement, with NO Yi-UFE
   indexation and NO options. Handy for a fast sanity check; not a filing.
"""

import io
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET

import requests
from openpyxl import load_workbook

import tax_tool   # the full, verified engine

# Core functions and implementation would go here
