# IRS-990-Extractor
Pull key variables in Form 990s and output into csv. 
"""
irs990_part1_scraper.py
================================================================================
Generalized IRS Form 990 *Part I (Summary)* scraper -> Google Sheets (+ CSV).
 
The ProPublica Nonprofit Explorer API exposes the IRS SOI *financial* extract.
Even for full-990 filers, that extract contains ONLY:
    total revenue, total expenses, end-of-year assets/liabilities/net assets,
    and a Y/N "had unrelated business income" flag.
 
COVERAGE CAVEATS (read these)
-----------------------------
* XML exists only for ELECTRONICALLY filed returns. Paper filers and most
  pre-2017 returns have no XML — only a PDF. For those, columns that are
  XML-only stay blank and you fall back to OCR. 
* Without the IRS index (USE_IRS_INDEX = False), only the *latest* filing gets
  XML (via ProPublica's latest_object_id); all earlier years are API-only.
* The prior-year (PY) column is the filer's own restatement of the prior year.
  It can differ from that year's originally-filed CY figure (amendments,
  reclassifications).
* The IRS index URL/format and the XML mirror have shifted historically. Both
  are configurable below. 
 
DEPENDENCIES
------------
    pip3 install --break-system-packages requests gspread google-auth
 
USAGE
-----
    python3 irs990_part1_scraper.py
================================================================================
"""
