import csv
import io
import os
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
 
import requests
 
# ------------------------------------------------------------------ CONFIG ----
EINS = [
    "843118937",   # <-- replace with your target EIN(s). Leading zeros OK. Current is Student Orgs of HLS. 
]
 
OUTPUT_CSV          = "irs990_part1.csv"
 
# Google Sheets output (set WRITE_TO_SHEETS=False to only produce the CSV)
WRITE_TO_SHEETS     = False
GOOGLE_SHEET_NAME   = "IRS 990 Part I"          # an existing Sheet you own/shared
GOOGLE_WORKSHEET    = "data"
GSPREAD_CREDS_JSON  = "service_account.json"     # path to your service-account key
 
# e-file XML resolution
USE_IRS_INDEX       = False                      # True = full historical XML (slow first run)
IRS_INDEX_YEARS     = list(range(2017, 2027))    # processing years to pull object_ids from
CACHE_DIR           = "irs990_cache"
 
# XML mirrors, tried in order. {oid} = object_id.
XML_SOURCES = [
    "https://opendata.grantseeker.io/data/{oid}_public.xml",
    "https://s3.amazonaws.com/irs-form-990/{oid}_public.xml",   # legacy; pre-2021 only
]
 
REQUEST_PAUSE_SEC   = 0.3                         # be kind to ProPublica/mirrors
HEADERS             = {"User-Agent": "part1-scraper/1.0 (research; contact: you@example.edu)"}
# ------------------------------------------------------------------------------
 
 
# ============================ small HTTP helper ===============================
def http_get(url, *, tries=3, expect="json"):
    """GET with simple backoff on 429/5xx. Returns parsed json / text / bytes / None."""
    for attempt in range(1, tries + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=60)
        except requests.RequestException as e:
            print(f"    ! network error ({e}); retry {attempt}/{tries}", file=sys.stderr)
            time.sleep(1.5 * attempt)
            continue
        if r.status_code == 200:
            time.sleep(REQUEST_PAUSE_SEC)
            if expect == "json":  return r.json()
            if expect == "bytes": return r.content
            return r.text
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(1.5 * attempt)
            continue
        return None  # 404 etc.
    return None
 
 
# ============================ ProPublica API ==================================
def fetch_org(ein):
    """Return (organization dict, list of filing dicts) or (None, [])."""
    ein_int = str(int(ein))  # ProPublica wants the integer form (no leading zeros)
    url = f"https://projects.propublica.org/nonprofits/api/v2/organizations/{ein_int}.json"
    data = http_get(url, expect="json")
    if not data:
        return None, []
    return data.get("organization", {}), data.get("filings_with_data", [])
 
 
def _first(d, keys):
    for k in keys:
        if d.get(k) is not None:
            return d[k]
    return None
 
 
def api_row_from_filing(f):
    """Pull the variables ProPublica actually provides (current/EOY only)."""
    rev = _first(f, ["totrevenue"])
    exp = _first(f, ["totfuncexpns"])
    rle = (rev - exp) if isinstance(rev, (int, float)) and isinstance(exp, (int, float)) else None
    return {
        "total_revenue":          rev,
        "total_expenses":         exp,
        "revenue_less_expenses":  rle,
        "total_assets_eoy":       _first(f, ["totassetsend"]),
        "total_liabilities_eoy":  _first(f, ["totliabend"]),
        # full-990 uses "totnetassetend"; 990-EZ uses "totnetassetsend"
        "net_assets_eoy":         _first(f, ["totnetassetend", "totnetassetsend"]),
        "_ubi_flag":              _first(f, ["unrelbusinccd", "unrelbusincd"]),
    }
 
 
# ============================ e-file XML parsing ==============================
def _local(tag):
    return tag.split("}", 1)[1] if "}" in tag else tag
 
 
def _find_irs990(root):
    for el in root.iter():
        if _local(el.tag) == "IRS990":
            return el
    return None
 
 
def _first_text(scope, names):
    """First descendant (in candidate-name priority order) with non-empty text."""
    if scope is None:
        return None
    found = {}
    for el in scope.iter():
        ln = _local(el.tag)
        if ln in names and ln not in found and el.text and el.text.strip():
            found[ln] = el.text.strip()
    for n in names:
        if n in found:
            return found[n]
    return None
 
 
# Part I element names (MeF schema). Lists allow tolerance across schema versions.
# All of these names are unique to Part I, so descendant search is safe.
PART1_XML = {
    "voting_members":             ["VotingMembersGoverningBodyCnt"],
    "independent_voting_members": ["VotingMembersIndependentCnt"],
    "total_employees":            ["TotalEmployeeCnt"],
    "total_volunteers":           ["TotalVolunteersCnt"],
    "ubi_gross_revenue":          ["TotalGrossUBIAmt"],
    "ubi_net_taxable_income":     ["NetUnrelatedBusTxblIncmAmt", "NetUnrelatedBusinessTxblIncmAmt"],
    "total_revenue":             ["CYTotalRevenueAmt"],
    "total_revenue_py":          ["PYTotalRevenueAmt"],
    "total_expenses":            ["CYTotalExpensesAmt"],
    "total_expenses_py":         ["PYTotalExpensesAmt"],
    "revenue_less_expenses":     ["CYRevenuesLessExpensesAmt"],
    "revenue_less_expenses_py":  ["PYRevenuesLessExpensesAmt"],
    "total_assets_boy":          ["TotalAssetsBOYAmt"],
    "total_assets_eoy":          ["TotalAssetsEOYAmt"],
    "total_liabilities_boy":     ["TotalLiabilitiesBOYAmt"],
    "total_liabilities_eoy":     ["TotalLiabilitiesEOYAmt"],
    "net_assets_boy":            ["NetAssetsOrFundBalancesBOYAmt"],
    "net_assets_eoy":            ["NetAssetsOrFundBalancesEOYAmt"],
}
 
_NUMERIC = set(PART1_XML) - set()  # all Part I fields here are numeric
 
 
def _to_num(v):
    if v is None:
        return None
    try:
        return int(v)
    except ValueError:
        try:
            return float(v)
        except ValueError:
            return v
 
 
def fetch_and_parse_xml(object_id):
    """Fetch e-file XML by object_id from the mirrors and parse Part I. Returns dict or None."""
    for tmpl in XML_SOURCES:
        raw = http_get(tmpl.format(oid=object_id), expect="bytes")
        if not raw:
            continue
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            continue
        irs990 = _find_irs990(root)
        if irs990 is None:
            continue
        out = {k: _to_num(_first_text(irs990, names)) for k, names in PART1_XML.items()}
        return out
    return None
 
 
# ============================ object_id resolution ============================
def _index_cache_path(year):
    return os.path.join(CACHE_DIR, f"index_{year}.csv")
 
 
def download_irs_index(year):
    """
    Download & cache the IRS e-file index for a processing YEAR.
    The IRS distributes these as TEOS ZIPs; the index CSV inside maps EIN->OBJECT_ID.
    NOTE: IRS has moved these URLs over time. If this 404s, grab the 'Index file (CSV)'
    link from https://www.irs.gov/charities-non-profits/form-990-series-downloads
    and drop it at irs990_cache/index_<year>.csv yourself.
    """
    path = _index_cache_path(year)
    if os.path.exists(path):
        return path
    os.makedirs(CACHE_DIR, exist_ok=True)
    candidates = [
        f"https://apps.irs.gov/pub/epostcard/990/xml/{year}/index_{year}.csv",
        f"https://apps.irs.gov/pub/epostcard/990/xml/{year}/{year}_TEOS_XML_CT.zip",
    ]
    for url in candidates:
        blob = http_get(url, expect="bytes")
        if not blob:
            continue
        if url.endswith(".zip"):
            try:
                zf = zipfile.ZipFile(io.BytesIO(blob))
                name = next((n for n in zf.namelist() if n.lower().endswith(".csv")), None)
                if not name:
                    continue
                blob = zf.read(name)
            except zipfile.BadZipFile:
                continue
        with open(path, "wb") as fh:
            fh.write(blob)
        return path
    print(f"    ! could not auto-download IRS index for {year} "
          f"(place it at {path} manually)", file=sys.stderr)
    return None
 
 
def build_objectid_map(eins):
    """{ein_int_str: {tax_year_or_period: object_id}} from cached IRS indexes."""
    wanted = {str(int(e)) for e in eins}
    mapping = {e: {} for e in wanted}
    for year in IRS_INDEX_YEARS:
        path = download_irs_index(year)
        if not path:
            continue
        with open(path, newline="", encoding="utf-8", errors="replace") as fh:
            reader = csv.DictReader(fh)
            # tolerate column-name variants across IRS index vintages
            cols = {c.lower(): c for c in (reader.fieldnames or [])}
            ein_c = cols.get("ein")
            oid_c = cols.get("object_id") or cols.get("objectid")
            tp_c  = cols.get("tax_period") or cols.get("taxperiod")
            if not (ein_c and oid_c):
                print(f"    ! unexpected index columns in {path}: {reader.fieldnames}",
                      file=sys.stderr)
                continue
            for row in reader:
                e = str(int(row[ein_c])) if row.get(ein_c, "").strip().isdigit() else None
                if e in wanted:
                    key = (row.get(tp_c) or "").strip() or row[oid_c]
                    mapping[e][key] = row[oid_c].strip()
    return mapping
 
 
# ============================ output: Google Sheets ===========================
COLUMNS = [
    "ein", "org_name", "fiscal_year", "tax_period_end", "form_type", "source",
    "total_revenue", "total_revenue_py",
    "total_expenses", "total_expenses_py",
    "revenue_less_expenses", "revenue_less_expenses_py",
    "total_assets_boy", "total_assets_eoy",
    "total_liabilities_boy", "total_liabilities_eoy",
    "net_assets_boy", "net_assets_eoy",
    "voting_members", "independent_voting_members",
    "total_employees", "total_volunteers",
    "ubi_gross_revenue", "ubi_net_taxable_income",
    "had_ubi_flag", "xml_object_id", "pdf_url",
]
 
 
def write_csv(rows, path):
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=COLUMNS)
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c) for c in COLUMNS})
    print(f"  wrote {len(rows)} rows -> {path}")
 
 
def write_sheets(rows):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("  ! gspread not installed; skipping Sheets. "
              "pip3 install --break-system-packages gspread google-auth", file=sys.stderr)
        return
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(GSPREAD_CREDS_JSON, scopes=scopes)
    gc = gspread.authorize(creds)
    try:
        sh = gc.open(GOOGLE_SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        sh = gc.create(GOOGLE_SHEET_NAME)
        print(f"  created new spreadsheet '{GOOGLE_SHEET_NAME}' "
              f"(share it from the service-account email if you can't see it)")
    try:
        ws = sh.worksheet(GOOGLE_WORKSHEET)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=GOOGLE_WORKSHEET, rows=2, cols=len(COLUMNS))
    table = [COLUMNS] + [[("" if r.get(c) is None else r.get(c)) for c in COLUMNS] for r in rows]
    ws.update(table, value_input_option="RAW")
    print(f"  wrote {len(rows)} rows -> Google Sheet '{GOOGLE_SHEET_NAME}' / '{GOOGLE_WORKSHEET}'")
 
 
# ============================ main pipeline ===================================
FORMTYPE = {0: "990", 1: "990-EZ", 2: "990-PF"}
 
 
def scrape(eins):
    oid_map = build_objectid_map(eins) if USE_IRS_INDEX else {}
    all_rows = []
 
    for ein in eins:
        org, filings = fetch_org(ein)
        if org is None:
            print(f"[{ein}] not found on ProPublica", file=sys.stderr)
            continue
        name = org.get("name", "")
        ein_int = str(int(ein))
        latest_oid = org.get("latest_object_id")
        latest_yr = max((f.get("tax_prd_yr") for f in filings), default=None)
        print(f"[{ein}] {name} — {len(filings)} filings with data")
 
        for f in filings:
            yr = f.get("tax_prd_yr")
            row = {c: None for c in COLUMNS}
            row.update({
                "ein": ein_int, "org_name": name, "fiscal_year": yr,
                "tax_period_end": f.get("tax_prd"),
                "form_type": FORMTYPE.get(f.get("formtype"), f.get("formtype")),
                "pdf_url": f.get("pdf_url"),
            })
 
            # 1) API-available figures (always)
            api = api_row_from_filing(f)
            row["had_ubi_flag"] = api.pop("_ubi_flag")
            for k, v in api.items():
                row[k] = v
            row["source"] = "api-only"
 
            # 2) resolve an object_id for this year, then overlay XML if we can
            oid = None
            if USE_IRS_INDEX:
                ymap = oid_map.get(ein_int, {})
                oid = next((o for key, o in ymap.items() if str(f.get("tax_prd")) in str(key)),
                           None) or ymap.get(str(f.get("tax_prd")))
            if oid is None and latest_oid and yr == latest_yr:
                oid = latest_oid   # ProPublica gives us the newest filing's id for free
 
            if oid:
                row["xml_object_id"] = oid
                xml = fetch_and_parse_xml(oid)
                if xml:
                    for k, v in xml.items():
                        if v is not None:
                            row[k] = v          # XML wins (full CY/PY + BOY/EOY)
                    row["source"] = "xml+api"
                    print(f"    {yr}: XML parsed (oid {oid})")
                else:
                    print(f"    {yr}: XML unavailable for oid {oid}; API-only")
 
            all_rows.append(row)
 
    all_rows.sort(key=lambda r: (r["ein"], -(r["fiscal_year"] or 0)))
    return all_rows
 
 
def main():
    if not EINS or EINS == ["042103580"]:
        print("Edit the EINS list at the top of the file first.", file=sys.stderr)
    rows = scrape(EINS)
    if not rows:
        print("No data.", file=sys.stderr)
        return
    write_csv(rows, OUTPUT_CSV)
    if WRITE_TO_SHEETS:
        write_sheets(rows)
 
 
if __name__ == "__main__":
    main()
