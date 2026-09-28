#!/usr/bin/env python3
"""
Travis County (Texas) Motivated Seller Lead Scraper
===================================================
Travis County's recorder is the County Clerk Web Search at tccsearch.org
(Aumentum Recorder / Harris Recording Solutions, anonymous, no login).
It is an ASP.NET WebForms app, so the search form is driven with
Playwright (headless chromium) rather than raw requests.

Sources
  1. tccsearch.org Real Estate index (anonymous):
       https://www.tccsearch.org/RealEstate/SearchEntry.aspx
     Flow per doc type (fresh browser context each): GET / -> click the
     disclaimer-acknowledge link -> GET SearchEntry.aspx -> check the
     doc-type checkbox (#cphNoMargin_f_dclDocType input[value=CODE]) ->
     type the Date Filed From/To (Infragistics masked inputs inside
     #cphNoMargin_f_ddcDateFiledFrom / ...To) -> click
     #cphNoMargin_SearchButtons2_btnSearch -> parse the results grid
     (20 rows/page; #OptionsBar1_imgNext paginates, URL gains ?pg=N).
     Row: instrument link (9 digits), Date Filed, Document Type, a names
     cell "[R] NAME (+) [E] NAME (+)", Legal Description, Status.
     [R] = creditor / trustee / filer, [E] = debtor / property owner --
     EXCEPT NOTICE OF SUBSTITUTE TRUSTEE SALE where the first [E] entry
     is the SALE DATE and the owner only appears on the detail page
     ("Sale Date/Owner" section), so FC rows get a capped detail-page
     pass to pull the owner + lender.
     Legal Description often embeds the property address as
     "... LOC 1421 DEXFORD DR AUSTIN TX 78753 ..." -> parsed out.
  2. Enrichment: TCAD parcels via Travis County GIS (public ArcGIS):
       https://gis.traviscountytx.gov/server1/rest/services/
         Boundaries_and_Jurisdictions/TCAD/MapServer/0/query
     py_owner_name = "LAST FIRST M" uppercase; py_address = ONE string
     ("2012 SALIENTE WAY CARLSBAD CA 92009") parsed from the END
     (zip -> state -> city-by-suffix); situs in structured fields
     (situs_street excludes prefix/suffix); value in market_value.

Run:
    python scraper/fetch.py                # default lookbacks
    python scraper/fetch.py --skip-parcel --skip-detail
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import logging
import os
import re
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
COUNTY = "Travis"
STATE = "TX"

REC_BASE = "https://www.tccsearch.org"
REC_ENTRY = f"{REC_BASE}/RealEstate/SearchEntry.aspx"
REC_RESULTS = f"{REC_BASE}/RealEstate/SearchResults.aspx"

PARCEL_API_URL = ("https://gis.traviscountytx.gov/server1/rest/services/"
                  "Boundaries_and_Jurisdictions/TCAD/MapServer/0/query")

LOOKBACK_DAYS = int(os.environ.get("LOOKBACK_DAYS", "7"))
LIEN_LOOKBACK_DAYS = 30      # liens / LP / FC notices trickle in slower
PRO_LOOKBACK_DAYS = 60       # probate/heirship: 60-day window
REQUEST_TIMEOUT = 45
ARCGIS_MAX_LOOKUPS = 1500
MAX_PAGES_PER_TYPE = 40      # 20 rows/page
MAX_DETAIL_VISITS = 150      # FC owner detail-page pulls per run

# Recorder document types -> (checkbox value, label, cat, cat_label, lookback)
REC_DOC_TYPES = [
    ("FORECLOSURE", "NOTICE OF SUBSTITUTE TRUSTEE SALE", "FC",
     "Notice of Trustee Sale", LIEN_LOOKBACK_DAYS),
    ("AJ",          "ABSTRACT OF JUDGMENT",  "JUD",  "Abstract of Judgment", LOOKBACK_DAYS),
    ("LIS PEND",    "LIS PENDENS",           "LP",   "Lis Pendens",          LIEN_LOOKBACK_DAYS),
    ("HEIRSHIP",    "AFFIDAVIT OF HEIRSHIP", "PRO",  "Affidavit of Heirship", PRO_LOOKBACK_DAYS),
    ("FED TAX",     "FEDERAL TAX LIENS AND NOTICES", "LIEN", "Federal Tax Lien", LIEN_LOOKBACK_DAYS),
    ("ST TAX LIEN", "STATE TAX LIEN",        "LIEN", "State Tax Lien",       LIEN_LOOKBACK_DAYS),
    ("ML",          "MECHANICS LIEN",        "LIEN", "Mechanic's Lien",      LOOKBACK_DAYS),
    ("HOSP LIEN",   "HOSPITAL LIEN",         "LIEN", "Hospital Lien",        LIEN_LOOKBACK_DAYS),
    ("CHILD SL",    "CHILD SUPPORT LIEN",    "LIEN", "Child Support Lien",   LIEN_LOOKBACK_DAYS),
    ("ASSESS LIEN", "ASSESSMENT LIEN",       "LIEN", "HOA/Assessment Lien",  LIEN_LOOKBACK_DAYS),
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("travis_scraper")

GHL_FIELDS = [
    "doc_num","doc_type","cat","cat_label","filed","owner","grantee",
    "amount","prop_address","prop_city","prop_state","prop_zip",
    "mail_address","mail_city","mail_state","mail_zip","legal","clerk_url","score","flags",
    "first_seen","status",
]
GHL_HEADERS = {f: f.replace("_", " ").title() for f in GHL_FIELDS}
GHL_HEADERS["first_seen"] = "Date Entered System"
GHL_HEADERS["status"] = "Status"

@dataclass
class LeadRecord:
    doc_num: str = ""
    doc_type: str = ""
    cat: str = ""
    cat_label: str = ""
    filed: str = ""
    owner: str = ""
    grantee: str = ""
    amount: float = 0.0
    legal: str = ""
    prop_address: str = ""
    prop_city: str = ""
    prop_state: str = STATE
    prop_zip: str = ""
    mail_address: str = ""
    mail_city: str = ""
    mail_state: str = STATE
    mail_zip: str = ""
    clerk_url: str = ""
    flags: list = field(default_factory=list)
    score: int = 0
    status: str = ""
    first_seen: str = ""
    rid: str = ""
    content_hash: str = ""

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def normalize_date(raw: str) -> str:
    raw = _norm_ws(raw).split(" ")[0]
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y", "%B %d, %Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return raw


def _norm_ws(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip())


def _arc_val(x) -> str:
    s = _norm_ws(x)
    return "" if s.upper() in ("NULL", "NONE") else s


def _sql_lit(s: str) -> str:
    return s.upper().replace("'", "''")


ENTITY_RE = re.compile(
    r"\b(LLC|L\.?L\.?C|INC|CORP|CORPORATION|COMPANY|CO|BANK|N\.?A|TRUST|TR|LP|L\.?P|LLP|"
    r"ASSOCIATION|ASSN|ASSOC|FUND|FUNDING|CREDIT UNION|CU|SYSTEM|COUNTY|CITY OF|"
    r"STATE OF|UNITED STATES|IRS|DEPARTMENT|DEPT|ISD|UNIVERSITY|COLLEGE|"
    r"HOSPITAL|MEDICAL|SERVICES|CAPITAL|MORTGAGE|FINANCIAL|HOLDINGS|HOA|"
    r"HOMEOWNERS|CONDOMINIUM|COMMUNITY|MANAGEMENT|PARTNERS|PROPERTIES|GROUP)\b", re.I)

DATE_TXT_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")
DOCNUM_RE = re.compile(r"^\d{9}$")


def _looks_like_entity(name: str) -> bool:
    return bool(ENTITY_RE.search(name or ""))


def _pick_owner(r_names: list, e_names: list) -> tuple:
    """[R] is the creditor/plaintiff/trustee, [E] is the debtor/owner.
    Prefer the non-entity side so reversed indexing still resolves the
    individual. Returns (owner, counterparty)."""
    e = next((n for n in e_names if n), "").strip()
    r = next((n for n in r_names if n), "").strip()
    e_ent, r_ent = _looks_like_entity(e), _looks_like_entity(r)
    if e and not e_ent:
        return e, r
    if r and not r_ent:
        return r, e
    return (e or r), (r if e else "")


def normalize_owner_for_parcel(name: str) -> str:
    """TCAD py_owner_name is 'LAST FIRST M' uppercase. Produce a
    'LAST FIRST' two-token prefix for LIKE matching."""
    if not name:
        return ""
    n = _norm_ws(name).upper()
    n = re.sub(r"\b(JR|SR|II|III|IV)\.?$", "", n).strip().rstrip(",")
    n = re.sub(r"\b(AKA|A/K/A|DBA|D/B/A|FKA|F/K/A)\b.*$", "", n).strip().rstrip(",")
    if _looks_like_entity(n):
        return n
    if "," in n:
        last, first = n.split(",", 1)
        first_tok = first.strip().split()
        return f"{last.strip()} {first_tok[0] if first_tok else ''}".strip()
    parts = n.split()
    if len(parts) >= 2:
        return f"{parts[0]} {parts[1]}"
    return n

# ---------------------------------------------------------------------------
# Address parsing (legal-description "LOC ..." + TCAD single-string mailing)
# ---------------------------------------------------------------------------
SUFFIXES = ("(?:RD|ROAD|ST|STREET|DR|DRIVE|LN|LANE|CT|COURT|CIR|CIRCLE|TRL|TRAIL|"
            "AVE|AVENUE|BLVD|BOULEVARD|WAY|PASS|PATH|LOOP|RUN|CV|COVE|BND|BEND|"
            "PKWY|PARKWAY|TER|TERRACE|PL|PLACE|PT|POINT|XING|CROSSING|HWY|"
            "HIGHWAY|EXPY|FM|SQ|SQUARE|VW|VIEW|VIS|VISTA|HL|HILL|HOLW|HOLLOW|"
            "RDG|RIDGE|MDW|MEADOW|GLN|GLEN|CRK|CREEK|TRCE|TRACE|WALK|ROW|CLF|"
            "CLIFF|SPGS?|SPRINGS?|EST|ESTATES?|GDNS?|GARDENS?|PLZ|PLAZA|OAKS?|"
            "MNR|MANOR|LNDG|LANDING|CYN|CANYON|SKWY|CSWY|OVERLOOK|BLF|BLUFF)")

# "... LOC 20204 KEARNEY HILL RD PFLUGERVILLE TX 78660 PDV 4"
LOC_RE = re.compile(
    r"\bLOC\s+(\d{1,6}\s+.+?\s+TX\s+(\d{5})(?:-\d{4})?)(?=\s|$)", re.I)

TRAVIS_CITIES = [
    "WEST LAKE HILLS", "DRIPPING SPRINGS", "MUSTANG RIDGE", "POINT VENTURE",
    "ROLLINGWOOD", "PFLUGERVILLE", "CEDAR PARK", "ROUND ROCK", "LAGO VISTA",
    "JONESTOWN", "BRIARCLIFF", "WEBBERVILLE", "DEL VALLE", "CREEDMOOR",
    "SPICEWOOD", "THE HILLS", "BEE CAVE", "GARFIELD", "LEANDER", "VOLENTE",
    "LAKEWAY", "MCNEIL", "AUSTIN", "MANOR", "ELGIN", "MANCHACA", "BUDA",
    "KYLE", "COUPLAND", "HUTTO",
]


def parse_us_address_tail(s: str) -> tuple:
    """Parse 'STREET... CITY ST ZIP' from the END. Returns
    (street, city, state, zip) best-effort."""
    s = _norm_ws(s).upper().rstrip(",.")
    if not s:
        return "", "", "", ""
    m = re.search(r"\b(\d{5})(?:-\d{4})?$", s)
    zp = m.group(1) if m else ""
    if m:
        s = s[:m.start()].strip().rstrip(",")
    st = ""
    m = re.search(r"\b([A-Z]{2})$", s)
    if m:
        st = m.group(1)
        s = s[:m.start()].strip().rstrip(",")
    city = ""
    for c in sorted(TRAVIS_CITIES, key=len, reverse=True):
        if s.endswith(" " + c):
            city, s = c, s[: -len(c) - 1].strip().rstrip(",")
            break
    if not city:
        sfx = re.compile(r"\b" + SUFFIXES + r"\.?\b", re.I)
        last = None
        for mm in sfx.finditer(s):
            last = mm
        if last and last.end() < len(s):
            city = s[last.end():].strip(" ,")
            s = s[: last.end()].strip()
        else:
            parts = s.rsplit(" ", 1)
            if len(parts) == 2 and not parts[1].isdigit() and len(parts[1]) > 2:
                s, city = parts[0], parts[1]
    return s.title(), city.title(), st, zp


def address_from_legal(legal: str) -> tuple:
    m = LOC_RE.search(legal or "")
    if not m:
        return "", "", "", ""
    street, city, st, zp = parse_us_address_tail(m.group(1))
    return street, city, (st or STATE), zp

# ---------------------------------------------------------------------------
# tccsearch.org recorder scraper (Playwright)
# ---------------------------------------------------------------------------
class TccSearchRecorder:
    _UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
           "AppleWebKit/537.36 (KHTML, like Gecko) "
           "Chrome/124.0.0.0 Safari/537.36")

    def __init__(self, end: datetime, skip_detail: bool = False):
        self.end = end
        self.skip_detail = skip_detail
        self._detail_visits = 0

    # -- session ----------------------------------------------------------
    def _enter_site(self, page) -> None:
        page.goto(f"{REC_BASE}/", wait_until="domcontentloaded", timeout=45000)
        try:
            page.click("text=acknowledge the disclaimer", timeout=8000)
            page.wait_for_load_state("domcontentloaded", timeout=15000)
        except Exception:
            pass  # some sessions land straight on the tabs

    # -- one doc-type search ----------------------------------------------
    def _fill_date(self, page, container_id: str, value: str) -> None:
        sel = f"#{container_id} input.igte_ElectricBlueEditInContainer"
        page.click(sel, timeout=15000)
        page.wait_for_timeout(300)
        page.keyboard.type(value, delay=40)
        page.wait_for_timeout(200)

    def _run_search(self, page, code: str, start: datetime, end: datetime) -> bool:
        page.goto(REC_ENTRY, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_selector("#cphNoMargin_f_dclDocType", timeout=30000)
        cb = page.locator(f"#cphNoMargin_f_dclDocType input[value='{code}']")
        if not cb.count():
            log.warning("tccsearch: doc-type checkbox %r not found", code)
            return False
        cb.check()
        self._fill_date(page, "cphNoMargin_f_ddcDateFiledFrom",
                        start.strftime("%m/%d/%Y"))
        self._fill_date(page, "cphNoMargin_f_ddcDateFiledTo",
                        end.strftime("%m/%d/%Y"))
        page.click("#cphNoMargin_SearchButtons2_btnSearch")
        try:
            page.wait_for_url("**/SearchResults.aspx*", timeout=60000)
        except Exception:
            pass
        try:
            page.wait_for_selector("text=records found", timeout=45000)
        except Exception:
            log.warning("tccsearch %s: results banner not seen", code)
        page.wait_for_timeout(800)
        return "SearchResults" in page.url

    # -- results parsing ---------------------------------------------------
    @staticmethod
    def _total_records(html: str) -> int:
        m = re.search(r"\(\s*([\d,]+)\s+records?\s+found", html, re.I)
        return int(m.group(1).replace(",", "")) if m else 0

    @staticmethod
    def _parse_rows(html: str) -> list:
        """Yield dicts: doc_num, href, filed, row_type, r_names, e_names,
        legal. Grid rows are the <tr>s that contain a 9-digit instrument
        anchor; cells are matched by content, not position."""
        soup = BeautifulSoup(html, "lxml")
        out = []
        seen = set()
        for a in soup.find_all("a"):
            txt = _norm_ws(a.get_text())
            if not DOCNUM_RE.match(txt):
                continue
            tr = a.find_parent("tr")
            if tr is None or id(tr) in seen:
                continue
            seen.add(id(tr))
            tds = tr.find_all("td")
            filed = row_type = legal = ""
            r_names, e_names = [], []
            names_idx = status_idx = None
            texts = [_norm_ws(td.get_text(" ")) for td in tds]
            for i, t in enumerate(texts):
                if not filed and DATE_TXT_RE.match(t):
                    filed = t
                if "[R]" in t and "[E]" in t and names_idx is None:
                    names_idx = i
                if t in ("Temp", "Perm") and status_idx is None and i > 4:
                    status_idx = i
            if names_idx is not None:
                cell = texts[names_idx]
                m = re.search(r"\[R\](.*?)\[E\](.*)$", cell)
                if m:
                    r_raw, e_raw = m.group(1), m.group(2)
                else:
                    r_raw, e_raw = cell, ""
                clean = lambda s: [x for x in
                                   (_norm_ws(p) for p in re.split(r"\(\+\)", s))
                                   if x and x not in ("[R]", "[E]")]
                r_names, e_names = clean(r_raw), clean(e_raw)
            # doc type: the non-date, non-names capitalized cell before names
            for i, t in enumerate(texts):
                if (t and t not in ("View", "Temp", "Perm") and
                        not DATE_TXT_RE.match(t) and not DOCNUM_RE.match(t.split()[0] if t.split() else "") and
                        "[R]" not in t and len(t) > 6 and t.upper() == t and
                        (names_idx is None or i < names_idx) and
                        not re.match(r"^\d", t)):
                    row_type = t
                    break
            if status_idx is not None and status_idx > 0:
                cand = texts[status_idx - 1]
                if cand and "[R]" not in cand and not DATE_TXT_RE.match(cand):
                    legal = cand
            href = a.get("href") or ""
            out.append({"doc_num": txt, "href": href, "filed": filed,
                        "row_type": row_type, "r_names": r_names,
                        "e_names": e_names, "legal": legal})
        return out

    def _paginate(self, page) -> list:
        html = page.content()
        total = self._total_records(html)
        rows = self._parse_rows(html)
        pages_needed = min((total + 19) // 20 if total else 1, MAX_PAGES_PER_TYPE)
        for _ in range(2, pages_needed + 1):
            try:
                nxt = page.locator("#OptionsBar1_imgNext")
                if not nxt.count():
                    break
                nxt.click()
                page.wait_for_timeout(2500)
                page.wait_for_selector("text=records found", timeout=30000)
                new = self._parse_rows(page.content())
                if not new or (rows and new[0]["doc_num"] == rows[-19 if len(rows) >= 19 else 0]["doc_num"]):
                    pass
                rows.extend(new)
            except Exception as exc:
                log.debug("pagination stopped: %s", exc)
                break
        # de-dup by doc_num (pagination overlap safety)
        seen, uniq = set(), []
        for r in rows:
            if r["doc_num"] not in seen:
                seen.add(r["doc_num"])
                uniq.append(r)
        return uniq

    # -- FC detail page: owner + lender ------------------------------------
    def _fc_detail(self, page, href: str) -> tuple:
        """Return (owner_names, lender, sale_date) from a NOTICE OF
        SUBSTITUTE TRUSTEE SALE detail page. The grid [E] column only
        shows the sale date; the detail page's 'Sale Date/Owner' section
        lists sale date(s) + owner name(s), 'Lender/Trustee' the
        trustee + lender."""
        url = href if href.startswith("http") else f"{REC_BASE}/RealEstate/{href.lstrip('/')}"
        page.goto(url, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(1200)
        body = page.inner_text("body")
        owners, lender, sale_date = [], "", ""
        sec = re.search(r"Sale Date/Owner\s*(.*?)(?:Returnee|Notary|Legal Description|$)",
                        body, re.S)
        if sec:
            for line in sec.group(1).splitlines():
                t = _norm_ws(re.sub(r"^\d+\s+", "", line))
                if not t:
                    continue
                if DATE_TXT_RE.match(t):
                    sale_date = sale_date or t
                elif len(t) > 2 and not t.startswith("Sale Date"):
                    owners.append(t)
        sec = re.search(r"Lender/Trustee\s*(.*?)(?:Sale Date/Owner|Returnee|$)",
                        body, re.S)
        if sec:
            cands = [_norm_ws(re.sub(r"^\d+\s+", "", l))
                     for l in sec.group(1).splitlines()]
            cands = [c for c in cands if c]
            ent = [c for c in cands if _looks_like_entity(c)]
            lender = ent[-1] if ent else (cands[-1] if cands else "")
        return owners, lender, sale_date

    # -- main run -----------------------------------------------------------
    def run(self) -> list:
        records = []
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            for code, label, cat, cat_label, days in REC_DOC_TYPES:
                start = self.end - timedelta(days=days)
                ctx = browser.new_context(
                    user_agent=self._UA,
                    viewport={"width": 1500, "height": 900})
                page = ctx.new_page()
                try:
                    self._enter_site(page)
                    if not self._run_search(page, code, start, self.end):
                        log.warning("tccsearch %-18s: search did not reach "
                                    "results", code)
                        ctx.close()
                        continue
                    rows = self._paginate(page)
                    lab_u = label.upper()
                    typed = [r for r in rows if r["row_type"]
                             and lab_u in r["row_type"].upper()]
                    if rows and not typed:
                        typed = rows  # type cell parse miss; trust criteria
                    n = 0
                    fc_pend = []
                    for row in typed:
                        e_dates = [x for x in row["e_names"] if DATE_TXT_RE.match(x)]
                        e_names = [x for x in row["e_names"] if not DATE_TXT_RE.match(x)]
                        if cat == "FC":
                            owner, counter = _pick_owner([], e_names)
                            sale = e_dates[0] if e_dates else ""
                        else:
                            owner, counter = _pick_owner(row["r_names"], e_names)
                            sale = ""
                        ps, pc, pst, pz = address_from_legal(row["legal"])
                        rec = LeadRecord(
                            doc_num=row["doc_num"], doc_type=label,
                            cat=cat,
                            cat_label=(f"Trustee Sale {normalize_date(sale)}"
                                       if sale else cat_label),
                            filed=normalize_date(row["filed"]),
                            owner=owner, grantee=counter,
                            legal=row["legal"],
                            prop_address=ps, prop_city=pc,
                            prop_state=pst or STATE, prop_zip=pz,
                            clerk_url=REC_ENTRY,
                        )
                        records.append(rec)
                        n += 1
                        if (cat == "FC" and not owner and row["href"]
                                and not self.skip_detail):
                            fc_pend.append((rec, row["href"]))
                    # FC owner pull from detail pages (capped)
                    got = 0
                    for rec, href in fc_pend:
                        if self._detail_visits >= MAX_DETAIL_VISITS:
                            break
                        try:
                            owners, lender, sale = self._fc_detail(page, href)
                            self._detail_visits += 1
                            if owners:
                                rec.owner = owners[0]
                                got += 1
                            if lender and not rec.grantee:
                                rec.grantee = lender
                            if sale and "Trustee Sale" not in rec.cat_label:
                                rec.cat_label = f"Trustee Sale {normalize_date(sale)}"
                            time.sleep(0.4)
                        except Exception as exc:
                            log.debug("FC detail %s failed: %s", rec.doc_num, exc)
                    log.info("tccsearch %-18s: %d records (%d FC owners from "
                             "detail)", label, n, got)
                except Exception as exc:
                    log.warning("tccsearch search %s failed: %s", label, exc)
                finally:
                    ctx.close()
        return records

# ---------------------------------------------------------------------------
# TCAD parcel enrichment (Travis County GIS public ArcGIS)
# ---------------------------------------------------------------------------
PARCEL_FIELDS = ("py_owner_name,py_address,situs_num,situs_street_prefx,"
                 "situs_street,situs_street_suffix,situs_city,situs_zip,"
                 "situs_address,market_value,PROP_ID")


def _situs_from(att: dict) -> tuple:
    city = _arc_val(att.get("situs_city")).title()
    zp = _arc_val(att.get("situs_zip"))[:5]
    street = " ".join(x for x in (
        _arc_val(att.get("situs_num")), _arc_val(att.get("situs_street_prefx")),
        _arc_val(att.get("situs_street")), _arc_val(att.get("situs_street_suffix"))) if x)
    street = _norm_ws(street)
    if street:
        return street.title(), city, zp
    disp = _arc_val(att.get("situs_address"))
    if disp:
        s, c, _, z = parse_us_address_tail(disp)
        return s, c or city, z or zp
    return "", city, zp


def _mailing_from(att: dict) -> tuple:
    """py_address is one string: '2012 SALIENTE WAY CARLSBAD CA 92009'."""
    raw = _arc_val(att.get("py_address"))
    if not raw:
        return "", "", "", ""
    street, city, st, zp = parse_us_address_tail(raw)
    return street, city, (st or STATE), zp


def _arcgis_query(session, where: str, count: int = 5) -> list:
    params = {
        "where": where, "outFields": PARCEL_FIELDS,
        "returnGeometry": "false", "f": "json", "resultRecordCount": count,
    }
    try:
        r = session.get(PARCEL_API_URL, params=params, timeout=REQUEST_TIMEOUT)
        return r.json().get("features", []) or []
    except Exception as exc:
        log.debug("ArcGIS query error: %s", exc)
        return []


def _addr_key(addr: str) -> tuple:
    m = re.match(r"\s*(\d+)\s+(.*)", addr or "")
    if not m:
        return "", ""
    rest = _norm_ws(m.group(2))
    rest = re.sub(r"\s+(#|APT|UNIT|STE|SUITE|BLDG|LOT)\b.*$", "", rest, flags=re.I)
    rest = re.sub(r"\s+" + SUFFIXES + r"\.?$", "", rest, flags=re.I).strip()
    rest = re.sub(r"^(N|S|E|W|NE|NW|SE|SW)\s+", "", rest, flags=re.I).strip()
    return m.group(1), rest


def enrich_parcels(records: list) -> None:
    session = requests.Session()
    session.headers["User-Agent"] = "TravisLeadScraper/1.0"

    fwd = [r for r in records if r.owner and (not r.prop_address or not r.mail_address)]
    log.info("TCAD owner-lookup for %d records...", len(fwd))
    hits = 0
    for rec in fwd[:ARCGIS_MAX_LOOKUPS]:
        norm = normalize_owner_for_parcel(rec.owner)
        if not norm or len(norm) < 5:
            continue
        feats = _arcgis_query(session,
                              f"UPPER(py_owner_name) LIKE '{_sql_lit(norm)}%'")
        if not feats:
            continue
        att = feats[0].get("attributes", {})
        if not rec.prop_address and len(feats) == 1:
            ps, pc, pz = _situs_from(att)
            if ps:
                rec.prop_address, rec.prop_city, rec.prop_zip = ps, pc or rec.prop_city, pz
        if not rec.mail_address:
            ms, mc, mst, mz = _mailing_from(att)
            if ms:
                rec.mail_address, rec.mail_city, rec.mail_state, rec.mail_zip = ms, mc, mst, mz
                hits += 1
        if not rec.amount:
            try:
                rec.amount = float(att.get("market_value") or 0)
            except (TypeError, ValueError):
                pass
        time.sleep(0.12)
    log.info("TCAD owner-lookup: %d mailing fills", hits)

    rev = [r for r in records if r.prop_address and not r.owner]
    log.info("TCAD address-lookup for %d records...", len(rev))
    hits = 0
    for rec in rev[:ARCGIS_MAX_LOOKUPS]:
        num, core = _addr_key(rec.prop_address)
        if not num or not core:
            continue
        feats = _arcgis_query(
            session, f"situs_num = '{_sql_lit(num)}' AND "
                     f"UPPER(situs_street) LIKE '{_sql_lit(core)}%'")
        if len(feats) == 1:
            att = feats[0].get("attributes", {})
            owner = _arc_val(att.get("py_owner_name"))
            if owner:
                rec.owner = owner
                hits += 1
            if not rec.mail_address:
                ms, mc, mst, mz = _mailing_from(att)
                if ms:
                    rec.mail_address, rec.mail_city, rec.mail_state, rec.mail_zip = ms, mc, mst, mz
            if not rec.amount:
                try:
                    rec.amount = float(att.get("market_value") or 0)
                except (TypeError, ValueError):
                    pass
        time.sleep(0.12)
    log.info("TCAD address-lookup: %d owner fills", hits)

# ---------------------------------------------------------------------------
# Hash / dedupe + NEW-CHANGED detection
# ---------------------------------------------------------------------------
def _repo_base() -> Path:
    return Path(__file__).parent.parent


def _record_rid(r) -> str:
    basis = r.doc_num or f"{r.owner}|{r.filed}|{r.doc_type}|{r.prop_address}"
    return hashlib.sha1(f"travis|{basis}".encode()).hexdigest()[:16]


def _record_chash(r) -> str:
    fields = "|".join(str(x or "") for x in (
        r.doc_num, r.doc_type, r.filed, r.owner, r.grantee, r.legal,
        r.amount, r.prop_address, r.mail_address))
    return hashlib.sha1(fields.encode()).hexdigest()[:16]


def detect_changes(records: list) -> None:
    state_path = _repo_base() / "data" / "state.json"
    try:
        state = json.loads(state_path.read_text(encoding="utf-8")) if state_path.exists() else {}
    except Exception:
        state = {}
    today = datetime.now().strftime("%Y-%m-%d")
    n_new = n_chg = n_exist = 0
    for r in records:
        r.rid = _record_rid(r)
        r.content_hash = _record_chash(r)
        prev = state.get(r.rid)
        if prev is None:
            r.status, r.first_seen = "NEW", today
            n_new += 1
        elif prev.get("content_hash") != r.content_hash:
            r.status = "CHANGED"
            r.first_seen = prev.get("first_seen", today)
            n_chg += 1
        else:
            r.status = "EXISTING"
            r.first_seen = prev.get("first_seen", today)
            n_exist += 1
        state[r.rid] = {"content_hash": r.content_hash,
                        "first_seen": r.first_seen, "last_seen": today}
    state_path.parent.mkdir(parents=True, exist_ok=True)
    state_path.write_text(json.dumps(state, indent=1), encoding="utf-8")
    log.info("NEW/CHANGED: NEW=%d CHANGED=%d EXISTING=%d (state=%d ids)",
             n_new, n_chg, n_exist, len(state))

# ---------------------------------------------------------------------------
# Scoring
# ---------------------------------------------------------------------------
def score_records(records: list, start: datetime) -> None:
    for r in records:
        s, flags = 30, []
        if r.cat == "LP": s += 10; flags.append("LIS_PENDENS")
        if r.cat == "FC": s += 15; flags.append("FORECLOSURE")
        if r.cat == "TAXFC": s += 18; flags.append("TAX_FORECLOSURE")
        if r.cat == "TAXDEED": s += 10; flags.append("TAX_DEED")
        if r.cat in ("LP","FC","TAXFC"): s += 5
        if r.cat == "JUD": s += 8; flags.append("JUDGMENT")
        if r.cat == "LIEN": s += 7; flags.append("LIEN")
        if r.cat == "PRO": s += 12; flags.append("PROBATE")
        if r.amount > 100000: s += 15; flags.append("HIGH_AMOUNT")
        elif r.amount > 50000: s += 10; flags.append("MID_AMOUNT")
        if r.filed:
            try:
                if datetime.strptime(r.filed, "%Y-%m-%d") >= start:
                    s += 5; flags.append("NEW_THIS_WEEK")
            except ValueError:
                pass
        if r.prop_address:
            s += 5; flags.append("HAS_ADDRESS")
        r.score = min(s, 100)
        r.flags = flags

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
DASH_CAT = {
    "LP": "foreclosure", "FC": "foreclosure", "TAXFC": "foreclosure",
    "TAXDEED": "tax_lien", "LIEN": "tax_lien",
    "JUD": "judgment", "PRO": "probate",
}
FLAG_NICE = {
    "LIS_PENDENS": "Lis pendens", "FORECLOSURE": "Pre-foreclosure",
    "TAX_FORECLOSURE": "Tax foreclosure", "TAX_DEED": "Tax deed",
    "JUDGMENT": "Judgment lien", "LIEN": "Tax lien",
    "PROBATE": "Probate / estate", "HIGH_AMOUNT": "Amount > $100k",
    "MID_AMOUNT": "Amount > $50k", "NEW_THIS_WEEK": "New this week",
    "HAS_ADDRESS": "Has address",
}


def write_outputs(records: list, start: datetime, end: datetime) -> None:
    base = _repo_base()
    for d in [base / "dashboard", base / "data"]:
        d.mkdir(parents=True, exist_ok=True)
    week_ago = (end - timedelta(days=7)).strftime("%Y-%m-%d")
    recs_out = []
    for r in records:
        d = asdict(r)
        d["cat_code"] = r.cat
        d["cat"] = DASH_CAT.get(r.cat, "tax_lien")
        d["flags"] = [FLAG_NICE.get(f, f) for f in (r.flags or [])]
        d["absentee"] = bool(
            r.prop_address and r.mail_address
            and r.prop_address.upper() != r.mail_address.upper())
        d["out_of_state"] = bool(r.mail_state and r.mail_state.upper() != STATE)
        recs_out.append(d)
    payload = {
        "fetched_at": datetime.utcnow().isoformat(),
        "county": COUNTY,
        "source": f"{COUNTY} County, {STATE} -- tccsearch.org Recorder + TCAD Parcels",
        "date_range": {"start": start.strftime("%Y-%m-%d"), "end": end.strftime("%Y-%m-%d")},
        "total": len(records),
        "new_7d": sum(1 for r in records if (r.first_seen or "") >= week_ago),
        "with_address": sum(1 for r in records if r.prop_address),
        "by_cat": {c: sum(1 for r in records if r.cat == c) for c in ("FC","TAXFC","TAXDEED","LP","JUD","LIEN","PRO")},
        "records": recs_out,
    }
    for path in [base / "dashboard" / "records.json", base / "data" / "records.json"]:
        path.write_text(json.dumps(payload, indent=2, default=str))
        log.info("JSON written: %s (%d records)", path, len(records))
    csv_path = base / "data" / "ghl_export.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(GHL_HEADERS.values()))
        writer.writeheader()
        for r in records:
            d = asdict(r)
            writer.writerow({GHL_HEADERS[k]: ("|".join(d[k]) if k=="flags" else d[k]) for k in GHL_FIELDS})
    log.info("GHL CSV written: %s (%d records)", csv_path, len(records))
    skip_path = base / "data" / "skiptrace_export.csv"
    skip_cols = ["First Name", "Last Name", "Mailing Address", "Mailing City",
                 "Mailing State", "Mailing Zip", "Property Address",
                 "Property City", "Property State", "Property Zip"]
    with open(skip_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=skip_cols)
        writer.writeheader()
        for r in records:
            owner = (r.owner or "").strip()
            if "," in owner:
                p = owner.split(",", 1)
                first, last = p[1].strip().title(), p[0].strip().title()
            elif _looks_like_entity(owner):
                first, last = "", owner.title()
            else:
                p = owner.split()
                if p and owner == owner.upper() and len(p) > 1:
                    # recorder/TCAD "LAST FIRST M" style
                    first, last = p[1].title(), p[0].title()
                else:
                    first = p[0].title() if p else ""
                    last = p[-1].title() if len(p) > 1 else ""
            writer.writerow({
                "First Name": first, "Last Name": last,
                "Mailing Address": r.mail_address, "Mailing City": r.mail_city,
                "Mailing State": r.mail_state, "Mailing Zip": r.mail_zip,
                "Property Address": r.prop_address, "Property City": r.prop_city,
                "Property State": r.prop_state, "Property Zip": r.prop_zip,
            })
    log.info("Skip trace CSV written: %s (%d records)", skip_path, len(records))

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="Travis County lead scraper")
    parser.add_argument("--days", type=int, default=LOOKBACK_DAYS)
    parser.add_argument("--skip-parcel", action="store_true")
    parser.add_argument("--skip-detail", action="store_true")
    args = parser.parse_args()
    end = datetime.now()
    start = end - timedelta(days=max(args.days, LIEN_LOOKBACK_DAYS))
    log.info("=" * 60)
    log.info("Travis County Motivated Seller Lead Scraper")
    log.info("Lookback: judgments/mech %dd, liens/LP/FC %dd, probate %dd",
             args.days, LIEN_LOOKBACK_DAYS, PRO_LOOKBACK_DAYS)
    log.info("=" * 60)

    recorder = TccSearchRecorder(end, skip_detail=args.skip_detail)
    records = recorder.run()

    # dedupe on doc_num
    seen, unique = set(), []
    for r in records:
        key = r.doc_num or f"{r.owner}|{r.filed}|{r.doc_type}"
        if key not in seen:
            seen.add(key)
            unique.append(r)
    records = unique

    if not args.skip_parcel:
        enrich_parcels(records)
    detect_changes(records)
    score_records(records, start)
    records.sort(key=lambda r: (r.status != "NEW", -r.score))
    if not records:
        log.warning("No records found. Writing empty output files.")
    else:
        log.info("Total after dedup + enrichment: %d", len(records))
    write_outputs(records, start, end)
    log.info("=" * 60)
    log.info("SUMMARY")
    log.info("  Total records  : %d", len(records))
    log.info("  With address   : %d", sum(1 for r in records if r.prop_address))
    log.info("  Score >= 70    : %d", sum(1 for r in records if r.score >= 70))
    log.info("  Score >= 50    : %d", sum(1 for r in records if r.score >= 50))
    for c in ("FC","LP","JUD","LIEN","PRO"):
        log.info("  cat %-5s      : %d", c, sum(1 for r in records if r.cat == c))


if __name__ == "__main__":
    main()
