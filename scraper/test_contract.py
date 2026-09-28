#!/usr/bin/env python3
"""Mocked pipeline-contract test for the Travis scraper.

Fakes the recorder (TccSearchRecorder.run) with rows shaped from live
tccsearch.org samples and the TCAD ArcGIS API with live-sampled parcel
attributes, then runs main() and asserts the records.json contract,
CSV row counts, dedupe, state.json, and absentee/out-of-state logic.

Run:  python scraper/test_contract.py
"""
import csv
import json
import sys
from pathlib import Path
from unittest import mock

import fetch as F


# --- fixtures: shaped from live tccsearch.org rows (2026-09-28) -----------
def fake_recorder_records():
    return [
        # FC row: owner pulled from detail page; legal has LOC address
        F.LeadRecord(
            doc_num="202641317", doc_type="NOTICE OF SUBSTITUTE TRUSTEE SALE",
            cat="FC", cat_label="Trustee Sale 2026-11-03", filed="2026-09-24",
            owner="ZAVALA ANGELA", grantee="21ST MORTGAGE CORPORATION",
            legal="LT 31 BLK C MEADOWS OF BLACKHAWK LOC 20204 KEARNEY HILL "
                  "RD PFLUGERVILLE TX 78660 PDV 4",
            prop_address="20204 Kearney Hill Rd", prop_city="Pflugerville",
            prop_zip="78660", clerk_url=F.REC_ENTRY),
        # duplicate doc_num -> must be deduped
        F.LeadRecord(
            doc_num="202641317", doc_type="NOTICE OF SUBSTITUTE TRUSTEE SALE",
            cat="FC", cat_label="Trustee Sale 2026-11-03", filed="2026-09-24",
            owner="ZAVALA ANGELA", clerk_url=F.REC_ENTRY),
        # JUD row: [E] debtor is the owner, no address until enrichment
        F.LeadRecord(
            doc_num="202641101", doc_type="ABSTRACT OF JUDGMENT",
            cat="JUD", cat_label="Abstract of Judgment", filed="2026-09-23",
            owner="MORGAN JOHN B", grantee="MIDLAND CREDIT MANAGEMENT INC",
            legal="", clerk_url=F.REC_ENTRY),
        # LIEN row (federal tax lien), entity-vs-person picking already done
        F.LeadRecord(
            doc_num="202640990", doc_type="FEDERAL TAX LIENS AND NOTICES",
            cat="LIEN", cat_label="Federal Tax Lien", filed="2026-09-22",
            owner="HARDY ROBERT S III", grantee="INTERNAL REVENUE SERVICE",
            legal="", clerk_url=F.REC_ENTRY),
        # PRO row (heirship)
        F.LeadRecord(
            doc_num="202640800", doc_type="AFFIDAVIT OF HEIRSHIP",
            cat="PRO", cat_label="Affidavit of Heirship", filed="2026-09-10",
            owner="FINDEISEN LILLIAN GRACE", grantee="",
            legal="LOT 47 BLK N VILLAGE TWENTY AT ANDERSON MILL PHS 2",
            clerk_url=F.REC_ENTRY),
        # LP row with LOC address but NO owner -> reverse enrichment fills
        F.LeadRecord(
            doc_num="202640700", doc_type="LIS PENDENS",
            cat="LP", cat_label="Lis Pendens", filed="2026-09-20",
            owner="", grantee="WILMINGTON SAVINGS FUND SOCIETY",
            legal="LT 15 BLK C SEC 1 HARRIS RIDGE LOC 1421 DEXFORD DR "
                  "AUSTIN TX 78753 PDV 3",
            prop_address="1421 Dexford Dr", prop_city="Austin",
            prop_zip="78753", clerk_url=F.REC_ENTRY),
    ]


# live-sampled TCAD attribute shapes (gis.traviscountytx.gov TCAD layer)
TCAD_BY_OWNER = {
    "MORGAN JOHN": {
        "py_owner_name": "MORGAN JOHN B & MICHELE",
        "py_address": "11504 TANGLEBRIAR TRL AUSTIN TX 78750",
        "situs_num": "11504", "situs_street_prefx": None,
        "situs_street": "TANGLEBRIAR", "situs_street_suffix": "TRL",
        "situs_city": "AUSTIN", "situs_zip": "78750",
        "situs_address": "11504 TANGLEBRIAR TRL AUSTIN 78750",
        "market_value": 409695, "PROP_ID": 177372,
    },
    "HARDY ROBERT": {
        "py_owner_name": "HARDY ROBERT S III",
        # out-of-state mailing -> out_of_state + absentee must be True
        "py_address": "2012 SALIENTE WAY CARLSBAD CA 92009",
        "situs_num": "11500", "situs_street_prefx": None,
        "situs_street": "TANGLEBRIAR", "situs_street_suffix": "TRL",
        "situs_city": "AUSTIN", "situs_zip": "78750",
        "situs_address": "11500 TANGLEBRIAR TRL AUSTIN 78750",
        "market_value": 406515, "PROP_ID": 177374,
    },
    "ZAVALA ANGELA": {
        "py_owner_name": "ZAVALA ANGELA",
        "py_address": "20204 KEARNEY HILL RD PFLUGERVILLE TX 78660",
        "situs_num": "20204", "situs_street_prefx": None,
        "situs_street": "KEARNEY HILL", "situs_street_suffix": "RD",
        "situs_city": "PFLUGERVILLE", "situs_zip": "78660",
        "situs_address": "20204 KEARNEY HILL RD PFLUGERVILLE 78660",
        "market_value": 355000, "PROP_ID": 500001,
    },
    "FINDEISEN LILLIAN": {
        "py_owner_name": "FINDEISEN LILLIAN GRACE",
        "py_address": "11501 FENCE POST TRL AUSTIN TX 78750",
        "situs_num": "11501", "situs_street_prefx": None,
        "situs_street": "FENCE POST", "situs_street_suffix": "TRL",
        "situs_city": "AUSTIN", "situs_zip": "78750",
        "situs_address": "11501 FENCE POST TRL AUSTIN 78750",
        "market_value": 402764, "PROP_ID": 177377,
    },
}
TCAD_BY_SITUS = {
    ("1421", "DEXFORD"): {
        "py_owner_name": "VACEK JOSEPH",
        "py_address": "PO BOX 141534 AUSTIN TX 78714",
        "situs_num": "1421", "situs_street_prefx": None,
        "situs_street": "DEXFORD", "situs_street_suffix": "DR",
        "situs_city": "AUSTIN", "situs_zip": "78753",
        "situs_address": "1421 DEXFORD DR AUSTIN 78753",
        "market_value": 310000, "PROP_ID": 600001,
    },
}


def fake_arcgis_query(session, where, count=5):
    import re
    m = re.search(r"py_owner_name\) LIKE '([^']+)%'", where)
    if m:
        att = TCAD_BY_OWNER.get(m.group(1))
        return [{"attributes": att}] if att else []
    m = re.search(r"situs_num = '([^']+)' AND UPPER\(situs_street\) "
                  r"LIKE '([^']+)%'", where)
    if m:
        att = TCAD_BY_SITUS.get((m.group(1), m.group(2)))
        return [{"attributes": att}] if att else []
    return []


def run():
    base = Path(F.__file__).parent.parent
    sp = base / "data" / "state.json"
    if sp.exists():
        sp.unlink()

    with mock.patch.object(F.TccSearchRecorder, "run",
                           lambda self: fake_recorder_records()), \
         mock.patch.object(F, "_arcgis_query", fake_arcgis_query), \
         mock.patch.object(sys, "argv", ["fetch.py"]):
        F.main()

    rj = json.loads((base / "dashboard" / "records.json").read_text())
    rj2 = json.loads((base / "data" / "records.json").read_text())
    assert rj["total"] == rj2["total"]

    # top-level contract
    for k in ("fetched_at", "county", "date_range", "total", "new_7d",
              "with_address", "records"):
        assert k in rj, f"missing top-level key {k}"
    assert rj["county"] == "Travis"
    assert set(rj["date_range"]) == {"start", "end"}
    assert rj["total"] == 5, f"dedupe failed: total={rj['total']}"

    cats = {"foreclosure", "tax_lien", "judgment", "probate"}
    req = ("status first_seen cat cat_code cat_label doc_type score flags "
           "absentee out_of_state owner grantee prop_address prop_city "
           "prop_state prop_zip mail_address mail_city mail_state mail_zip "
           "amount legal doc_num clerk_url filed").split()
    for r in rj["records"]:
        for k in req:
            assert k in r, f"record missing {k}: {r['doc_num']}"
        assert r["cat"] in cats, r["cat"]
        assert r["status"] == "NEW"
        assert 0 <= r["score"] <= 100
        assert isinstance(r["flags"], list)

    by_num = {r["doc_num"]: r for r in rj["records"]}
    # enrichment: JUD got situs + mailing + value from TCAD
    jud = by_num["202641101"]
    assert jud["prop_address"] == "11504 Tanglebriar Trl", jud["prop_address"]
    assert jud["mail_address"] == "11504 Tanglebriar Trl"
    assert jud["mail_city"] == "Austin" and jud["mail_state"] == "TX"
    assert jud["mail_zip"] == "78750"
    assert jud["amount"] == 409695
    assert jud["cat"] == "judgment"
    # out-of-state mailing -> absentee + out_of_state
    lien = by_num["202640990"]
    assert lien["mail_state"] == "CA", lien["mail_state"]
    assert lien["mail_city"] == "Carlsbad"
    assert lien["out_of_state"] is True
    assert lien["absentee"] is True
    assert "Amount > $100k" in lien["flags"]
    # in-county owner-occupant -> not absentee
    fc = by_num["202641317"]
    assert fc["absentee"] is False, fc
    assert fc["out_of_state"] is False
    assert fc["cat"] == "foreclosure"
    # reverse enrichment filled owner from situs
    lp = by_num["202640700"]
    assert lp["owner"] == "VACEK JOSEPH", lp["owner"]
    assert lp["mail_address"] == "Po Box 141534"
    assert lp["cat"] == "foreclosure"  # LP maps to foreclosure
    # probate mapping
    assert by_num["202640800"]["cat"] == "probate"

    # LOC address parser directly
    s, c, st, z = F.address_from_legal(
        "LT 31 BLK C MEADOWS OF BLACKHAWK LOC 20204 KEARNEY HILL RD "
        "PFLUGERVILLE TX 78660 PDV 4")
    assert (s, c, st, z) == ("20204 Kearney Hill Rd", "Pflugerville", "TX",
                             "78660"), (s, c, st, z)
    s, c, st, z = F.parse_us_address_tail("2012 SALIENTE WAY CARLSBAD CA 92009")
    assert (s, c, st, z) == ("2012 Saliente Way", "Carlsbad", "CA", "92009"), (s, c, st, z)

    # CSVs
    with open(base / "data" / "ghl_export.csv") as f:
        ghl = list(csv.DictReader(f))
    assert len(ghl) == rj["total"]
    assert "Date Entered System" in ghl[0] and "Status" in ghl[0]
    with open(base / "data" / "skiptrace_export.csv") as f:
        sk = list(csv.DictReader(f))
    assert len(sk) == rj["total"]
    row = next(r for r in sk if r["Property Address"] == "11504 Tanglebriar Trl")
    assert row["First Name"] == "John" and row["Last Name"] == "Morgan", row

    # state.json persisted
    state = json.loads(sp.read_text())
    assert len(state) == rj["total"]

    # second run -> EXISTING
    with mock.patch.object(F.TccSearchRecorder, "run",
                           lambda self: fake_recorder_records()), \
         mock.patch.object(F, "_arcgis_query", fake_arcgis_query), \
         mock.patch.object(sys, "argv", ["fetch.py"]):
        F.main()
    rj = json.loads((base / "dashboard" / "records.json").read_text())
    assert all(r["status"] == "EXISTING" for r in rj["records"]), \
        [r["status"] for r in rj["records"]]

    print("\nALL CONTRACT TESTS PASSED")


if __name__ == "__main__":
    run()
