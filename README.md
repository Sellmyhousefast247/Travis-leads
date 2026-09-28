# Travis-leads

Motivated-seller lead scraper for **Travis County, TX** (Austin).

- **Recorder**: [tccsearch.org](https://www.tccsearch.org/) County Clerk Web Search (Aumentum, anonymous) — trustee-sale notices, abstracts of judgment, lis pendens, heirship affidavits, federal/state tax liens, mechanic's/hospital/child-support/HOA liens
- **Enrichment**: TCAD parcels via Travis County GIS ArcGIS (owner, situs, mailing address, market value)
- **Pipeline**: scrape → normalize → hash/dedupe → NEW/CHANGED detection → score → export
- **Outputs**: `dashboard/records.json` (live dashboard), `data/ghl_export.csv`, `data/skiptrace_export.csv`
- **Automation**: GitHub Actions daily at 13:00 UTC + manual dispatch; dashboard deployed to GitHub Pages

Dashboard: https://sellmyhousefast247.github.io/Travis-leads/
