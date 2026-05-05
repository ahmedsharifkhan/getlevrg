# Get Levrg — Lead Scraper

Production-ready Python scraper that pulls B2B lead candidates from genuinely
public sources (Y Combinator, HubSpot Solutions Partners, Clutch), enriches
each one by visiting its website, applies the Get Levrg ICP filter, scores
each lead 0–10, and produces a 16-column XLSX ready for outreach.

---

## What you get

After a successful run, three files appear in your output folder:

| File | Purpose |
|---|---|
| `leads_raw.json` | Every candidate pulled from the directories, before filtering. Useful for debugging or re-scoring later. |
| `leads_final.csv` | Filtered + scored leads (score ≥ 7), CSV format — drag into HubSpot, Pipedrive, or Sheets. |
| `leads_final.xlsx` | Same data, formatted Excel with frozen header, autofilter, and sized columns. |

Columns produced (matches the original brief exactly):

1. Company Name
2. Website
3. Country
4. Industry
5. Estimated Employee Size
6. ICP Segment
7. Buying Signal
8. Why This Is a Quality Lead
9. Best Service to Pitch
10. Decision Maker Name
11. Decision Maker Title
12. Decision Maker LinkedIn/Profile URL
13. Public Email or Contact Page URL
14. Source URL 1
15. Source URL 2
16. Lead Score 1–10

---

## Setup (one time)

```bash
# 1. Make sure you have Python 3.10 or newer
python3 --version

# 2. Install dependencies
pip install requests beautifulsoup4 lxml pandas openpyxl

# 3. (Optional) Use a virtual environment
python3 -m venv venv && source venv/bin/activate
pip install requests beautifulsoup4 lxml pandas openpyxl
```

---

## Run

Default run (50 leads, all 3 sources, score ≥ 7):

```bash
python3 getlevrg_lead_scraper.py
```

Common variations:

```bash
# Only YC + HubSpot (skip Clutch — they often block scrapers)
python3 getlevrg_lead_scraper.py --sources yc hubspot

# Pull more pages from each directory
python3 getlevrg_lead_scraper.py --yc-pages 10 --hubspot-pages 10

# Looser score threshold (gets more leads, lower quality)
python3 getlevrg_lead_scraper.py --min-score 6 --target-leads 100

# Save into a specific folder
python3 getlevrg_lead_scraper.py --outdir ./output_2026_05
```

All flags:

```
--target-leads N      How many leads to keep at the end (default: 50)
--min-score N         Drop anything below this score (default: 7)
--delay SECONDS       Pause between HTTP requests (default: 1.5; bump to 3 if blocked)
--sources yc hubspot clutch    Which directories to pull from
--yc-pages N          Pages of YC results (100 companies/page; default: 5)
--hubspot-pages N     Pages of HubSpot partner results (default: 5)
--outdir PATH         Where to write output files
```

---

## How long does it take?

| Stage | Time (default settings) |
|---|---|
| Stage 1 — pull directories | 1–3 minutes |
| Stage 2 — visit each website | 30–60 minutes (1.5s delay × ~6 page hits per company × 100–200 candidates) |
| Stage 3 — write CSV/XLSX | seconds |

If you're in a hurry: drop `--yc-pages 2 --hubspot-pages 2` for a faster pass
(~15 minutes total, ~25–30 final leads).

---

## What the ICP filter actually does

A lead is scored 0–10 based on these signals; only ≥ 7 is kept:

| Signal | Points |
|---|---|
| Industry text matches B2B/SaaS/agency/consulting/RevOps | +1 |
| Employee size 10–250 | +2 |
| Employee size <500 (broader band) | +1 |
| Located in US/CA/UK/AU | +1 |
| Hiring signal (marketing/CRM/RevOps/SDR roles on careers page) | +2 |
| Active blog | +1 |
| Mentions HubSpot on site | +1 |
| Has podcast or case studies | +1 |
| Decision maker name + title found on team page | +1 |
| Industry contains ecommerce / restaurant / salon / etc. | EXCLUDED |

Leads that match B2B + active marketing + hiring + decision-maker-found tend
to land in the 8–10 range — those are the highest-priority outreach targets.

---

## Honest expectations

This is a free, public-sources scraper. Be realistic about what it CAN and
CANNOT find:

**It will reliably populate:**
- Company name, website, country, industry, employee size
- Whether they have a blog / podcast / mention HubSpot
- Whether they're hiring marketing/RevOps roles
- Contact page URL (always available)
- Public email IF they list one (~30–50% of companies do)
- Decision maker name + title IF they have a public team page (~40–60%)
- Decision maker LinkedIn IF they link it on the team page (~25–35%)

**It will frequently leave blank:**
- Decision maker LinkedIn URL (most companies don't link team LinkedIns publicly)
- Direct decision-maker email (almost never on websites)

For LinkedIn URLs and verified emails, plug in:
- **Apollo.io free tier** (50 email reveals/mo, unlimited LinkedIn URLs)
- **Hunter.io** for email patterns
- Or pass the CSV through Clay / Apify's LinkedIn actor

---

## Troubleshooting

**"Clutch returns 403"** — Clutch has aggressive anti-bot. Either skip
(`--sources yc hubspot`) or run behind a proxy. Their data isn't critical;
YC + HubSpot covers the highest-quality candidates.

**"YC API returned 0 hits"** — Their Algolia keys occasionally rotate. If
this happens, open https://www.ycombinator.com/companies in a browser, open
DevTools → Network tab, click any filter, find the request to
`*-dsn.algolia.net/1/indexes/*/queries`, copy the `X-Algolia-API-Key` and
`X-Algolia-Application-Id` headers, paste them into the script
(`fetch_yc_companies` function).

**"Many leads have empty Decision Maker"** — Expected. See "Honest
expectations" above. Run an Apollo / Lusha enrichment pass on the CSV
to fill those columns.

**"Script seems hung"** — It isn't. Default delay is 1.5s × ~6 fetches per
company × 200 companies = ~30 minutes. Watch the console output — you'll see
each company being processed.

**"Hitting rate limits"** — Bump `--delay 3` or `--delay 5`.

---

## Where each lead comes from (transparency)

- **Y Combinator** — pulls from their public Algolia search index (the same
  one ycombinator.com/companies uses). Filtered to `tags=B2B`. These are
  funded startups that are growing fast — strong fractional-marketing-team
  candidates.
- **HubSpot Solutions Partners** — public partner directory. Every entry is
  a marketing/CRM agency. Premium fit for Get Levrg's agency-fulfillment
  service line. Bigger partners (Platinum, Diamond, Elite) are dealing with
  fulfillment pressure that Get Levrg directly solves.
- **Clutch** — verified B2B service providers. Best categories for our ICP:
  `digital-marketing`, `content-marketing`, `inbound-marketing`. Anti-bot is
  aggressive, so this source is best-effort.

---

## Customizing for your sales motion

Edit the top of `getlevrg_lead_scraper.py`:

- `TARGET_COUNTRIES` — adjust if you want EU, India, etc.
- `GOOD_KEYWORDS` — add industries you specifically target.
- `BAD_KEYWORDS` — add anything you want to exclude (industries, business
  models, etc.).
- `DECISION_TITLES` — tighten or loosen the regex if you want to find only
  CEOs, or also include Marketing Directors at the IC level, etc.

---

## License / use

This script only fetches data from publicly visible directory pages and
public company websites. It respects rate limits (1.5s default delay) and
sets a normal browser User-Agent. No login walls are bypassed, no LinkedIn
scraping, no terms-of-service violations. Use it inside your normal sales
prospecting workflow.