"""
Get Levrg — B2B Lead Scraper
============================

Scrapes high-quality B2B leads from genuinely public sources:
  1. Y Combinator company directory (public Algolia API)
  2. HubSpot Solutions Partner directory
  3. Clutch.co B2B service provider listings (best-effort, anti-bot resilient)
  4. Each company's website (about / team / contact / careers)

ICP filter (must match >= 3 signals):
  - B2B company or B2B agency
  - 10–250 employees
  - Based in US / CA / UK / AU / English-speaking market
  - Active marketing (blog, podcast, content, LinkedIn)
  - Hiring for marketing/CRM/RevOps/SDR/content roles
  - Lean team with growth signals
  - Uses HubSpot, CRM, outbound, content-led growth
  - Agency that may need fulfillment support
  - SaaS / consulting / coaching / RevOps / professional services / agency model

Outputs:
  - leads_raw.json        (intermediate dump, all candidates with raw fields)
  - leads_final.csv       (filtered, scored, ready-to-import)
  - leads_final.xlsx      (formatted Excel with autofilter + frozen header)

Usage:
  pip install requests beautifulsoup4 lxml pandas openpyxl tldextract
  python getlevrg_lead_scraper.py --target-leads 50 --min-score 7

Notes:
  - Respect robots.txt and rate limits. Default 1.5s delay between requests.
  - LinkedIn URLs are NOT scraped (ToS + bot detection). Decision-maker LinkedIn
    URLs are populated from team pages where companies publicly link them.
  - Public emails are extracted only when explicitly listed on the site.
    Otherwise the contact-page URL is recorded (per the user's prompt rule).
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import random
import re
import sys
import time
from dataclasses import dataclass, asdict, field
from typing import Iterable, Optional
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup

try:
    import pandas as pd
except ImportError:
    pd = None

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0 Safari/537.36"
)

HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

TARGET_COUNTRIES = {"United States", "Canada", "United Kingdom", "Australia",
                    "USA", "US", "UK", "CA", "AU", "GB"}

# Industries we WANT
GOOD_KEYWORDS = [
    "saas", "b2b", "marketing", "sales", "crm", "hubspot", "revops",
    "consulting", "professional services", "agency", "coaching",
    "analytics", "fintech", "edtech", "healthtech", "proptech",
    "cybersecurity", "devtools", "developer tools", "data",
    "lead gen", "growth", "automation", "operations",
]

# Industries we DON'T want (strict exclusions from the user's prompt)
BAD_KEYWORDS = [
    "restaurant", "salon", "clinic", "dental", "gym", "real estate broker",
    "ecommerce", "e-commerce", "shopify store", "marketplace consumer",
    "consumer apparel", "fashion brand", "food delivery",
]

# Title regex for finding decision makers on team pages
DECISION_TITLES = re.compile(
    r"\b(founder|co-?founder|ceo|chief\s+executive|cmo|chief\s+marketing|"
    r"vp\s+(of\s+)?marketing|vice\s+president\s+of\s+marketing|"
    r"head\s+of\s+(marketing|growth|revops|revenue\s+operations|sales|content)|"
    r"director\s+of\s+(marketing|operations|growth|revops|revenue\s+operations|sales)|"
    r"revops|operations\s+director|managing\s+director|owner|partner|principal)\b",
    re.IGNORECASE,
)

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

PHONE_RE = re.compile(
    r"(?<!\d)"
    r"(?:\+?1[\s\-.]?)?\(?\d{3}\)?[\s\-.]?\d{3}[\s\-.]?\d{4}"
    r"|"
    r"\+\d{1,3}[\s\-.]?\(?\d{2,4}\)?[\s\-.]?\d{3,4}[\s\-.]?\d{3,4}"
    r"(?!\d)"
)

FAKE_EMAIL_FRAGMENTS = (
    "example.", "@2x.", ".png", ".jpg", "sentry.",
    "you@", "name@", "email@", "user@", "test@",
    "@company.", "@domain.", "@yourdomain.", "noreply",
    "no-reply", "donotreply", "wixpress", "sentry-next",
    "schema.org", "@w3.", "@example", "@acme.", "@placeholder",
    "someone@", "person@", "john@", "jane@",
)

# Simple person-name validator: 2-4 capitalized words, no digits, no company noise
_NAME_BAD = re.compile(
    r"\b(the|and|or|for|inc|llc|ltd|our|meet|team|about|read|bio|"
    r"more|view|see|learn|click|get|sign|join|contact|follow|"
    r"connect|watch|play|download)\b",
    re.IGNORECASE,
)

def looks_like_person_name(s: str) -> bool:
    words = s.strip().split()
    if not 2 <= len(words) <= 5:
        return False
    if any(c.isdigit() or c in "@/\\<>{}[]|" for c in s):
        return False
    if _NAME_BAD.search(s):
        return False
    # Each word should start with uppercase
    if not all(w[0].isupper() for w in words if w):
        return False
    # Reject if too long (company names tend to be longer strings)
    if len(s) > 40:
        return False
    return True

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Lead:
    company_name: str = ""
    website: str = ""
    country: str = ""
    industry: str = ""
    employee_size: str = ""
    icp_segment: str = ""
    buying_signal: str = ""
    why_quality: str = ""
    best_service: str = ""
    decision_maker_name: str = "Not found"
    decision_maker_title: str = ""
    decision_maker_linkedin: str = ""
    public_email_or_contact: str = ""
    phone_number: str = ""
    source_url_1: str = ""
    source_url_2: str = ""
    lead_score: int = 0
    # internal fields
    icp_signals_matched: list = field(default_factory=list)
    notes: str = ""

    def to_row(self) -> dict:
        return {
            "Company Name": self.company_name,
            "Website": self.website,
            "Country": self.country,
            "Industry": self.industry,
            "Estimated Employee Size": self.employee_size,
            "ICP Segment": self.icp_segment,
            "Buying Signal": self.buying_signal,
            "Why This Is a Quality Lead": self.why_quality,
            "Best Service to Pitch": self.best_service,
            "Decision Maker Name": self.decision_maker_name,
            "Decision Maker Title": self.decision_maker_title,
            "Decision Maker LinkedIn/Profile URL": self.decision_maker_linkedin,
            "Public Email or Contact Page URL": self.public_email_or_contact,
            "Phone Number": self.phone_number,
            "Source URL 1": self.source_url_1,
            "Source URL 2": self.source_url_2,
            "Lead Score 1-10": self.lead_score,
        }


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

class Fetcher:
    """Polite HTTP fetcher with rate limiting + retries."""

    def __init__(self, delay: float = 1.5, max_retries: int = 2):
        self.delay = delay
        self.max_retries = max_retries
        self.session = requests.Session()
        self.session.headers.update(HEADERS)

    def get(self, url: str, **kw) -> Optional[requests.Response]:
        for attempt in range(self.max_retries + 1):
            try:
                time.sleep(self.delay + random.uniform(0, 0.5))
                r = self.session.get(url, timeout=20, **kw)
                if r.status_code == 200:
                    return r
                if r.status_code in (429, 503):
                    time.sleep(8 + attempt * 4)
                    continue
                if 400 <= r.status_code < 500:
                    return None
            except requests.exceptions.ConnectionError as e:
                # DNS failure or refused connection — no point retrying
                print(f"  [fetch error] {url}: {e}", file=sys.stderr)
                return None
            except requests.RequestException as e:
                print(f"  [fetch error] {url}: {e}", file=sys.stderr)
                time.sleep(2)
        return None


# ---------------------------------------------------------------------------
# Source 1: Y Combinator companies (public Algolia search index)
# ---------------------------------------------------------------------------

def fetch_yc_api_key() -> tuple[str, str]:
    """Fetch current Algolia app ID and API key from YC companies page (rotates periodically)."""
    try:
        r = requests.get(
            "https://www.ycombinator.com/companies",
            headers={"User-Agent": USER_AGENT},
            timeout=20,
        )
        import json as _json
        scripts = re.findall(r"<script[^>]*>(.*?)</script>", r.text, re.DOTALL)
        for s in scripts:
            if "AlgoliaOpts" in s:
                m = re.search(r"window\.AlgoliaOpts\s*=\s*(\{[^;]+\})", s)
                if m:
                    opts = _json.loads(m.group(1))
                    return opts["app"], opts["key"]
    except Exception as e:
        print(f"  [warn] could not fetch YC API key dynamically: {e}", file=sys.stderr)
    # fallback (may be expired)
    return "45BWZJ1SGC", "fa5638a7042525116f6e8527d9b4c11a"


def fetch_yc_companies(fetcher: Fetcher, max_pages: int = 5) -> list[dict]:
    """
    YC publishes their company directory via an Algolia public search index.
    The API key is fetched dynamically from the YC page to handle rotation.
    Returns dicts with: name, website, one_liner, batch, team_size, industries,
    regions, tags.
    """
    app_id, api_key = fetch_yc_api_key()
    url = f"https://{app_id.lower()}-dsn.algolia.net/1/indexes/*/queries"
    headers = {
        "X-Algolia-API-Key": api_key,
        "X-Algolia-Application-Id": app_id,
        "Content-Type": "application/json",
    }
    out = []
    for page in range(max_pages):
        body = {
            "requests": [{
                "indexName": "YCCompany_production",
                "params": (
                    f"facetFilters=%5B%5B%22tags%3AB2B%22%5D%5D"
                    f"&hitsPerPage=100"
                    f"&page={page}"
                ),
            }]
        }
        try:
            time.sleep(fetcher.delay)
            r = requests.post(url, headers=headers, json=body, timeout=20)
            if r.status_code != 200:
                print(f"  YC API page {page} returned {r.status_code}", file=sys.stderr)
                break
            hits = r.json()["results"][0].get("hits", [])
            if not hits:
                break
            out.extend(hits)
            print(f"  YC page {page}: {len(hits)} companies")
        except Exception as e:
            print(f"  YC fetch error: {e}", file=sys.stderr)
            break
    return out


# ---------------------------------------------------------------------------
# Source 2: DesignRush agency directory
# ---------------------------------------------------------------------------

DESIGNRUSH_CATEGORIES = [
    "digital-marketing",
    "content-marketing",
    "inbound-marketing",
    "marketing",
]

DESIGNRUSH_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}


def fetch_designrush_agencies(fetcher: Fetcher, category: str = "digital-marketing",
                               max_pages: int = 3) -> list[dict]:
    """
    DesignRush lists verified B2B agencies with employee size, location, and
    service categories — ideal for finding marketing agencies to pitch
    Get Levrg fulfillment services to.
    """
    out = []
    for page in range(1, max_pages + 1):
        url = f"https://www.designrush.com/agency/{category}?page={page}" if page > 1 \
              else f"https://www.designrush.com/agency/{category}"
        try:
            time.sleep(fetcher.delay)
            r = requests.get(url, headers=DESIGNRUSH_HEADERS, timeout=20)
            if r.status_code != 200:
                print(f"  DesignRush {category} page {page}: status {r.status_code}",
                      file=sys.stderr)
                break
            soup = BeautifulSoup(r.text, "lxml")
            articles = soup.select("article[data-agency-name]")
            if not articles:
                break
            for a in articles:
                name = a.get("data-agency-name", "").strip()
                profile_url = ""
                link = a.select_one("a.gtm-agency-profile-link")
                if link:
                    profile_url = link.get("href", "")
                # Location: first span that has a comma (city,state pattern)
                location = ""
                for span in a.select("span"):
                    t = span.get_text(strip=True)
                    if "," in t and len(t) < 60 and not any(
                            c.isdigit() for c in t[:3]):
                        location = t
                        break
                # Employee size: span matching common size patterns
                size = ""
                for span in a.select("span"):
                    t = span.get_text(strip=True)
                    if re.match(r"(Under \d+|\d+ - \d+|\d+\+)", t):
                        size = t
                        break
                if name:
                    out.append({
                        "name": name,
                        "website": "",
                        "location": location,
                        "employee_size": size,
                        "profile_url": profile_url,
                        "source": f"designrush_{category}",
                    })
            print(f"  DesignRush {category} page {page}: {len(articles)} agencies")
            if len(articles) < 10:
                break
        except Exception as e:
            print(f"  DesignRush fetch error: {e}", file=sys.stderr)
            break
    return out


# ---------------------------------------------------------------------------
# Source 3: (Clutch removed — blocked by anti-bot; replaced by DesignRush above)


# ---------------------------------------------------------------------------
# Company website parser — find decision makers, contact, signals
# ---------------------------------------------------------------------------

TEAM_PATHS = ["/team", "/about", "/about-us", "/people", "/leadership",
              "/our-team", "/company", "/who-we-are"]
CONTACT_PATHS = ["/contact", "/contact-us", "/get-in-touch", "/hello"]
CAREERS_PATHS = ["/careers", "/jobs", "/join-us", "/work-with-us", "/hiring"]
BLOG_PATHS = ["/blog", "/insights", "/resources", "/articles"]


def parse_company_site(fetcher: Fetcher, website: str) -> dict:
    """
    Visit a company website and extract:
      - Decision maker name + title + LinkedIn (from team / about pages)
      - Public email OR contact page URL
      - Buying signals: hiring (careers page open jobs), active blog,
        HubSpot/CRM mentions, podcast, video content
      - Country (from contact page if shown)
    """
    if not website.startswith("http"):
        website = "https://" + website.lstrip("/")
    parsed = urlparse(website)
    base = f"{parsed.scheme}://{parsed.netloc}"

    result = {
        "decision_maker_name": "",
        "decision_maker_title": "",
        "decision_maker_linkedin": "",
        "email": "",
        "phone": "",
        "contact_url": "",
        "signals": [],
        "country_hint": "",
        "team_url": "",
    }

    def clean_emails(raw: list[str]) -> list[str]:
        return [e for e in raw if not any(f in e.lower() for f in FAKE_EMAIL_FRAGMENTS)]

    def clean_phones(raw: list[str]) -> list[str]:
        out = []
        for p in raw:
            digits = re.sub(r"\D", "", p)
            if not (10 <= len(digits) <= 15):
                continue
            if len(set(digits)) <= 2:
                continue
            if digits.startswith("000") or digits.startswith("111"):
                continue
            if len(digits) == 10 and digits[0] in "01":
                continue
            # Reject INT_MAX and similar code constants
            if digits in ("2147483647", "4294967295", "9999999999", "1234567890"):
                continue
            # Must have at least 4 distinct digits (real phone numbers vary more)
            if len(set(digits)) < 4:
                continue
            out.append(p.strip())
        return out

    def extract_tel_links(soup_or_text) -> str:
        """Prefer <a href='tel:...'> tags — most reliable phone source."""
        if hasattr(soup_or_text, "select"):
            for a in soup_or_text.select("a[href^='tel:']"):
                num = a["href"].replace("tel:", "").strip()
                digits = re.sub(r"\D", "", num)
                if 10 <= len(digits) <= 15:
                    return num
        return ""

    def best_email(emails: list[str]) -> str:
        preferred = [e for e in emails if any(
            p in e.lower() for p in ("info@", "hello@", "contact@", "sales@", "hi@"))]
        return (preferred or emails)[0] if emails else ""

    # Try homepage first for general signals; bail entirely if DNS fails
    home = fetcher.get(base)
    if not home:
        return result
    home_text = home.text
    text = home_text.lower()
    if "hubspot" in text:
        result["signals"].append("mentions_hubspot")
    if "podcast" in text:
        result["signals"].append("has_podcast")
    if "case stud" in text:
        result["signals"].append("publishes_case_studies")
    if "newsletter" in text or "subscribe" in text:
        result["signals"].append("has_newsletter")

    # Extract email + phone from homepage (footer often has them)
    home_soup = BeautifulSoup(home_text, "lxml")
    home_emails = clean_emails(EMAIL_RE.findall(home_text))
    if home_emails:
        result["email"] = best_email(home_emails)
    # tel: links are the most reliable phone source
    tel = extract_tel_links(home_soup)
    if tel:
        result["phone"] = tel
    else:
        home_phones = clean_phones(PHONE_RE.findall(home_text))
        if home_phones:
            result["phone"] = home_phones[0]

    # Team page -> decision maker
    for path in TEAM_PATHS:
        url = base + path
        r = fetcher.get(url)
        if not r:
            continue
        result["team_url"] = url
        soup = BeautifulSoup(r.text, "lxml")

        # Strategy 1: find title element, then look for adjacent name
        candidates = []
        for el in soup.find_all(["h1", "h2", "h3", "h4", "p", "div", "span"]):
            txt = el.get_text(" ", strip=True)
            if 0 < len(txt) <= 80 and DECISION_TITLES.search(txt):
                candidates.append((el, txt))

        if candidates:
            el, title_text = candidates[0]
            name = ""

            # Strategy 1: adjacent heading sibling
            for sib in [
                el.find_previous(["h1", "h2", "h3", "h4"]),
                el.find_next(["h1", "h2", "h3", "h4"]),
                el.parent.find(["h1", "h2", "h3", "h4"]) if el.parent else None,
            ]:
                if sib:
                    s = sib.get_text(" ", strip=True)
                    if looks_like_person_name(s) and not DECISION_TITLES.search(s):
                        name = s
                        break

            # Strategy 2: card wrapper pattern
            if not name:
                for card_sel in ("[class*=card]", "[class*=member]", "[class*=person]",
                                 "[class*=team]", "li", "article"):
                    for card in soup.select(card_sel):
                        card_txt = card.get_text(" ", strip=True)
                        if not DECISION_TITLES.search(card_txt):
                            continue
                        for tag in card.find_all(["h1", "h2", "h3", "h4", "strong"]):
                            s = tag.get_text(" ", strip=True)
                            if looks_like_person_name(s) and not DECISION_TITLES.search(s):
                                name = s
                                break
                        if name:
                            break
                    if name:
                        break

            # LinkedIn near this element
            li_url = ""
            container = el.parent or el
            for a in container.find_all("a", href=True):
                if "linkedin.com/in/" in a["href"]:
                    li_url = a["href"]
                    break

            result["decision_maker_name"] = name
            result["decision_maker_title"] = title_text.strip()
            result["decision_maker_linkedin"] = li_url

            # Also grab email / phone from team page if not found yet
            if not result["email"]:
                te = clean_emails(EMAIL_RE.findall(r.text))
                if te:
                    result["email"] = best_email(te)
            if not result["phone"]:
                t_tel = extract_tel_links(soup)
                if t_tel:
                    result["phone"] = t_tel
                else:
                    tp = clean_phones(PHONE_RE.findall(r.text))
                    if tp:
                        result["phone"] = tp[0]
            break

    # Contact page -> email, phone, URL (overrides homepage values if better)
    for path in CONTACT_PATHS:
        url = base + path
        r = fetcher.get(url)
        if not r:
            continue
        result["contact_url"] = url
        c_soup = BeautifulSoup(r.text, "lxml")
        c_emails = clean_emails(EMAIL_RE.findall(r.text))
        if c_emails:
            result["email"] = best_email(c_emails)
        # Prefer tel: links over regex matches on contact page
        c_tel = extract_tel_links(c_soup)
        if c_tel:
            result["phone"] = c_tel
        elif not result["phone"]:
            c_phones = clean_phones(PHONE_RE.findall(r.text))
            if c_phones:
                result["phone"] = c_phones[0]
        # Country hint
        for word in TARGET_COUNTRIES:
            if word.lower() in r.text.lower():
                result["country_hint"] = word
                break
        break

    # Careers page -> hiring signal
    for path in CAREERS_PATHS:
        url = base + path
        r = fetcher.get(url)
        if not r or r.status_code != 200:
            continue
        text = r.text.lower()
        marketing_roles = ["marketing", "content", "social media", "video editor",
                           "crm", "hubspot", "sales ops", "revops",
                           "revenue operations", "sdr", "demand gen"]
        matched = [t for t in marketing_roles if t in text]
        if matched:
            result["signals"].append(f"hiring:{','.join(matched[:3])}")
        break

    # Blog presence
    for path in BLOG_PATHS:
        url = base + path
        r = fetcher.get(url)
        if r and r.status_code == 200:
            result["signals"].append("active_blog")
            break

    return result


# ---------------------------------------------------------------------------
# ICP scoring
# ---------------------------------------------------------------------------

def score_lead(lead: Lead, raw: dict, site: dict) -> int:
    """
    Score 0-10. We only KEEP leads scoring 7+.
    """
    score = 0
    matched = []

    # Signal 1: B2B
    industry_text = (lead.industry + " " + raw.get("one_liner", "")).lower()
    if any(k in industry_text for k in ("b2b", "saas", "agency", "consulting",
                                          "professional service", "revops", "marketing")):
        score += 1
        matched.append("b2b")

    # Signal 2: Employee size band
    sz = raw.get("team_size") or raw.get("employee_size") or ""
    try:
        size_n = int(re.findall(r"\d+", str(sz))[0]) if sz else 0
    except Exception:
        size_n = 0
    if 10 <= size_n <= 250:
        score += 2
        matched.append("size_10_250")
    elif size_n and size_n < 500:
        score += 1
        matched.append("size_under_500")

    # Signal 3: target country
    country = lead.country or raw.get("regions", [""])[0] if isinstance(raw.get("regions"), list) else lead.country
    if any(c.lower() in country.lower() for c in TARGET_COUNTRIES):
        score += 1
        matched.append("target_country")

    # Signal 4-6: site signals (active marketing, hiring, hubspot)
    sigs = site.get("signals", [])
    if any(s.startswith("hiring:") for s in sigs):
        score += 2
        matched.append("hiring_signal")
    if "active_blog" in sigs:
        score += 1
        matched.append("active_blog")
    if "mentions_hubspot" in sigs:
        score += 1
        matched.append("hubspot_user")
    if "has_podcast" in sigs or "publishes_case_studies" in sigs:
        score += 1
        matched.append("content_marketing")

    # Signal 7: decision maker found
    if site.get("decision_maker_name"):
        score += 1
        matched.append("decision_maker_found")

    # Hard exclusions
    text_blob = (industry_text + " " + raw.get("one_liner", "")).lower()
    if any(b in text_blob for b in BAD_KEYWORDS):
        return 0  # exclude

    lead.icp_signals_matched = matched
    return min(score, 10)


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

def normalize_yc(c: dict) -> dict:
    """Normalize a YC Algolia hit into our common candidate format."""
    return {
        "name": c.get("name", ""),
        "website": c.get("website", ""),
        "country": (c.get("regions") or [""])[0] if isinstance(c.get("regions"), list) else "",
        "industry": ", ".join(c.get("industries") or [])[:100],
        "employee_size": str(c.get("team_size") or ""),
        "one_liner": c.get("one_liner", ""),
        "batch": c.get("batch", ""),
        "tags": c.get("tags", []),
        "regions": c.get("regions", []),
        "source": "yc",
        "source_url": f"https://www.ycombinator.com/companies/{c.get('slug', '')}",
    }


def normalize_designrush(c: dict) -> dict:
    category = c.get("source", "designrush").replace("designrush_", "").replace("-", " ").title()
    return {
        "name": c["name"],
        "website": c.get("website", ""),
        "country": c.get("location", ""),
        "industry": f"Marketing Agency ({category})",
        "employee_size": c.get("employee_size", ""),
        "one_liner": f"Verified {category} agency listed on DesignRush.",
        "source": c.get("source", "designrush"),
        "source_url": c.get("profile_url", ""),
    }


def build_lead(raw: dict, site: dict) -> Lead:
    lead = Lead()
    lead.company_name = raw["name"]
    lead.website = raw.get("website", "") or site.get("homepage", "")
    lead.country = raw.get("country", "") or site.get("country_hint", "")
    lead.industry = raw.get("industry", "")
    lead.employee_size = raw.get("employee_size", "")

    # ICP segment
    src = raw.get("source", "")
    if "designrush" in src:
        lead.icp_segment = "Marketing services agency"
        lead.best_service = "Agency fulfillment + content production + video editing"
    else:
        lead.icp_segment = "B2B SaaS / professional services"
        lead.best_service = "Fractional marketing team + HubSpot/CRM ops"

    # Buying signal narrative
    sigs = site.get("signals", [])
    if any(s.startswith("hiring:") for s in sigs):
        roles = [s.split(":", 1)[1] for s in sigs if s.startswith("hiring:")][0]
        lead.buying_signal = f"Hiring for: {roles}. Indicates active marketing investment without enough internal capacity."
    elif "active_blog" in sigs and "mentions_hubspot" in sigs:
        lead.buying_signal = "Publishes actively + uses HubSpot. Likely needs content/CRM execution support."
    elif "active_blog" in sigs:
        lead.buying_signal = "Active blog and content output — content production help is plausible."
    else:
        lead.buying_signal = "Visible content marketing activity per site."

    lead.why_quality = (
        f"Matches {len(lead.icp_signals_matched) if hasattr(lead,'icp_signals_matched') else 0} ICP signals: "
        + ", ".join(getattr(lead, "icp_signals_matched", []))
    )

    # Decision maker
    if site.get("decision_maker_name"):
        lead.decision_maker_name = site["decision_maker_name"]
        lead.decision_maker_title = site["decision_maker_title"]
        lead.decision_maker_linkedin = site["decision_maker_linkedin"]

    # Contact
    if site.get("email"):
        lead.public_email_or_contact = site["email"]
    elif site.get("contact_url"):
        lead.public_email_or_contact = site["contact_url"]

    # Phone
    lead.phone_number = site.get("phone", "")

    # Sources
    lead.source_url_1 = raw.get("source_url", "")
    lead.source_url_2 = site.get("team_url", "") or site.get("contact_url", "")

    return lead


def run(args):
    fetcher = Fetcher(delay=args.delay)
    candidates: list[dict] = []

    print("=" * 60)
    print("STAGE 1: Collect candidates from public directories")
    print("=" * 60)

    if "yc" in args.sources:
        print("\n[YC] Fetching B2B companies from Y Combinator...")
        for c in fetch_yc_companies(fetcher, max_pages=args.yc_pages):
            candidates.append(normalize_yc(c))

    if "designrush" in args.sources:
        print("\n[DesignRush] Fetching B2B marketing agency directory...")
        for cat in ["digital-marketing", "content-marketing", "inbound-marketing"]:
            for c in fetch_designrush_agencies(fetcher, category=cat, max_pages=args.designrush_pages):
                candidates.append(normalize_designrush(c))

    # Dedupe by website / name
    seen = set()
    deduped = []
    for c in candidates:
        key = (c.get("website") or c.get("name", "")).lower().strip()
        if key and key not in seen:
            seen.add(key)
            deduped.append(c)
    print(f"\nTotal unique candidates after dedupe: {len(deduped)}")

    # Save raw
    raw_path = os.path.join(args.outdir, "leads_raw.json")
    with open(raw_path, "w") as f:
        json.dump(deduped, f, indent=2)
    print(f"Saved raw candidates -> {raw_path}")

    print("\n" + "=" * 60)
    print("STAGE 2: Enrich each candidate by visiting their website")
    print("=" * 60)

    leads: list[Lead] = []
    for i, raw in enumerate(deduped):
        if len(leads) >= args.target_leads * 3:
            print("Enough enriched candidates; stopping enrichment.")
            break
        site_url = raw.get("website") or ""
        if not site_url:
            continue
        print(f"\n[{i+1}/{len(deduped)}] {raw.get('name','?')} — {site_url}")
        try:
            site = parse_company_site(fetcher, site_url)
        except Exception as e:
            print(f"  parse error: {e}", file=sys.stderr)
            continue
        lead = build_lead(raw, site)
        score = score_lead(lead, raw, site)
        lead.lead_score = score
        if score >= args.min_score:
            leads.append(lead)
            print(f"  [KEPT] score={score}")
        else:
            print(f"  [skip] score={score}")

    print("\n" + "=" * 60)
    print(f"STAGE 3: Final filter — keep top {args.target_leads} by score")
    print("=" * 60)

    leads.sort(key=lambda x: x.lead_score, reverse=True)
    final = leads[: args.target_leads]
    print(f"Final count: {len(final)} leads")

    # Save CSV
    csv_path = os.path.join(args.outdir, "leads_final.csv")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        if final:
            writer = csv.DictWriter(f, fieldnames=list(final[0].to_row().keys()))
            writer.writeheader()
            for lead in final:
                writer.writerow(lead.to_row())
    print(f"CSV -> {csv_path}")

    # Save XLSX with formatting
    if pd is not None and final:
        xlsx_path = os.path.join(args.outdir, "leads_final.xlsx")
        df = pd.DataFrame([lead.to_row() for lead in final])
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Leads")
            ws = writer.sheets["Leads"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            # Set column widths
            widths = {
                "A": 28, "B": 32, "C": 16, "D": 28, "E": 14, "F": 32,
                "G": 50, "H": 50, "I": 35, "J": 24, "K": 28, "L": 36,
                "M": 36, "N": 20, "O": 36, "P": 36, "Q": 8,
            }
            for col, w in widths.items():
                ws.column_dimensions[col].width = w
        print(f"XLSX -> {xlsx_path}")


def main():
    ap = argparse.ArgumentParser(description="Get Levrg lead scraper")
    ap.add_argument("--target-leads", type=int, default=50)
    ap.add_argument("--min-score", type=int, default=7)
    ap.add_argument("--delay", type=float, default=1.5,
                    help="Seconds between requests (be polite)")
    ap.add_argument("--sources", nargs="+",
                    default=["yc", "designrush"],
                    choices=["yc", "designrush"])
    ap.add_argument("--yc-pages", type=int, default=5)
    ap.add_argument("--designrush-pages", type=int, default=3)
    ap.add_argument("--outdir", default=".")
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    run(args)


if __name__ == "__main__":
    main()