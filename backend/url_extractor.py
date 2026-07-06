"""
url_extractor.py – Scrape a company website URL and extract:
  • company_name  – the organisation name
  • about_us      – a short description / about-us paragraph

Strategy (in order):
  1. Fetch the homepage HTML.
  2. Look for structured data (JSON-LD Organization / WebSite) → name + description.
  3. Fallback: <meta og:site_name>, <meta og:description>, <title>, <meta description>.
  4. If an /about or /about-us page exists, fetch it and extract the first meaty paragraph.
  5. Final fallback: return what we could find, empty strings if nothing.
"""

import re
import json
import requests
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}
TIMEOUT = 10  # seconds


def _fetch(url: str) -> BeautifulSoup | None:
    """Fetch URL and return BeautifulSoup object, or None on error."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        resp.raise_for_status()
        return BeautifulSoup(resp.text, "lxml")
    except Exception as e:
        print(f"[url_extractor] fetch error for {url}: {e}")
        return None


def _clean(text: str) -> str:
    """Collapse whitespace and strip."""
    return re.sub(r"\s+", " ", text or "").strip()


def _from_json_ld(soup: BeautifulSoup) -> dict:
    """Extract name/description from JSON-LD Organization or WebSite schema."""
    result = {"company_name": "", "about_us": ""}
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or "")
            # data can be a list or dict
            items = data if isinstance(data, list) else [data]
            for item in items:
                graph = item.get("@graph", [item])
                for node in graph:
                    t = node.get("@type", "")
                    types = t if isinstance(t, list) else [t]
                    if any(x in ("Organization", "Corporation", "LocalBusiness", "WebSite") for x in types):
                        if not result["company_name"]:
                            result["company_name"] = _clean(node.get("name", ""))
                        if not result["about_us"]:
                            result["about_us"] = _clean(
                                node.get("description", "") or node.get("slogan", "")
                            )
        except Exception:
            continue
    return result


def _from_meta(soup: BeautifulSoup) -> dict:
    """Extract name/description from meta tags and <title>."""
    def _meta(prop_or_name: str) -> str:
        tag = soup.find("meta", attrs={"property": prop_or_name}) or \
              soup.find("meta", attrs={"name": prop_or_name})
        return _clean(tag["content"]) if tag and tag.get("content") else ""

    og_site   = _meta("og:site_name")
    og_desc   = _meta("og:description")
    meta_desc = _meta("description")
    title_tag = soup.find("title")
    title     = _clean(title_tag.get_text()) if title_tag else ""

    # Company name heuristic: og:site_name > first part of <title> before | or -
    name = og_site
    if not name and title:
        name = re.split(r"\s[|\-–—]\s", title)[0].strip()

    about = og_desc or meta_desc
    return {"company_name": name, "about_us": about}


def _first_meaty_paragraph(soup: BeautifulSoup, min_words: int = 20) -> str:
    """Return the first paragraph with at least min_words words."""
    # Remove nav, header, footer, aside, script, style noise
    for tag in soup(["nav", "header", "footer", "aside", "script", "style", "noscript"]):
        tag.decompose()

    for p in soup.find_all("p"):
        text = _clean(p.get_text())
        if len(text.split()) >= min_words:
            return text
    return ""


def _find_about_url(base_url: str, soup: BeautifulSoup) -> str | None:
    """Heuristically find an /about or /about-us link."""
    patterns = re.compile(r"about[\-_]?(us|company|operisoft)?$", re.I)
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        path = urlparse(href).path.rstrip("/").split("/")[-1]
        if patterns.search(path):
            return urljoin(base_url, href)
    return None


def extract_company_info(url: str) -> dict:
    """
    Main entry point.  Returns:
      {
        "company_name": str,
        "about_us":     str,
        "source_url":   str,
        "error":        str | None
      }
    """
    # Normalise URL
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    parsed   = urlparse(url)
    base_url = f"{parsed.scheme}://{parsed.netloc}"

    soup = _fetch(url)
    if soup is None:
        return {"company_name": "", "about_us": "", "source_url": url,
                "error": "Could not fetch the URL. Check the address and try again."}

    # 1. JSON-LD
    info = _from_json_ld(soup)

    # 2. Meta tags fill gaps
    meta = _from_meta(soup)
    if not info["company_name"]:
        info["company_name"] = meta["company_name"]
    if not info["about_us"]:
        info["about_us"] = meta["about_us"]

    # 3. Try /about page for a richer description
    about_url = _find_about_url(base_url, soup)
    if about_url:
        about_soup = _fetch(about_url)
        if about_soup:
            # Try JSON-LD on about page first
            about_ld = _from_json_ld(about_soup)
            if about_ld["about_us"]:
                info["about_us"] = about_ld["about_us"]
            elif not info["about_us"]:
                info["about_us"] = _first_meaty_paragraph(about_soup)
            # Grab name from about page if still missing
            if not info["company_name"]:
                info["company_name"] = _from_meta(about_soup)["company_name"]

    # 4. Last-resort: first meaty paragraph from homepage
    if not info["about_us"]:
        info["about_us"] = _first_meaty_paragraph(soup)

    # Trim about_us to a reasonable length (~500 chars) for the SOW prompt
    if len(info["about_us"]) > 600:
        info["about_us"] = info["about_us"][:600].rsplit(" ", 1)[0] + "…"

    info["source_url"] = url
    info["error"] = None
    return info
