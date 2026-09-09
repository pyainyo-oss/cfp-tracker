"""
============================================================
 Myanmar CFP Tracker – v2.0
 Sources: MIMU, fundsforNGOs, Advance Africa, DevelopmentAid,
          K4DM, UNFPA, US Embassy, Japan GGP, Lorcan Lovett,
          GrantStation
 Output:  CFP_Tracker_Myanmar.xlsx
============================================================
"""

import requests
import re
import os
import xml.etree.ElementTree as ET
from datetime import datetime, date
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

try:
    import feedparser
except ImportError:
    feedparser = None

# ──────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────
EXCEL_FILE = "CFP_Tracker_Myanmar.xlsx"
SHEET_NAME = "CFPs"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
}

# All sources with their type
SOURCES = [
    # ─── Aggregators (HTML table/list) ───
    {"name": "MIMU", "url": "https://themimu.info/calls-for-proposals", "type": "html"},
    {"name": "fundsforNGOs", "url": "https://www2.fundsforngos.org/tag/burmamyanmar/", "type": "html"},
    {"name": "Advance Africa", "url": "https://www.advance-africa.com/Grants-for-NGOs-in-Myanmar.html", "type": "html"},
    {"name": "DevelopmentAid", "url": "https://www.developmentaid.org/tenders/search?locations=143", "type": "html"},
    {"name": "GrantStation", "url": "https://grantstation.com/find-grants/myanmar-burma", "type": "html"},

    # ─── Government / Embassy (HTML) ───
    {"name": "US Embassy Burma", "url": "https://mm.usembassy.gov/grants-and-fellowships/", "type": "html"},
    {"name": "Japan GGP", "url": "https://www.mm.emb-japan.go.jp/profile/english/ggp.htm", "type": "html"},
    {"name": "UNFPA Myanmar", "url": "https://myanmar.unfpa.org/en/call-for-submissions", "type": "html"},

    # ─── Research (RSS feeds – easier to parse) ───
    {"name": "K4DM", "url": "https://k4dm.ca/feed/", "type": "rss"},
    {"name": "Lorcan Lovett", "url": "https://lorcanlovett.substack.com/feed", "type": "rss"},
]


# ──────────────────────────────────────────────
# FETCH
# ──────────────────────────────────────────────

def fetch(url):
    """Fetch URL content with browser-like headers."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        r.raise_for_status()
        return r.text
    except Exception as e:
        print(f"  [ERROR] {url}: {e}")
        return ""


# ──────────────────────────────────────────────
# PARSERS – HTML Sources
# ──────────────────────────────────────────────

def parse_mimu(html):
    """Parse MIMU CFP table (Drupal views table)."""
    entries = []
    if not html:
        return entries
    soup = BeautifulSoup(html, "html.parser")

    # MIMU uses a table with columns: Title | App Form | Geo | Source | Format | Size | Deadline
    table = soup.find("table", class_="views-table")
    if not table:
        # Fallback: find any table with CFP-related content
        for t in soup.find_all("table"):
            text = t.get_text().lower()
            if "deadline" in text and ("call" in text or "proposal" in text or "tender" in text):
                table = t
                break
    if not table:
        return entries

    for row in table.find_all("tr")[1:]:  # skip header row
        cells = row.find_all("td")
        if len(cells) < 4:
            continue

        # Column 0: Title (with link)
        link = cells[0].find("a")
        if not link:
            continue
        title = link.get_text(strip=True)
        url = link.get("href", "")
        if not url.startswith("http"):
            url = "https://themimu.info" + url

        # Column 2: Geographic Coverage
        geo = cells[2].get_text(strip=True) if len(cells) > 2 else ""

        # Column 3: Data Source (funder)
        funder = cells[3].get_text(strip=True) if len(cells) > 3 else ""

        # Last column: Deadline
        deadline_raw = cells[-1].get_text(strip=True)
        m = re.search(r'(\d{1,2}-[A-Za-z]{3}-\d{4})', deadline_raw)
        deadline = m.group(1) if m else ""

        entries.append({
            "title": title,
            "funder": funder,
            "location": geo,
            "deadline": deadline,
            "amount": "",
            "url": url,
            "source": "MIMU"
        })
    return entries


def parse_fundsforngos(html):
    """Parse fundsforNGOs Myanmar tag page."""
    entries = []
    if not html:
        return entries
    soup = BeautifulSoup(html, "html.parser")

    # Look for article titles (h2 or h3 with links)
    for heading in soup.find_all(["h2", "h3"]):
        a = heading.find("a")
        if not a:
            continue
        title = a.get_text(strip=True)
        url = a.get("href", "")
        if not url.startswith("http"):
            url = "https://www2.fundsforngos.org" + url

        # Check relevance
        if not re.search(r'myanmar|burma', title, re.I):
            # Check parent context
            parent = a.find_parent("div", class_=re.compile(r"post|article|card|teaser"))
            if not parent or not re.search(r'myanmar|burma', parent.get_text(), re.I):
                continue

        # Try to extract deadline from surrounding text
        parent = a.find_parent("div", class_=re.compile(r"post|article|card|teaser"))
        body_text = parent.get_text() if parent else ""

        deadline = ""
        m = re.search(r'[Dd]eadline:?\s*(\d{1,2}-[A-Za-z]{3}-\d{4})', body_text)
        if m:
            deadline = m.group(1)

        # Extract funder (heuristic)
        funder = "Unknown"
        fm = re.search(r'(?:by|from|The)\s+([A-Z][A-Za-z&.\-]+(?:\s+[A-Z][A-Za-z&.\-]+){0,4})', body_text)
        if fm:
            funder = fm.group(1)

        # Extract amount
        amount = ""
        am = re.search(r'(?:USD|US\$|\$|MMK|€|EUR)\s?[\d,]+(?:\s*(?:to|-|–)\s*(?:USD|US\$|\$|MMK|€|EUR)\s?[\d,]+)?', body_text)
        if am:
            amount = am.group(0)

        entries.append({
            "title": title,
            "funder": funder,
            "location": "Myanmar",
            "deadline": deadline,
            "amount": amount,
            "url": url,
            "source": "fundsforNGOs"
        })
    return entries


def parse_advance_africa(html):
    """Parse Advance Africa Myanmar grants page."""
    entries = []
    if not html:
        return entries
    soup = BeautifulSoup(html, "html.parser")
    seen = set()

    for a in soup.find_all("a", href=True):
        url = a["href"]
        title = a.get_text(strip=True)

        if not re.search(r'myanmar|burma', title, re.I):
            continue
        if len(title) < 10 or url in seen:
            continue
        if not url.startswith("http"):
            url = "https://www.advance-africa.com" + url
        # Skip nav/footer
        if re.search(r'(home|about|contact|privacy|terms)', url, re.I):
            continue

        seen.add(url)
        entries.append({
            "title": title,
            "funder": "Advance Africa (aggregator)",
            "location": "Myanmar",
            "deadline": "",
            "amount": "",
            "url": url,
            "source": "Advance Africa"
        })
    return entries


def parse_developmentaid(html):
    """Parse DevelopmentAid Myanmar tenders/grants page."""
    entries = []
    if not html:
        return entries
    soup = BeautifulSoup(html, "html.parser")

    # DevelopmentAid uses card-based layout with title links
    # Look for links that contain tender/grant/consult keywords
    seen = set()
    for a in soup.find_all("a", href=True):
        url = a["href"]
        title = a.get_text(strip=True)

        # Must be a tender/grant detail page
        if not re.search(r'/tenders/view/', url):
            continue
        if len(title) < 10 or url in seen:
            continue

        # Get context from parent card
        parent = a.find_parent("div", class_=re.compile(r"card|item|result|tender"))
        context = parent.get_text() if parent else ""

        # Check if Myanmar-related
        if not re.search(r'myanmar|burma', context, re.I):
            continue

        seen.add(url)
        if not url.startswith("http"):
            url = "https://www.developmentaid.org" + url

        # Extract deadline
        deadline = ""
        m = re.search(r'Deadline:?\s*([A-Z][a-z]{2}\s+\d{1,2},\s*\d{4})', context)
        if m:
            deadline = m.group(1)

        # Extract budget
        amount = ""
        bm = re.search(r'Budget:?\s*(USD\s?[\d,]+|N/A)', context)
        if bm and bm.group(1) != "N/A":
            amount = bm.group(1)

        # Extract funding agency
        funder = "Unknown"
        fm = re.search(r'Funding agency:?\s*([A-Z][A-Za-z&\.\-]+(?:\s+[A-Z][A-Za-z&\.\-]+)*)', context)
        if fm:
            funder = fm.group(1)

        # Extract category
        category = ""
        cm = re.search(r'Category:?\s*([A-Za-z &]+?)(?:\n|Posted:|$)', context)
        if cm:
            category = cm.group(1).strip()

        entries.append({
            "title": title,
            "funder": funder,
            "location": "Myanmar",
            "deadline": deadline,
            "amount": amount,
            "url": url,
            "source": f"DevelopmentAid ({category})" if category else "DevelopmentAid"
        })
    return entries


def parse_grantstation(html):
    """Parse GrantStation Myanmar grants page."""
    entries = []
    if not html:
        return entries
    soup = BeautifulSoup(html, "html.parser")

    seen = set()
    for a in soup.find_all("a", href=True):
        url = a["href"]
        title = a.get_text(strip=True)

        if not re.search(r'myanmar|burma', title, re.I):
            continue
        if len(title) < 15 or url in seen:
            continue
        if not url.startswith("http"):
            url = "https://grantstation.com" + url
        # Skip non-grant links
        if re.search(r'(login|register|pricing|about|blog|faq)', url, re.I):
            continue

        seen.add(url)
        entries.append({
            "title": title,
            "funder": "GrantStation (aggregator)",
            "location": "Myanmar",
            "deadline": "",
            "amount": "",
            "url": url,
            "source": "GrantStation"
        })
    return entries


def parse_us_embassy(html):
    """Parse US Embassy Burma grants & fellowships page."""
    entries = []
    if not html:
        return entries
    soup = BeautifulSoup(html, "html.parser")

    seen = set()
    for a in soup.find_all("a", href=True):
        url = a["href"]
        title = a.get_text(strip=True)

        if not re.search(r'grant|fellowship|NOFO|call for|competition|small grant', title, re.I):
            continue
        if len(title) < 10 or url in seen:
            continue
        if not url.startswith("http"):
            url = "https://mm.usembassy.gov" + url
        if re.search(r'(about|contact|faqs|embassy|staff)', url, re.I):
            continue

        seen.add(url)
        entries.append({
            "title": title,
            "funder": "US Embassy Burma",
            "location": "Myanmar",
            "deadline": "",
            "amount": "",
            "url": url,
            "source": "US Embassy Burma"
        })
    return entries


def parse_japan_ggp(html):
    """Parse Japan GGP page."""
    entries = []
    if not html:
        return entries
    soup = BeautifulSoup(html, "html.parser")

    # GGP page is simple – look for call/application links
    for a in soup.find_all("a", href=True):
        url = a["href"]
        title = a.get_text(strip=True)

        if not re.search(r'call|application|GGP|grant|2025|2026', title, re.I):
            continue
        if len(title) < 5:
            continue
        if not url.startswith("http"):
            url = "https://www.mm.emb-japan.go.jp" + url

        entries.append({
            "title": title,
            "funder": "Embassy of Japan (GGP)",
            "location": "Myanmar",
            "deadline": "",
            "amount": "Up to ¥20,000,000",
            "url": url,
            "source": "Japan GGP"
        })
    return entries


def parse_unfpa(html):
    """Parse UNFPA Myanmar call for submissions page."""
    entries = []
    if not html:
        return entries
    soup = BeautifulSoup(html, "html.parser")

    # UNFPA uses article/card layout with dates
    for a in soup.find_all("a", href=True):
        url = a["href"]
        title = a.get_text(strip=True)

        if not re.search(r'call|EOI|grant|CSO|submission|proposal|RFQ|RFP', title, re.I):
            continue
        if len(title) < 10:
            continue
        if not url.startswith("http"):
            url = "https://myanmar.unfpa.org" + url
        if re.search(r'(vacanc|staff|about|contact)', url, re.I):
            continue

        # Try to get deadline from parent
        parent = a.find_parent("div", class_=re.compile(r"card|item|post|article"))
        context = parent.get_text() if parent else ""
        deadline = ""
        m = re.search(r'(\d{1,2}\s+[A-Z][a-z]{2}\s+\d{4})', context)
        if m:
            deadline = m.group(1)

        entries.append({
            "title": title,
            "funder": "UNFPA Myanmar",
            "location": "Myanmar",
            "deadline": deadline,
            "amount": "Up to USD 25,000",
            "url": url,
            "source": "UNFPA Myanmar"
        })
    return entries


# ──────────────────────────────────────────────
# PARSERS – RSS Sources (much simpler)
# ──────────────────────────────────────────────

def parse_rss_k4dm(feed_url):
    """Parse K4DM RSS feed for research grant calls."""
    entries = []
    if not feedparser:
        # Fallback: parse XML directly
        return _parse_rss_xml(feed_url, "K4DM",
                              keywords=[r'grant', r'research', r'call for', r'proposal', r'funding'])

    feed = feedparser.parse(feed_url, agent=HEADERS["User-Agent"])
    for entry in feed.entries:
        title = entry.get("title", "")
        link = entry.get("link", "")
        summary = entry.get("summary", "")

        # Filter: only grant/research related posts
        if not re.search(r'grant|research|call for|proposal|funding|fellowship', title + summary, re.I):
            continue

        # Extract deadline from summary
        deadline = ""
        m = re.search(r'(?:deadline|due|close[sd]?)[^a-z]*(\d{1,2}\s+[A-Z][a-z]{2}\s+\d{4})', summary, re.I)
        if m:
            deadline = m.group(1)

        entries.append({
            "title": title,
            "funder": "K4DM (Knowledge for Democracy Myanmar)",
            "location": "Myanmar",
            "deadline": deadline,
            "amount": "Up to CAD 10,000",
            "url": link,
            "source": "K4DM"
        })
    return entries


def parse_rss_lorcanlovett(feed_url):
    """Parse Lorcan Lovett Substack RSS for Myanmar grants/opportunities."""
    entries = []
    if not feedparser:
        return _parse_rss_xml(feed_url, "Lorcan Lovett",
                              keywords=[r'grant', r'funding', r'research', r'call for',
                                        r'proposal', r'opportunity', r'myanmar', r'burma'])

    feed = feedparser.parse(feed_url, agent=HEADERS["User-Agent"])
    for entry in feed.entries:
        title = entry.get("title", "")
        link = entry.get("link", "")
        summary = entry.get("summary", "")

        # Filter: Myanmar-related grants/opportunities
        if not re.search(r'myanmar|burma', title + summary, re.I):
            continue
        if not re.search(r'grant|funding|research|call for|proposal|opportunity|fellowship',
                         title + summary, re.I):
            continue

        deadline = ""
        m = re.search(r'(?:deadline|due|close[sd]?)[^a-z]*(\d{1,2}\s+[A-Z][a-z]{2}\s+\d{4})', summary, re.I)
        if m:
            deadline = m.group(1)

        entries.append({
            "title": title,
            "funder": "Lorcan Lovett (curated)",
            "location": "Myanmar",
            "deadline": deadline,
            "amount": "",
            "url": link,
            "source": "Lorcan Lovett"
        })
    return entries


def _parse_rss_xml(feed_url, source_name, keywords):
    """Fallback RSS parser using xml.etree (no feedparser needed)."""
    entries = []
    try:
        html = fetch(feed_url)
        if not html:
            return entries
        root = ET.fromstring(html)
        # Handle both RSS 2.0 and Atom
        ns = {"atom": "http://www.w3.org/2005/Atom"}
        items = root.findall(".//item") or root.findall(".//atom:entry", ns)
        for item in items:
            title_el = item.find("title") or item.find("atom:title", ns)
            link_el = item.find("link") or item.find("atom:link", ns)
            desc_el = item.find("description") or item.find("atom:summary", ns)

            title = title_el.text if title_el is not None and title_el.text else ""
            link = link_el.text if link_el is not None and link_el.text else ""
            if link_el is not None and link_el.get("href"):
                link = link_el.get("href")
            summary = desc_el.text if desc_el is not None and desc_el.text else ""

            # Filter by keywords
            text = (title + summary).lower()
            if not any(re.search(kw, text, re.I) for kw in keywords):
                continue

            deadline = ""
            m = re.search(r'(?:deadline|due|close[sd]?)[^a-z]*(\d{1,2}\s+[A-Z][a-z]{2}\s+\d{4})', summary, re.I)
            if m:
                deadline = m.group(1)

            entries.append({
                "title": title.strip(),
                "funder": source_name,
                "location": "Myanmar",
                "deadline": deadline,
                "amount": "",
                "url": link.strip(),
                "source": source_name
            })
    except Exception as e:
        print(f"  [RSS ERROR] {feed_url}: {e}")
    return entries


# ──────────────────────────────────────────────
# EXCEL OUTPUT
# ──────────────────────────────────────────────

def deadline_passed(deadline_str):
    """Check if a deadline has passed. Handles multiple formats."""
    if not deadline_str:
        return False
    formats = ["%d-%b-%Y", "%b %d, %Y", "%d %B %Y", "%Y-%m-%d"]
    for fmt in formats:
        try:
            d = datetime.strptime(deadline_str.strip(), fmt).date()
            return d < date.today()
        except ValueError:
            continue
    return False


def log_to_excel(entries, filepath):
    """Write all entries to a formatted Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    # Headers
    headers = ["Date Found", "Title", "Funder", "Location",
               "Deadline", "Amount", "Source", "URL", "Status"]
    ws.append(headers)

    # Style headers
    header_fill = PatternFill("solid", fgColor="2F5496")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Data rows
    today = date.today().strftime("%Y-%m-%d")
    open_count = 0
    closed_count = 0

    for e in entries:
        status = "Closed" if deadline_passed(e["deadline"]) else "Open"
        if status == "Open":
            open_count += 1
        else:
            closed_count += 1

        ws.append([
            today, e["title"], e["funder"], e["location"],
            e["deadline"], e["amount"], e["source"], e["url"], status
        ])

        # Style URL as hyperlink
        url_cell = ws.cell(row=ws.max_row, column=8)
        url_cell.hyperlink = e["url"]
        url_cell.font = Font(color="0563C1", underline="single", size=10)

        # Color-code status
        status_cell = ws.cell(row=ws.max_row, column=9)
        if status == "Open":
            status_cell.font = Font(color="006100", bold=True)
            status_cell.fill = PatternFill("solid", fgColor="C6EFCE")
        else:
            status_cell.font = Font(color="9C0006")
            status_cell.fill = PatternFill("solid", fgColor="FFC7CE")

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.value:
                max_len = max(max_len, min(len(str(cell.value)), 55))
        ws.column_dimensions[col[0].column_letter].width = max_len + 3

    # Freeze header row
    ws.freeze_panes = "A2"

    # Add summary row
    ws.append([])
    ws.append([f"Summary: {len(entries)} total | {open_count} Open | {closed_count} Closed | Generated: {today}"])
    ws.cell(row=ws.max_row, column=1).font = Font(italic=True, color="666666")

    wb.save(filepath)
    print(f"  → Saved {len(entries)} entries ({open_count} open, {closed_count} closed) to {filepath}")


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main():
    print(f"\n{'='*60}")
    print(f"  Myanmar CFP Tracker v2.0")
    print(f"  Run: {datetime.now().strftime('%Y-%m-%d %H:%M UTC')}")
    print(f"{'='*60}\n")

    # Map source names to their parsers
    html_parsers = {
        "MIMU": parse_mimu,
        "fundsforNGOs": parse_fundsforngos,
        "Advance Africa": parse_advance_africa,
        "DevelopmentAid": parse_developmentaid,
        "GrantStation": parse_grantstation,
        "US Embassy Burma": parse_us_embassy,
        "Japan GGP": parse_japan_ggp,
        "UNFPA Myanmar": parse_unfpa,
    }

    rss_parsers = {
        "K4DM": parse_rss_k4dm,
        "Lorcan Lovett": parse_rss_lorcanlovett,
    }

    all_entries = []
    source_stats = []

    for source in SOURCES:
        name = source["name"]
        url = source["url"]
        stype = source["type"]

        print(f"  [{stype.upper()}] {name}...")
        try:
            if stype == "rss":
                parser = rss_parsers.get(name)
                entries = parser(url) if parser else []
            else:
                html = fetch(url)
                parser = html_parsers.get(name)
                entries = parser(html) if parser else []
        except Exception as e:
            print(f"         [SKIP] {name} failed: {e}")
            entries = []   

        print(f"         → {len(entries)} items found")
        source_stats.append(f"    • {name}: {len(entries)}")
        all_entries.extend(entries)

    # Deduplicate by title (case-insensitive)
    seen_titles = set()
    unique_entries = []
    for e in all_entries:
        key = e["title"].lower().strip()[:80]  # first 80 chars as key
        if key not in seen_titles:
            seen_titles.add(key)
            unique_entries.append(e)

    # Sort: Open first, then by deadline (earliest first)
    def sort_key(e):
        is_open = 0 if not deadline_passed(e["deadline"]) else 1
        # Try to parse deadline for sorting
        dl = e["deadline"]
        for fmt in ["%d-%b-%Y", "%b %d, %Y"]:
            try:
                d = datetime.strptime(dl.strip(), fmt)
                return (is_open, d)
            except:
                pass
        return (is_open, datetime(2099, 1, 1))  # no deadline = last

    unique_entries.sort(key=sort_key)

    print(f"\n  {'─'*40}")
    print(f"  TOTAL: {len(unique_entries)} unique entries")
    for stat in source_stats:
        print(stat)
    print(f"  {'─'*40}\n")

    # Write to Excel
    if unique_entries:
        log_to_excel(unique_entries, EXCEL_FILE)
    else:
        log_to_excel([{"title": "No data found – check sources", "funder": "-",
                       "location": "-", "deadline": "", "amount": "",
                       "url": "", "source": "-"}], EXCEL_FILE)

    print(f"\n  ✅ Done. Output: {EXCEL_FILE}\n")


if __name__ == "__main__":
    main()   
