import os
import re
import xml.etree.ElementTree as ET
from datetime import date, datetime
from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

try:
    import feedparser
except ImportError:
    feedparser = None

# ──────────────────────────────────────────────
# CONFIGURATION
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "CFP_Tracker_Myanmar.xlsx")
SHEET_NAME = "CFPs"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

SOURCES = [
    {"name": "MIMU", "url": "https://themimu.info/calls-for-proposals", "type": "html"},
    {"name": "fundsforNGOs", "url": "https://www2.fundsforngos.org/tag/burmamyanmar/", "type": "html"},
    {"name": "Advance Africa", "url": "https://www.advance-africa.com/Grants-for-NGOs-in-Myanmar.html", "type": "html"},
    {"name": "GrantStation", "url": "https://grantstation.com/find-grants/myanmar-burma", "type": "html"},
    {"name": "UNFPA Myanmar", "url": "https://myanmar.unfpa.org/en/call-for-submissions", "type": "html"},
    {"name": "K4DM", "url": "https://k4dm.ca/feed/", "type": "rss"},
    {"name": "Lorcan Lovett", "url": "https://lorcanlovett.substack.com/feed", "type": "rss"},
]

def fetch(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=25)
        r.raise_for_status()
        return r.text
    except Exception as e:
        print(f"  [ERROR] {url}: {e}")
        return ""

def parse_date_safely(date_str):
    if not date_str:
        return None
    date_patterns = [
        r'\d{1,2}-[A-Za-z]{3}-\d{4}',
        r'[A-Za-z]{3}\s+\d{1,2},\s*\d{4}',
        r'\d{1,2}\s+[A-Za-z]+\s+\d{4}',
        r'\d{4}-\d{2}-\d{2}'
    ]
    for pattern in date_patterns:
        m = re.search(pattern, str(date_str).strip())
        if m:
            extracted = m.group(0)
            for fmt in ["%d-%b-%Y", "%b %d, %Y", "%d %B %Y", "%Y-%m-%d"]:
                try:
                    return datetime.strptime(extracted, fmt).date()
                except ValueError:
                    continue
    return None

def deadline_passed(deadline_str):
    parsed = parse_date_safely(deadline_str)
    return parsed < date.today() if parsed else False

def parse_mimu(html):
    entries = []
    if not html: return entries
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", class_="views-table") or soup.find("table")
    if not table: return entries

    for row in table.find_all("tr")[1:]:
        cells = row.find_all("td")
        if len(cells) < 3: continue
        try:
            link = cells[0].find("a")
            if not link: continue
            title = link.get_text(strip=True)
            url = link.get("href", "")
            if not url.startswith("http"): url = "https://themimu.info" + url

            geo = cells[2].get_text(strip=True) if len(cells) > 2 else "Myanmar"
            funder = cells[3].get_text(strip=True) if len(cells) > 3 else "Unknown"
            deadline_raw = cells[-1].get_text(strip=True) if len(cells) > 4 else ""

            entries.append({
                "title": title, "funder": funder, "location": geo,
                "deadline": deadline_raw, "amount": "", "url": url, "source": "MIMU"
            })
        except Exception:
            continue
    return entries

def parse_advance_africa(html):
    entries = []
    if not html: return entries
    soup = BeautifulSoup(html, "html.parser")
    seen = set()
    for a in soup.find_all("a", href=True):
        url, title = a["href"], a.get_text(strip=True)
        if not re.search(r'myanmar|burma', title, re.I) or len(title) < 12 or url in seen:
            continue
        if not url.startswith("http"): url = "https://www.advance-africa.com" + url
        seen.add(url)
        entries.append({
            "title": title, "funder": "Advance Africa", "location": "Myanmar",
            "deadline": "", "amount": "", "url": url, "source": "Advance Africa"
        })
    return entries

def parse_grantstation(html):
    entries = []
    if not html: return entries
    soup = BeautifulSoup(html, "html.parser")
    seen = set()
    for a in soup.find_all("a", href=True):
        url, title = a["href"], a.get_text(strip=True)
        if not re.search(r'myanmar|burma', title, re.I) or len(title) < 12 or url in seen:
            continue
        if not url.startswith("http"): url = "https://grantstation.com" + url
        seen.add(url)
        entries.append({
            "title": title, "funder": "GrantStation", "location": "Myanmar",
            "deadline": "", "amount": "", "url": url, "source": "GrantStation"
        })
    return entries

def parse_unfpa(html):
    entries = []
    if not html: return entries
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        url, title = a["href"], a.get_text(strip=True)
        if not re.search(r'call|EOI|grant|CSO|submission|proposal', title, re.I) or len(title) < 10:
            continue
        if not url.startswith("http"): url = "https://myanmar.unfpa.org" + url
        entries.append({
            "title": title, "funder": "UNFPA Myanmar", "location": "Myanmar",
            "deadline": "", "amount": "", "url": url, "source": "UNFPA Myanmar"
        })
    return entries

def parse_rss_feed(feed_url, source_name):
    entries = []
    if feedparser:
        feed = feedparser.parse(feed_url, agent=HEADERS["User-Agent"])
        for entry in feed.entries:
            title = entry.get("title", "")
            link = entry.get("link", "")
            summary = entry.get("summary", "")
            if not re.search(r'grant|funding|research|call|proposal|opportunity|myanmar', title + summary, re.I):
                continue
            entries.append({
                "title": title, "funder": source_name, "location": "Myanmar",
                "deadline": "", "amount": "", "url": link, "source": source_name
            })
    return entries

def log_to_excel(entries, filepath):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    headers = ["Date Found", "Title", "Funder", "Location", "Deadline", "Amount", "Source", "URL", "Status"]
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="2F5496")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    today = date.today().strftime("%Y-%m-%d")
    open_cnt, closed_cnt = 0, 0

    for e in entries:
        status = "Closed" if deadline_passed(e["deadline"]) else "Open"
        if status == "Open": open_cnt += 1
        else: closed_cnt += 1

        ws.append([today, e["title"], e["funder"], e["location"], e["deadline"], e["amount"], e["source"], e["url"], status])

        url_cell = ws.cell(row=ws.max_row, column=8)
        url_cell.hyperlink = e["url"]
        url_cell.font = Font(color="0563C1", underline="single", size=10)

        status_cell = ws.cell(row=ws.max_row, column=9)
        if status == "Open":
            status_cell.font = Font(color="006100", bold=True)
            status_cell.fill = PatternFill("solid", fgColor="C6EFCE")
        else:
            status_cell.font = Font(color="9C0006")
            status_cell.fill = PatternFill("solid", fgColor="FFC7CE")

    # Column Width သတ်မှတ်ခြင်း အပိုင်း ပြင်ဆင်ချက်
    for col in ws.columns:
        col_letter = col[0].column_letter
        if col_letter == 'H':  # URL Column ကို Width ၃၀ အသေထားခြင်း
            ws.column_dimensions[col_letter].width = 30
        else:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

    ws.freeze_panes = "A2"
    wb.save(filepath)
    print(f"\n ✅ Saved {len(entries)} items ({open_cnt} Open, {closed_cnt} Closed) to {filepath}")

def main():
    html_parsers = {
        "MIMU": parse_mimu,
        "Advance Africa": parse_advance_africa,
        "GrantStation": parse_grantstation,
        "UNFPA Myanmar": parse_unfpa,
    }

    all_entries = []
    for source in SOURCES:
        name, url, stype = source["name"], source["url"], source["type"]
        print(f" Processing {name}...")
        try:
            if stype == "rss":
                entries = parse_rss_feed(url, name)
            else:
                html = fetch(url)
                parser = html_parsers.get(name)
                entries = parser(html) if parser else []
        except Exception as e:
            print(f"   [SKIP] Failed {name}: {e}")
            entries = []

        all_entries.extend(entries)

    seen, unique_entries = set(), []
    for e in all_entries:
        key = e["title"].lower().strip()[:80]
        if key not in seen and len(key) > 5:
            seen.add(key)
            unique_entries.append(e)

    def sort_key(e):
        is_closed = 1 if deadline_passed(e["deadline"]) else 0
        parsed = parse_date_safely(e["deadline"])
        return (is_closed, parsed or date(2099, 1, 1))

    unique_entries.sort(key=sort_key)
    log_to_excel(unique_entries, EXCEL_FILE)

if __name__ == "__main__":
    main()
