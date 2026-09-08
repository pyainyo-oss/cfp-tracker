"""
Myanmar CFP Tracker – runs on GitHub Actions
Scrapes MIMU, fundsforNGOs, Advance Africa
Outputs: CFP_Tracker_Myanmar.xlsx
"""

import requests
import re
import os
from datetime import datetime, date
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
EXCEL_FILE = "CFP_Tracker_Myanmar.xlsx"
SHEET_NAME = "CFPs"

SOURCES = [
    {"name": "MIMU", "url": "https://themimu.info/calls-for-proposals"},
    {"name": "fundsforNGOs", "url": "https://www2.fundsforngos.org/tag/burmamyanmar/"},
    {"name": "Advance Africa", "url": "https://www.advance-africa.com/Grants-for-NGOs-in-Myanmar.html"},
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}

# ──────────────────────────────────────────────
# FETCH
# ──────────────────────────────────────────────

def fetch(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        r.raise_for_status()
        return r.text
    except Exception as e:
        print(f"  [ERROR] {url}: {e}")
        return ""


# ──────────────────────────────────────────────
# PARSERS
# ──────────────────────────────────────────────

def parse_mimu(html):
    entries = []
    if not html:
        return entries
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", class_="views-table")
    if not table:
        for t in soup.find_all("table"):
            if "call" in t.get_text().lower() or "proposal" in t.get_text().lower():
                table = t
                break
    if not table:
        return entries

    for row in table.find_all("tr")[1:]:
        cells = row.find_all("td")
        if len(cells) < 4:
            continue
        link = cells[0].find("a")
        if not link:
            continue
        title = link.get_text(strip=True)
        url = link.get("href", "")
        if not url.startswith("http"):
            url = "https://themimu.info" + url

        geo = cells[1].get_text(strip=True) if len(cells) > 1 else ""
        funder = cells[2].get_text(strip=True) if len(cells) > 2 else ""
        deadline_raw = cells[-1].get_text(strip=True) if cells else ""
        m = re.search(r'\d{1,2}-[A-Za-z]{3}-\d{4}', deadline_raw)
        deadline = m.group(0) if m else ""

        entries.append({
            "title": title, "funder": funder, "location": geo,
            "deadline": deadline, "amount": "", "url": url, "source": "MIMU"
        })
    return entries


def parse_fundsforngos(html):
    entries = []
    if not html:
        return entries
    soup = BeautifulSoup(html, "html.parser")

    for h2 in soup.find_all("h2"):
        a = h2.find("a")
        if not a:
            continue
        title = a.get_text(strip=True)
        url = a.get("href", "")
        if not url.startswith("http"):
            url = "https://www2.fundsforngos.org" + url

        parent = a.find_parent("div", class_=re.compile(r"post|article|card"))
        body_text = parent.get_text() if parent else ""
        if not re.search(r'myanmar|burma', title + body_text, re.I):
            continue

        deadline = ""
        m = re.search(r'[Dd]eadline:?\s*(\d{1,2}-[A-Za-z]{3}-\d{4})', body_text)
        if m:
            deadline = m.group(1)

        funder = "Unknown"
        fm = re.search(r'(?:by|from|The)\s+([A-Z][A-Za-z&]+(?:\s+[A-Z][A-Za-z&]+){0,4})', body_text)
        if fm:
            funder = fm.group(1)

        amount = ""
        am = re.search(r'(?:USD|US\$|\$|MMK|€)\s?[\d,]+', body_text)
        if am:
            amount = am.group(0)

        entries.append({
            "title": title, "funder": funder, "location": "Myanmar",
            "deadline": deadline, "amount": amount, "url": url, "source": "fundsforNGOs"
        })
    return entries


def parse_advance_africa(html):
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
        seen.add(url)
        entries.append({
            "title": title, "funder": "Advance Africa", "location": "Myanmar",
            "deadline": "", "amount": "", "url": url, "source": "Advance Africa"
        })
    return entries


# ──────────────────────────────────────────────
# EXCEL
# ──────────────────────────────────────────────

def deadline_passed(deadline_str):
    if not deadline_str:
        return False
    try:
        d = datetime.strptime(deadline_str, "%d-%b-%Y").date()
        return d < date.today()
    except:
        return False


def log_to_excel(entries, filepath):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    headers = ["Date Found", "Title", "Funder", "Location",
               "Deadline", "Amount", "Source", "URL", "Status"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")

    today = date.today().strftime("%Y-%m-%d")
    for e in entries:
        status = "Closed" if deadline_passed(e["deadline"]) else "Open"
        ws.append([today, e["title"], e["funder"], e["location"],
                   e["deadline"], e["amount"], e["source"], e["url"], status])
        url_cell = ws.cell(row=ws.max_row, column=8)
        url_cell.hyperlink = e["url"]
        url_cell.font = Font(color="0563C1", underline="single")

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

    wb.save(filepath)
    print(f"  → Saved {len(entries)} entries to {filepath}")


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main():
    print(f"=== Myanmar CFP Tracker – {datetime.now().strftime('%Y-%m-%d %H:%M')} ===\n")

    parsers = {
        "MIMU": parse_mimu,
        "fundsforNGOs": parse_fundsforngos,
        "Advance Africa": parse_advance_africa,
    }

    all_entries = []

    for source in SOURCES:
        name = source["name"]
        print(f"  Scraping {name}...")
        html = fetch(source["url"])
        entries = parsers[name](html)
        print(f"    Found {len(entries)} items.")
        all_entries.extend(entries)

    print(f"\n  Total entries: {len(all_entries)}")

    if all_entries:
        log_to_excel(all_entries, EXCEL_FILE)
    else:
        # Create empty file so artifact upload doesn't fail
        log_to_excel([{"title": "No data found", "funder": "-", "location": "-",
                       "deadline": "", "amount": "", "url": "", "source": "-"}], EXCEL_FILE)

    print("\n  Done.")


if __name__ == "__main__":
    main()   
