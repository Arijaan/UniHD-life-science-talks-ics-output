import os
import re
from bs4 import BeautifulSoup

HTML_PATH = r"C:\Users\psoor\OneDrive\Desktop\Life Science Talks on Campus - Heidelberg University.html"

def main():
    with open(HTML_PATH, encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    rows = soup.select("table:nth-of-type(2) > tbody > tr")
    for idx, row in enumerate(rows):
        cells = row.find_all("td")
        if not cells:
            continue

        strong_texts = [strong.get_text(strip=True) for strong in row.select("strong")]
        link_tag = row.select_one("a[href]")
        if strong_texts and not link_tag:
            continue

        date_strings = list(cells[0].stripped_strings)
        if not date_strings:
            continue

        joined = " ".join(date_strings)
        if not any(ch.isdigit() for ch in joined):
            continue

        time_strings = [s for s in date_strings[1:] if any(ch.isdigit() for ch in s)]
        if any(re.search(r"[\-\u2013\u2014]", s) for s in time_strings):
            continue

        if any("am" in s.lower() or "pm" in s.lower() for s in time_strings):
            print(f"Row {idx} problematic time strings: {time_strings}")
            for c_idx, cell in enumerate(cells):
                print(f"  cell{c_idx}: {list(cell.stripped_strings)}")
            print("-" * 80)

if __name__ == "__main__":
    main()
