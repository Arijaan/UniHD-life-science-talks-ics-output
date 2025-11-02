import os
import re
from datetime import datetime, timedelta, time
from bs4 import BeautifulSoup
from ics import Calendar, Event
import tkinter as tk
from tkinter import filedialog

try:
    from zoneinfo import ZoneInfo
except ImportError: 
    ZoneInfo = None

'''
This script is meant for exporting appointment times of life science talks at the university of Heidelberg,
listed on the page with the same name:
(www.uni-heidelberg.de/en/research/research-profile/fields-of-focus/field-of-focus-i/life-science-talks-on-campus). 
As an output, a .ics file will be generated in the root directory, which can be imported into calendar apps.

How to:

1) Download the page as HTML.

2) run the "local" script, selecting the html file after the file dialog OR
run the "online" script

'''
root = tk.Tk()
root.withdraw()
HTML_PATH = filedialog.askopenfilename(
    title="Select the HTML file of the events page",
    filetypes=[("HTML files", "*.html;*.htm"), ("All files", "*.*")],
)
if not HTML_PATH:
    raise SystemExit("No HTML file selected; aborting.")

OUTPUT_PATH = "life_science_talks.ics"

DEFAULT_TIMEZONE_NAME = "Europe/Berlin"
if ZoneInfo:
    try:
        DEFAULT_TIMEZONE = ZoneInfo(DEFAULT_TIMEZONE_NAME)
    except Exception:
        DEFAULT_TIMEZONE = None
else:
    DEFAULT_TIMEZONE = None

DEFAULT_SINGLE_EVENT_DURATION = timedelta(hours=1)

MERIDIEM_PATTERN = re.compile(r"(a\.?m\.?|p\.?m\.?)", re.IGNORECASE)
TZ_LABEL_PATTERN = re.compile(
    r"\b(?:cet|cest|gmt|utc(?:[+\-]\d{1,2})?|mez|mesz)\b", re.IGNORECASE
)
RANGE_CONNECTOR_PATTERN = re.compile(r"\b(?:to|bis)\b", re.IGNORECASE)
EXTRA_LABEL_PATTERN = re.compile(r"\b(?:hrs?|hours?|uhr|o'clock|h)\b", re.IGNORECASE)
TIME_CORE_PATTERN = re.compile(r"(?P<hour>\d{1,2})(?::(?P<minute>\d{2}))?")


def clean_text(node):
    if not node:
        return ""
    return " ".join(node.stripped_strings)


def extract_month_and_year(text, fallback_year):
    sanitized = re.sub(r"[\u2013\u2014,/]", " ", text).strip()
    sanitized = re.sub(r"\s+", " ", sanitized)
    if not sanitized:
        return None, fallback_year

    for fmt in ("%B %Y", "%b %Y"):
        try:
            parsed = datetime.strptime(sanitized, fmt)
            return parsed.strftime("%B"), parsed.year
        except ValueError:
            continue

    for token in sanitized.split():
        token_clean = token.strip(". ")
        if not token_clean:
            continue
        try:
            parsed = datetime.strptime(token_clean, "%B")
            month_name = parsed.strftime("%B")
            year = fallback_year
            year_match = re.search(r"\b(20\d{2})\b", sanitized)
            if year_match:
                year = int(year_match.group(1))
            return month_name, year
        except ValueError:
            try:
                parsed = datetime.strptime(token_clean, "%b")
                month_name = parsed.strftime("%B")
                year = fallback_year
                year_match = re.search(r"\b(20\d{2})\b", sanitized)
                if year_match:
                    year = int(year_match.group(1))
                return month_name, year
            except ValueError:
                continue

    year_match = re.search(r"\b(20\d{2})\b", sanitized)
    if year_match:
        return None, int(year_match.group(1))

    return None, fallback_year


def parse_date(date_text, current_month_name, current_year):
    cleaned = re.sub(r"[\u00a0,]", " ", date_text).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = re.sub(
        r"\b(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b",
        "",
        cleaned,
        flags=re.IGNORECASE,
    ).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    if not cleaned:
        raise ValueError("Date text is empty")

    candidates = []
    if current_month_name and current_month_name.lower() not in cleaned.lower():
        candidates.append(f"{current_month_name} {cleaned}")
    candidates.append(cleaned)

    for candidate in candidates:
        for fmt in ("%B %d %Y", "%b %d %Y", "%d %B %Y", "%d %b %Y"):
            try:
                return datetime.strptime(candidate, fmt).date()
            except ValueError:
                continue

        for fmt in ("%B %d", "%b %d", "%d %B", "%d %b"):
            try:
                parsed = datetime.strptime(candidate, fmt)
                return parsed.replace(year=current_year).date()
            except ValueError:
                continue

        for fmt in ("%d.%m.%Y", "%d.%m.%y", "%d.%m."):
            try:
                parsed = datetime.strptime(candidate, fmt)
                if "%Y" not in fmt and "%y" not in fmt:
                    parsed = parsed.replace(year=current_year)
                return parsed.date()
            except ValueError:
                continue

    if cleaned.isdigit() and current_month_name:
        month_dt = datetime.strptime(current_month_name, "%B")
        return datetime(current_year, month_dt.month, int(cleaned)).date()

    raise ValueError(f"Unable to parse date '{date_text}'")


def sanitize_time_fragment(fragment):
    fragment = fragment.replace("p.m.", "pm").replace("a.m.", "am")
    fragment = re.sub(r"(?<=\d)h(?=\d)", ":", fragment, flags=re.IGNORECASE)
    fragment = re.sub(r"(?<=\d)\.(?=\d)", ":", fragment)
    fragment = TZ_LABEL_PATTERN.sub("", fragment)
    fragment = EXTRA_LABEL_PATTERN.sub("", fragment)
    fragment = fragment.replace("midday", "noon")
    fragment = fragment.replace("·", " ")
    fragment = re.sub(r"[()]", " ", fragment)
    fragment = re.sub(r"\s+", " ", fragment)
    return fragment.strip()


def extract_meridiem_label(text):
    match = MERIDIEM_PATTERN.search(text)
    if not match:
        return None
    return "pm" if "p" in match.group(0).lower() else "am"


def convert_to_24_hour(hour, meridiem):
    hour %= 24
    if meridiem is None:
        return hour
    if meridiem == "am":
        return 0 if hour == 12 else hour
    return 12 if hour == 12 else hour + 12


def parse_time_fragment(fragment, default_meridiem=None, prefer_12_hour=False):
    sanitized = sanitize_time_fragment(fragment)
    if not sanitized:
        raise ValueError("Empty time fragment")

    lowered = sanitized.lower()
    if lowered == "noon":
        return 12, 0
    if lowered == "midnight":
        return 0, 0

    match = TIME_CORE_PATTERN.search(sanitized)
    if not match:
        raise ValueError(f"Could not parse time from '{fragment}'")

    hour = int(match.group("hour"))
    minute = int(match.group("minute") or 0)
    meridiem = extract_meridiem_label(sanitized)

    if meridiem:
        prefer_12_hour = True

    if (
        meridiem is None
        and prefer_12_hour
        and default_meridiem
        and hour <= 12
    ):
        meridiem = default_meridiem

    hour_24 = convert_to_24_hour(hour, meridiem)
    return hour_24, minute


def assemble_datetime(event_date, hour, minute):
    base_time = time(hour=hour % 24, minute=minute)
    if DEFAULT_TIMEZONE:
        base_time = base_time.replace(tzinfo=DEFAULT_TIMEZONE)
    return datetime.combine(event_date, base_time)


def looks_like_time_range(text):
    lowered = text.lower()
    if re.search(r"[\-\u2013\u2014]", text) and re.search(r"\d", text):
        return True
    if " to " in lowered or " bis " in lowered:
        return True
    return False


def parse_time_components(time_range, event_date):
    normalized = RANGE_CONNECTOR_PATTERN.sub("-", time_range)
    normalized = normalized.replace("\u2013", "-").replace("\u2014", "-")
    normalized = TZ_LABEL_PATTERN.sub("", normalized)
    normalized = EXTRA_LABEL_PATTERN.sub("", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip(" -")

    if not normalized:
        raise ValueError(f"Empty time range: {time_range}")

    parts = [part.strip() for part in normalized.split("-") if part.strip()]

    if len(parts) == 1:
        start_part = parts[0]
        start_hint = extract_meridiem_label(start_part)
        start_hour, start_minute = parse_time_fragment(
            start_part, default_meridiem=start_hint, prefer_12_hour=bool(start_hint)
        )
        start_dt = assemble_datetime(event_date, start_hour, start_minute)
        end_dt = start_dt + DEFAULT_SINGLE_EVENT_DURATION
        return start_dt, end_dt

    if len(parts) < 2:
        raise ValueError(f"Unexpected time range: {time_range}")

    start_part, end_part = parts[0], parts[1]
    start_hint = extract_meridiem_label(start_part)
    end_hint = extract_meridiem_label(end_part)
    start_default = start_hint or end_hint
    end_default = end_hint or start_hint
    prefer_12_hour = bool(start_hint or end_hint)

    start_hour, start_minute = parse_time_fragment(
        start_part, default_meridiem=start_default, prefer_12_hour=prefer_12_hour
    )
    end_hour, end_minute = parse_time_fragment(
        end_part, default_meridiem=end_default, prefer_12_hour=prefer_12_hour
    )

    start_dt = assemble_datetime(event_date, start_hour, start_minute)
    end_dt = assemble_datetime(event_date, end_hour, end_minute)

    if end_dt <= start_dt:
        end_dt += timedelta(days=1)

    return start_dt, end_dt


def format_duration(delta):
    total_minutes = int(delta.total_seconds() // 60)
    hours, minutes = divmod(total_minutes, 60)
    parts = []
    if hours:
        parts.append(f"{hours}h")
    if minutes or not parts:
        parts.append(f"{minutes}m")
    return " ".join(parts)


def build_description(speaker, venue, link, duration_text):
    chunks = []
    if speaker:
        chunks.append(f"Speaker: {speaker}")
    if venue:
        chunks.append(f"Venue: {venue}")
    if link:
        chunks.append(f"Link: {link}")
    if duration_text:
        chunks.append(f"Duration: {duration_text}")
    return "\n".join(chunks)


def main():
    if DEFAULT_TIMEZONE is None:
        print(
            f"Warning: timezone '{DEFAULT_TIMEZONE_NAME}' not available; exporting floating times."
        )

    with open(HTML_PATH, encoding="utf-8") as handle:
        soup = BeautifulSoup(handle.read(), "html.parser")

    table = soup.select_one("table:nth-of-type(2) > tbody")
    if table is None:
        raise RuntimeError("Could not find the expected table with event data.")

    cal = Calendar()
    current_month_name = None
    current_year = datetime.now().year
    created_events = 0
    skipped_details = []

    for row in table.select("tr"):
        cells = row.find_all("td")
        if not cells:
            continue

        strong_texts = [strong.get_text(strip=True) for strong in row.select("strong")]
        link_tag = row.select_one("a[href]")

        if strong_texts and not link_tag:
            for text in strong_texts:
                month_candidate, year_candidate = extract_month_and_year(text, current_year)
                if year_candidate != current_year:
                    current_year = year_candidate
                if month_candidate:
                    current_month_name = month_candidate
            continue

        date_strings = list(cells[0].stripped_strings)
        if not date_strings:
            continue

        joined_date = " ".join(date_strings)
        if not any(ch.isdigit() for ch in joined_date):
            continue

        date_token = next(
            (value for value in date_strings if any(char.isdigit() for char in value)),
            date_strings[0],
        )

        try:
            event_date = parse_date(date_token, current_month_name, current_year)
        except ValueError as exc:
            skipped_details.append(f"date parse failed for '{date_token}': {exc}")
            continue

        try:
            date_index = date_strings.index(date_token)
        except ValueError:
            date_index = 0

        trailing_strings = date_strings[date_index + 1 :]
        time_candidates = [value for value in trailing_strings if looks_like_time_range(value)]
        if not time_candidates:
            fallback = next(
                (value for value in trailing_strings if any(char.isdigit() for char in value)),
                "",
            )
            if fallback:
                time_candidates = [fallback]
        if not time_candidates:
            skipped_details.append(f"missing time range for date '{date_token}'")
            continue

        time_text = time_candidates[0]

        try:
            start_dt, end_dt = parse_time_components(time_text, event_date)
        except ValueError as exc:
            skipped_details.append(f"time parse failed for '{time_text}': {exc}")
            continue

        title = clean_text(cells[2]) if len(cells) > 2 else ""
        speaker = clean_text(cells[4]) if len(cells) > 4 else ""
        venue = clean_text(cells[6]) if len(cells) > 6 else (clean_text(cells[-1]) if len(cells) else "")
        link = link_tag.get("href", "") if link_tag else ""

        if not title:
            title = "Untitled Event"

        event = Event(name=title, begin=start_dt, end=end_dt)
        event.description = build_description(
            speaker,
            venue,
            link,
            format_duration(end_dt - start_dt),
        )
        if venue:
            event.location = venue
        if link:
            event.url = link

        cal.events.add(event)
        created_events += 1

    with open(OUTPUT_PATH, "w", encoding="utf-8") as handle:
        handle.write(cal.serialize())

    print(f"Exported {created_events} events to {OUTPUT_PATH}")
    if skipped_details:
        preview = "; ".join(skipped_details[:5])
        more = "" if len(skipped_details) <= 5 else f" and {len(skipped_details) - 5} more"
        print(f"Skipped {len(skipped_details)} rows: {preview}{more}")


if __name__ == "__main__":
    main()
