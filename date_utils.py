import re
from datetime import date

CURRENT_YEAR = date.today().year
_TWO_DIGIT_PIVOT = CURRENT_YEAR % 100  # e.g. 26 for 2026

_MONTH_ABBREV_RE = r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?"


def _two_digit_to_four(yy: int) -> int:
    """Convert a 2-digit year to a plausible 4-digit year using a rolling
    pivot around the current year (yy <= pivot -> 20yy, else 19yy). Safe
    assumption for career-history dates, which won't span >100 years."""
    return 2000 + yy if yy <= _TWO_DIGIT_PIVOT else 1900 + yy


def parse_year_range(text: str):
    """Extract (start_year, end_year, is_current) from any of:
    '2001-2003', '2001–2003', '2009-Present', 'Since 2009',
    'Apr 2016 to Apr 2019', '2007' (single year), 'Jan 2017 – Nov 23'
    (an abbreviated 2-digit year after a month name, common in Indian CVs
    that write the end of a range as just the month + last two digits),
    or DD/MM/YY(YY)-style dates such as '4/09/18 to 30/04/19', including
    compound strings with several sub-ranges joined by commas/'and' (e.g.
    multiple stints in the same role) — the overall span across all
    sub-ranges is returned."""
    if not text:
        return None, None, False

    t = text.strip().lower()
    is_current = bool(re.search(r"present|current|till date|ongoing|now|onwards", t))

    years = [int(m) for m in re.findall(r"(?:19|20)\d{2}", t)]

    # Also catch abbreviated 2-digit years written directly after a month
    # name with no 4-digit year of their own (e.g. "Jan 2017 - Nov 23").
    # Must NOT match a day-of-month in a full "Month DD, YYYY" date (e.g.
    # "Nov 07, 2017" or "Dec 26, 2023") — the negative lookahead excludes
    # any 2-digit number immediately followed by a comma and a 4-digit year,
    # since that pattern means the 2 digits are a day, not a year.
    for m in re.finditer(
        _MONTH_ABBREV_RE + r"\s+(\d{2})(?!\d)(?!\s*,\s*(?:19|20)\d{2})", t
    ):
        yy = int(m.group(1))
        years.append(_two_digit_to_four(yy))

    if not years:
        # Fall back to DD/MM/YY(YY)-style dates (tolerant of typos like a
        # doubled slash "30//04/16" or a missing space before the date like
        # "to30/07/09" — neither is anchored to a word boundary).
        for _d, _m, y in re.findall(r"(\d{1,2})/+(\d{1,2})/+(\d{2,4})", t):
            yi = int(y)
            if yi < 100:
                yi = _two_digit_to_four(yi)
            years.append(yi)

    if not years:
        return None, None, is_current

    # Use min/max rather than first/last so compound multi-sub-range strings
    # (several date pairs in one date_range field) still produce a sane
    # overall span regardless of extraction order.
    start = min(years)
    end = max(years)

    if is_current:
        return start, CURRENT_YEAR, True

    # Handle single-year entries that actually indicate an ongoing range
    # due to a trailing connector (e.g. '2010 -', 'May 2010 –', '2010 to').
    # In such cases treat as current rather than a single-year event.
    if start == end and len(years) == 1:
        # If the text explicitly uses 'since', it's current
        if "since" in t:
            return start, CURRENT_YEAR, True

        # Remove the matched year(s) from the text and check whether the
        # remainder ends with a dash/connector with nothing after it.
        t_no_years = re.sub(r"(?:19|20)\d{2}", "", t).strip()
        # Connectors that imply an open-ended range when trailing with nothing after
        if re.search(r"(?:[-–—]|to|until|till)\s*$", t_no_years):
            return start, CURRENT_YEAR, True

        return start, start, False
    return start, end, False