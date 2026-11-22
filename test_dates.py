"""
Unit tests for the two source-specific date parsers in job_scraper.py.

Naukri's "Posted N days ago" text is handled by parse_days_old() and covered in
test_helpers.py. LinkedIn and ZipRecruiter don't use that phrasing, so each has
its own parser, and both feed the same max_days_old freshness filter:

    days = _li_days_old(...) / _zr_parse_date(...)
    if days is not None and days > max_days_old:
        skip

Note the `is not None`. An unreadable date is deliberately let through, so that
a parser bug never silently loses a real job. The cost of that choice is that
*every* parse failure on an old posting is a stale job in the sheet, and it
looks like a successful run -- nothing is logged, nothing raises. That is the
failure these tests are here to catch.

Pure stdlib -- no network, no browser, no API key.

Run:
    python3 -m unittest discover -s . -p "test_*.py"
"""

import unittest
from datetime import date, datetime, timedelta
from unittest import mock

import job_scraper as js


def days_ago(n):
    """The date n days before today."""
    return date.today() - timedelta(days=n)


def frozen_today(y, m, d):
    """A stand-in for job_scraper's `datetime` whose now() is fixed.

    The absolute-date branch of _zr_parse_date() reasons about the current
    year, so tests for it would otherwise pass or skip depending on the day
    they run -- the year-rollover case can only be built out of real dates for
    part of the year. Subclassing datetime keeps strptime() intact and only
    pins down now().
    """
    class _Frozen(datetime):
        @classmethod
        def now(cls, tz=None):
            return cls(y, m, d, 12, 0, 0)
    return _Frozen


class TestLinkedInDaysOld(unittest.TestCase):
    """_li_days_old() reads the ISO date out of LinkedIn's
    <time datetime="YYYY-MM-DD"> element."""

    def test_missing_attribute_is_unknown(self):
        self.assertIsNone(js._li_days_old(None))
        self.assertIsNone(js._li_days_old(""))

    def test_today_is_zero(self):
        self.assertEqual(js._li_days_old(date.today().isoformat()), 0)

    def test_a_plain_date_counts_back_from_today(self):
        for n in (1, 3, 10, 45):
            with self.subTest(days=n):
                self.assertEqual(js._li_days_old(days_ago(n).isoformat()), n)

    def test_a_full_timestamp_is_truncated_to_its_date(self):
        # LinkedIn sometimes serves a full ISO timestamp in the same attribute.
        stamp = days_ago(4).isoformat() + "T09:30:00Z"
        self.assertEqual(js._li_days_old(stamp), 4)

    def test_time_of_day_does_not_shift_the_day_count(self):
        day = days_ago(2).isoformat()
        self.assertEqual(js._li_days_old(day + "T00:00:01Z"),
                         js._li_days_old(day + "T23:59:59Z"))

    def test_garbage_is_unknown_not_zero(self):
        # Zero would mean "posted today" and pass the freshness filter as fresh.
        for text in ("not a date", "2026-13-45", "yesterday", "20261120"):
            with self.subTest(text=text):
                self.assertIsNone(js._li_days_old(text))

    def test_us_format_is_not_misread(self):
        # 11/20/2026 must not be read as year 11 -- better unknown than wrong.
        self.assertIsNone(js._li_days_old("11/20/2026"))


class TestLinkedInDaysToText(unittest.TestCase):
    """_li_days_to_text() renders the parsed age for the Excel 'Posted' column,
    in Naukri's phrasing so one column reads the same for every source."""

    def test_unknown_renders_empty(self):
        self.assertEqual(js._li_days_to_text(None), "")

    def test_zero_is_today(self):
        self.assertEqual(js._li_days_to_text(0), "Today")

    def test_one_day_is_singular(self):
        self.assertEqual(js._li_days_to_text(1), "1 day ago")

    def test_several_days_are_plural(self):
        self.assertEqual(js._li_days_to_text(2), "2 days ago")
        self.assertEqual(js._li_days_to_text(30), "30 days ago")

    def test_a_future_date_reads_as_today(self):
        # A clock-skewed or scheduled posting shouldn't print "-2 days ago".
        self.assertEqual(js._li_days_to_text(-2), "Today")

    def test_output_round_trips_through_parse_days_old(self):
        # The 'Posted' column is written by this function and read back by
        # parse_days_old() in later tooling, so the two must agree.
        for n in (0, 1, 5, 21, 31):
            with self.subTest(days=n):
                self.assertEqual(js.parse_days_old(js._li_days_to_text(n)), n)


class TestZipRecruiterParseDate(unittest.TestCase):
    """_zr_parse_date() handles both shapes ZipRecruiter's jobList-date div
    shows: a relative phrase, or an absolute day-and-month."""

    def test_missing_text_is_unknown(self):
        self.assertIsNone(js._zr_parse_date(None))
        self.assertIsNone(js._zr_parse_date(""))

    def test_relative_day_counts(self):
        self.assertEqual(js._zr_parse_date("2 days ago"), 2)
        self.assertEqual(js._zr_parse_date("1 day ago"), 1)

    def test_same_day_phrasings(self):
        # The article forms ("an hour ago") carry no digit for the numeric
        # branches to find, and ZipRecruiter uses all of these.
        for text in ("Today", "just now", "just posted", "an hour ago",
                     "5 hours ago", "20 minutes ago"):
            with self.subTest(text=text):
                self.assertEqual(js._zr_parse_date(text), 0)

    def test_yesterday(self):
        self.assertEqual(js._zr_parse_date("Yesterday"), 1)

    def test_plus_suffix_is_not_dropped(self):
        # Regression: the old regex r"(\d+)\s*day" had nothing to match the
        # '+', so ZR's most common stale label parsed as None -- and an unknown
        # age is let through, so every 30+ day posting reached the sheet.
        self.assertEqual(js._zr_parse_date("30+ days ago"), 30)
        self.assertEqual(js._zr_parse_date("30 + days ago"), 30)

    def test_weeks_and_months_convert_to_days(self):
        # Regression: no week or month branch existed, so these were unknown
        # and let through as if fresh.
        self.assertEqual(js._zr_parse_date("2 weeks ago"), 14)
        self.assertEqual(js._zr_parse_date("3 months ago"), 90)

    def test_an_absolute_date_counts_back_from_today(self):
        with mock.patch.object(js, "datetime", frozen_today(2026, 6, 20)):
            self.assertEqual(js._zr_parse_date("14 Jun"), 6)
            self.assertEqual(js._zr_parse_date("20 Jun"), 0)
            self.assertEqual(js._zr_parse_date("1 Mar"), 111)

    def test_an_absolute_date_with_a_year_is_used_as_given(self):
        with mock.patch.object(js, "datetime", frozen_today(2026, 6, 20)):
            self.assertEqual(js._zr_parse_date("14 Jun 2025"), 371)

    def test_a_date_later_in_the_year_is_last_years(self):
        # Regression: a bare '%d %b' was stamped with the *current* year, so a
        # December date read in January became a future date and scored about
        # -350 days -- under every freshness limit, so a posting from last year
        # looked brand new. It is last year's date, and nearly a year old.
        with mock.patch.object(js, "datetime", frozen_today(2026, 1, 10)):
            self.assertEqual(js._zr_parse_date("14 Dec"), 27)
            self.assertEqual(js._zr_parse_date("2 Jan"), 8)

    def test_tomorrow_is_not_rolled_back_a_year(self):
        # One day of leeway for a timezone edge, so a job ZipRecruiter dates
        # one day ahead isn't reported as a year old and filtered out.
        with mock.patch.object(js, "datetime", frozen_today(2026, 6, 20)):
            self.assertEqual(js._zr_parse_date("21 Jun"), -1)
            self.assertEqual(js._zr_parse_date("22 Jun"), 363)

    def test_month_case_is_ignored(self):
        with mock.patch.object(js, "datetime", frozen_today(2026, 6, 20)):
            self.assertEqual(js._zr_parse_date("14 JUN"), 6)
            self.assertEqual(js._zr_parse_date("14 jun"), 6)

    def test_surrounding_whitespace_is_ignored(self):
        self.assertEqual(js._zr_parse_date("  4 days ago  "), 4)

    def test_unreadable_text_is_unknown_not_zero(self):
        # None means "let it through"; zero would claim it was posted today.
        for text in ("", "Featured", "Easy Apply", "—", "Jun"):
            with self.subTest(text=text):
                self.assertIsNone(js._zr_parse_date(text))

    def test_an_impossible_date_does_not_raise(self):
        # 31 Feb / 29 Feb in a non-leap year must come back unknown, not blow
        # up a scrape that has already cost a page load.
        for text in ("31 Feb", "32 Jan", "29 Feb"):
            with self.subTest(text=text):
                self.assertIsNone(js._zr_parse_date(text))

    def test_a_stale_posting_is_never_reported_as_fresh(self):
        # The invariant behind all of the above: for every phrasing ZR shows
        # for a month-old job, the parser must not return a small day count.
        for text in ("30+ days ago", "45 days ago", "2 months ago",
                     "6 weeks ago"):
            with self.subTest(text=text):
                days = js._zr_parse_date(text)
                self.assertIsNotNone(days)
                self.assertGreater(days, 7)


if __name__ == "__main__":
    unittest.main()
