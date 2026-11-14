"""
Unit tests for the pure helpers in job_scraper.py.

These cover the stdlib-only functions -- freshness parsing, URL
normalization -- that every scraped job passes through before it reaches a
network call, a Gemini call or the spreadsheet. They need no network, no
browser and no API key.

Run:
    python3 -m unittest discover -s . -p "test_*.py"
"""

import unittest

import job_scraper as js


class TestParseDaysOld(unittest.TestCase):
    """parse_days_old() feeds the max_days_old freshness filter, so a wrong
    answer either drops fresh jobs or lets stale ones into the sheet."""

    def test_missing_text_is_unknown(self):
        self.assertIsNone(js.parse_days_old(None))
        self.assertIsNone(js.parse_days_old(""))

    def test_same_day_phrasings(self):
        for text in ("Just now", "Few hours ago", "Few minutes ago", "Posted today"):
            with self.subTest(text=text):
                self.assertEqual(js.parse_days_old(text), 0)

    def test_hours_round_down_to_same_day(self):
        self.assertEqual(js.parse_days_old("5 hours ago"), 0)
        self.assertEqual(js.parse_days_old("23 hours ago"), 0)

    def test_yesterday_and_one_day(self):
        self.assertEqual(js.parse_days_old("Yesterday"), 1)
        self.assertEqual(js.parse_days_old("1 day ago"), 1)
        self.assertEqual(js.parse_days_old("Posted 1 day ago"), 1)

    def test_plain_day_counts(self):
        self.assertEqual(js.parse_days_old("Posted 3 days ago"), 3)
        self.assertEqual(js.parse_days_old("7 days ago"), 7)

    def test_day_counts_containing_a_literal_one_day(self):
        """Regression: "21 days ago" contains the substring "1 day".

        The old implementation checked that substring before matching the
        number, so 11/21/31-day-old postings were reported as 1 day old and
        passed the freshness filter. Every one of these must report its real age.
        """
        self.assertEqual(js.parse_days_old("11 days ago"), 11)
        self.assertEqual(js.parse_days_old("Posted 21 days ago"), 21)
        self.assertEqual(js.parse_days_old("31 days ago"), 31)

    def test_plus_suffix(self):
        self.assertEqual(js.parse_days_old("30+ days ago"), 30)
        self.assertEqual(js.parse_days_old("30 + days ago"), 30)

    def test_weeks_and_months_convert_to_days(self):
        self.assertEqual(js.parse_days_old("1 week ago"), 7)
        self.assertEqual(js.parse_days_old("2 weeks ago"), 14)
        self.assertEqual(js.parse_days_old("1 month ago"), 30)
        self.assertEqual(js.parse_days_old("3 months ago"), 90)

    def test_case_and_whitespace_are_ignored(self):
        self.assertEqual(js.parse_days_old("JUST NOW"), 0)
        self.assertEqual(js.parse_days_old("Posted 4  Days  Ago"), 4)

    def test_days_win_over_a_trailing_hour_count(self):
        self.assertEqual(js.parse_days_old("2 days 6 hours ago"), 2)

    def test_unparseable_text_is_unknown_not_zero(self):
        """None means "no idea" and the caller lets the job through; returning 0
        would silently claim every unrecognised format is fresh."""
        for text in ("Posted on 12/11/2026", "Be an early applicant", "   "):
            with self.subTest(text=text):
                self.assertIsNone(js.parse_days_old(text))


class TestNormalizeJobUrl(unittest.TestCase):
    """normalize_job_url() is the fallback dedup key for every source, so two
    spellings of one posting have to collapse to the same string."""

    def test_missing_url_is_empty_string(self):
        self.assertEqual(js.normalize_job_url(None), "")
        self.assertEqual(js.normalize_job_url(""), "")

    def test_strips_query_and_fragment(self):
        self.assertEqual(
            js.normalize_job_url("https://example.com/jobs/abc?utm_source=x&ref=y"),
            "https://example.com/jobs/abc",
        )
        self.assertEqual(
            js.normalize_job_url("https://example.com/jobs/abc#apply"),
            "https://example.com/jobs/abc",
        )

    def test_strips_trailing_slashes(self):
        self.assertEqual(
            js.normalize_job_url("https://example.com/jobs/abc/"),
            "https://example.com/jobs/abc",
        )
        self.assertEqual(
            js.normalize_job_url("https://example.com/jobs/abc///"),
            "https://example.com/jobs/abc",
        )

    def test_lowercases(self):
        self.assertEqual(
            js.normalize_job_url("HTTPS://Example.COM/Jobs/Data-Analyst"),
            "https://example.com/jobs/data-analyst",
        )

    def test_same_posting_different_tracking_params_collapses(self):
        a = "https://example.com/jobs/data-analyst-123?src=search&page=2"
        b = "https://example.com/jobs/data-analyst-123/?src=email"
        self.assertEqual(js.normalize_job_url(a), js.normalize_job_url(b))

    def test_different_postings_stay_distinct(self):
        self.assertNotEqual(
            js.normalize_job_url("https://example.com/jobs/analyst-123"),
            js.normalize_job_url("https://example.com/jobs/analyst-124"),
        )


class TestModuleImportsWithoutOptionalDeps(unittest.TestCase):
    """These tests exist at all because importing job_scraper no longer requires
    Playwright or the Gemini SDK; require_deps() enforces them at run time."""

    def test_require_deps_is_callable_and_config_driven(self):
        self.assertTrue(callable(js.require_deps))
        self.assertIsInstance(js._MISSING_DEPS, list)


if __name__ == "__main__":
    unittest.main()
