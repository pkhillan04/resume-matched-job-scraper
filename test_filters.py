"""
Unit tests for the two pre-fetch filters in job_scraper.py.

``passes_relevance()`` and ``ziprecruiter_location_ok()`` run *before* a job's
JD is fetched and before it reaches the analyzer, so both of them fail silently
in the expensive direction:

  * too strict -> a real, wanted posting is dropped and never appears in the
    sheet. The user cannot tell the difference between "that job does not
    exist" and "the filter ate it".
  * too loose -> off-target jobs burn a Gemini call each and clutter the
    output.

Both read their term lists out of the module-level ``CONFIG`` dict, which the
real script mutates from CLI flags. Every case here therefore patches CONFIG
explicitly and restores it in tearDown, so the tests pin the *filter logic*
rather than today's default keyword lists.

Pure stdlib -- no network, no browser, no API key.

Run:
    python3 -m unittest discover -s . -p "test_*.py"
"""

import unittest

import job_scraper as js


class ConfigPatchMixin:
    """Patch CONFIG keys for one test and put the originals back afterwards."""

    def patch_config(self, **values):
        if not hasattr(self, "_saved_config"):
            self._saved_config = {}
        for key, value in values.items():
            if key not in self._saved_config:
                self._saved_config[key] = js.CONFIG.get(key)
            js.CONFIG[key] = value

    def tearDown(self):
        for key, value in getattr(self, "_saved_config", {}).items():
            js.CONFIG[key] = value
        self._saved_config = {}


class TestPassesRelevance(ConfigPatchMixin, unittest.TestCase):
    """The coarse title gate. Keep-list decides what is in scope, drop-list
    overrides it for titles that are in scope but at the wrong level."""

    def setUp(self):
        self.patch_config(
            relevance_filter=True,
            title_keep=["analyst", "analytics", "business intelligence", "mis"],
            title_drop=["senior", "sr", "manager", "data engineer"],
        )

    def test_filter_disabled_keeps_everything(self):
        """relevance_filter=False must be a true bypass, not a softer filter --
        it is the escape hatch when the keep-list is cutting too much."""
        self.patch_config(relevance_filter=False)
        for title in ("Senior Data Engineer", "", None, "Chief Forestry Officer"):
            with self.subTest(title=title):
                self.assertTrue(js.passes_relevance(title))

    def test_missing_title_is_dropped(self):
        """A job with no title cannot be judged, and scraping produces these
        when a card's markup changes. Dropping is the safe direction: it costs
        one listing, where keeping it costs an analyzer call on nothing."""
        self.assertFalse(js.passes_relevance(None))
        self.assertFalse(js.passes_relevance(""))

    def test_keeps_target_titles(self):
        for title in ("Data Analyst", "Business Analyst", "Analytics Consultant",
                      "MIS Executive", "Business Intelligence Developer"):
            with self.subTest(title=title):
                self.assertTrue(js.passes_relevance(title))

    def test_case_and_surrounding_whitespace_are_ignored(self):
        for title in ("DATA ANALYST", "data analyst", "  Data Analyst  "):
            with self.subTest(title=title):
                self.assertTrue(js.passes_relevance(title))

    def test_drops_titles_outside_the_keep_list(self):
        for title in ("Software Developer", "Account Executive", "Chef de Partie"):
            with self.subTest(title=title):
                self.assertFalse(js.passes_relevance(title))

    def test_drop_list_overrides_a_keep_match(self):
        """These all match the keep-list on "analyst" and must still go."""
        for title in ("Senior Data Analyst", "Sr. Analyst - MIS",
                      "Analytics Manager", "Data Engineer / Analyst"):
            with self.subTest(title=title):
                self.assertFalse(js.passes_relevance(title))

    def test_keep_terms_match_whole_words_only(self):
        """Substring matching on a short keep term would pull in unrelated
        roles; "mis" inside "Dismissal" or "Commissioning" is the obvious
        trap, and "mis" is short enough that plural tolerance stays off."""
        for title in ("Dismissal Case Officer", "Commissioning Engineer",
                      "Missionary Coordinator"):
            with self.subTest(title=title):
                self.assertFalse(js.passes_relevance(title))

    def test_multi_word_terms_match_as_substrings(self):
        """Multi-word terms are matched loosely on purpose -- punctuation and
        joining words vary too much between sources to require word breaks."""
        self.assertTrue(js.passes_relevance("Business Intelligence/Analytics Lead-In"))
        self.assertFalse(js.passes_relevance("Data Engineering Analyst"))

    def test_plural_titles_are_treated_like_the_singular(self):
        """Postings hiring for more than one seat title the role in the plural
        ("Business Analysts - 3 positions"). Those are the same jobs the
        keep-list is for, so a plural must not slip past the gate."""
        for title in ("Business Analysts", "Data Analysts - 3 Positions",
                      "MIS / Reporting Analysts"):
            with self.subTest(title=title):
                self.assertTrue(js.passes_relevance(title))

    def test_plural_drop_terms_still_drop(self):
        """Same leniency has to apply to the drop-list, or a plural title
        sneaks an out-of-level role through the keep-list unchallenged."""
        for title in ("Analytics Managers", "Senior Analysts"):
            with self.subTest(title=title):
                self.assertFalse(js.passes_relevance(title))

    def test_empty_keep_list_drops_everything(self):
        """An empty keep-list means "nothing is in scope". Treating it as
        "everything is in scope" would quietly flood the sheet."""
        self.patch_config(title_keep=[])
        self.assertFalse(js.passes_relevance("Data Analyst"))

    def test_empty_drop_list_keeps_every_keep_match(self):
        self.patch_config(title_drop=[])
        self.assertTrue(js.passes_relevance("Senior Data Analyst"))

    def test_regex_metacharacters_in_a_term_are_literal(self):
        """Config terms are user-editable, so a term like "c++" or "(bi)" must
        be matched literally instead of blowing up re.compile."""
        self.patch_config(title_keep=["c++"], title_drop=[])
        self.assertTrue(js.passes_relevance("Analyst (c++)"))
        self.assertFalse(js.passes_relevance("Analyst (python)"))


class TestZipRecruiterLocationOk(ConfigPatchMixin, unittest.TestCase):
    """ZipRecruiter is an aggregator with loose geo tagging -- it returns
    all-India and foreign postings under a Delhi search -- so this is the only
    thing keeping out-of-area jobs out of the sheet."""

    def setUp(self):
        self.patch_config(
            ziprecruiter_location_allow=["delhi", "gurgaon", "gurugram", "noida", "ncr"],
        )

    def test_empty_allow_list_disables_the_check(self):
        for allow in ([], None):
            with self.subTest(allow=allow):
                self.patch_config(ziprecruiter_location_allow=allow)
                self.assertTrue(js.ziprecruiter_location_ok("Kuala Lumpur"))

    def test_keeps_allowed_places(self):
        for loc in ("New Delhi, Delhi", "Gurugram, Haryana", "Noida",
                    "Delhi / NCR", "gurgaon"):
            with self.subTest(loc=loc):
                self.assertTrue(js.ziprecruiter_location_ok(loc))

    def test_drops_out_of_area_places(self):
        for loc in ("Bengaluru, Karnataka", "Kuala Lumpur", "Pune", "Dubai"):
            with self.subTest(loc=loc):
                self.assertFalse(js.ziprecruiter_location_ok(loc))

    def test_unknown_location_is_kept(self):
        """No location text means the card did not carry one -- a selector
        change, not an out-of-area job. Dropping those would silently empty a
        whole source, so an unknown location defers to the relevance gate and
        the score instead."""
        self.assertTrue(js.ziprecruiter_location_ok(None))
        self.assertTrue(js.ziprecruiter_location_ok(""))

    def test_whitespace_only_location_counts_as_unknown(self):
        """Location is read out of the DOM, so "no location" arrives as
        padding and newlines far more often than as a clean empty string.
        Those have to take the same path as "" or the unknown-location rule
        above does nothing in practice."""
        for loc in ("   ", "\n", "\n  \t "):
            with self.subTest(loc=repr(loc)):
                self.assertTrue(js.ziprecruiter_location_ok(loc))
