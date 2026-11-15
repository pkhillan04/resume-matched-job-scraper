"""
Unit tests for the per-source dedup keys in job_scraper.py.

Dedup is the one piece of this scraper whose failures are invisible. A key that
is too *loose* collapses two real postings onto one string and the second job is
silently dropped as a duplicate -- the user never learns it existed. A key that
is too *strict* lets the same posting through on every run, so the sheet fills
with jobs already seen. Neither shows up as an error, which is exactly why these
functions need tests.

The same key function runs in two places and the two must agree:

  * at scrape time, on whatever identifier the page yielded
    (a bare ``data-jk`` for Indeed, a full link for LinkedIn and ZipRecruiter);
  * in ``load_seen_jobs()``, on the apply URL stored in seen_jobs.json.

If those two disagree for one posting, dedup fails open. The
``TestKeysAgreeAcrossCallSites`` cases below pin that invariant down.

Pure stdlib -- no network, no browser, no API key.

Run:
    python3 -m unittest discover -s . -p "test_*.py"
"""

import unittest

import job_scraper as js


class TestLinkedinDedupKey(unittest.TestCase):
    """LinkedIn guest URLs are not stable across runs: the same posting comes
    back on different hosts (in.linkedin.com vs www.linkedin.com), with
    different slug text and rotating refId/trk params. Only the trailing numeric
    job ID is stable, so every spelling has to canonicalize to it."""

    CANON = "https://www.linkedin.com/jobs/view/4012345678"

    def test_missing_url_is_empty_string(self):
        self.assertEqual(js.linkedin_dedup_key(None), "")
        self.assertEqual(js.linkedin_dedup_key(""), "")

    def test_every_spelling_of_one_posting_collapses(self):
        spellings = [
            "https://in.linkedin.com/jobs/view/data-analyst-at-acme-4012345678?refId=abc",
            "https://www.linkedin.com/jobs/view/data-analyst-at-acme-4012345678",
            "https://www.linkedin.com/jobs/view/business-analyst-at-acme-4012345678",
            "https://www.linkedin.com/jobs/view/data-analyst-at-acme-4012345678/?trk=x",
            "https://www.linkedin.com/jobs/view/4012345678",
            "https://www.linkedin.com/jobs/view/4012345678/",
            "https://www.linkedin.com/jobs/view/4012345678#apply",
        ]
        for url in spellings:
            with self.subTest(url=url):
                self.assertEqual(js.linkedin_dedup_key(url), self.CANON)

    def test_different_postings_stay_distinct(self):
        self.assertNotEqual(
            js.linkedin_dedup_key("https://www.linkedin.com/jobs/view/analyst-at-acme-4012345678"),
            js.linkedin_dedup_key("https://www.linkedin.com/jobs/view/analyst-at-acme-4012345679"),
        )

    def test_id_is_taken_from_the_end_not_the_slug(self):
        """A slug can contain its own digits (a "2026" intake, a "123" in the
        company name). The job ID is the trailing group, not the first one."""
        self.assertEqual(
            js.linkedin_dedup_key(
                "https://www.linkedin.com/jobs/view/analyst-2026-batch-at-acme-123-4012345678"
            ),
            self.CANON,
        )

    def test_non_job_url_falls_back_without_crashing(self):
        """Not every link the scraper sees is a posting. The fallback just has to
        be stable and distinct -- it must not raise, and must not collapse."""
        a = js.linkedin_dedup_key("https://www.linkedin.com/jobs/collections/recommended")
        b = js.linkedin_dedup_key("https://www.linkedin.com/jobs/collections/top-applicant")
        self.assertTrue(a)
        self.assertNotEqual(a, b)


class TestIndeedDedupKey(unittest.TestCase):
    """Indeed links vary by route (/viewjob?jk=, /rc/clk?jk=, /pagead/...&jk=)
    and carry rotating tracking params, but the jk is the stable per-posting ID.
    The scraper reads a bare jk off ``data-jk``, while load_seen_jobs() only ever
    sees the stored apply URL, so both inputs must produce one key."""

    CANON = "https://in.indeed.com/viewjob?jk=a1b2c3d4e5f6a7b8"

    def test_missing_input_is_empty_string(self):
        self.assertEqual(js.indeed_dedup_key(None), "")
        self.assertEqual(js.indeed_dedup_key(""), "")

    def test_bare_jk_and_every_url_route_collapse(self):
        inputs = [
            "a1b2c3d4e5f6a7b8",
            "https://in.indeed.com/viewjob?jk=a1b2c3d4e5f6a7b8",
            "https://in.indeed.com/viewjob?jk=a1b2c3d4e5f6a7b8&tk=xyz&from=serp",
            "https://in.indeed.com/rc/clk?jk=a1b2c3d4e5f6a7b8&fccid=abc&vjs=3",
            "https://in.indeed.com/pagead/clk?mo=r&ad=longtrackingblob&jk=a1b2c3d4e5f6a7b8",
            "https://www.indeed.com/viewjob?jk=a1b2c3d4e5f6a7b8",
        ]
        for value in inputs:
            with self.subTest(value=value):
                self.assertEqual(js.indeed_dedup_key(value), self.CANON)

    def test_different_jks_stay_distinct(self):
        self.assertNotEqual(
            js.indeed_dedup_key("https://in.indeed.com/viewjob?jk=a1b2c3d4e5f6a7b8"),
            js.indeed_dedup_key("https://in.indeed.com/viewjob?jk=a1b2c3d4e5f6a7b9"),
        )

    def test_whitespace_around_a_bare_jk_is_ignored(self):
        """data-jk is read straight off the page and can arrive padded."""
        self.assertEqual(js.indeed_dedup_key("  a1b2c3d4e5f6a7b8\n"), self.CANON)

    # ── Regressions: the jk is an opaque token, not necessarily hex ──────────

    def test_non_hex_jk_is_matched_in_full(self):
        """Regression: the old pattern was [0-9a-fA-F]+ and unanchored, so
        "jk=abz123456789" matched only the leading "ab" and the whole posting was
        keyed on a 2-character prefix -- any other jk starting "ab" collided with
        it and was dropped as a duplicate."""
        self.assertEqual(
            js.indeed_dedup_key("https://in.indeed.com/viewjob?jk=abz123456789"),
            "https://in.indeed.com/viewjob?jk=abz123456789",
        )

    def test_jk_starting_with_a_non_hex_char_does_not_collapse(self):
        """Regression: a jk whose first character was not hex matched nothing, so
        the fallback stripped the query and EVERY such posting collapsed onto the
        bare "https://in.indeed.com/viewjob". Only the first one survived dedup;
        all the rest were silently discarded."""
        a = js.indeed_dedup_key("https://in.indeed.com/viewjob?jk=zz111111")
        b = js.indeed_dedup_key("https://in.indeed.com/viewjob?jk=zz222222")
        self.assertNotEqual(a, b)
        self.assertNotEqual(a, "https://in.indeed.com/viewjob")
        self.assertIn("zz111111", a)

    def test_jks_sharing_a_hex_prefix_stay_distinct(self):
        """The failure mode the unanchored pattern produced, stated directly."""
        keys = {
            js.indeed_dedup_key("https://in.indeed.com/viewjob?jk=abc_one"),
            js.indeed_dedup_key("https://in.indeed.com/viewjob?jk=abc_two"),
            js.indeed_dedup_key("https://in.indeed.com/viewjob?jk=abc_three"),
        }
        self.assertEqual(len(keys), 3)

    def test_case_is_preserved_not_folded(self):
        """A jk is opaque. Folding case would merge two postings that differ only
        in case, and losing a real job is worse than showing one twice."""
        self.assertNotEqual(
            js.indeed_dedup_key("A1B2C3D4E5F6A7B8"),
            js.indeed_dedup_key("a1b2c3d4e5f6a7b8"),
        )

    def test_indeed_url_without_a_jk_falls_back_without_collapsing(self):
        a = js.indeed_dedup_key("https://in.indeed.com/company/acme/jobs/analyst-111")
        b = js.indeed_dedup_key("https://in.indeed.com/company/acme/jobs/analyst-222")
        self.assertTrue(a)
        self.assertNotEqual(a, b)


class TestZiprecruiterDedupKey(unittest.TestCase):
    """ZipRecruiter India posting URLs are /jobs/<numeric-id>-<slug>. The numeric
    ID is the stable part; the slug and the tracking params are not. ZR is an
    aggregator, so the same posting also arrives from the other sources -- a
    loose key here means duplicate rows in the sheet."""

    CANON = "https://www.ziprecruiter.in/jobs/544036357"

    def test_missing_url_is_empty_string(self):
        self.assertEqual(js.ziprecruiter_dedup_key(None), "")
        self.assertEqual(js.ziprecruiter_dedup_key(""), "")

    def test_every_spelling_of_one_posting_collapses(self):
        spellings = [
            "https://www.ziprecruiter.in/jobs/544036357-data-analyst",
            "https://www.ziprecruiter.in/jobs/544036357-data-analyst?utm_source=search",
            "https://www.ziprecruiter.in/jobs/544036357-data-analyst-noida/",
            "https://www.ziprecruiter.in/jobs/544036357-business-analyst-acme#apply",
            "https://www.ziprecruiter.in/job/544036357",
        ]
        for url in spellings:
            with self.subTest(url=url):
                self.assertEqual(js.ziprecruiter_dedup_key(url), self.CANON)

    def test_different_postings_stay_distinct(self):
        self.assertNotEqual(
            js.ziprecruiter_dedup_key("https://www.ziprecruiter.in/jobs/544036357-analyst"),
            js.ziprecruiter_dedup_key("https://www.ziprecruiter.in/jobs/544036358-analyst"),
        )

    def test_slug_digits_do_not_become_the_id(self):
        """"/jobs/544036357-analyst-2026" must key on 544036357, not 2026."""
        self.assertEqual(
            js.ziprecruiter_dedup_key(
                "https://www.ziprecruiter.in/jobs/544036357-data-analyst-2026-intake"
            ),
            self.CANON,
        )

    def test_jid_param_is_the_second_choice_key(self):
        """Company-page links carry no path ID but do carry ?jid=."""
        self.assertEqual(
            js.ziprecruiter_dedup_key("https://www.ziprecruiter.in/c/Acme/Job/Data-Analyst?jid=ab12cd34"),
            "https://www.ziprecruiter.in/jobs/ab12cd34",
        )

    def test_jid_is_case_insensitive_for_one_posting(self):
        self.assertEqual(
            js.ziprecruiter_dedup_key("https://www.ziprecruiter.in/c/Acme/Job/A?jid=AB12CD34"),
            js.ziprecruiter_dedup_key("https://www.ziprecruiter.in/c/Acme/Job/A?jid=ab12cd34"),
        )

    def test_path_fallback_is_stable_and_distinct(self):
        """With neither a path ID nor a jid, the query-stripped path is the key.
        It only has to be deterministic and not collapse two different pages."""
        a = js.ziprecruiter_dedup_key("https://www.ziprecruiter.in/c/Acme/Job/Data-Analyst?x=1")
        b = js.ziprecruiter_dedup_key("https://www.ziprecruiter.in/c/Acme/Job/Data-Analyst?x=2")
        c = js.ziprecruiter_dedup_key("https://www.ziprecruiter.in/c/Acme/Job/BI-Analyst")
        self.assertEqual(a, b)
        self.assertNotEqual(a, c)


class TestKeysAgreeAcrossCallSites(unittest.TestCase):
    """Each key function is called once at scrape time and again in
    load_seen_jobs() on the URL stored in seen_jobs.json. If the two calls
    disagree for one posting, the job is re-shown on every single run -- the
    exact thing the seen file exists to prevent.
    """

    def test_linkedin_scrape_key_matches_stored_url_key(self):
        scraped = "https://in.linkedin.com/jobs/view/data-analyst-at-acme-4012345678?refId=abc"
        stored = js.linkedin_dedup_key(scraped)          # what goes in seen_jobs.json
        self.assertEqual(js.linkedin_dedup_key(stored), stored)

    def test_indeed_bare_jk_key_matches_stored_apply_url_key(self):
        """Regression: the scraper keys on the bare data-jk while the seen file
        holds "…/viewjob?jk=<jk>". For a non-hex jk the two used to produce
        different keys, so the posting was never recognised as seen."""
        for jk in ("a1b2c3d4e5f6a7b8", "abz123456789", "zz111111", "a-b_c1"):
            with self.subTest(jk=jk):
                apply_url = f"https://in.indeed.com/viewjob?jk={jk}"
                self.assertEqual(js.indeed_dedup_key(jk), js.indeed_dedup_key(apply_url))

    def test_ziprecruiter_scrape_key_matches_stored_url_key(self):
        scraped = "https://www.ziprecruiter.in/jobs/544036357-data-analyst?utm_source=x"
        stored = js.ziprecruiter_dedup_key(scraped)
        self.assertEqual(js.ziprecruiter_dedup_key(stored), stored)

    def test_keys_are_idempotent(self):
        """Applying a key function to its own output must be a no-op, or a job
        re-keyed on a later run drifts to a new identity."""
        pairs = [
            (js.linkedin_dedup_key, "https://www.linkedin.com/jobs/view/analyst-at-acme-4012345678"),
            (js.indeed_dedup_key, "https://in.indeed.com/rc/clk?jk=a1b2c3d4e5f6a7b8&vjs=3"),
            (js.ziprecruiter_dedup_key, "https://www.ziprecruiter.in/jobs/544036357-analyst/"),
        ]
        for fn, url in pairs:
            with self.subTest(fn=fn.__name__):
                once = fn(url)
                self.assertEqual(fn(once), once)

    def test_sources_do_not_collide_with_each_other(self):
        """All three keys share one seen-jobs set, so their namespaces must not
        overlap -- a LinkedIn key must never equal an Indeed or ZR key."""
        keys = {
            js.linkedin_dedup_key("https://www.linkedin.com/jobs/view/analyst-4012345678"),
            js.indeed_dedup_key("https://in.indeed.com/viewjob?jk=4012345678"),
            js.ziprecruiter_dedup_key("https://www.ziprecruiter.in/jobs/4012345678-analyst"),
        }
        self.assertEqual(len(keys), 3)


if __name__ == "__main__":
    unittest.main()
