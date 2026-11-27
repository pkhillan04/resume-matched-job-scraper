"""
Unit tests for the seen-jobs store in job_scraper.py
(``load_seen_jobs`` / ``save_seen_jobs``).

This pair is the scraper's only memory. Everything else in a run is
recomputed from the live pages, so seen_jobs.json is the single piece of
state that has to survive between runs, and both directions fail quietly:

  * read it wrong and the user is handed jobs they already rejected, or
    (worse) the history is dropped and *every* job comes back as new;
  * write it wrong and a crash mid-write loses the whole record.

``load_seen_jobs`` is written to fail open — the docstring says so and the
try/except is there to make it true. It wasn't. Two parseable-but-malformed
shapes got past the try/except and took the run down with them, before a
single job had been scraped:

  1. A non-string entry (``null``, a number) in ``seen_urls``. ``set()``
     accepted it, then the LinkedIn/Indeed/ZR migration loop — which sits
     *outside* the try/except — did ``"linkedin.com" in u`` and raised
     ``TypeError: argument of type 'NoneType' is not iterable``. The same
     entry also broke the other direction: ``save_seen_jobs`` sorts the set,
     and ``sorted({None, "x"})`` raises.
  2. A bare JSON array instead of an object. ``data.get`` raised
     AttributeError into the except, so the file read as "no history at all"
     and the next run re-reported everything.

The ``TestMalformedFilesFailOpen`` cases below pin both down, plus the
ordinary corrupt-file paths that always did work.

Pure stdlib — no network, no browser, no API key. Every case writes to its
own tmpdir, so nothing here can touch a real seen_jobs.json.

Run:
    python3 -m unittest discover -s . -p "test_*.py"
"""

import json
import os
import shutil
import tempfile
import unittest

import job_scraper as js


class SeenJobsTempFile(unittest.TestCase):
    """Base: each test gets a private directory and a path inside it."""

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="seen_jobs_test_")
        self.path = os.path.join(self.dir, "seen_jobs.json")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def write_raw(self, text):
        with open(self.path, "w") as f:
            f.write(text)

    def write_json(self, obj):
        with open(self.path, "w") as f:
            json.dump(obj, f)


class TestRoundTrip(SeenJobsTempFile):
    """save → load has to return exactly what went in. These two functions are
    each other's only caller in practice, so a disagreement about the on-disk
    shape would silently reset the history on the next run."""

    def test_missing_file_is_empty_set(self):
        self.assertEqual(js.load_seen_jobs(self.path), set())

    def test_empty_set_round_trips(self):
        js.save_seen_jobs(self.path, set())
        self.assertEqual(js.load_seen_jobs(self.path), set())

    def test_urls_round_trip(self):
        # Naukri URLs are already canonical, so nothing is added on load and
        # the set comes back exactly as stored.
        urls = {
            "https://www.naukri.com/job-listings-analyst-acme-123",
            "https://www.naukri.com/job-listings-engineer-acme-456",
        }
        js.save_seen_jobs(self.path, urls)
        self.assertEqual(js.load_seen_jobs(self.path), urls)

    def test_round_trip_is_a_superset_for_migratable_urls(self):
        # Not an equality round-trip by design: loading a LinkedIn/Indeed/ZR
        # URL also adds its canonical key (see TestLoadMigratesToStableKeys),
        # so the set grows. It must never *shrink* — dropping a stored URL
        # would re-report a job the user has already seen.
        urls = {
            "https://www.ziprecruiter.in/jobs/544036357-data-analyst",
            "https://in.linkedin.com/jobs/view/analyst-at-acme-4012345678",
        }
        js.save_seen_jobs(self.path, urls)
        self.assertTrue(urls.issubset(js.load_seen_jobs(self.path)))

    def test_save_is_idempotent(self):
        urls = {"https://www.naukri.com/job-listings-x-1"}
        js.save_seen_jobs(self.path, urls)
        js.save_seen_jobs(self.path, urls)
        self.assertEqual(js.load_seen_jobs(self.path), urls)

    def test_second_save_replaces_rather_than_appends(self):
        js.save_seen_jobs(self.path, {"https://www.naukri.com/a-1"})
        js.save_seen_jobs(self.path, {"https://www.naukri.com/b-2"})
        self.assertEqual(
            js.load_seen_jobs(self.path), {"https://www.naukri.com/b-2"}
        )

    def test_written_file_is_valid_json_with_expected_keys(self):
        js.save_seen_jobs(self.path, {"https://www.naukri.com/a-1"})
        with open(self.path) as f:
            payload = json.load(f)
        self.assertEqual(payload["seen_urls"], ["https://www.naukri.com/a-1"])
        self.assertEqual(payload["total_seen"], 1)
        self.assertIn("last_updated", payload)

    def test_total_seen_matches_url_count(self):
        urls = {f"https://www.naukri.com/job-{i}" for i in range(7)}
        js.save_seen_jobs(self.path, urls)
        with open(self.path) as f:
            payload = json.load(f)
        self.assertEqual(payload["total_seen"], len(payload["seen_urls"]))
        self.assertEqual(payload["total_seen"], 7)

    def test_urls_are_sorted_on_disk(self):
        # Sorted output keeps the file's diff readable and stable between runs,
        # instead of reshuffling with Python's set iteration order.
        js.save_seen_jobs(self.path, {"https://b.example/2", "https://a.example/1"})
        with open(self.path) as f:
            payload = json.load(f)
        self.assertEqual(payload["seen_urls"], sorted(payload["seen_urls"]))


class TestMalformedFilesFailOpen(SeenJobsTempFile):
    """The docstring promises an empty set rather than an exception for a
    damaged file. A raised exception here kills the run at startup, before
    anything has been scraped — strictly worse than forgetting the history."""

    def test_unparseable_json_is_empty_set(self):
        self.write_raw("{not json at all")
        self.assertEqual(js.load_seen_jobs(self.path), set())

    def test_truncated_json_is_empty_set(self):
        self.write_raw('{"seen_urls": ["https://www.naukri.com/a-1",')
        self.assertEqual(js.load_seen_jobs(self.path), set())

    def test_empty_file_is_empty_set(self):
        self.write_raw("")
        self.assertEqual(js.load_seen_jobs(self.path), set())

    def test_missing_seen_urls_key_is_empty_set(self):
        self.write_json({"last_updated": "2026-01-01T00:00:00", "total_seen": 0})
        self.assertEqual(js.load_seen_jobs(self.path), set())

    def test_json_scalar_is_empty_set(self):
        self.write_raw("42")
        self.assertEqual(js.load_seen_jobs(self.path), set())

    # --- regression: non-string entries used to raise, not fail open ---------

    def test_null_entry_does_not_raise(self):
        # Was: TypeError: argument of type 'NoneType' is not iterable, from the
        # migration loop outside the try/except.
        self.write_json({"seen_urls": ["https://www.naukri.com/a-1", None]})
        self.assertEqual(
            js.load_seen_jobs(self.path), {"https://www.naukri.com/a-1"}
        )

    def test_numeric_entry_does_not_raise(self):
        self.write_json({"seen_urls": [12345, "https://www.naukri.com/a-1"]})
        self.assertEqual(
            js.load_seen_jobs(self.path), {"https://www.naukri.com/a-1"}
        )

    def test_nested_entry_does_not_raise(self):
        self.write_json({"seen_urls": [{"url": "https://www.naukri.com/a-1"}]})
        self.assertEqual(js.load_seen_jobs(self.path), set())

    def test_all_entries_non_string_is_empty_set(self):
        self.write_json({"seen_urls": [None, 1, True, []]})
        self.assertEqual(js.load_seen_jobs(self.path), set())

    def test_non_string_entry_survives_a_full_reload_cycle(self):
        # The entry has to be dropped on *load*, not just tolerated: the set
        # goes straight back into save_seen_jobs(), whose sorted() raises on a
        # set mixing None with str.
        self.write_json({"seen_urls": [None, "https://www.naukri.com/a-1"]})
        seen = js.load_seen_jobs(self.path)
        js.save_seen_jobs(self.path, seen)           # must not raise
        self.assertEqual(
            js.load_seen_jobs(self.path), {"https://www.naukri.com/a-1"}
        )

    def test_non_string_entry_in_a_linkedin_file_does_not_raise(self):
        # Worst case for the old code: a LinkedIn URL present (so the migration
        # loop has real work to do) *and* a null entry in the same list.
        self.write_json({
            "seen_urls": [
                None,
                "https://in.linkedin.com/jobs/view/analyst-at-acme-4012345678",
            ],
        })
        seen = js.load_seen_jobs(self.path)
        self.assertIn("https://www.linkedin.com/jobs/view/4012345678", seen)
        self.assertNotIn(None, seen)

    # --- regression: a bare array silently dropped the whole history ---------

    def test_bare_json_array_is_read_as_urls(self):
        # Was: data.get -> AttributeError -> except -> set(), i.e. the entire
        # dedup history discarded and every job re-reported as new.
        self.write_json(["https://www.naukri.com/a-1", "https://www.naukri.com/b-2"])
        self.assertEqual(
            js.load_seen_jobs(self.path),
            {"https://www.naukri.com/a-1", "https://www.naukri.com/b-2"},
        )

    def test_bare_json_array_also_filters_non_strings(self):
        self.write_json(["https://www.naukri.com/a-1", None, 7])
        self.assertEqual(
            js.load_seen_jobs(self.path), {"https://www.naukri.com/a-1"}
        )

    def test_empty_bare_array_is_empty_set(self):
        self.write_json([])
        self.assertEqual(js.load_seen_jobs(self.path), set())


class TestLoadMigratesToStableKeys(SeenJobsTempFile):
    """Older files stored jobs by whatever URL the page gave, which for
    LinkedIn and ZipRecruiter was not stable between runs. load_seen_jobs adds
    the canonical key alongside the stored URL so a job already seen under the
    old spelling is still recognized. The keys themselves are tested in
    test_dedup_keys.py; what matters here is that loading applies them."""

    def test_linkedin_url_gains_its_canonical_key(self):
        stored = "https://in.linkedin.com/jobs/view/data-analyst-at-acme-4012345678"
        self.write_json({"seen_urls": [stored]})
        seen = js.load_seen_jobs(self.path)
        self.assertIn(stored, seen)                       # original kept
        self.assertIn(js.linkedin_dedup_key(stored), seen)  # canonical added

    def test_indeed_url_with_jk_gains_its_canonical_key(self):
        stored = "https://in.indeed.com/rc/clk?jk=abz1234567890xy&fccid=xyz"
        self.write_json({"seen_urls": [stored]})
        seen = js.load_seen_jobs(self.path)
        self.assertIn("https://in.indeed.com/viewjob?jk=abz1234567890xy", seen)

    def test_ziprecruiter_url_gains_its_canonical_key(self):
        stored = "https://www.ziprecruiter.in/jobs/544036357-analyst?utm_source=x"
        self.write_json({"seen_urls": [stored]})
        seen = js.load_seen_jobs(self.path)
        self.assertIn("https://www.ziprecruiter.in/jobs/544036357", seen)

    def test_naukri_url_is_left_alone(self):
        # Naukri URLs are already stable; nothing should be invented for them.
        stored = "https://www.naukri.com/job-listings-analyst-acme-123"
        self.write_json({"seen_urls": [stored]})
        self.assertEqual(js.load_seen_jobs(self.path), {stored})

    def test_migration_is_idempotent_across_a_save_reload(self):
        # The migrated set is what gets saved, so a second load must not keep
        # growing the file. Two runs in a row have to reach a fixed point.
        stored = "https://in.linkedin.com/jobs/view/analyst-at-acme-4012345678"
        self.write_json({"seen_urls": [stored]})

        first = js.load_seen_jobs(self.path)
        js.save_seen_jobs(self.path, first)
        second = js.load_seen_jobs(self.path)
        js.save_seen_jobs(self.path, second)
        third = js.load_seen_jobs(self.path)

        self.assertEqual(first, second)
        self.assertEqual(second, third)

    def test_already_canonical_linkedin_url_adds_nothing(self):
        canon = "https://www.linkedin.com/jobs/view/4012345678"
        self.write_json({"seen_urls": [canon]})
        self.assertEqual(js.load_seen_jobs(self.path), {canon})

    def test_two_spellings_of_one_linkedin_job_share_a_key(self):
        # The point of the migration: the same posting stored twice under
        # different hosts/slugs collapses onto one canonical key.
        a = "https://in.linkedin.com/jobs/view/analyst-at-acme-4012345678"
        b = "https://www.linkedin.com/jobs/view/senior-analyst-acme-4012345678"
        self.write_json({"seen_urls": [a, b]})
        seen = js.load_seen_jobs(self.path)
        self.assertIn("https://www.linkedin.com/jobs/view/4012345678", seen)
        self.assertEqual(
            len([u for u in seen if u.endswith("/jobs/view/4012345678")]), 1
        )


class TestSaveIsAtomic(SeenJobsTempFile):
    """The temp-file + rename dance exists so a crash or Ctrl+C mid-write can
    never publish a half-written file. Tested here because the failure only
    shows up on the *next* run, as an unreadable history."""

    def test_no_temp_file_left_after_a_successful_save(self):
        js.save_seen_jobs(self.path, {"https://www.naukri.com/a-1"})
        self.assertFalse(os.path.exists(self.path + ".tmp"))

    def test_existing_file_survives_a_failed_save(self):
        good = {"https://www.naukri.com/a-1"}
        js.save_seen_jobs(self.path, good)
        with self.assertRaises(TypeError):
            js.save_seen_jobs(self.path, {object()})   # unsortable
        self.assertEqual(js.load_seen_jobs(self.path), good)

    def test_no_temp_file_left_after_a_failed_save(self):
        # Was: the partial .tmp stayed on disk next to seen_jobs.json, looking
        # like recoverable state when its contents are arbitrary.
        with self.assertRaises(TypeError):
            js.save_seen_jobs(self.path, {object()})
        self.assertFalse(os.path.exists(self.path + ".tmp"))

    def test_failed_save_on_a_fresh_path_creates_nothing(self):
        with self.assertRaises(TypeError):
            js.save_seen_jobs(self.path, {object()})
        self.assertFalse(os.path.exists(self.path))
        self.assertFalse(os.path.exists(self.path + ".tmp"))

    def test_stale_temp_file_does_not_affect_the_next_load(self):
        js.save_seen_jobs(self.path, {"https://www.naukri.com/a-1"})
        self.write_raw_tmp('{"seen_urls": ["https://www.naukri.com/GARBAGE"')
        self.assertEqual(
            js.load_seen_jobs(self.path), {"https://www.naukri.com/a-1"}
        )

    def write_raw_tmp(self, text):
        with open(self.path + ".tmp", "w") as f:
            f.write(text)


if __name__ == "__main__":
    unittest.main()
