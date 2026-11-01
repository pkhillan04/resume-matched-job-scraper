# resume-matched-job-scraper

A daily job-search pipeline that scrapes four job boards, scores every posting
against a real resume instead of a hand-typed keyword list, and writes a ranked
spreadsheet.

I built it because someone close to me was applying to dozens of Delhi-NCR data
analyst roles with no callbacks, and the bottleneck wasn't effort — it was
spending evenings re-reading listings that were never a fit in the first place.

**It ran 95 times between 2026-05-31 and 2026-10-30 and has deduplicated 5,405
unique postings.** The engineering below is mostly what it took to survive that:
five months of live sites changing their markup underneath it.

## What it does

1. Reads a resume PDF **once**, extracts a structured profile (skills, years,
   domains, target titles), caches it to `profile.json`, and invalidates that
   cache by MD5 only when the PDF actually changes.
2. Scrapes Naukri, LinkedIn, Indeed India, and ZipRecruiter India.
3. Filters by freshness, relevance, and location before spending anything
   expensive on a posting.
4. Scores each job against the cached profile.
5. Writes a dated `.xlsx` with scores, gaps, and a colour-coded source column.

## Design decisions worth explaining

**Four sources, four access strategies.** Naukri, Indeed, and ZipRecruiter need
a real browser (Playwright, non-headless) because they block plain HTTP —
ZipRecruiter answers 403, Indeed throws Cloudflare challenges. LinkedIn goes
through its public guest endpoint over plain HTTP instead, which means the tool
never touches a logged-in account and can't get one flagged.

**Two-layer deduplication.** URL-based dedup alone failed: boards repost the
same job with a fresh URL and it comes back as new. So a second persistent layer
keys on a normalised `(title, company)` signature, checked at scrape time across
all four sources, which also collapses the overlap an aggregator like
ZipRecruiter creates. The state file is backward-compatible with the old
URL-only format.

**Cheap filters before expensive calls.** A title-relevance filter runs before
the job description is fetched or the LLM is called. Scraping a JD costs a page
load; scoring it costs quota. Dropping an off-target title early saves both.

**Spending scarce quota where it counts.** The free Gemini tier allows roughly 20
requests a day. Rather than burning that on the first 20 jobs scraped, a keyword
baseline scores *everything* offline first, then the remaining quota upgrades
only the highest-scoring candidates. When quota runs out mid-run, a
`QuotaExhausted` exception switches the rest of the run to keyword scoring
silently instead of crashing — the run still produces a usable sheet.

**Failure isolation.** Each source is wrapped independently. Indeed hitting a
Cloudflare wall doesn't cost you the other three sources' results.

**Atomic state writes.** `seen_jobs.json` writes to a temp file and renames, so
interrupting a run can't corrupt five months of dedup history.

**Stable per-board job IDs.** Each board gets its own dedup key function, because
each exposes its ID differently and LinkedIn's full URL carries tracking
parameters that change between runs.

## Running it

```bash
pip3 install playwright openpyxl pypdf google-genai
python3 -m playwright install chromium

export GEMINI_API_KEY=...        # optional — omit to run fully offline
python3 job_scraper.py
```

Put the resume PDF next to the script as `resume.pdf`, or point `resume_path` in
`CONFIG` at it. Everything else — search keywords, locations, experience level,
freshness window, per-source toggles, relevance filters — is in the `CONFIG`
block at the top of the file.

Two scoring modes: `keyword` is offline and free forever; `gemini` is sharper and
falls back to `keyword` automatically when quota runs out.

`profile.example.json` shows the profile shape. The real `profile.json` is
gitignored because it caches your resume text verbatim.

## Honest limits

- **Selectors break.** Four job boards change their markup whenever they like.
  Lines marked `UPDATE WHEN CHANGES` are the ones that rot first. If a source
  reports `→ 0 cards found`, that's where to look.
- **Scraping is against most job boards' terms of service.** This is a personal
  tool for a single person's own job search, it runs at human pace with no
  parallelism, and it deliberately uses LinkedIn's public guest endpoint rather
  than automating a logged-in session. Know what you're agreeing to before you
  point it at anything.
- **The LLM scores, it doesn't decide.** A high score means the description
  overlaps the resume, not that the job is worth taking.
- **India-specific.** The location filters, the boards, and the date parsing all
  assume the Indian market.

## Credits

Built with Claude as a coding partner. The source additions and most of the
fixes came out of real debugging sessions against live markup — captured DOM
dumps, failing selectors, quota exhaustion mid-run — rather than from a design
written up front.

## License

MIT
