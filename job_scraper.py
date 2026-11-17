"""
Naukri + LinkedIn Data Analyst Scraper + Resume-Aware JD Matcher — FREE
========================================================================
Reads the resume ONCE, extracts a profile ONCE (cached to profile.json),
then scrapes Naukri AND LinkedIn jobs and scores each JD against THE CANDIDATE'S ACTUAL RESUME.

LinkedIn uses the PUBLIC guest endpoint (no login) — it never touches the
candidate's account, so there's zero risk of the candidate's LinkedIn getting flagged/banned. Both
sources flow through the SAME freshness filter, the SAME seen_jobs dedup, the
SAME matcher, and land in ONE Excel (a "Source" column marks Naukri vs LinkedIn).

Efficiency design:
  - Resume PDF is read once at startup.
  - Profile extraction (Gemini) happens once, then cached to profile.json.
    On future runs it reuses the cache unless the resume file changes.
  - Each job is matched against the cached profile — no resume re-processing.

TWO MODES (set "analyzer" in CONFIG):
  - "keyword" : offline, free forever. Matches JD against skills pulled from resume.
  - "gemini"  : free Google API key. Smarter JD<->resume comparison per job.

USAGE:
    export GEMINI_API_KEY=...        # only needed for gemini mode
    python3 naukri_scraper.py
"""

import os
import re
import sys
import json
import time
import hashlib
import asyncio
import functools
from datetime import datetime

# ============================================================================
# CONFIG — EDIT THESE
# ============================================================================
CONFIG = {
    # ── Analyzer ────────────────────────────────────────────────────────────
    "analyzer": "gemini",          # "keyword" or "gemini"

    "gemini_api_key": os.environ.get("GEMINI_API_KEY", ""),
    "gemini_model": "gemini-2.5-flash-lite",

    # ── Resume ──────────────────────────────────────────────────────────────
    # Path to the candidate's resume PDF. Put the PDF next to this script, or give full path.
    "resume_path": "resume.pdf",
    "profile_cache": "profile.json",   # extracted profile is cached here

    # ── Search params ───────────────────────────────────────────────────────
    "search_keyword": "data analyst",
    "locations": ["delhi-ncr", "gurgaon", "noida"],
    "experience_years": 2,             # target years of experience
    "max_jobs_per_location": 20,
    "max_days_old": 7,                 # skip jobs posted more than N days ago

    # ── LinkedIn (public guest listings — no login, no account risk) ──────────
    "include_linkedin": True,          # set False to skip LinkedIn entirely
    "linkedin_locations": ["Delhi, India", "Gurugram, India", "Noida, India"],
    "max_jobs_per_location_linkedin": 20,
    # LinkedIn experience levels: 1=Intern 2=Entry 3=Associate 4=Mid-Senior
    # 5=Director. "2,3,4" keeps it around the target level. "" = no filter.
    "linkedin_exp_filter": "2,3,4",

    # ── Indeed (in.indeed.com via Playwright — real browser, anti-bot heavy) ───
    # Best India coverage of the three, but the most fragile: Indeed throws
    # Cloudflare/bot challenges, so this is the source most likely to return 0 or
    # need selector tweaks. Set include_indeed=False to skip it.
    "include_indeed": True,
    "indeed_locations": ["Delhi", "Gurgaon", "Noida"],
    "max_jobs_per_location_indeed": 20,

    # ── ZipRecruiter India (ziprecruiter.in via Playwright — real browser) ────
    # Well-populated and sortable by publish date, but blocks plain HTTP (403),
    # so it renders in a real browser like Naukri/Indeed. It's an aggregator, so
    # (a) expect overlap with the other sources (dedup collapses it) and (b) loose
    # location tagging — hence the extra geo check.
    "include_ziprecruiter": True,
    "ziprecruiter_locations": ["Delhi", "Gurgaon", "Noida"],
    "max_jobs_per_location_ziprecruiter": 20,
    # Drop ZipRecruiter jobs whose location doesn't mention an NCR place. ZR is an
    # aggregator that bleeds in all-India and foreign listings (e.g. a Kuala Lumpur
    # job under a Delhi search); this keeps it to the actual target area. Re-add
    # "india" if you ever want all-India coverage.
    "ziprecruiter_location_allow": ["delhi", "gurgaon", "gurugram", "noida", "ncr",
                                    "ghaziabad", "faridabad"],

    # ── Relevance filter (cuts noise from broad keyword search) ───────────────
    # Applied to BOTH Naukri & LinkedIn, BEFORE fetching the JD or analyzing — so
    # off-target roles never burn a Gemini call or clutter the sheet. Toggle off
    # with relevance_filter=False to keep every result.
    "relevance_filter": True,
    # Keep a job only if its TITLE contains one of these (whole-word, case-insensitive):
    "title_keep": ["analyst", "analytics", "business intelligence", "mis", "reporting"],
    # ...and DROP it if the title contains any of these (too senior / wrong role):
    "title_drop": ["senior", "sr", "lead", "principal", "head", "manager", "director",
                   "vp", "architect", "intern", "internship", "data engineer",
                   "ml engineer", "machine learning engineer", "compensation",
                   "talent acquisition", "recruiter", "recruitment", "forestry"],
    # Only WRITE jobs scoring >= this to the Excel. Weaker matches are still
    # remembered (so they're not re-analyzed daily) — just kept out of the sheet.
    # Set to 0 to show everything.
    "min_score_to_show": 55,

    # ── Dedup ───────────────────────────────────────────────────────────────
    # Tracks every job already shown. Never shows the same job twice.
    "seen_jobs_file": "seen_jobs.json",

    # ── Output ──────────────────────────────────────────────────────────────
    # Each run writes a NEW dated file: naukri_jobs_YYYY-MM-DD.xlsx
    "output_prefix": "naukri_jobs",
    "resume_tailor_url": "https://claude.ai/",
}


# ============================================================================
# DEPS
# ============================================================================
# Third-party imports are recorded rather than fatal. The pure helpers in this
# module (freshness parsing, dedup keys, the relevance gate, scoring) are plain
# stdlib, so keeping the imports non-fatal means the module can be imported --
# and unit-tested -- on a machine that has no Playwright or Gemini SDK. A real
# run still fails fast with install instructions: require_deps() does that check
# at the top of main(), before any scraping starts.
_MISSING_DEPS = []

try:
    from playwright.async_api import async_playwright
except ImportError:
    async_playwright = None
    _MISSING_DEPS.append("playwright")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    Workbook = Font = PatternFill = Alignment = Border = Side = None
    _MISSING_DEPS.append("openpyxl")

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None
    _MISSING_DEPS.append("pypdf")

try:
    import requests
except ImportError:
    requests = None
    _MISSING_DEPS.append("requests")

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None
    _MISSING_DEPS.append("beautifulsoup4")

try:
    from google import genai
except ImportError:
    genai = None
    _MISSING_DEPS.append("google-genai")


def require_deps():
    """Exit with install instructions if a package THIS run needs is missing.

    Which packages are required depends on CONFIG: requests/beautifulsoup4 only
    matter when LinkedIn is enabled, google-genai only in gemini analyzer mode.
    A package that is missing but unused for the configured run is not an error.
    """
    needed = {"playwright", "openpyxl", "pypdf"}
    if CONFIG.get("include_linkedin"):
        needed |= {"requests", "beautifulsoup4"}
    if CONFIG.get("analyzer") == "gemini":
        needed.add("google-genai")

    missing = sorted(p for p in _MISSING_DEPS if p in needed)
    if not missing:
        return

    print(f"\n❌ Missing packages: {', '.join(missing)}")
    print(f"Install:  pip3 install {' '.join(missing)}")
    if "playwright" in missing:
        print("Then:     python3 -m playwright install chromium")
    sys.exit(1)


# ============================================================================
# FRESHNESS + DEDUP HELPERS
# ============================================================================
def parse_days_old(posted_text):
    """
    Parse Naukri's 'Posted X days ago' / 'Few hours ago' / 'Just now' text
    into an integer number of days. Returns None if unparseable.

    The numeric "N day" pattern is matched BEFORE the "yesterday" shortcut on
    purpose. The shortcut used to also test for the literal substring "1 day",
    which is contained in "21 days ago", "11 days ago" and "31 days ago" -- so
    month-old postings were read as one day old and sailed straight through the
    max_days_old freshness filter. Matching the number first removes the overlap.
    """
    if not posted_text:
        return None
    t = posted_text.lower()
    if "just now" in t or "few hours" in t or "few minutes" in t or "today" in t:
        return 0
    m = re.search(r"(\d+)\s*\+?\s*day", t)
    if m:
        return int(m.group(1))
    if "yesterday" in t:
        return 1
    m = re.search(r"(\d+)\s*\+?\s*hour", t)
    if m:
        return 0
    m = re.search(r"(\d+)\s*\+?\s*week", t)
    if m:
        return int(m.group(1)) * 7
    m = re.search(r"(\d+)\s*\+?\s*month", t)
    if m:
        return int(m.group(1)) * 30
    return None  # unknown format — let it through, don't lose jobs to parser bugs


def normalize_job_url(url):
    """Strip query params/fragments so the same job URL always hashes the same."""
    if not url:
        return ""
    url = url.split("?")[0].split("#")[0].rstrip("/")
    return url.lower()


def linkedin_dedup_key(url):
    """Stable dedup key for a LinkedIn job, based on its numeric job ID.

    LinkedIn's guest URLs are NOT stable across runs — the same job comes back
    on different domains (in.linkedin.com vs www.linkedin.com) and with different
    slug text, so keying on the full URL lets the same job slip past dedup. The
    trailing numeric job ID is the only stable identifier, so we canonicalize to
    it. Falls back to plain URL normalization if no ID is found.
    """
    if not url:
        return ""
    base = url.split("?")[0].split("#")[0]
    m = re.search(r"-(\d+)/?$", base) or re.search(r"/view/(\d+)", base)
    if m:
        return f"https://www.linkedin.com/jobs/view/{m.group(1)}"
    return normalize_job_url(url)


# A jk is an opaque per-posting token. In practice Indeed serves 16 lowercase
# hex characters, but nothing guarantees that: data-jk is read straight off the
# page, and treating the token as hex-only was wrong in two ways that both broke
# dedup (see test_dedup_keys.py for the regression tests).
#
#   1. The old URL pattern [0-9a-fA-F]+ was unanchored, so "jk=abz123" matched
#      only the leading "ab" -- distinct postings sharing a hex prefix collapsed
#      onto one key and every one after the first was silently dropped as a
#      duplicate. A jk starting with a non-hex character matched nothing at all,
#      so the fallback stripped the query and EVERY such job collapsed onto the
#      single key "https://in.indeed.com/viewjob".
#   2. The bare-jk and URL branches disagreed for the same token: the scraper
#      keys a fresh job on the bare data-jk while load_seen_jobs() keys the
#      stored apply_link URL, so a non-hex jk produced two different keys for
#      one posting and the job was re-shown on every run.
#
# Matching the whole token fixes both. Case is preserved rather than lowercased:
# if Indeed ever serves a case-sensitive token, folding case would merge two
# genuinely different postings, which loses a job -- the worse failure.
_INDEED_JK_CHARS = r"[0-9A-Za-z_-]+"
_INDEED_BARE_JK_RE = re.compile(rf"{_INDEED_JK_CHARS}\Z")
_INDEED_URL_JK_RE = re.compile(rf"[?&]jk=({_INDEED_JK_CHARS})")


def indeed_dedup_key(url_or_jk):
    """Stable dedup key for an Indeed job, based on its 'jk' job key.

    Indeed links vary (/viewjob?jk=, /rc/clk?jk=, /pagead/...&jk=) and carry
    rotating tracking params, but the jk value is the stable per-posting ID.
    Accepts either a raw jk or any Indeed URL containing jk=, and canonicalizes
    both to one form so the scrape-time key and the key rebuilt from the seen
    file always agree. Falls back to URL normalization if no jk is found.
    """
    if not url_or_jk:
        return ""
    token = url_or_jk.strip()
    # No URL can fullmatch the jk charset (it has no ':' '/' '?' or '.'), so an
    # exact match here means the caller handed us a raw jk.
    if _INDEED_BARE_JK_RE.fullmatch(token):
        return f"https://in.indeed.com/viewjob?jk={token}"
    m = _INDEED_URL_JK_RE.search(token)
    if m:
        return f"https://in.indeed.com/viewjob?jk={m.group(1)}"
    return normalize_job_url(token)


def ziprecruiter_dedup_key(url):
    """Stable, collision-free dedup key for a ZipRecruiter job.

    Real ZR India job URLs are /jobs/<numeric-id>-<slug> — the numeric id is the
    stable per-posting identifier, so we key on it. Falls back to a 'jid' param,
    then the query-stripped path. Rotating tracking params are ignored.
    """
    if not url:
        return ""
    m = re.search(r"/jobs?/(\d+)", url)            # /jobs/544036357-... or /job/123
    if m:
        return f"https://www.ziprecruiter.in/jobs/{m.group(1)}"
    j = re.search(r"[?&]jid=([0-9a-zA-Z_-]+)", url.lower())
    if j:
        return f"https://www.ziprecruiter.in/jobs/{j.group(1)}"
    base = url.split("?")[0].split("#")[0].rstrip("/").lower()
    mm = re.search(r"ziprecruiter\.in(/.*)$", base)
    return f"https://www.ziprecruiter.in{mm.group(1) if mm else base}"


def ziprecruiter_location_ok(loc_text):
    """Keep only ZR jobs whose location mentions an allowed place (kills the
    mis-tagged out-of-city / abroad listings the aggregator mixes in)."""
    allow = CONFIG.get("ziprecruiter_location_allow")
    if not allow:
        return True
    t = (loc_text or "").strip().lower()
    if not t:
        # No location text → don't over-filter, let relevance/score decide.
        # .strip() matters: the text comes out of the DOM, so an absent
        # location usually arrives as padding or a newline rather than as a
        # clean empty string, and an unstripped check dropped those jobs.
        return True
    return any(a in t for a in allow)


# Short terms keep strict whole-word matching: "mis" must not match "miss" or
# "dismissal", and a 2-3 character term is almost always an abbreviation whose
# plural is not a real job title.
_TITLE_PLURAL_MIN_LEN = 4


def _bounded(body, term):
    """Wrap an escaped pattern in word boundaries, but only where one works.

    ``\\b`` needs a word character on its side of the boundary, so a plain
    ``\\bc\\+\\+\\b`` can never match anything at all: the boundary after "+"
    demands a word character that "c++" does not have. Any term beginning or
    ending in punctuation was therefore invisible to the matcher. The boundary
    is applied only on the sides that actually have a word character.
    """
    lead = r"\b" if term[:1].isalnum() or term[:1] == "_" else ""
    trail = r"\b" if term[-1:].isalnum() or term[-1:] == "_" else ""
    return f"{lead}{body}{trail}"


@functools.lru_cache(maxsize=512)
def _title_term_re(term):
    """Compile the whole-word matcher for one single-word title keyword.

    A plain ``\\b<term>\\b`` got two cases wrong, and both of them failed
    silently in the direction that loses jobs:

    * Plural titles. A posting hiring for several seats titles the role in the
      plural ("Business Analysts — 3 positions"), which never matched the
      singular keep-term, so the whole posting was gated out. The same gap let
      a plural drop-term ("Analytics Managers") past the drop-list. Terms of
      at least _TITLE_PLURAL_MIN_LEN characters now tolerate a trailing "s".
    * Terms that begin or end in punctuation — see _bounded().
    """
    body = re.escape(term)
    if len(term) >= _TITLE_PLURAL_MIN_LEN and term[-1].isalnum():
        body += "s?"
    return re.compile(_bounded(body, term))


@functools.lru_cache(maxsize=512)
def _skill_term_re(term):
    """Compile the whole-word matcher for one skill or tool name.

    Same boundary handling as _title_term_re, but without the plural
    tolerance: a trailing "s" on a title is another seat for the same role,
    while on a skill it is usually a different word ("excel" vs "excels at").
    """
    return re.compile(_bounded(re.escape(term), term))


def title_has_term(title, term):
    """Does ``title`` contain ``term``, matching single words whole?

    Shared by the relevance gate and the scorer so the two cannot drift. A
    multi-word term is a substring check, since its own spaces already keep it
    from matching inside another word.
    """
    term = term.lower()
    if " " in term:
        return term in title
    return _title_term_re(term).search(title) is not None


def passes_relevance(title):
    """Coarse title gate, run BEFORE any JD fetch or analysis.

    Keeps only titles that look like the target roles (analyst / analytics / BI /
    MIS / reporting) and drops obviously-wrong ones (senior, lead, principal,
    manager, data/ML engineer, compensation, recruiting, etc.). This is what turns
    a 30-job noise dump into a handful of real matches — and it saves Gemini quota,
    since junk never reaches the analyzer. Toggle via CONFIG['relevance_filter'].
    """
    if not CONFIG.get("relevance_filter"):
        return True
    t = (title or "").lower()

    def has(term):
        return title_has_term(t, term)

    if not any(has(k) for k in CONFIG.get("title_keep", [])):
        return False
    if any(has(d) for d in CONFIG.get("title_drop", [])):
        return False
    return True


def load_seen_jobs(path):
    """Load the set of already-shown job URLs from disk."""
    if not os.path.exists(path):
        return set()
    try:
        with open(path) as f:
            data = json.load(f)
        seen = set(data.get("seen_urls", []))
    except Exception:
        return set()
    # One-time migration: earlier runs stored LinkedIn jobs by full URL, whose
    # domain/slug varied run-to-run and slipped past dedup. Add a stable job-id
    # key for every LinkedIn URL already on record, so jobs you've already seen
    # won't reappear once. Idempotent — safe to run every time.
    for u in list(seen):
        if "linkedin.com" in u:
            seen.add(linkedin_dedup_key(u))
        elif "indeed." in u and "jk=" in u:
            seen.add(indeed_dedup_key(u))
        elif "ziprecruiter." in u:
            seen.add(ziprecruiter_dedup_key(u))
    return seen


def save_seen_jobs(path, seen_set):
    """Persist the seen-jobs set atomically (temp file + rename), so a crash or
    Ctrl+C mid-write can't leave a half-written, corrupt JSON behind."""
    payload = {
        "seen_urls": sorted(seen_set),
        "last_updated": datetime.now().isoformat(),
        "total_seen": len(seen_set),
    }
    tmp = f"{path}.tmp"
    with open(tmp, "w") as f:
        json.dump(payload, f, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, path)   # atomic on the same filesystem


# ============================================================================
# RESUME → PROFILE (runs ONCE, cached)
# ============================================================================
def read_resume_text(path):
    """Extract raw text from the resume PDF."""
    if not os.path.exists(path):
        print(f"\n❌ Resume not found at: {path}")
        print("   Put the PDF next to this script, or fix 'resume_path' in CONFIG.")
        sys.exit(1)
    reader = PdfReader(path)
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def file_fingerprint(path):
    """Hash the resume so we know if it changed since last extraction."""
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def build_profile(resume_text, gemini_client, model):
    """Extract a structured profile from the resume. ONE Gemini call (or regex fallback)."""
    if gemini_client:
        prompt = f"""Extract a structured candidate profile from this resume. Return ONLY valid JSON, no markdown.

RESUME:
{resume_text[:6000]}

Return this exact shape:
{{
  "name": "<full name>",
  "years_experience": <number, e.g. 2 or 1.5>,
  "seniority": "<junior/mid/senior>",
  "skills": ["lowercase skill", ...],
  "domains": ["industry/domain the candidate worked in", ...],
  "strengths": ["2-4 standout strengths from the candidate's achievements"],
  "target_titles": ["job titles that fit the candidate, e.g. data analyst, bi analyst"]
}}

Extract ALL technical skills, tools, and methods mentioned (SQL, Python, specific libraries, BI tools, statistical methods, etc.). Be thorough on skills."""
        try:
            resp = gemini_client.models.generate_content(model=model, contents=prompt)
            text = re.sub(r"^```(?:json)?|```$", "", resp.text.strip(), flags=re.MULTILINE).strip()
            profile = json.loads(text)
            profile["skills"] = [s.lower() for s in profile.get("skills", [])]
            return profile
        except Exception as e:
            print(f"  ⚠ Gemini profile extraction failed ({e}); using regex fallback.")

    # Regex fallback — pull known skills from resume text
    known = ["sql", "mysql", "postgresql", "python", "pandas", "numpy", "matplotlib",
             "r", "dplyr", "ggplot2", "excel", "vlookup", "xlookup", "pivot tables",
             "index-match", "tableau", "power bi", "powerbi", "google analytics",
             "regression", "predictive modelling", "predictive modeling", "hypothesis testing",
             "trend analysis", "data cleaning", "etl", "data pipelines", "data warehousing",
             "data validation", "dashboard", "kpi tracking", "data visualization",
             "data storytelling", "customer segmentation", "campaign", "funnel",
             "a/b testing", "statistical modelling"]
    rt = resume_text.lower()
    found = sorted({s for s in known if s in rt})
    return {
        "name": resume_text.split("\n")[0].strip()[:40],
        "years_experience": CONFIG["experience_years"],
        "seniority": "mid",
        "skills": found,
        "domains": [],
        "strengths": [],
        "target_titles": ["data analyst", "bi analyst", "business analyst"],
    }


def load_or_build_profile(gemini_client, model):
    """Load cached profile if resume unchanged, else build fresh."""
    resume_path = CONFIG["resume_path"]
    cache_path = CONFIG["profile_cache"]
    fp = file_fingerprint(resume_path)

    if os.path.exists(cache_path):
        try:
            with open(cache_path) as f:
                cached = json.load(f)
            if cached.get("_resume_fingerprint") == fp:
                print(f"  ✓ Using cached profile ({cached['name']}, "
                      f"{cached['years_experience']}y, {len(cached['skills'])} skills)")
                return cached
        except Exception:
            pass

    print("  → Extracting profile from resume (one time)...")
    resume_text = read_resume_text(resume_path)
    profile = build_profile(resume_text, gemini_client, model)
    profile["_resume_fingerprint"] = fp
    profile["_resume_text"] = resume_text[:6000]  # kept for gemini per-job matching
    with open(cache_path, "w") as f:
        json.dump(profile, f, indent=2)
    print(f"  ✓ Profile built: {profile['name']}, {profile['years_experience']}y, "
          f"{len(profile['skills'])} skills. Cached to {cache_path}.")
    return profile


# ============================================================================
# SCRAPER
# ============================================================================
async def scrape_naukri_listings(location, keyword, max_jobs, exp_years,
                                  seen_urls, max_days_old):
    keyword_slug = keyword.replace(" ", "-")
    url = f"https://www.naukri.com/{keyword_slug}-jobs-in-{location}?experience={exp_years}"
    print(f"  → {url}")
    jobs = []
    skipped_old = 0
    skipped_seen = 0
    skipped_irrel = 0
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = await browser.new_context(
            user_agent=("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/120.0.0.0 Safari/537.36"),
            viewport={"width": 1366, "height": 768},
        )
        page = await context.new_page()
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=30000)
            await page.wait_for_timeout(3000)
            cards = await page.query_selector_all("div.srp-jobtuple-wrapper")  # UPDATE WHEN CHANGES
            print(f"  → {len(cards)} cards found")
            for i, card in enumerate(cards[:max_jobs]):
                try:
                    title_el = await card.query_selector("a.title")
                    company_el = await card.query_selector("a.comp-name")
                    loc_el = await card.query_selector("span.locWdth")
                    exp_el = await card.query_selector("span.expwdth")
                    sal_el = await card.query_selector("span.sal-wrap span")
                    desc_el = await card.query_selector("span.job-desc")
                    posted_el = await card.query_selector("span.job-post-day")

                    title = (await title_el.inner_text()).strip() if title_el else ""
                    company = (await company_el.inner_text()).strip() if company_el else ""
                    loc = (await loc_el.inner_text()).strip() if loc_el else ""
                    exp = (await exp_el.inner_text()).strip() if exp_el else ""
                    sal = (await sal_el.inner_text()).strip() if sal_el else "Not disclosed"
                    desc = (await desc_el.inner_text()).strip() if desc_el else ""
                    posted = (await posted_el.inner_text()).strip() if posted_el else ""
                    link = await title_el.get_attribute("href") if title_el else ""
                    if not title or not link:
                        continue

                    # ── Freshness filter ────────────────────────────────────
                    days_old = parse_days_old(posted)
                    if days_old is not None and days_old > max_days_old:
                        skipped_old += 1
                        continue

                    # ── Dedup filter ────────────────────────────────────────
                    norm_url = normalize_job_url(link)
                    if norm_url in seen_urls:
                        skipped_seen += 1
                        continue

                    # ── Relevance filter (drop off-target titles cheaply) ────
                    if not passes_relevance(title):
                        skipped_irrel += 1
                        continue

                    # Passed all filters — now fetch full JD (expensive step)
                    full_jd = await fetch_full_jd(context, link)
                    jobs.append({"title": title, "company": company, "location": loc,
                                 "experience": exp, "salary": sal, "short_desc": desc,
                                 "full_jd": full_jd or desc, "posted": posted,
                                 "apply_link": link, "_norm_url": norm_url,
                                 "source": "Naukri"})
                    print(f"    [+] {title} @ {company} · {posted}")
                except Exception as e:
                    print(f"    ⚠ skipped: {e}")
        except Exception as e:
            print(f"  ❌ {e}")
        finally:
            await browser.close()
    print(f"  → kept {len(jobs)} · skipped {skipped_old} (too old) · {skipped_seen} (already seen) · {skipped_irrel} (off-target)")
    return jobs


async def fetch_full_jd(context, url):
    try:
        page = await context.new_page()
        await page.goto(url, wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(2000)
        jd_el = await page.query_selector("section.styles_job-desc-container__txpYf")
        if not jd_el:
            jd_el = await page.query_selector("div.dang-inner-html")
        text = (await jd_el.inner_text()).strip() if jd_el else ""
        await page.close()
        return text
    except Exception:
        return ""


# ============================================================================
# LINKEDIN SCRAPER  (public guest endpoint — no login, no account risk)
# ============================================================================
# Mirrors scrape_naukri_listings: same freshness filter, same dedup against
# seen_urls, same job-dict shape so jobs flow through the identical matcher
# and Excel writer. Uses LinkedIn's logged-out "jobs-guest" API, so it never
# touches the candidate's account.
LINKEDIN_GUEST_SEARCH = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
LINKEDIN_GUEST_JOB = "https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{job_id}"
LINKEDIN_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/120.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Language": "en-US,en;q=0.9",
}


def _li_days_old(datetime_attr):
    """LinkedIn <time datetime='YYYY-MM-DD'> → integer days old (or None)."""
    if not datetime_attr:
        return None
    try:
        d = datetime.strptime(datetime_attr[:10], "%Y-%m-%d").date()
        return (datetime.now().date() - d).days
    except Exception:
        return None


def _li_days_to_text(days):
    """Render days-old as Naukri-style text so the Excel 'Posted' column matches."""
    if days is None:
        return ""
    if days <= 0:
        return "Today"
    if days == 1:
        return "1 day ago"
    return f"{days} days ago"


def fetch_linkedin_jd(job_id):
    """Fetch the full JD text from LinkedIn's guest job-posting endpoint."""
    try:
        r = requests.get(LINKEDIN_GUEST_JOB.format(job_id=job_id),
                         headers=LINKEDIN_HEADERS, timeout=20)
        if r.status_code != 200:
            return ""
        soup = BeautifulSoup(r.text, "html.parser")
        desc = soup.find("div", class_="show-more-less-html__markup")
        return desc.get_text(" ", strip=True) if desc else ""
    except requests.RequestException:
        return ""


def scrape_linkedin_listings(keyword, location, max_jobs, seen_urls,
                             max_days_old, exp_filter=""):
    """Scrape LinkedIn public guest listings. Same filters + dict shape as Naukri.

    f_TPR=r{seconds} restricts to the freshness window at the source itself,
    so we pull mostly-fresh cards and then re-check each one locally too.
    """
    print(f"  → LinkedIn guest search: '{keyword}' in {location}")
    jobs, start = [], 0
    skipped_old = skipped_seen = skipped_irrel = 0
    pages, max_pages = 0, 12   # bound pagination so a sparse pool can't loop forever

    while len(jobs) < max_jobs and pages < max_pages:
        pages += 1
        params = {"keywords": keyword, "location": location,
                  "f_TPR": f"r{max_days_old * 86400}", "start": start}
        if exp_filter:
            params["f_E"] = exp_filter
        try:
            r = requests.get(LINKEDIN_GUEST_SEARCH, params=params,
                             headers=LINKEDIN_HEADERS, timeout=20)
        except requests.RequestException as e:
            print(f"  ⚠ LinkedIn request failed: {e}")
            break
        if r.status_code != 200 or not r.text.strip():
            if r.status_code == 429:
                print("  ⚠ LinkedIn rate-limited (429). Stopping LinkedIn for this run.")
            break

        cards = BeautifulSoup(r.text, "html.parser").find_all("li")
        if not cards:
            break

        for card in cards:
            if len(jobs) >= max_jobs:
                break
            link_el = card.find("a", class_="base-card__full-link")
            title_el = card.find("h3", class_="base-search-card__title")
            comp_el = card.find("h4", class_="base-search-card__subtitle")
            loc_el = card.find("span", class_="job-search-card__location")
            time_el = card.find("time")
            if not (link_el and title_el):
                continue

            link = link_el.get("href", "").split("?")[0]
            if not link:
                continue
            # Extract the stable numeric job ID and use it as the dedup key,
            # so domain/slug variation across runs can't sneak the job past dedup.
            m = re.search(r"-(\d+)/?$", link) or re.search(r"/view/(\d+)", link)
            job_id = m.group(1) if m else None
            norm_url = linkedin_dedup_key(link)

            # ── Freshness filter (same rule as Naukri) ──────────────────────
            days_old = _li_days_old(time_el.get("datetime") if time_el else None)
            if days_old is not None and days_old > max_days_old:
                skipped_old += 1
                continue

            # ── Dedup filter (shared seen_urls memory, keyed on job ID) ─────
            if norm_url in seen_urls:
                skipped_seen += 1
                continue

            title = title_el.get_text(strip=True)
            company = comp_el.get_text(strip=True) if comp_el else ""
            loc = loc_el.get_text(strip=True) if loc_el else location

            # ── Relevance filter (drop off-target titles before the JD fetch) ─
            if not passes_relevance(title):
                skipped_irrel += 1
                continue

            full_jd = fetch_linkedin_jd(job_id) if job_id else ""
            time.sleep(1.5)  # polite pause between detail fetches

            jobs.append({
                "title": title, "company": company, "location": loc,
                "experience": "", "salary": "Not disclosed",
                "short_desc": "", "full_jd": full_jd or title,
                "posted": _li_days_to_text(days_old), "apply_link": link,
                "_norm_url": norm_url, "source": "LinkedIn",
            })
            print(f"    [+] {title} @ {company} · {_li_days_to_text(days_old)}")

        start += 25
        time.sleep(2)

    print(f"  → kept {len(jobs)} · skipped {skipped_old} (too old) · {skipped_seen} (already seen) · {skipped_irrel} (off-target)")
    return jobs


# ============================================================================
# INDEED SCRAPER  (in.indeed.com via Playwright — real browser)
# ============================================================================
# Mirrors the Naukri scraper: same freshness filter (Indeed's own fromage param
# + posted-text parse), same dedup (keyed on the stable 'jk' job id), same
# relevance gate before the JD fetch, same job-dict shape so jobs flow through
# the identical matcher + Excel writer. Indeed is anti-bot heavy, so selectors
# may need occasional updates and a run can come back empty if challenged.
INDEED_HEADERS_UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                     "AppleWebKit/537.36 (KHTML, like Gecko) "
                     "Chrome/120.0.0.0 Safari/537.36")


async def fetch_indeed_jd(context, jk):
    """Open a job's viewjob page and pull the full description. Best-effort."""
    if not jk:
        return ""
    page = await context.new_page()
    try:
        await page.goto(f"https://in.indeed.com/viewjob?jk={jk}",
                        wait_until="domcontentloaded", timeout=25000)
        await page.wait_for_timeout(1500)
        el = await page.query_selector("#jobDescriptionText")
        return (await el.inner_text()).strip() if el else ""
    except Exception:
        return ""
    finally:
        await page.close()


async def scrape_indeed_listings(location, keyword, max_jobs, seen_urls, max_days_old):
    """Scrape in.indeed.com listings. Same filters + dict shape as the others."""
    q = keyword.replace(" ", "+")
    base = f"https://in.indeed.com/jobs?q={q}&l={location}&fromage={max_days_old}"
    print(f"  → {base}")
    jobs = []
    skipped_old = skipped_seen = skipped_irrel = skipped_noparse = 0

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = await browser.new_context(
            user_agent=INDEED_HEADERS_UA,
            viewport={"width": 1366, "height": 768},
            locale="en-IN",
        )
        page = await context.new_page()
        try:
            start = 0
            while len(jobs) < max_jobs:
                await page.goto(f"{base}&start={start}",
                                wait_until="domcontentloaded", timeout=30000)
                await page.wait_for_timeout(3000)
                cards = await page.query_selector_all("div.job_seen_beacon")  # UPDATE WHEN CHANGES
                if not cards:
                    # Likely a bot challenge or markup change — bail out cleanly.
                    if start == 0:
                        print("  ⚠ Indeed returned 0 cards (bot challenge or selector "
                              "change). Skipping Indeed this run.")
                    break
                print(f"  → {len(cards)} cards on page (start={start})")

                for card in cards:
                    if len(jobs) >= max_jobs:
                        break
                    try:
                        # The job link is an anchor whose href carries the jk
                        # (…/rc/clk?jk=… or …/viewjob?jk=…); its text is the title.
                        link_el = (await card.query_selector("h2.jobTitle a")
                                   or await card.query_selector("a.jcs-JobTitle")
                                   or await card.query_selector("a[href*='/rc/clk?jk=']")
                                   or await card.query_selector("a[href*='/viewjob?jk=']")
                                   or await card.query_selector("a[data-jk]"))
                        title = (await link_el.inner_text()).strip() if link_el else ""
                        if not title and link_el:
                            title = (await link_el.get_attribute("title")
                                     or await link_el.get_attribute("aria-label") or "").strip()
                        href = await link_el.get_attribute("href") if link_el else ""

                        jk = await card.get_attribute("data-jk")
                        if not jk and link_el:
                            jk = await link_el.get_attribute("data-jk")
                        if not jk and href:
                            mm = re.search(r"[?&]jk=([0-9a-fA-F]+)", href)
                            jk = mm.group(1) if mm else None
                        if not title or not jk:
                            skipped_noparse += 1
                            continue

                        comp_el = await card.query_selector(
                            '[data-testid="company-name"], span.companyName, [data-company-name="true"]')
                        loc_el = await card.query_selector(
                            '[data-testid="text-location"], div.companyLocation')
                        snip_el = await card.query_selector(
                            "div.job-snippet, [data-testid='jobsnippet_footer'], ul[style*='list-style-type']")
                        date_el = await card.query_selector(
                            "span.date, [data-testid='myJobsStateDate']")

                        company = (await comp_el.inner_text()).strip() if comp_el else ""
                        loc = (await loc_el.inner_text()).strip() if loc_el else location
                        snippet = (await snip_el.inner_text()).strip() if snip_el else ""
                        posted = (await date_el.inner_text()).strip() if date_el else ""

                        norm_url = indeed_dedup_key(jk)
                        apply_link = f"https://in.indeed.com/viewjob?jk={jk}"

                        # ── Freshness filter (fromage=7 already bounds at source;
                        #     missing per-card date → kept, not wrongly dropped) ──
                        days_old = parse_days_old(posted)
                        if days_old is not None and days_old > max_days_old:
                            skipped_old += 1
                            continue
                        # ── Dedup filter (shared memory, keyed on jk) ───────
                        if norm_url in seen_urls:
                            skipped_seen += 1
                            continue
                        # ── Relevance filter (before the JD fetch) ──────────
                        if not passes_relevance(title):
                            skipped_irrel += 1
                            continue

                        full_jd = await fetch_indeed_jd(context, jk)
                        await page.wait_for_timeout(800)
                        jobs.append({
                            "title": title, "company": company, "location": loc,
                            "experience": "", "salary": "Not disclosed",
                            "short_desc": snippet, "full_jd": full_jd or snippet or title,
                            "posted": posted, "apply_link": apply_link,
                            "_norm_url": norm_url, "source": "Indeed",
                        })
                        print(f"    [+] {title} @ {company} · {posted or '?'}")
                    except Exception as e:
                        print(f"    ⚠ skipped: {e}")
                start += 10
        except Exception as e:
            print(f"  ❌ {e}")
        finally:
            await browser.close()

    print(f"  → kept {len(jobs)} · skipped {skipped_old} (too old) · {skipped_seen} (already seen) · "
          f"{skipped_irrel} (off-target) · {skipped_noparse} (couldn't parse)")
    return jobs


# ============================================================================
# ZIPRECRUITER INDIA SCRAPER  (ziprecruiter.in via Playwright — real browser)
# ============================================================================
# Plain HTTP gets a 403 here (bot detection), so we render in a real browser
# like the other heavy sources. Same freshness/dedup/relevance treatment, PLUS
# a location check to drop the aggregator's mis-tagged out-of-city jobs.
# Selectors are best-effort (marked UPDATE WHEN CHANGES) with broad fallbacks.
ZIPRECRUITER_UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/120.0.0.0 Safari/537.36")


def _zr_parse_date(text):
    """ZR shows dates like '14 Jun' or '2 days ago' → integer days old (or None)."""
    if not text:
        return None
    t = text.strip().lower()
    m = re.search(r"(\d+)\s*day", t)
    if m:
        return int(m.group(1))
    if "today" in t or "hour" in t or "just" in t or "min" in t:
        return 0
    if "yesterday" in t:
        return 1
    for fmt in ("%d %b %Y", "%d %b"):
        try:
            d = datetime.strptime(text.strip(), fmt)
            if fmt == "%d %b":
                d = d.replace(year=datetime.now().year)
            return (datetime.now().date() - d.date()).days
        except ValueError:
            continue
    return None


async def scrape_ziprecruiter_listings(keyword, location, max_jobs, seen_urls, max_days_old):
    """Scrape ziprecruiter.in listings via Playwright. Same dict shape as the rest.

    Real DOM (confirmed): li.job-listing › a.jobList-title (href /jobs/<id>-slug,
    title in <strong>) › ul.jobList-introMeta li[0]=company, li[1]=location ›
    div.jobList-description (snippet) › div.jobList-date (e.g. '14 Jun').
    """
    base = "https://www.ziprecruiter.in/jobs/search"
    print(f"  → ZipRecruiter: '{keyword}' in {location}")
    jobs = []
    skipped_old = skipped_seen = skipped_irrel = skipped_geo = skipped_noparse = 0
    picked = set()   # intra-run dedup, in case pagination loops back to page 1

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False, args=["--disable-blink-features=AutomationControlled"])
        context = await browser.new_context(
            user_agent=ZIPRECRUITER_UA, viewport={"width": 1366, "height": 768},
            locale="en-IN")
        page = await context.new_page()
        try:
            pg = 1
            while len(jobs) < max_jobs and pg <= 10:
                url = (f"{base}?q={keyword.replace(' ', '+')}"
                       f"&l={location}&sort=published_at&page={pg}")
                await page.goto(url, wait_until="domcontentloaded", timeout=30000)
                await page.wait_for_timeout(2500)

                cards = await page.query_selector_all("li.job-listing")
                if not cards:
                    if pg == 1:
                        print("  ⚠ ZipRecruiter: no job cards (blocked or markup change). Skipping.")
                    break
                print(f"  → {len(cards)} cards on page {pg}")

                new_keys = 0
                for card in cards:
                    if len(jobs) >= max_jobs:
                        break
                    try:
                        link_el = await card.query_selector("a.jobList-title, a.job-link")
                        if not link_el:
                            skipped_noparse += 1
                            continue
                        title = (await link_el.inner_text()).strip()
                        link = (await link_el.get_attribute("href") or "").strip()
                        if not title or not link:
                            skipped_noparse += 1
                            continue
                        if link.startswith("/"):
                            link = "https://www.ziprecruiter.in" + link
                        norm_url = ziprecruiter_dedup_key(link)
                        if norm_url in picked:
                            continue          # same job again (pagination looped)
                        picked.add(norm_url)
                        new_keys += 1

                        meta = await card.query_selector_all("ul.jobList-introMeta li")
                        company = (await meta[0].inner_text()).strip() if len(meta) >= 1 else ""
                        loc = (await meta[1].inner_text()).strip() if len(meta) >= 2 else location
                        snip_el = await card.query_selector("div.jobList-description")
                        date_el = await card.query_selector("div.jobList-date")
                        snippet = (await snip_el.inner_text()).strip() if snip_el else ""
                        date_txt = (await date_el.inner_text()).strip() if date_el else ""

                        # ── Freshness filter ────────────────────────────────
                        days_old = _zr_parse_date(date_txt)
                        if days_old is not None and days_old > max_days_old:
                            skipped_old += 1
                            continue
                        # ── Dedup filter (shared memory) ────────────────────
                        if norm_url in seen_urls:
                            skipped_seen += 1
                            continue
                        # ── Relevance filter ────────────────────────────────
                        if not passes_relevance(title):
                            skipped_irrel += 1
                            continue
                        # ── Location filter (drop mis-tagged jobs) ──────────
                        if not ziprecruiter_location_ok(loc):
                            skipped_geo += 1
                            continue

                        jobs.append({
                            "title": title, "company": company, "location": loc,
                            "experience": "", "salary": "Not disclosed",
                            "short_desc": snippet,
                            "full_jd": (title + ". " + snippet) if snippet else title,
                            "posted": date_txt or "", "apply_link": link,
                            "_norm_url": norm_url, "source": "ZipRecruiter",
                        })
                        print(f"    [+] {title} @ {company} · {date_txt or '?'}")
                    except Exception as e:
                        print(f"    ⚠ skipped: {e}")

                if new_keys == 0:
                    break          # page returned nothing new → stop paginating
                pg += 1
                await page.wait_for_timeout(1500)
        except Exception as e:
            print(f"  ❌ {e}")
        finally:
            await browser.close()

    print(f"  → kept {len(jobs)} · skipped {skipped_old} (too old) · {skipped_seen} (already seen) · "
          f"{skipped_irrel} (off-target) · {skipped_geo} (wrong location) · {skipped_noparse} (couldn't parse)")
    return jobs


# ============================================================================
# MATCHING (against the cached profile)
# ============================================================================
# Title terms the scorer reacts to. These are matched with title_has_term(),
# i.e. whole words for single-word terms. A bare substring check read
# "Internal Reporting Analyst" as an internship (-30 and a false red flag),
# "Analyst, Headquarters Finance" as a head-of role (-18), and any title
# containing "mis" ("Misconduct") as an MIS role (+12).
TARGET_TITLE_TELLS = ["data analyst", "business analyst", "bi analyst",
                      "analytics", "reporting analyst", "mis"]
SENIOR_TITLE_TELLS = ["senior", "sr.", "sr ", "lead", "principal", "head",
                      "manager"]
# "intern" alone would miss "Internship", which _title_term_re's optional
# trailing "s" does not cover.
INTERN_TITLE_TELLS = ["intern", "internship"]

DATA_ENGINEER_TELLS = ["spark", "hadoop", "kafka", "scala", "airflow", "data pipeline engineering"]
ML_ENGINEER_TELLS = ["tensorflow", "pytorch", "deep learning", "mlops", "model deployment"]
DEV_TELLS = ["react", "node.js", "frontend", "backend", "microservices", "java spring"]


def match_keyword(job, profile):
    """Offline match: JD vs resume-extracted skills."""
    jd_text = (job["title"] + " " + job["full_jd"]).lower()
    skills = profile["skills"]
    yoe = profile["years_experience"]

    matched = [s for s in skills if _skill_term_re(s).search(jd_text)]
    anti = [s for s in (DATA_ENGINEER_TELLS + ML_ENGINEER_TELLS + DEV_TELLS)
            if _skill_term_re(s).search(jd_text)]

    # Score = how well the candidate covers the skills THIS JD actually asks for.
    # We look for a broad set of data-analyst skills in the JD, then check
    # how many of those the candidate has. (Scoring against the full skill list
    # would always look low, since any one JD mentions only a few.)
    da_universe = ["sql", "mysql", "postgresql", "python", "pandas", "r",
                   "excel", "advanced excel", "vlookup", "xlookup", "pivot tables",
                   "tableau", "power bi", "powerbi", "looker", "google analytics",
                   "snowflake", "bigquery", "dbt", "etl", "data cleaning",
                   "data visualization", "dashboard", "kpi", "statistics",
                   "regression", "hypothesis testing", "a/b testing",
                   "predictive modelling", "predictive modeling",
                   "customer segmentation", "data storytelling"]
    jd_wants = [s for s in da_universe if _skill_term_re(s).search(jd_text)]
    cand_has = [s for s in jd_wants if s in skills]

    if jd_wants:
        coverage = len(cand_has) / len(jd_wants)   # % of JD's asks the candidate meets
        # Map coverage to a 0-85 band so titles/experience can nudge above/below.
        # Reward absolute count too, so a JD asking 5 skills the candidate has all of
        # outranks one asking 2 the candidate has all of.
        score = int(coverage * 70) + min(len(cand_has) * 4, 20)
    else:
        score = 40 + min(len(matched) * 4, 25)

    score -= len(anti) * 10

    title = job["title"].lower()
    is_intern = any(title_has_term(title, t) for t in INTERN_TITLE_TELLS)
    if any(title_has_term(title, t) for t in TARGET_TITLE_TELLS):
        score += 12
    if any(title_has_term(title, t) for t in SENIOR_TITLE_TELLS):
        score -= 18
    if is_intern:
        score -= 30

    exp_text = job.get("experience", "").lower()
    m = re.search(r"(\d+)\s*-\s*(\d+)", exp_text)
    if m:
        lo, hi = int(m.group(1)), int(m.group(2))
        if lo > yoe + 2:
            score -= 20
        elif lo <= yoe <= hi + 1:
            score += 8
    # An absent salary is not a disclosed one. LinkedIn and ZipRecruiter
    # postings often carry no salary field at all, and scoring "" as disclosed
    # handed them a free +5 over Naukri postings that honestly say
    # "Not disclosed".
    salary_text = (job.get("salary") or "").strip().lower()
    if salary_text and "not disclosed" not in salary_text:
        score += 5

    score = max(0, min(100, score))

    common = ["sql", "python", "tableau", "power bi", "r", "looker", "snowflake",
              "bigquery", "dbt", "airflow", "spark", "a/b testing", "statistics",
              "machine learning", "google analytics", "advanced excel"]
    jd_skills = [s for s in common if _skill_term_re(s).search(jd_text)]
    gaps = [s for s in jd_skills if s not in skills][:4]

    if score >= 80:
        fit = f"Strong match — JD hits {len(matched)} of the candidate's skills."
    elif score >= 60:
        fit = f"Decent fit — {len(matched)} skill matches, tailor resume."
    elif score >= 40:
        fit = f"Stretch — {len(matched)} matches only."
    else:
        fit = "Weak fit — different stack."

    flags = []
    if anti:
        flags.append(f"Mentions {', '.join(anti[:2])} (different role)")
    if is_intern:
        flags.append("Internship — below target level")
    if m and int(m.group(1)) > yoe + 2:
        flags.append(f"Wants {m.group(0)}y — too senior")
    red = " · ".join(flags) if flags else "None"

    summary = re.sub(r"\s+", " ", job["full_jd"])[:250]
    if len(job["full_jd"]) > 250:
        summary += "..."

    return {"match_score": score, "fit_summary": fit,
            "key_skills_required": matched[:6] or ["See JD"],
            "skills_gap": gaps, "jd_summary": summary, "red_flags": red,
            "analyzed_by": "keyword"}


# Sentinel raised when Gemini's daily quota is exhausted, so the caller
# can permanently switch to keyword mode for the rest of the run.
class QuotaExhausted(Exception):
    pass


def match_gemini(client, model, job, profile):
    """Smart match: full resume context vs this JD."""
    prompt = f"""You are matching a candidate to a job. Return ONLY valid JSON, no markdown.

CANDIDATE PROFILE:
Name: {profile['name']}
Experience: {profile['years_experience']} years ({profile['seniority']})
Skills: {', '.join(profile['skills'])}
Domains: {', '.join(profile.get('domains', []))}
Strengths: {', '.join(profile.get('strengths', []))}

JOB:
Title: {job['title']}
Company: {job['company']}
Experience required: {job['experience']}
Salary: {job['salary']}
JD: {job['full_jd'][:3000]}

Return this exact JSON:
{{
  "match_score": <int 0-100>,
  "fit_summary": "<one sentence, max 20 words, specific to THE CANDIDATE'S background>",
  "key_skills_required": ["skill", ...],
  "skills_gap": ["skill the candidate lacks", ...],
  "jd_summary": "<2-3 sentences>",
  "red_flags": "<concerns or 'None'>"
}}

Scoring: 80+=must apply, 60-79=apply w/ tailored resume, 40-59=stretch, <40=skip.
Consider the candidate's ~{profile['years_experience']}y experience — don't over-penalize 2-4y roles, do penalize 5y+ senior roles."""
    try:
        resp = client.models.generate_content(model=model, contents=prompt)
        text = re.sub(r"^```(?:json)?|```$", "", resp.text.strip(), flags=re.MULTILINE).strip()
        result = json.loads(text)
        result["analyzed_by"] = "gemini"
        return result
    except Exception as e:
        msg = str(e)
        # Quota / rate-limit errors → signal caller to switch to keyword for the rest
        if "429" in msg or "RESOURCE_EXHAUSTED" in msg or "quota" in msg.lower():
            raise QuotaExhausted(msg)
        # Any other error (bad JSON, network blip) → just keyword this one job
        print(f"    ⚠ Gemini match failed on this job, using keyword: {e}")
        return match_keyword(job, profile)


# ============================================================================
# EXCEL OUTPUT
# ============================================================================
def write_excel(jobs, output_path, tailor_url, profile):
    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"
    headers = ["Match", "Title", "Company", "Location", "Salary", "Posted",
               "Fit Summary", "JD Summary", "Key Skills", "Skills Gap",
               "Red Flags", "Source", "Analyzed By", "Apply", "Tailor Resume"]
    ws.append(headers)

    hf = PatternFill("solid", start_color="1C1917")
    hfont = Font(name="Arial", bold=True, color="FAFAF9", size=11)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="78716C")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = hf; cell.font = hfont; cell.alignment = center; cell.border = border

    for job in sorted(jobs, key=lambda x: -x["analysis"]["match_score"]):
        a = job["analysis"]; s = a["match_score"]
        engine = a.get("analyzed_by", "keyword")
        ws.append([s, job["title"], job["company"], job["location"], job["salary"],
                   job["posted"], a["fit_summary"], a["jd_summary"],
                   ", ".join(a["key_skills_required"][:6]),
                   ", ".join(a["skills_gap"][:4]), a["red_flags"],
                   job.get("source", "Naukri"), engine, "Apply →", "Tailor →"])
        r = ws.max_row
        # Source cell (col 12) — color-coded per job board
        src = job.get("source", "Naukri")
        sc = ws.cell(row=r, column=12)
        sc.alignment = Alignment(horizontal="center", vertical="center")
        if src == "LinkedIn":
            sc.fill = PatternFill("solid", start_color="E0F2FE")
            sc.font = Font(name="Arial", size=9, bold=True, color="0A66C2")
        elif src == "Indeed":
            sc.fill = PatternFill("solid", start_color="E3F0FD")
            sc.font = Font(name="Arial", size=9, bold=True, color="2557A7")
        elif src == "ZipRecruiter":
            sc.fill = PatternFill("solid", start_color="E7F6EC")
            sc.font = Font(name="Arial", size=9, bold=True, color="1A7F37")
        else:
            sc.fill = PatternFill("solid", start_color="EDE9FE")
            sc.font = Font(name="Arial", size=9, bold=True, color="5B21B6")
        # Engine cell (col 13) — color-coded
        ec = ws.cell(row=r, column=13)
        ec.alignment = Alignment(horizontal="center", vertical="center")
        if engine == "gemini":
            ec.fill = PatternFill("solid", start_color="DBEAFE")
            ec.font = Font(name="Arial", size=9, bold=True, color="1E40AF")
        else:
            ec.fill = PatternFill("solid", start_color="F5F5F4")
            ec.font = Font(name="Arial", size=9, color="78716C")
        ac = ws.cell(row=r, column=14); ac.hyperlink = job["apply_link"]
        ac.font = Font(name="Arial", color="0563C1", underline="single", bold=True)
        ac.alignment = Alignment(horizontal="center", vertical="center")
        tl = f"{tailor_url}?company={job['company'][:40].replace(' ','%20')}"
        tc = ws.cell(row=r, column=15); tc.hyperlink = tl
        tc.font = Font(name="Arial", color="D97706", underline="single", bold=True)
        tc.alignment = Alignment(horizontal="center", vertical="center")
        mc = ws.cell(row=r, column=1); mc.alignment = Alignment(horizontal="center", vertical="center")
        if s >= 80:
            mc.fill = PatternFill("solid", start_color="10B981"); mc.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
        elif s >= 60:
            mc.fill = PatternFill("solid", start_color="F59E0B"); mc.font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
        elif s >= 40:
            mc.fill = PatternFill("solid", start_color="FEF3C7"); mc.font = Font(name="Arial", bold=True, size=12, color="78350F")
        else:
            mc.fill = PatternFill("solid", start_color="FEE2E2"); mc.font = Font(name="Arial", bold=True, size=12, color="991B1B")
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            if c not in (1, 12, 13, 14, 15):
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.font = Font(name="Arial", size=10)
            cell.border = border

    widths = {"A": 8, "B": 32, "C": 22, "D": 22, "E": 16, "F": 12, "G": 36,
              "H": 50, "I": 32, "J": 26, "K": 28, "L": 12, "M": 12, "N": 10, "O": 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 70
    ws.freeze_panes = "A2"

    st = wb.create_sheet("Profile + Stats")
    st["A1"] = "Matched against resume:"; st["A1"].font = Font(bold=True, size=13)
    st["A2"] = profile["name"]
    st["A3"] = f"{profile['years_experience']} years · {profile.get('seniority','')}"
    st["A4"] = "Skills: " + ", ".join(profile["skills"][:20])
    st["A4"].alignment = Alignment(wrap_text=True)
    st["A6"] = "Scraped at"; st["B6"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    st["A7"] = "Analyzer"; st["B7"] = CONFIG["analyzer"]
    st["A8"] = "Total jobs"; st["B8"] = "=COUNTA(Jobs!B2:B1000)"
    st["A9"] = "Strong (80+)"; st["B9"] = '=COUNTIF(Jobs!A2:A1000,">=80")'
    st["A10"] = "Decent (60-79)"; st["B10"] = '=COUNTIFS(Jobs!A2:A1000,">=60",Jobs!A2:A1000,"<80")'
    st["A11"] = "Avg score"; st["B11"] = "=ROUND(AVERAGE(Jobs!A2:A1000),1)"
    st.column_dimensions["A"].width = 30; st.column_dimensions["B"].width = 22

    wb.save(output_path)
    print(f"\n✓ Saved {output_path}")


# ============================================================================
# MAIN
# ============================================================================
async def main():
    require_deps()

    print("=" * 60)
    print(f"Naukri Resume-Aware Scraper · Analyzer: {CONFIG['analyzer']}")
    print("=" * 60)

    gemini_client = None
    if CONFIG["analyzer"] == "gemini":
        if not CONFIG["gemini_api_key"]:
            print("\n❌ Gemini mode needs a key. Get one free at https://aistudio.google.com/apikey")
            print("   Then:  export GEMINI_API_KEY=...")
            sys.exit(1)
        gemini_client = genai.Client(api_key=CONFIG["gemini_api_key"])

    # Resume → profile (once, cached)
    print("\n📄 Loading resume profile...")
    profile = load_or_build_profile(gemini_client, CONFIG["gemini_model"])

    # Load seen-jobs history (so we never show the same job twice)
    seen_urls = load_seen_jobs(CONFIG["seen_jobs_file"])
    print(f"📚 Dedup memory: {len(seen_urls)} jobs already shown in past runs.")

    # Scrape all sources. Each source is isolated in try/except: if one blocks or
    # crashes, the others still run and we still match whatever we collected.
    all_jobs = []

    for loc in CONFIG["locations"]:
        print(f"\n📍 Naukri · {loc}")
        try:
            all_jobs += await scrape_naukri_listings(
                loc, CONFIG["search_keyword"], CONFIG["max_jobs_per_location"],
                CONFIG["experience_years"], seen_urls, CONFIG["max_days_old"])
        except Exception as e:
            print(f"  ❌ Naukri ({loc}) failed: {e} — continuing with other sources.")
        time.sleep(2)

    if CONFIG.get("include_linkedin"):
        for li_loc in CONFIG["linkedin_locations"]:
            print(f"\n🔗 LinkedIn · {li_loc}")
            try:
                all_jobs += scrape_linkedin_listings(
                    CONFIG["search_keyword"], li_loc,
                    CONFIG["max_jobs_per_location_linkedin"], seen_urls,
                    CONFIG["max_days_old"], CONFIG.get("linkedin_exp_filter", ""))
            except Exception as e:
                print(f"  ❌ LinkedIn ({li_loc}) failed: {e} — continuing.")
            time.sleep(2)

    if CONFIG.get("include_indeed"):
        for in_loc in CONFIG["indeed_locations"]:
            print(f"\n🟦 Indeed · {in_loc}")
            try:
                all_jobs += await scrape_indeed_listings(
                    in_loc, CONFIG["search_keyword"],
                    CONFIG["max_jobs_per_location_indeed"], seen_urls,
                    CONFIG["max_days_old"])
            except Exception as e:
                print(f"  ❌ Indeed ({in_loc}) failed: {e} — continuing.")
            time.sleep(2)

    if CONFIG.get("include_ziprecruiter"):
        for zr_loc in CONFIG["ziprecruiter_locations"]:
            print(f"\n🟩 ZipRecruiter · {zr_loc}")
            try:
                all_jobs += await scrape_ziprecruiter_listings(
                    CONFIG["search_keyword"], zr_loc,
                    CONFIG["max_jobs_per_location_ziprecruiter"], seen_urls,
                    CONFIG["max_days_old"])
            except Exception as e:
                print(f"  ❌ ZipRecruiter ({zr_loc}) failed: {e} — continuing.")
            time.sleep(2)

    # Within-run dedup across ALL sources by (title, company) — collapses the same
    # role cross-posted to multiple boards. Whitespace-normalized for safety.
    seen_pairs, unique = set(), []
    for j in all_jobs:
        k = (re.sub(r"\s+", " ", j["title"].lower()).strip(),
             re.sub(r"\s+", " ", j["company"].lower()).strip())
        if k not in seen_pairs:
            seen_pairs.add(k)
            unique.append(j)

    if not unique:
        print("\n✅ No new jobs found. Either nothing fresh this run, or all results were already shown.")
        return

    print(f"\n📊 {len(unique)} NEW jobs (fresh, unseen). Matching against {profile['name']}'s resume...")

    # Pass 1 — instant, free keyword score for EVERY job, so nothing is left
    # unanalyzed regardless of Gemini quota.
    analyzed = [{**job, "analysis": match_keyword(job, profile)} for job in unique]

    # Pass 2 — spend the scarce Gemini quota on the MOST PROMISING jobs first
    # (highest keyword score), upgrading their analysis to the smarter model.
    # Interrupt-safe: Ctrl+C just stops upgrading and proceeds to write/save.
    use_gemini = (CONFIG["analyzer"] == "gemini" and gemini_client is not None)
    if use_gemini:
        order = sorted(range(len(analyzed)),
                       key=lambda idx: -analyzed[idx]["analysis"]["match_score"])
        upgraded = 0
        try:
            for rank, idx in enumerate(order, 1):
                job = analyzed[idx]
                try:
                    a = match_gemini(gemini_client, CONFIG["gemini_model"], job, profile)
                    analyzed[idx] = {**job, "analysis": a}
                    upgraded += 1
                    print(f"  [{rank}/{len(order)}] {job['title'][:45]}  · gemini ({a['match_score']})")
                    time.sleep(4.5)  # stay under free-tier rate limit
                except QuotaExhausted:
                    print(f"\n  ⚡ Gemini daily quota used up after {upgraded} jobs. "
                          f"The remaining {len(order) - rank + 1} keep their keyword scores.\n")
                    break
        except KeyboardInterrupt:
            print(f"\n  ⏹ Interrupted — keeping the {upgraded} Gemini upgrades done so far "
                  f"and writing results now.")

    # Score threshold: only WRITE strong-enough matches to the sheet. Everything
    # analyzed is still remembered (below), so weak matches don't get re-analyzed
    # tomorrow — they just stay out of the Excel.
    min_score = CONFIG.get("min_score_to_show", 0)
    to_show = [j for j in analyzed if j["analysis"]["match_score"] >= min_score]
    hidden = len(analyzed) - len(to_show)

    # Dated output filename — each run gets its own file, never overwrites
    today = datetime.now().strftime("%Y-%m-%d")
    output_path = f"{CONFIG['output_prefix']}_{today}.xlsx"
    if to_show:
        write_excel(to_show, output_path, CONFIG["resume_tailor_url"], profile)
    else:
        print(f"\n  ℹ No jobs scored ≥ {min_score} this run — nothing written to Excel.")

    # Update seen-jobs memory with EVERYTHING analyzed (shown or hidden), so
    # nothing gets re-fetched or re-analyzed on future runs. Atomic save.
    for job in analyzed:
        seen_urls.add(job["_norm_url"])
    save_seen_jobs(CONFIG["seen_jobs_file"], seen_urls)

    strong = sum(1 for j in to_show if j["analysis"]["match_score"] >= 80)
    decent = sum(1 for j in to_show if 60 <= j["analysis"]["match_score"] < 80)
    n_gem = sum(1 for j in analyzed if j["analysis"].get("analyzed_by") == "gemini")
    n_kw = len(analyzed) - n_gem
    n_naukri = sum(1 for j in to_show if j.get("source") == "Naukri")
    n_li = sum(1 for j in to_show if j.get("source") == "LinkedIn")
    n_in = sum(1 for j in to_show if j.get("source") == "Indeed")
    n_zr = sum(1 for j in to_show if j.get("source") == "ZipRecruiter")
    print(f"\n✅ {len(to_show)} jobs shown ({strong} strong + {decent} decent), "
          f"{hidden} weak hidden (scored < {min_score}).")
    print(f"   Shown sources: {n_naukri} Naukri + {n_li} LinkedIn + {n_in} Indeed + {n_zr} ZipRecruiter.")
    print(f"   Analyzed by: {n_gem} gemini (smart) + {n_kw} keyword (offline).")
    print(f"   Dedup memory now: {len(seen_urls)} jobs.")
    if to_show:
        print(f"   📁 Open {output_path}.")


if __name__ == "__main__":
    asyncio.run(main())
