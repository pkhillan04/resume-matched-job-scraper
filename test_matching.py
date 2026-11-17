"""Tests for the offline keyword matcher (``match_keyword``).

This is the function that decides which postings reach the spreadsheet when
Gemini is unavailable or out of quota, so a silent scoring bug here is a job
the candidate never sees. The tests below pin down the contract: the score
bands, the anti-tell penalty, the title and experience adjustments, and the
shape of the returned dict.

Where a test checks a title or salary adjustment it compares against a control
job that differs only in the field under test, so the assertion is about the
adjustment itself rather than the absolute score, which the coverage formula
may legitimately change later.
"""
import unittest

import job_scraper as js


def _profile(skills, yoe=2):
    return {"skills": list(skills), "years_experience": yoe}


def _job(title="Data Analyst", jd="", experience="", salary="Not disclosed"):
    return {"title": title, "full_jd": jd, "experience": experience,
            "salary": salary}


def _score(title="Data Analyst", jd="", experience="", salary="Not disclosed",
           skills=("sql",), yoe=2):
    return js.match_keyword(_job(title, jd, experience, salary),
                            _profile(skills, yoe))["match_score"]


# A JD that mentions exactly three known analyst skills and nothing else the
# scorer reacts to. Used as the neutral body for title/salary/experience tests.
PLAIN_JD = "The role uses sql, python and excel day to day."
PLAIN_SKILLS = ("sql", "python", "excel")


class TestScoreBounds(unittest.TestCase):
    def test_score_is_an_int(self):
        self.assertIsInstance(_score(jd=PLAIN_JD, skills=PLAIN_SKILLS), int)

    def test_score_never_exceeds_100(self):
        score = _score(title="Data Analyst", jd=PLAIN_JD, salary="8-12 LPA",
                       experience="1-3 Yrs", skills=PLAIN_SKILLS)
        self.assertLessEqual(score, 100)

    def test_score_never_goes_below_zero(self):
        # Every penalty at once: senior intern title, five anti-tells, an
        # experience range far above the candidate, no matching skills.
        score = _score(
            title="Senior Data Engineering Intern",
            jd="spark hadoop kafka scala pytorch tensorflow microservices",
            experience="10-15 Yrs", skills=("excel",))
        self.assertGreaterEqual(score, 0)


class TestCoverageScoring(unittest.TestCase):
    def test_covering_every_asked_skill_beats_covering_none(self):
        jd = "You will use sql, python and tableau."
        covered = _score(jd=jd, skills=("sql", "python", "tableau"))
        uncovered = _score(jd=jd, skills=("excel",))
        self.assertGreater(covered, uncovered)

    def test_more_absolute_matches_win_at_equal_coverage(self):
        # Both candidates cover 100% of what the JD asks; the one answering a
        # JD with more asks should rank higher.
        five = _score(jd="sql, python, tableau, excel and looker.",
                      skills=("sql", "python", "tableau", "excel", "looker"))
        two = _score(jd="sql and python.", skills=("sql", "python"))
        self.assertGreater(five, two)

    def test_jd_naming_no_known_skill_uses_the_fallback_band(self):
        # Nothing from the analyst skill universe appears, so the score falls
        # back to a mid band rather than to zero coverage.
        score = _score(jd="Own the weekly reporting rhythm for the region.",
                       skills=("sql",))
        self.assertGreaterEqual(score, 40)

    def test_fallback_band_rewards_resume_skill_mentions(self):
        # "jira" and "confluence" are resume skills but not in the analyst
        # universe, so they only move the fallback band.
        jd = "Day to day you will live in jira and confluence."
        with_matches = _score(jd=jd, skills=("jira", "confluence"))
        without = _score(jd=jd, skills=("cobol",))
        self.assertGreater(with_matches, without)


class TestAntiTells(unittest.TestCase):
    def test_engineering_tells_lower_the_score(self):
        clean = _score(jd=PLAIN_JD, skills=PLAIN_SKILLS)
        tell = _score(jd=PLAIN_JD + " Pipelines run on airflow and spark.",
                      skills=PLAIN_SKILLS)
        self.assertLess(tell, clean)

    def test_more_tells_lower_it_further(self):
        one = _score(jd=PLAIN_JD + " We use spark.", skills=PLAIN_SKILLS)
        three = _score(jd=PLAIN_JD + " We use spark, kafka and scala.",
                       skills=PLAIN_SKILLS)
        self.assertLess(three, one)

    def test_red_flag_names_the_tells(self):
        out = js.match_keyword(
            _job(jd=PLAIN_JD + " Built on spark and kafka."),
            _profile(PLAIN_SKILLS))
        self.assertIn("spark", out["red_flags"])
        self.assertIn("different role", out["red_flags"])


class TestTitleAdjustments(unittest.TestCase):
    """Title terms must match whole words.

    A substring check reads "Internal" as "intern" and "Headquarters" as
    "head", which drops a real analyst posting by 30 points and tags it with a
    red flag that is simply untrue.
    """

    def test_target_titles_get_a_bonus(self):
        self.assertGreater(
            _score(title="Data Analyst", jd=PLAIN_JD, skills=PLAIN_SKILLS),
            _score(title="Associate", jd=PLAIN_JD, skills=PLAIN_SKILLS))

    def test_senior_titles_are_penalised(self):
        self.assertLess(
            _score(title="Senior Data Analyst", jd=PLAIN_JD,
                   skills=PLAIN_SKILLS),
            _score(title="Data Analyst", jd=PLAIN_JD, skills=PLAIN_SKILLS))

    def test_internships_are_penalised_and_flagged(self):
        out = js.match_keyword(_job(title="Data Analyst Intern", jd=PLAIN_JD),
                               _profile(PLAIN_SKILLS))
        self.assertLess(out["match_score"],
                        _score(title="Data Analyst", jd=PLAIN_JD,
                               skills=PLAIN_SKILLS))
        self.assertIn("Internship", out["red_flags"])

    def test_the_word_internship_is_also_caught(self):
        out = js.match_keyword(
            _job(title="Data Analyst Internship", jd=PLAIN_JD),
            _profile(PLAIN_SKILLS))
        self.assertIn("Internship", out["red_flags"])

    def test_internal_is_not_an_internship(self):
        out = js.match_keyword(
            _job(title="Internal Reporting Analyst", jd=PLAIN_JD),
            _profile(PLAIN_SKILLS))
        self.assertEqual("None", out["red_flags"])
        self.assertEqual(out["match_score"],
                         _score(title="Reporting Analyst", jd=PLAIN_JD,
                                skills=PLAIN_SKILLS))

    def test_headquarters_is_not_a_head_of_role(self):
        self.assertEqual(
            _score(title="Analyst, Headquarters Finance", jd=PLAIN_JD,
                   skills=PLAIN_SKILLS),
            _score(title="Analyst, Regional Finance", jd=PLAIN_JD,
                   skills=PLAIN_SKILLS))

    def test_mis_does_not_match_a_longer_word(self):
        self.assertEqual(
            _score(title="Analyst, Misconduct Review", jd=PLAIN_JD,
                   skills=PLAIN_SKILLS),
            _score(title="Analyst, Conduct Review", jd=PLAIN_JD,
                   skills=PLAIN_SKILLS))

    def test_mis_titles_still_get_the_bonus(self):
        self.assertGreater(
            _score(title="MIS Executive", jd=PLAIN_JD, skills=PLAIN_SKILLS),
            _score(title="Operations Executive", jd=PLAIN_JD,
                   skills=PLAIN_SKILLS))

    def test_plural_target_titles_get_the_bonus(self):
        self.assertGreater(
            _score(title="Analytics Associates - 3 openings", jd=PLAIN_JD,
                   skills=PLAIN_SKILLS),
            _score(title="Associates - 3 openings", jd=PLAIN_JD,
                   skills=PLAIN_SKILLS))


class TestExperienceRange(unittest.TestCase):
    def test_range_far_above_the_candidate_is_penalised_and_flagged(self):
        out = js.match_keyword(
            _job(jd=PLAIN_JD, experience="10-15 Yrs"), _profile(PLAIN_SKILLS))
        self.assertLess(out["match_score"],
                        _score(jd=PLAIN_JD, skills=PLAIN_SKILLS))
        self.assertIn("too senior", out["red_flags"])

    def test_range_around_the_candidate_gets_a_bonus(self):
        self.assertGreater(
            _score(jd=PLAIN_JD, experience="1-3 Yrs", skills=PLAIN_SKILLS),
            _score(jd=PLAIN_JD, skills=PLAIN_SKILLS))

    def test_no_range_means_no_adjustment(self):
        self.assertEqual(
            _score(jd=PLAIN_JD, experience="", skills=PLAIN_SKILLS),
            _score(jd=PLAIN_JD, experience="Not specified",
                   skills=PLAIN_SKILLS))

    def test_a_single_figure_is_not_read_as_a_range(self):
        # Documents current behaviour: only "lo-hi" is parsed.
        self.assertEqual(
            _score(jd=PLAIN_JD, experience="8 years", skills=PLAIN_SKILLS),
            _score(jd=PLAIN_JD, experience="", skills=PLAIN_SKILLS))


class TestSalaryBonus(unittest.TestCase):
    def test_a_disclosed_salary_adds_a_bonus(self):
        self.assertGreater(
            _score(jd=PLAIN_JD, salary="6-9 LPA", skills=PLAIN_SKILLS),
            _score(jd=PLAIN_JD, salary="Not disclosed", skills=PLAIN_SKILLS))

    def test_an_absent_salary_is_treated_like_not_disclosed(self):
        # LinkedIn and ZipRecruiter postings often carry no salary at all.
        # Scoring an empty string as "disclosed" handed them a free bonus over
        # Naukri postings that honestly say "Not disclosed".
        baseline = _score(jd=PLAIN_JD, salary="Not disclosed",
                          skills=PLAIN_SKILLS)
        self.assertEqual(baseline, _score(jd=PLAIN_JD, salary="",
                                          skills=PLAIN_SKILLS))
        self.assertEqual(baseline, _score(jd=PLAIN_JD, salary="   ",
                                          skills=PLAIN_SKILLS))
        missing = js.match_keyword(
            {"title": "Data Analyst", "full_jd": PLAIN_JD, "experience": ""},
            _profile(PLAIN_SKILLS))
        self.assertEqual(baseline, missing["match_score"])


class TestSkillMatching(unittest.TestCase):
    def test_skills_match_whole_words_only(self):
        out = js.match_keyword(
            _job(jd="Own the reporting rhythm."), _profile(("r",)))
        self.assertNotIn("r", out["key_skills_required"])

    def test_a_skill_inside_a_longer_word_does_not_match(self):
        out = js.match_keyword(
            _job(jd="Experience with nosql stores."), _profile(("sql",)))
        self.assertNotIn("sql", out["key_skills_required"])

    def test_skills_ending_in_punctuation_still_match(self):
        # "\b" needs a word character on its side, so a plain \bc\+\+\b can
        # never match anything: every such resume skill was invisible.
        out = js.match_keyword(
            _job(jd="The analytics tooling is written in c++."),
            _profile(("c++", "sql")))
        self.assertIn("c++", out["key_skills_required"])

    def test_multi_word_skills_match(self):
        out = js.match_keyword(
            _job(jd="Dashboards are built in power bi."),
            _profile(("power bi",)))
        self.assertIn("power bi", out["key_skills_required"])


class TestOutputShape(unittest.TestCase):
    def test_analyzed_by_is_keyword(self):
        out = js.match_keyword(_job(jd=PLAIN_JD), _profile(PLAIN_SKILLS))
        self.assertEqual("keyword", out["analyzed_by"])

    def test_every_documented_key_is_present(self):
        out = js.match_keyword(_job(jd=PLAIN_JD), _profile(PLAIN_SKILLS))
        for key in ("match_score", "fit_summary", "key_skills_required",
                    "skills_gap", "jd_summary", "red_flags", "analyzed_by"):
            self.assertIn(key, out)

    def test_key_skills_are_capped_at_six(self):
        skills = ["sql", "python", "excel", "tableau", "looker", "dbt",
                  "snowflake", "bigquery"]
        out = js.match_keyword(_job(jd=" ".join(skills)), _profile(skills))
        self.assertEqual(6, len(out["key_skills_required"]))

    def test_key_skills_fall_back_to_a_placeholder(self):
        out = js.match_keyword(_job(jd="Own the weekly rhythm."),
                               _profile(("cobol",)))
        self.assertEqual(["See JD"], out["key_skills_required"])

    def test_skills_gap_lists_what_the_candidate_lacks(self):
        out = js.match_keyword(
            _job(jd="Needs sql, python and tableau."), _profile(("sql",)))
        self.assertIn("python", out["skills_gap"])
        self.assertIn("tableau", out["skills_gap"])
        self.assertNotIn("sql", out["skills_gap"])

    def test_skills_gap_is_capped_at_four(self):
        out = js.match_keyword(
            _job(jd="sql python tableau power bi looker snowflake dbt"),
            _profile(("excel",)))
        self.assertEqual(4, len(out["skills_gap"]))

    def test_jd_summary_collapses_whitespace_and_truncates(self):
        jd = "word " * 200
        out = js.match_keyword(_job(jd=jd), _profile(PLAIN_SKILLS))
        self.assertNotIn("  ", out["jd_summary"])
        self.assertTrue(out["jd_summary"].endswith("..."))
        self.assertEqual(253, len(out["jd_summary"]))

    def test_a_short_jd_is_not_truncated(self):
        out = js.match_keyword(_job(jd=PLAIN_JD), _profile(PLAIN_SKILLS))
        self.assertEqual(PLAIN_JD, out["jd_summary"])

    def test_red_flags_reads_none_when_clean(self):
        out = js.match_keyword(_job(jd=PLAIN_JD), _profile(PLAIN_SKILLS))
        self.assertEqual("None", out["red_flags"])

    def test_fit_summary_tracks_the_score_band(self):
        strong = js.match_keyword(
            _job(jd="sql, python, excel and tableau.", salary="8-12 LPA"),
            _profile(("sql", "python", "excel", "tableau")))
        weak = js.match_keyword(
            _job(title="Senior Data Engineer",
                 jd="spark hadoop kafka scala airflow"),
            _profile(("excel",)))
        self.assertGreaterEqual(strong["match_score"], 80)
        self.assertIn("Strong match", strong["fit_summary"])
        self.assertLess(weak["match_score"], 40)
        self.assertIn("Weak fit", weak["fit_summary"])


if __name__ == "__main__":
    unittest.main()
