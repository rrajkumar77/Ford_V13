# Ford / Magnit CV Template Builder

A small Streamlit app that fills TEKsystems' Ford/Magnit "CV Template"
from a candidate resume — the same logic as the `ford-magnit-cv-template`
Claude Skill, packaged so it can run as a standalone web app.

- Upload a resume (PDF), optionally auto-extract the fields with Claude
  (bring your own Anthropic API key — entered in the app, never stored),
  review/edit the JSON, and download the filled `.docx`.
- The resume is appended to the CV as **editable text** — no email,
  phone, LinkedIn/GitHub, photo, or logos in that section.
- The CV Template's own photo field only gets filled if you explicitly
  upload a photo for it.

## Files

- `app.py` — the Streamlit UI.
- `cv_builder.py` — the actual docx-filling logic (no Streamlit
  dependency; also runs as a plain CLI — see the docstring at the top).
- `assets/Candidate_Template.docx` — the bundled blank template.
- `field-schema.md` — the full `candidate.json` schema, NA/TBC
  conventions, and formatting notes (also used as part of the Claude
  extraction prompt inside the app).
- `example_candidate.json` — a worked example matching the schema.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Opens at `http://localhost:8501`.

## Deploy via GitHub + Streamlit Community Cloud

1. **Create a GitHub repo and push this folder:**

   ```bash
   cd streamlit_app          # this folder
   git init
   git add .
   git commit -m "Ford/Magnit CV builder"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```

   (Create the empty repo on github.com first, or use `gh repo create`
   if you have the GitHub CLI.)

2. **Deploy on Streamlit Community Cloud** (free, no server to manage):
   - Go to <https://share.streamlit.io> and sign in with your GitHub account.
   - Click **"New app"**.
   - Pick your repo, branch `main`, and set **Main file path** to `app.py`.
   - Click **Deploy**.
   - Streamlit Cloud reads `requirements.txt` automatically and installs
     everything.

3. **The Anthropic API key is never part of the repo.** Each person using
   the deployed app pastes their own key into the sidebar at runtime — it
   lives only in that browser session. If you'd rather set one default
   key for your whole team instead of everyone entering their own:
   - In the Streamlit Cloud app's **Settings → Secrets**, add:
     ```toml
     ANTHROPIC_API_KEY = "sk-ant-..."
     ```
   - Then in `app.py`, change the `api_key = st.text_input(...)` line to
     fall back to `st.secrets.get("ANTHROPIC_API_KEY", "")` as its
     default, so the field is pre-filled but still overridable.
   - Never commit an actual key into `app.py` or any file in the repo —
     secrets belong in Streamlit Cloud's Secrets manager (or a local
     `.streamlit/secrets.toml`, which `.gitignore` already excludes).

4. **Updating later:** just `git push` again — Streamlit Cloud
   auto-redeploys on every push to the tracked branch.

## If Ford/Magnit issues a new template

`cv_builder.py` fills the docx by matching exact text/paragraph anchors
in `assets/Candidate_Template.docx`. If a new template revision changes
labels or structure, some fields may stop filling (the app will still
generate a file, but a field may come out blank). To fix:

1. Diff the new template against the bundled one — render both to PDF
   and compare page by page.
2. Open `cv_builder.py` and search for the anchor text that changed
   (e.g. a relabeled field), update it to match the new wording, and
   re-test with `example_candidate.json` before using it on a real
   candidate.
3. Replace `assets/Candidate_Template.docx` with the new file and
   re-test the whole flow.

You can also upload a replacement template directly in the app's sidebar
for a one-off run without changing the repo, but if it changes the
anchors that still needs a code update to actually fill correctly.

## Changelog

**2026-09-18** — Fixes after a real candidate run (Ford/Magnit, jumbled
two-column Naukri-style PDF) produced a broken docx: empty Employment
History, "NA" Education, contact details leaking into Top Skills/
Education, and a mangled Top Skills line ("2011 Roadmap 2009").
Root causes and fixes:
- `pdf_layout.py` only flushed its left/right column buffers once, at
  the very end of the whole document, so a multi-page sidebar (or body)
  column got concatenated into one giant run with unrelated pages'
  content sitting right next to it. Now flushes at each page boundary.
- `extract_candidate.py`'s `_merge_wrapped_lines` treated any short,
  space-free token as a wrapped continuation of the previous line —
  including a bare 4-digit timeline year ("2011"), which is how it
  ended up glued onto an adjacent, unrelated skill line. Bare years are
  now never merged, and are dropped outright if they still end up in
  the skills/certifications pool.
- Sidebar headers like "PERSONAL INFORMATION" and "OTHER PERSONAL
  DETAILS" weren't in `SECTION_HEADERS`, so unmatched, their contact
  content could bleed into whichever real section happened to be open.
  They're now a recognized `_ignore` section, and every extracted
  section is additionally passed through `_strip_pii()` as a hard
  backstop that drops any line matching an email or phone pattern,
  regardless of which section it lands in.
- `cv_builder.py`: the Certification box on page 1 is a fixed-height
  text box with no autofit, so more than ~3 lines visually overflow
  past its border. `build_docx()` now caps the box at 3 certifications
  and returns `(docx_bytes, warnings)` — the app surfaces a warning
  when it had to trim; the full list still goes into the appended
  resume section.

None of this makes heuristic extraction perfect on every layout — a
complex two-column/timeline resume can still confuse it. Always review
the extracted JSON (the app already prompts for this) before
generating, especially for non-standard layouts.

**2026-09-18 (follow-up)** — The previous fix caught umbrella sidebar
headers ("PERSONAL INFORMATION") but this candidate's PDF export labels
each contact/personal field individually ("Email", "Total work
experience", "Social Link", "City", ...), with no umbrella header at
all - so those labels (and the value line right after each) kept
falling through to whichever real section was open, contaminating
Education again. `SECTION_HEADERS["_ignore"]` now lists these individual
field labels explicitly.

Adding short, generic labels like "city"/"state"/"address" to a
prefix-matched header list is risky on its own - "State-of-the-art
platform" or "Addressed client concerns" would wrongly match "state"/
"address" as a header. `_match_header()` now requires a real word
boundary (whitespace, ':', or end of line - not a bare letter and not a
hyphen, which would still let "state-of-the-art" slip through) right
after the keyword before treating it as a header match.

Also fixed: "2011 Roadmap" (a stray timeline year fused onto an
unrelated skill word) was surviving as one *already-combined* token
straight out of pdfplumber's own word-to-row grouping - before any
Python-side line-merging ever ran, so the earlier bare-year guard
(which only caught a line that was *only* a year) couldn't catch it.
`_parse_skills_and_certs()` now strips a leading or trailing bare
4-digit year from every skills/certifications line before treating it
as an item.

**2026-09-19 (real PDF, ground truth)** — Previous two rounds of fixes
were made blind, from broken docx output alone. Given the actual source
PDF for the first time, found the real root causes (different from what
the earlier rounds guessed):

- **Column boundary was miscalibrated, not "sidebar vs main" logic
  itself.** The old calibration sampled rows with exactly 2 subgroups to
  find the sidebar/main-content gutter - but on a timeline-layout resume
  (date column + description column, both already inside "main"), most
  2-subgroup rows are that inner date/description split, not the true
  sidebar split. The calibrated boundary landed inside the main content
  region (verified: x=288 on the real file), so every timeline date got
  classified as sidebar content. Replaced with `_global_gap_boundary()`:
  project every subgroup's horizontal extent onto the x-axis, merge
  overlapping ones, and take the largest gap no text anywhere on the
  page ever crosses. Verified this lands at x=203, the actual gutter.
- The header zone (name/title above the columns) was buffered through
  the same left/right logic as the body, so a recognized section header
  landing in that zone (e.g. because the true gutter is now correctly
  further left) could jump ahead of the name in the output. Fixed by
  keeping the header zone in natural top-to-bottom order, with its
  cutoff computed dynamically (where the sidebar first appears on page
  1) instead of a fixed pixel guess.
- `_merge_wrapped_lines` was gluing entire independent Skills and
  Certifications lists into one or two blobs - every short, space-free
  line (a single-word skill like "Scrum") looks like a wrapped
  continuation to that heuristic. Both sections are one complete item
  per line on this export style; merging is no longer applied to either.
- **Employment History came back completely empty.** Two compounding
  causes: (1) date ranges render as a wrapped 2-line badge ("Feb 2022 -
  Mar" / "2024"), so `DATE_RANGE_RE` never matched either half alone -
  `_repair_split_date_ranges()` now rejoins them; (2) this resume's PDF
  only has ONE "WORK EXPERIENCE" header covering 5 jobs, and an
  unrelated header re-appearing mid-stream (a sidebar "LANGUAGES" list)
  stole `_find_sections`' "current" pointer before jobs 2-5 were
  captured - `_collect_experience_raw()` now recovers any section's
  content that contains a genuine date range, since a real
  language/skills/contact list never would. Also added a forward-scan
  for title/company (date, then title, then company - this template's
  actual order) since the existing backward-scan was occasionally
  reaching past a short line with no trailing punctuation into
  unrelated recovered content and confidently returning the wrong
  title/company instead of falling through.
- "COURSES & CERTIFICATIONS" didn't match any certifications keyword
  (it starts with "Courses", not "Certifications") - added.
- Education: a single section-wide `fallback_institution` (whichever
  institution-looking line came first in the whole section) was reused
  for every degree row missing an inline institution, so the bachelor's
  degree kept getting the master's institution. Now looks at the
  specific next line after each degree row first.
- `resume["projects"]` was hard-coded to always return empty - never
  implemented. Added a best-effort single-project extractor.
- Individual sidebar field labels ("Total work experience", "Social
  Link", "City", "Country", ~20 more) still weren't all recognized,
  continuing to leak into whichever section was open. Adding short,
  generic ones like "city"/"state" safely required tightening
  `_match_header()` to require a real word boundary after the keyword -
  otherwise "State-of-the-art platform" would wrongly match "state".

Net result on the real test file: all 5 jobs, both degrees with correct
institutions, all 5 certifications, and clean top skills - verified by
rendering the actual generated docx, not just inspecting the JSON.

Known minor remaining rough edges (cosmetic, not data loss): a
single-project extraction can merge one extra sentence into the title
when the description's first sentence doesn't end in punctuation before
wrapping; one short leftover phrase can survive the overflow cleanup in
the appended resume's Languages line. Both are documented here rather
than chased further given how narrow and low-impact they are.

**2026-09-19 (Ford/Magnit logo missing on overflow pages)** — With 5
jobs, Employment History now runs to a 3rd page, and the Ford/Magnit
logos didn't appear there. Root cause: the logos were never a real Word
header - `assets/Candidate_Template.docx` had them pasted as plain
inline images directly in the body, once near the top of page 1's
content and a second copy near page 2's ("EDUCATION HISTORY"). There
was no `header1.xml` part at all, so anything pushing content past 2
pages had nothing to carry a 3rd copy.

Fixed at the template level (once, not per-candidate) rather than by
guessing where page breaks will fall at generation time - a docx has no
real page count until something actually renders it, so any
"insert a 3rd logo if content looks long enough" heuristic would be
wrong right at the threshold. Converted the template to use a genuine
repeating header instead: removed both inline logo paragraphs from the
body, added a real header (`section.header`) containing the same Ford +
tab + Magnit layout (identical image bytes, sizes, and tab-stop
position as the originals), and adjusted `top_margin`/`header_distance`
to make room. Word/LibreOffice now repeat it on every page
automatically - 3, 4, 5, however many - with no page-counting logic in
`cv_builder.py` at all.

Verified: rebuilt the 5-job real-PDF candidate and confirmed the logo
now appears on pages 1-4 (Employment History spills to page 3;
"ORIGINAL RESUME" starts on page 4 - previously never had a logo
either). Also rebuilt `example_candidate.json` and confirmed every
existing text anchor in `cv_builder.py` still matches with the expected
count (no `[MISSING]`/`[WARN]` output) - the header conversion didn't
disturb anything else in the template.

**2026-09-19 (certification box height + smaller font for overflow)** —
When a candidate's other page-1 fields (Experience in Core Skill, Other
Qualifications) run long enough to wrap to multiple lines, they eat into
the same fixed-height box the Certification list shares, so even the
existing 3-cert cap could overflow. Two changes:
- `cv_builder.py` now drops the certification font from 12pt to 9pt
  whenever there are more than 3 certifications, instead of silently
  truncating the list - a real full list matters more than font
  consistency here, especially now that the appended "original resume"
  section (the previous fallback for showing the rest) is optional and
  often left out.
- `assets/Candidate_Template.docx`: grew the Certification box itself
  (~0.7in taller) and shifted the "Candidate's Top 3 Skills" box down by
  the same amount so they don't collide. This is a grouped floating
  shape (photo + text box together) - grown by increasing the group's
  outer extent AND its child-coordinate extent (`chExt`) by the same
  amount, keeping the internal 1:1 scale so the photo box and existing
  text position are completely undisturbed; the extra height just adds
  blank space at the bottom for more content.

Verified on the real 5-certification candidate: all 5 certifications
now render inside the box at 9pt, no overlap with the Top-3-Skills box,
and a fresh render of `example_candidate.json` (3 certs, fits at normal
size) confirms normal-length candidates are unaffected other than a
small amount of extra whitespace in the box.

**2026-09-19 (bullets split across multiple lines / merged together)** —
Real bug behind "bullet points overlap": each *physical PDF text line*
was being treated as its own bullet in `_guess_employment`, rather than
each *logical* bullet (the source resume uses "•" - or "+" on at least
one job - as an inline separator within a flowing paragraph, not as a
per-line marker). Since the PDF wraps lines mid-sentence with no regard
for where a logical bullet starts or ends, this both split one real
bullet across 2-3 output bullets AND merged the tail of one bullet with
the start of the next onto a single output bullet whenever a wrap
happened to land exactly on a "• " boundary.

Fixed with `_reconstruct_bullets()`: join all of a job's raw lines back
into one continuous string first, then split on the real inline marker
(`•`, or a `+` immediately followed by a letter). Verified on the real
resume - every bullet is now one complete, grammatically whole sentence
matching the source, for all 5 jobs.

Fixing this exposed a related contamination bug in
`_collect_experience_raw()`: when recovering a job's content from an
overflow section (see the 2026-09-19 "ground truth" entry above), it
was pulling in that ENTIRE section, including whatever unrelated lines
sat before that section's own date anchor (sidebar/language-list
leftovers) - previously invisible because those stray lines just became
their own extra, easy-to-ignore bullet; now that bullets get joined into
one string, they were fusing onto the end of the PRECEDING job's last
real bullet instead. Now only takes each recovered section's content
from its own first date anchor onward.

Known minor remaining rough edge: on this same resume, two stray words
("SQL Power BI") still tail-end athenahealth's last bullet - the
original PDF has that job's true final clause physically separated from
the rest of its own bullet by several sections' worth of sidebar
content, in a way this fix can't safely disentangle from genuinely
unrelated leftover content without risking the contamination bug above.
Documented rather than chased further; costs 2 words on one bullet for
one known-messy resume vs. a much worse regression.

**2026-09-20 (tested against 3 new resumes - found and fixed 4 more real bugs)** —
Ran the pipeline against 3 different candidates' real resumes (single-
column bullet-list layout, a "Title Company Date" one-liner layout, and
a 2-digit-year date convention) instead of only the one two-column
resume everything so far had been tuned against. Found genuine bugs in
each, none specific to those resumes - all are logic bugs that could hit
any candidate with a similar (very common) layout:

- **Critical regression**: the "no reliable page-wide gap" result from
  `_global_gap_boundary` (added 2026-09-19) was falling back to the old,
  since-removed per-row-sample method instead of being trusted. On a
  plain single-column, hanging-indent bullet list - arguably the most
  common resume layout there is - that fallback mixed two unrelated
  small gaps (bullet-marker-to-text, and title-to-right-aligned-date)
  into one meaningless "boundary," and confidently reordered the entire
  single-column document as if it were two-column, scrambling every
  section. Fixed by trusting `None` (no clear gap = no columns) and
  going straight to plain single-stream extraction. The old fallback
  method (`_robust_bound` and the per-row-sample loop) is now removed
  entirely rather than left as a latent trap.
- **Critical**: `PHONE_RE` (used by `_strip_pii`'s defense-in-depth
  contact-info stripping, added 2026-09-19) is deliberately loose - any
  long-enough run of digits/spaces/parens/dashes/dots - which means an
  ordinary employment date range like "09/2023 - 09/2024" matches it
  just as well as a real phone number does, and was getting silently
  deleted as if it were PII. This alone caused 4 of 5 jobs to vanish
  from one resume. Fixed: a line matching `DATE_RANGE_RE` is now always
  exempted from the phone check - a genuine phone number never also
  looks like a date range, so this loses no real protection.
- **Critical**: `DATE_RANGE_RE`'s optional month-name prefix accepted
  ANY 3-9 letter word (meant to catch "Mar", "September", etc.), so on
  a "Title Company MM/YYYY - MM/YYYY" layout (title and company with no
  delimiter before the date) it was greedily swallowing the company
  name itself into the date match ("KaptureCX 09/2024" parsed as one
  "date"), corrupting the duration and deleting the company from the
  remaining text used for title/company splitting. Restricted the
  prefix to an actual month-name alternation.
- Added support for a 2-digit apostrophe year convention ("Sep'20",
  "June'19") - previously only 4-digit years were recognized at all,
  so any resume using this (a common Indian-resume date style) got zero
  employment entries.
- `SECTION_HEADERS["skills"]` didn't recognize "TECH SKILLS" (only
  "technical skills") - added.
- Education discipline text wasn't stripped of leading punctuation
  ("- Civil Engineering..." instead of "Civil Engineering...") when a
  degree abbreviation touches the following dash/colon with no space
  ("B.Tech- Civil Engineering", "BE-Electronics & Communication") -
  fixed for both the with-institution and no-institution code paths.
- `_split_title_company`'s company result could keep a leading
  "-"/":"/"," left over from the split - stripped.

Verified with a 4-resume regression matrix (job count, employment
fields, top-skills, certifications) after every fix - no prior resume's
extraction regressed while fixing later ones.

Known remaining gap, not fixed: a title containing its own comma
("Team Lead, Business Analyst - Cotiviti - Healthcare") defeats the
comma-based title/company split heuristic - falls back to `company:
"NA"` (safe/honest) rather than a wrong guess. Also: a resume with no
dedicated Certifications section (only an unmapped "Achievements" list)
correctly returns `"NA"` rather than inventing content - this is
expected behavior, not a bug.

**2026-09-21 (reviewed against 3 generated outputs — 5 more real bugs found and fixed)** —
Reviewed three already-generated CVs (flagged by manual QA) against
their source resumes (recovered from the appended "ORIGINAL RESUME"
text where the source PDF wasn't directly available) and found five
more genuine extraction bugs, all logic errors rather than anything
specific to one candidate:

- **`_split_title_company` guessed wrong instead of saying "NA"** when
  a job header line had no recognized title keyword at all (e.g.
  "Speridian Technologies / Pvt. Ltd." - entirely a company name, no
  title on that line - the real title was embedded in the first
  bullet instead). The old fallback blindly took the first 1-2 words
  as "title" by word count alone, which is how "Speridian Technologies"
  became a fake job title and "Pvt. Ltd." became the "company." Now:
  no keyword found anywhere → the whole line is the company, title is
  left "NA" for manual entry rather than guessed wrong.
- Related, same function: a compound title with its own internal comma
  ("Team Lead, Business Analyst - Cotiviti - Healthcare") defeated the
  comma-first split (title got truncated to "Team Lead", and the actual
  company after the comma was never even looked at → "NA"). Now prefers
  splitting on " - " when present, which keeps the compound title
  intact and correctly finds the company in the very next segment,
  discarding a further trailing industry tag ("- Healthcare").
- **"Education & Certifications" as one heading wrapped across two
  physical lines** ("EDUCATION" / "& Certifications", nothing in
  between) was treated as two ordinary sequential headers - the second
  one immediately flushed an empty buffer and took over as "current,"
  silently losing "education" as a destination for everything that
  followed (which was a mix of the degree line AND the real
  certifications, with no further marker separating them). Certifications
  came back "NA" despite 5 real ones being present. Fixed in
  `_find_sections`: when a new header match arrives with nothing
  accumulated since the previous one, both categories now stay active
  together instead of one replacing the other.
- That fix's own naive first pass (excluding anything degree-shaped from
  the certifications pool) then wrongly excluded a genuine certification
  that happens to start with a degree-looking token ("MBA, PMI-GAC
  Certified" is a cert, not an education row). Tightened to require
  actual degree evidence (a named institution or a year) before
  excluding, not just the keyword prefix.
- `education_master_duration` was hard-coded to always return `"NA"` -
  never wired up to real extraction at all. Now pulled from the parsed
  master's degree line (which already carries a year range once
  present) - e.g. "2017-19" instead of always "NA".
- `YEAR_RANGE_RE` (education duration) only recognized a 4-digit start
  AND end year - a 2-digit end year ("2017-19", the same shorthand
  convention `DATE_RANGE_RE` already needed a fix for on 2026-09-20)
  went completely unmatched.

Verified with the full 4-resume regression matrix (job count,
duration/company/title per job, both degrees, certification count) -
no prior resume regressed while fixing these five.

Two issues from the same review are NOT fixed - they need the source
PDF (not just the already-generated, already-lossy docx output) to
diagnose with any confidence rather than guess: a candidate whose name
extracted as "NA NA," and the same candidate's employment history
merging multiple distinct jobs together on a "Title | Company | Date |
Location"-per-line pipe-delimited format. Both are logged, pending the
actual source file.
