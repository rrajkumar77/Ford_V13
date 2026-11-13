# candidate.json schema — Ford/Magnit CV Template

Feed this JSON into `scripts/build_cv.py --data candidate.json`. Every field
is optional; anything omitted falls back to `"NA"` (Ford-history / cert /
score fields) or `"TBC"` (scheduling fields), matching the convention
Rajkumar asked for in Sep 2026.

## Extraction rule

Use ONLY information present in the resume. Do not invent, infer, or
estimate anything not stated. If a field is not present in the resume, use
`"NA"` (or `"TBC"` for the four scheduling fields below) rather than
guessing.

## Top-level fields

| Key | Notes |
|---|---|
| `first_name` | |
| `last_name` | **Only fill from what the resume itself states as the candidate's full name.** Don't invent a surname/initial from a filename or guesswork — client onboarding systems check this against passport/Aadhaar, and a wrong guess is worse than a gap. If the resume gives one name with no clear split (e.g. just "Nandan"), set `last_name` to `"NA"` and flag it to Rajkumar so he can confirm the legal name with the candidate rather than submitting an unverified guess. |
| `notice_period` | Keep SHORT — see "Line-wrap warning" below. Default `"TBC"`. |
| `internal_external` | `"Internal"` / `"External"` / default `"TBC"` |
| `interview_availability` | Keep SHORT. Default `"TBC"` |
| `start_availability` | Keep SHORT. Default `"TBC"` |
| `worked_ford_direct` | `"Yes"` / `"No"`. Default `"No"` if not stated. |
| `worked_ford_agency` | `"Yes"` / `"No"`. Default `"No"` if not stated. |
| `cdsid` | Only relevant if either of the above is Yes. Default `"NA"` |
| `supervision_duration` | Default `"NA"` |
| `exit_reason` | Default `"NA"` |
| `overall_it_experience` | Only fill if the resume states a total, e.g. "6 Years". Default `"NA"` |
| `experience_core_skill` | Default `"NA"` if the resume doesn't break out a separate figure. But if the candidate's *entire* stated career has been in one consistent specialization (e.g. every role is Azure Data Engineering), it's reasonable to set this equal to `overall_it_experience` rather than leaving it NA — that's not inventing a number, it's recognizing the whole career total already answers the "core skill" question. Only do this when the resume genuinely shows one consistent specialization throughout; if the candidate switched domains partway through, leave it NA rather than guess which years count. |
| `other_qualifications` | Default `"NA"` |
| `hackerrank_score` | Default `"NA"` |
| `certifications` | **List** of certification strings, one per item — each renders as its own numbered line (not one run-on bullet). A single string still works (split on `;`/newlines) for backward compatibility, but prefer a list. Default `["NA"]`. |
| `top_skills` | List of exactly 3 short phrases (not sentences), most prominent skills/domains from the resume. |
| `education_bachelor` | One line: degree – college (affiliation) – completion date. |
| `education_master` | One line: degree(s) (specialization) – university name. If candidate has 2 postgrad degrees (e.g. M.Sc + M.Tech), combine with `;`. |
| `education_master_duration` | e.g. "NA – 05.2022 (M.Tech); NA – 06.2020 (M.Sc)" — resume rarely states a start date, so leading side is usually "NA". |
| `education_bachelor_detail` | Repeats bachelor's degree, college, and university affiliation (the template has a separate line for this). |
| `employment` | List of jobs, **most recent first**, see below. Not capped at 3 — the script rebuilds the whole section for however many entries you give it. |
| `resume` | Optional. If present, the resume is appended after the template as **editable text** (not an image). See "Resume" below for the schema and what must be excluded. |

### `employment[]` entry

```json
{
  "duration": "June 2022 \u2013 Present",
  "company": "Hewlett Packard (HP Inc), Bengaluru, India",
  "title": "Embedded System Engineer & AI Specialist",
  "technology": "Agentic AI, MCP server, Embedded Linux, Yocto Project",
  "role": ["Short bullet one", "Short bullet two", "Short bullet three"]
}
```

`role` can be a **list of short bullets** (preferred — recruiters scan this
table, and 4-8 short bullets read far faster than one dense paragraph) or
a single string (rendered as one paragraph — fine for a role with very
little to say, but prefer bullets once there's more than 2-3 sentences of
substance). Either way, condense the resume's own bullets rather than
rewriting their meaning — pick the most substantive points and drop
near-duplicates rather than paraphrasing what you keep.

**Company-name consistency check:** some resumes state one employer name
in a summary line (e.g. "Worked at Acme Corp, Jan 2022–Dec 2023") but a
*different* vendor/client name in the detailed role breakdown for the same
dates (common in staffing/consulting resumes with sub-vendors or client
engagements). When you see this, don't silently pick one name to make the
table look clean — use the same company/client name here that you used in
the `resume` block's `work_experience` for that date range, so the two
sections of the document agree with each other, and call the discrepancy
out to Rajkumar in your reply so he can verify the correct legal employer
name with the candidate before submission (background-verification teams
will flag a mismatch between what's submitted and what they find).

### Resume (appended as editable text — contact details, photo, and logos excluded)

Standing instruction from Rajkumar (Sep 2026): the appended resume must be
**editable Word text**, not a picture of the original file, and must
**never include**:
- the candidate's contact details — email, mobile number, LinkedIn,
  GitHub, or any other personal link (TEKsystems/Magnit handle candidate
  contact, so this doesn't go to the client)
- the candidate's photo
- any organisation/company logo

The candidate's **name** is fine to keep — it's not a contact method.
Everything else (roles, dates, bullets, education, skills, projects,
certificates) should be reproduced faithfully from the resume, condensed
lightly if needed but not rewritten in meaning — same rule as the
`employment[]` entries above.

```json
{
  "name": "AMESH A M",
  "summary": ["Optional — one string or a list of bullet lines, if the resume has a career-summary/profile section worth keeping."],
  "work_experience": [
    {"title": "...", "company": "...", "duration": "...", "bullets": ["...", "..."]}
  ],
  "education": ["MBA/PGDM \u2014 P E S Institute of Technology, Bangalore \u2014 2019", "..."],
  "skills": ["Roadmap", "Backlog Refinement", "Scrum", "..."],
  "projects": [{"title": "...", "description": "..."}],
  "certificates": ["POPM", "Google Data Analytics Professional Certificate", "..."],
  "languages": ["English", "Kannada", "Hindi"]
}
```

Do not add an `"email"`, `"phone"`, `"linkedin"`, `"github"`, `"contact"`,
or `"photo"` key to this block — `build_resume_section_xml()` in
`scripts/build_cv.py` doesn't render one even if present, but the cleanest
approach is to just never extract those fields from the resume into this
JSON in the first place.


## Line-wrap warning

`notice_period`, `internal_external`, `interview_availability`, and
`start_availability` sit in a text box that reserves right-side space for
the candidate photo. Long values (e.g. the literal phrase "To be
confirmed") wrap onto a second line and look broken. Use short values:
`"TBC"`, `"Immediate"`, `"7 Days"`, `"15 Days"`, `"2 Weeks"`. If a value
must be longer, check the rendered PDF and shorten it if it wraps.

## Photo (CV Template's own field only)

As of the Sep-2026 template revision, the CV Template's photo box (page 1)
ships **empty** — Ford/Magnit removed the old sample placeholder photo.
Pass `--photo /path/to/photo.jpg` to insert a real candidate photo into
that box. It gets center-cropped to the box's aspect ratio and inserted —
see `insert_photo()` in `scripts/build_cv.py` for how this works (it
places a new picture into the template's empty photo-box shape, since
there's no existing image to overwrite anymore).

Per Rajkumar's Sep-2026 instruction, don't source this photo from the
candidate's resume automatically — only pass `--photo` when he's
explicitly supplied one for this field. If `--photo` is omitted, the box
stays empty, which is now the correct default (not a sign something's
missing) — no need to flag it to Rajkumar unless he's separately supplied
a photo you forgot to pass in. (The resume section appended below the
template never has a photo, regardless.)
