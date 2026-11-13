"""
extract_candidate.py - Rule-based (regex/heuristics) extraction of the
Ford/Magnit candidate.json schema straight from resume text.

No LLM, no API key, no network call. Purely local parsing. It will not be
as accurate as an LLM extraction on messy/unusual resume layouts -
always review the JSON in the app before generating the docx.

Fields it cannot confidently find are left at the schema defaults
("NA" / "TBC"), per field-schema.md's extraction rule: never guess.
"""
import re

SECTION_HEADERS = {
    "summary": ["summary", "profile", "career summary", "professional summary", "objective"],
    "skills": ["skills", "technical skills", "tech skills", "core competencies", "key skills", "skills & certifications", "skills and certifications"],
    "certifications": ["certification", "certifications", "certificates", "courses", "courses & certifications", "courses and certifications", "& certifications", "and certifications"],
    "education": ["education", "academic", "qualifications", "education qualifications"],
    "experience": [
        "experience", "work experience", "employment", "employment history",
        "professional experience", "career history",
    ],
    "languages": ["languages", "language skills"],
    "projects": ["projects", "key projects"],
    # Sidebar blocks that hold exactly the contact/personal details this
    # tool must never carry into candidate.json. Recognizing them as their
    # own section - rather than leaving them unmatched - means their
    # content gets bucketed here and dropped, instead of silently bleeding
    # into whatever real section happened to be open when the (unmatched)
    # header line was seen.
    #
    # Naukri-style profile exports don't always group these under one
    # umbrella header - each field often has its own inline label
    # ("Email", "Total work experience", "Social Link", ...) sitting
    # directly in the sidebar. Every such label line has to be listed
    # here individually, or it (and the value line right after it) keeps
    # falling through to whatever real section happens to be open.
    "_ignore": [
        "personal information", "other personal details", "personal details",
        "contact", "contact details", "contact information", "personal info",
        "email", "e-mail", "mobile", "phone", "contact no", "contact number",
        "total work experience", "social link", "social links", "website",
        "current location", "current city", "home town", "hometown",
        "current salary", "expected salary", "current ctc", "expected ctc",
        "notice period", "date of birth", "dob", "gender", "marital status",
        "nationality", "passport", "visa status", "pan", "aadhar",
        "city", "country", "state", "pincode", "postal code", "address",
    ],
}

_MONTH_RE = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-zA-Z]*\.?"
# A 2-digit year attached directly to a month name ("Sep'20", "June'19")
# only - never a bare 2-digit number on its own, which is far too easy
# to confuse with something else in a resume. The (?!\d) guards against
# swallowing just the first two digits of an ordinary 4-digit year.
_MONTH_2DIGIT_YEAR_RE = rf"{_MONTH_RE}['\u2019]?\d{{2}}(?!\d)"
DATE_RANGE_RE = re.compile(
    rf"(?P<start>(?:{_MONTH_RE}[\s./-]*)?\d{{1,2}}[/.-]\d{{4}}|(?:{_MONTH_RE}[\s./-]*)?\d{{4}}|{_MONTH_2DIGIT_YEAR_RE}|Present)\s*"
    r"(?:[-\u2013\u2014to]{1,3})\s*"
    rf"(?P<end>(?:{_MONTH_RE}[\s./-]*)?\d{{1,2}}[/.-]\d{{4}}|(?:{_MONTH_RE}[\s./-]*)?\d{{4}}|{_MONTH_2DIGIT_YEAR_RE}|Present|Current|Till date|Till Date)",
    re.IGNORECASE,
)
YEAR_RANGE_RE = re.compile(r"\d{4}\s*[\u2013\u2014-]\s*(?:\d{4}|\d{2}(?!\d))")
DURATION_LINE_RE = re.compile(r"^(\d{1,2}\+?\s*Years?)(?:\s+(\d{1,2}\s*Months?))?$", re.IGNORECASE)
YEARS_EXPERIENCE_RE = re.compile(
    r"(\d{1,2}(?:\.\d)?\+?)\s*(?:years?|yrs?)\b[^.\n]{0,60}?\bexperience\b",
    re.IGNORECASE,
)
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
PHONE_RE = re.compile(r"(?:\+?\d[\d\s().-]{7,}\d)")
BULLET_RE = re.compile(r"^[\s]*[\u2022\-\*\u25cf\u25aa\u2023]\s*")
PAGE_NUM_RE = re.compile(r"^page\s+\d+\s+of\s+\d+$", re.IGNORECASE)

BACHELOR_PATTERNS = [r"\bb\.?\s?tech\b", r"\bb\.?\s?e\b", r"\bb\.?\s?sc\b", r"\bb\.?\s?com\b", r"\bbca\b", r"\bbachelor"]
MASTER_PATTERNS = [r"\bm\.?\s?tech\b", r"\bm\.?\s?sc\b", r"\bm\.?\s?com\b", r"\bmca\b", r"\bmba\b", r"\bpgdm\b", r"\bm\.?\s?s\b", r"\bmaster"]
EDU_HEADER_LABELS = {
    "degree", "discipline / stream", "discipline/stream", "stream",
    "institution / board", "institution/board", "institution", "board", "year",
}
CLASS_LABELS = {"x", "xi", "xii"}


def _lines(text):
    return [ln.rstrip() for ln in text.splitlines()]


def _match_header(line, keywords):
    """Return (remainder_text) if line starts with one of keywords, else None.
    Some headers share a line with real content (e.g. 'Professional
    Experience   6 Years 8 Months') - the leftover after the keyword is
    returned so it isn't silently dropped.

    Requires a word boundary right after the keyword, so a short generic
    label like "city" matches "City: Bengaluru" but not "Citywide rollout"
    or "State-of-the-art platform" for the keyword "state" - without this,
    short sidebar-field labels are too easy to accidentally match against
    the start of an unrelated sentence."""
    clean = line.strip().rstrip(":")
    low = clean.lower()
    if not low or len(low) > 60:
        return None
    matches = []
    for kw in keywords:
        if low == kw:
            matches.append(kw)
            continue
        if low.startswith(kw):
            nxt = low[len(kw):len(kw) + 1]
            # Whitespace, ':' or end-of-line only - NOT '-', since a
            # hyphen lets compound words like "state-of-the-art" slip
            # through as if "state" were a standalone label.
            if nxt == "" or nxt in " :":
                matches.append(kw)
    if not matches:
        return None
    kw = max(matches, key=len)
    remainder = clean[len(kw):].strip(" :-")
    return remainder if len(remainder) < 60 else ""


def _strip_pii(sections):
    """Defense in depth: drop any line containing an email or phone match
    from every extracted section, regardless of which section it landed
    in. This tool promises never to carry contact details into
    candidate.json - that promise shouldn't depend solely on sidebar
    headers being recognized and routed to "_ignore" correctly upstream;
    a contact line that slips through column/section misclassification
    (e.g. onto a "skills" or "education" buffer) must still never reach
    the output.

    PHONE_RE is deliberately loose (any long enough run of digits/
    spaces/parens/dashes/dots), which is exactly what a real employment
    date range also looks like ("09/2023 - 09/2024" matches it just as
    readily as a phone number does). A date range matching DATE_RANGE_RE
    is a well-defined, important pattern in its own right - checked
    first, and takes priority over the phone heuristic - so an
    employment anchor line is never discarded as if it were a phone
    number. A genuine phone number never also looks like a date range,
    so this loses no real PII protection."""
    for name, lines in list(sections.items()):
        sections[name] = [
            ln for ln in lines
            if DATE_RANGE_RE.search(ln) or (not EMAIL_RE.search(ln) and not PHONE_RE.search(ln))
        ]
    return sections


def _find_sections(lines):
    """Return {section_name: [lines]} by splitting on recognized headers.

    Handles a combined heading that wraps across two lines, e.g.
    "EDUCATION" / "& Certifications" (two separate physical lines, no
    content in between) - really just one heading, "Education &
    Certifications", that a narrow column wrapped in two. If it were
    treated as two ordinary sequential headers, the second would
    immediately flush an empty buffer and take over as "current",
    silently losing "education" as a destination for everything that
    follows - even though what follows is a mix of both degree info and
    actual certifications with no further marker separating them.
    Detected by: the new header match arrives with NOTHING accumulated
    since the previous header (buf is still empty) - in that specific
    case both categories stay active together, rather than one
    replacing the other, so subsequent content is filed under both.
    (A degree line ending up in the certifications bucket this way is
    harmless: the certifications parser skips anything shaped like a
    degree row - see _parse_skills_and_certs.)"""
    sections = {}
    current = []
    buf = []
    for line in lines:
        matched, remainder = None, ""
        for name, kws in SECTION_HEADERS.items():
            r = _match_header(line, kws)
            if r is not None:
                matched, remainder = name, r
                break
        if matched:
            if not buf and current and matched not in current:
                current.append(matched)
            else:
                for c in current:
                    sections.setdefault(c, []).extend(buf)
                current = [matched]
            buf = [remainder] if remainder else []
        elif current:
            buf.append(line)
    for c in current:
        sections.setdefault(c, []).extend(buf)
    return sections


def _clean_bullet(line):
    return BULLET_RE.sub("", line).strip()


def _split_name(top_lines):
    """Guess name from the first few non-empty lines (resumes usually lead with it)."""
    for line in top_lines[:8]:
        candidate = line.strip()
        if not candidate:
            continue
        if EMAIL_RE.search(candidate) or PHONE_RE.search(candidate):
            continue
        # Never treat a recognized section/sidebar header line (e.g.
        # "PERSONAL INFORMATION", "KEY SKILLS") as the candidate's name -
        # a real name never coincides with one of these keywords, but a
        # header can legitimately end up ahead of the name in the
        # flattened text on layouts where the sidebar is read first.
        if any(_match_header(candidate, kws) is not None for kws in SECTION_HEADERS.values()):
            continue
        words = candidate.split()
        if not (1 <= len(words) <= 5):
            continue
        cleaned = [w.strip(".,\u00ae\u2122()'\"") for w in words]
        if all(w and w.replace("-", "").replace("'", "").isalpha() for w in cleaned):
            if len(cleaned) == 1:
                return cleaned[0], "NA"
            return cleaned[0], " ".join(cleaned[1:])
    return "NA", "NA"


def _split_outside_parens(text):
    """Comma-split, but ignore commas inside (...) - e.g. 'Regulatory &
    Compliance (HIPAA, GDPR, CFPB)' should stay one item, not four."""
    parts, depth, current = [], 0, ""
    for ch in text:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        if ch == "," and depth == 0:
            parts.append(current)
            current = ""
        else:
            current += ch
    parts.append(current)
    return parts


def _merge_wrapped_lines(lines):
    """Rejoin a line that's really the tail of the previous one - common
    when a bullet/skill/cert wraps: 'Regulatory & Compliance (HIPAA, GDPR,'
    + 'FCRA)', or 'LCAP - LTIMindtree Certified Agile' + 'Professional'."""
    merged = []
    for raw in lines:
        line = raw.strip()
        if not line:
            continue
        if merged:
            prev = merged[-1]
            # A bare 4-digit year ("2019") is a common standalone token in
            # its own right - the date column of a timeline-style
            # education/experience layout - never a wrapped continuation
            # of the previous line. Without this guard it satisfies the
            # short-fragment heuristic below and gets glued onto whatever
            # (unrelated) line preceded it, e.g. a skill list item, which
            # is how "Roadmap" and "2011"/"2009" ended up on the same line.
            is_bare_year = bool(re.fullmatch(r"(19|20)\d{2}", line))
            prev_dangling = bool(re.search(r"[\u2013\u2014\-&/:,]$", prev)) or prev.count("(") > prev.count(")")
            is_fragment = (not is_bare_year and " " not in line and len(line) <= 20) or line.startswith("(") or line[:1].islower()
            if not is_bare_year and (prev_dangling or is_fragment):
                merged[-1] = f"{prev} {line}"
                continue
        merged.append(line)
    return merged


_LEADING_YEAR_RE = re.compile(r"^(?:19|20)\d{2}\s+(?=\S)")
_TRAILING_YEAR_RE = re.compile(r"\s+(?:19|20)\d{2}$")


def _strip_year_artifact(line):
    """Strip a bare 4-digit year glued onto the front or back of an
    otherwise-real line, e.g. "2011 Roadmap" or "Roadmap 2011".

    This happens even before any line-merging: pdfplumber's own
    word-to-row grouping (see pdf_layout.py) buckets words that share
    almost the same vertical position into one row/text-run, and on a
    timeline-style resume layout a stray year from a nearby date column
    can end up sharing a row with an unrelated sidebar word - producing
    one already-combined token like "2011 Roadmap" before any Python-side
    line ever gets merged. A full-line year match (see the caller) can't
    catch this because the line isn't *just* a year."""
    stripped = _LEADING_YEAR_RE.sub("", line)
    stripped = _TRAILING_YEAR_RE.sub("", stripped)
    return stripped.strip()


def _parse_skills_and_certs(sections):
    """Handles two common resume shapes:
    - bulleted 'Category: item, item, item' lines (skills), possibly with
      certs folded in as a 'Certifications: X' sub-line;
    - a dedicated CERTIFICATIONS section that's just one plain line per
      certification (no label, no commas to split on)."""
    skill_pool, cert_pool = [], []

    # NOT merged via _merge_wrapped_lines here (unlike certifications
    # below) - a sidebar "KEY SKILLS" list is almost always one complete,
    # independent skill per line (Scrum / Agile / JIRA / ...), and that
    # heuristic's short-fragment rule (needed to rejoin a genuinely
    # wrapped cert title) can't tell "Agile" apart from a wrapped
    # continuation - it was gluing 15+ separate skills into one or two
    # giant blobs.
    for line in sections.get("skills", []):
        line = _clean_bullet(line.strip())
        # A bare 4-digit year is never a real skill - it's a stray
        # timeline date that got column-misclassified into this section.
        # No longer glued onto a neighboring skill name (see
        # _merge_wrapped_lines), but still worth dropping outright rather
        # than letting it occupy one of the top-3 skill slots.
        if not line or re.fullmatch(r"(19|20)\d{2}", line):
            continue
        line = _strip_year_artifact(line)
        if not line:
            continue
        label, sep, rest = line.partition(":")
        is_cert_label = bool(sep) and "certif" in label.lower()
        items_src = rest if sep else line
        parts = [p.strip(" .") for p in _split_outside_parens(items_src) if p.strip(" .")]
        parts = [p for p in parts if len(p) < 60]
        (cert_pool if is_cert_label else skill_pool).extend(parts)

    # Not merged either, for the same reason as skills above - a
    # dedicated certifications section is also one complete item per
    # line (confirmed on a real resume: "Beginners in Business Analyst"
    # / "POPM" were two separate certifications that got glued into one
    # by the same short-fragment heuristic).
    for line in sections.get("certifications", []):
        line = _clean_bullet(line.strip())
        line = _strip_year_artifact(line)
        # A combined "Education & Certifications" heading (see
        # _find_sections) can leave a degree row sitting in this same
        # bucket alongside the real certifications. Only treat it as a
        # degree row to exclude when it's ALSO shaped like one (names an
        # institution or a year) - DEGREE_ROW_RE alone isn't enough,
        # since a genuine certification can start with a degree-looking
        # token too ("MBA, PMI-GAC Certified" is a certification, not an
        # education row, even though it starts with "MBA").
        looks_like_degree_row = (
            DEGREE_ROW_RE.match(line)
            and (INSTITUTION_RE.search(line) or re.search(r"\b(19|20)\d{2}\b", line))
        )
        if line and len(line) < 80 and not looks_like_degree_row:
            cert_pool.append(line)

    top_skills = skill_pool[:3] if skill_pool else ["NA", "NA", "NA"]
    certifications = cert_pool if cert_pool else ["NA"]
    return top_skills, certifications, skill_pool


CORE_SKILL_DOMAINS = [
    "product management", "program management", "project management",
    "product owner", "business analysis", "data science", "data analytics",
    "software engineering", "software development", "quality assurance",
    "devops", "cloud engineering", "cybersecurity", "sales", "marketing",
    "human resources", "finance", "supply chain", "operations",
    "customer success", "ui/ux design", "product design",
]


def _titles_share_a_word(employment):
    """True only if every job title has at least one word in common -
    a loose proxy for 'one consistent specialization throughout'."""
    titles = [j["title"].lower() for j in employment if j.get("title") not in (None, "NA", "")]
    if len(titles) < 2:
        return False
    word_sets = [set(t.split()) - {"and", "the", "of", "a", "an"} for t in titles]
    common = set.intersection(*word_sets)
    return bool(common)


def _guess_core_skill_experience(sections, full_text, employment):
    """Per field-schema.md: only fill this in when the resume states a
    separate figure AND the whole career is one consistent specialization -
    otherwise leave NA rather than guess."""
    if not _titles_share_a_word(employment):
        return "NA"
    summary_text = " ".join(sections.get("summary", [])) or full_text
    domain = next((d for d in CORE_SKILL_DOMAINS if d in summary_text.lower()), None)
    years_m = re.search(r"(\d{1,2}\+?)\s*(?:years?|yrs?)", summary_text, re.IGNORECASE)
    if domain and years_m:
        return f"{years_m.group(1)} Years in {domain.title()}"
    return "NA"


DEGREE_ROW_RE = re.compile(
    r"^\s*(?P<degree>B\.?\s?Tech|B\.?\s?E\.?|B\.?\s?Sc|B\.?\s?Com|BCA|Bachelor(?:'?s)?(?:\s+of\s+[\w\s]+?)?|"
    r"M\.?\s?Tech|M\.?\s?Sc|M\.?\s?Com|MCA|MBA|PGDM|Master(?:'?s)?(?:\s+of\s+[\w\s]+?)?|XII|XI|X)\b",
    re.IGNORECASE,
)
INSTITUTION_RE = re.compile(r"\b(University|College|Institute|School)\b.*$", re.IGNORECASE)


def _classify_degree(token):
    low = token.strip().lower().replace(".", "").replace(" ", "")
    if low in ("xii", "xi", "x"):
        return None
    if low.startswith("m") or "master" in low or "pgdm" in low:
        return "master"
    return "bachelor"


def _row_to_line(degree, rest, fallback_institution=None, fallback_year=None):
    year_m = YEAR_RANGE_RE.search(rest)
    year = year_m.group(0) if year_m else fallback_year
    rest_wo_year = rest[:year_m.start()].strip() if year_m else rest.strip()
    inst_m = INSTITUTION_RE.search(rest_wo_year)
    if inst_m:
        institution = inst_m.group(0).strip()
        discipline = rest_wo_year[:inst_m.start()].strip().lstrip("-:|,").strip() or None
    else:
        institution = fallback_institution
        discipline = rest_wo_year.strip().lstrip("/,-:|").strip() or None
    if discipline:
        discipline = re.sub(r"^(in|of)\s+", "", discipline, flags=re.IGNORECASE)
    degree_text = f"{degree} ({discipline})" if discipline else degree
    parts = [degree_text] + ([institution] if institution else []) + ([year] if year else [])
    return " \u2013 ".join(parts)


def _guess_education(sections):
    raw = [_clean_bullet(l) for l in sections.get("education", []) if _clean_bullet(l)]
    raw = [l for l in raw if not PAGE_NUM_RE.match(l)]
    raw = [l for l in raw if not (l.strip().lower().startswith("degree") and "discipline" in l.lower())]

    # Fallbacks for layouts where degree/institution/year are on separate
    # lines rather than one table row - search the whole section for them.
    fallback_institution = next(
        (l for l in raw if re.search(r"\b(University|College|Institute|School)\b", l, re.IGNORECASE)), None
    )
    fallback_year = next((m.group(0) for l in raw for m in [YEAR_RANGE_RE.search(l)] if m), None)

    bachelor_line, master_line = "NA", "NA"
    for i, line in enumerate(raw):
        m = DEGREE_ROW_RE.match(line)
        if not m:
            continue
        degree = m.group("degree").strip()
        kind = _classify_degree(degree)
        if kind is None:
            continue  # secondary/higher-secondary schooling row (X/XI/XII) - not part of this schema
        # This layout often puts institution on its own line right after
        # the degree line, rather than sharing one line - use THAT
        # specific line as the institution for THIS degree when the
        # degree's own line has nothing inline. A single section-wide
        # fallback_institution reused for every degree row is wrong here
        # (it kept assigning the master's institution to the bachelor's
        # row too, since it's just "whichever institution-looking line
        # came first in the whole section").
        near_institution = fallback_institution
        remainder_text = line[m.end():]
        if not re.search(r"\b(University|College|Institute|School)\b", remainder_text, re.IGNORECASE):
            nxt = raw[i + 1] if i + 1 < len(raw) else ""
            if re.search(r"\b(University|College|Institute|School)\b", nxt, re.IGNORECASE):
                near_institution = nxt
        line_out = _row_to_line(degree, line[m.end():], near_institution, fallback_year)
        if kind == "bachelor" and bachelor_line == "NA":
            bachelor_line = line_out
        elif kind == "master" and master_line == "NA":
            master_line = line_out
    return bachelor_line, master_line


def _guess_overall_experience(sections, full_text):
    for line in sections.get("experience", [])[:3]:
        if DURATION_LINE_RE.match(line.strip()):
            return line.strip()
    m = YEARS_EXPERIENCE_RE.search(full_text)
    if m:
        return f"{m.group(1)} Years"
    return "NA"


TITLE_KEYWORDS = {
    "manager", "owner", "lead", "director", "analyst", "engineer", "consultant",
    "specialist", "intern", "associate", "executive", "coordinator", "architect",
    "scientist", "officer", "president", "head", "developer", "designer",
    "administrator", "representative", "strategist", "vp", "evp", "svp", "avp",
    "founder", "co-founder", "trainee", "supervisor",
}


def _split_title_company(remainder):
    """remainder e.g. 'Product Manager Cloudphysician Healthcare, Bangalore'
    -> title='Product Manager', company='Cloudphysician Healthcare'.

    Prefers splitting on ' - ' when present: on
    'Team Lead, Business Analyst - Cotiviti - Healthcare' this keeps the
    compound title intact (with its own internal comma) and takes just
    the segment right after the first dash as company, discarding a
    further trailing industry tag like '- Healthcare' - a plain
    comma-first split would otherwise truncate the title at "Team Lead"
    and never even look at the text after the comma for a company.
    Only falls back to a keyword search within the comma-prefix when
    there's no dash-separated structure to use instead."""
    dash_parts = [p.strip() for p in remainder.split(" - ") if p.strip()]
    if len(dash_parts) >= 2:
        return (dash_parts[0].strip() or "NA",
                dash_parts[1].strip().lstrip("-:|,").strip() or "NA")

    before_sep = re.split(r"[,|]", remainder, maxsplit=1)[0]
    words = before_sep.split()
    last_kw_idx = None
    for i, w in enumerate(words):
        if w.strip(".,()").lower() in TITLE_KEYWORDS:
            last_kw_idx = i
    if last_kw_idx is None:
        # No recognized title keyword anywhere in this line - there's no
        # reliable basis to guess WHERE a title/company boundary would
        # even be. A line like "Speridian Technologies / Pvt. Ltd."
        # (entirely a company name, no title on this line at all)
        # previously got its first couple of words chopped off and
        # mislabeled as a job title purely by word-count. Treat the
        # whole thing as the company and leave title for manual review
        # instead of guessing wrong.
        return "NA", remainder.strip().lstrip("-:|,").strip() or "NA"

    title = " ".join(words[:last_kw_idx + 1])
    company = " ".join(words[last_kw_idx + 1:])
    if not company:
        # The keyword was the last word before the separator - the
        # company is likely on the OTHER side of that comma/pipe rather
        # than nonexistent.
        company = remainder[len(before_sep):].lstrip(",|").strip()
    return title.strip() or "NA", company.strip().lstrip("-:|,").strip() or "NA"


_INCOMPLETE_DATE_RE = re.compile(
    rf"^(?:{_MONTH_RE}[\s./-]*)?\d{{4}}\s*[-\u2013\u2014]\s*{_MONTH_RE}$"
)
_BARE_YEAR_LINE_RE = re.compile(r"^(19|20)\d{2}$")


def _repair_split_date_ranges(lines):
    """Rejoin a date range whose closing year wrapped onto its own line
    a few lines later - e.g. "Feb 2022 - Mar" / "Product Owner" / "2024"
    / "Cerner". A narrow timeline "date badge" column can wrap a range
    across two visual lines, and that second line's row can coincide
    with an unrelated column's row-bucket, landing it several lines away
    in the flattened text instead of right after the first line.
    DATE_RANGE_RE can't match either half alone, so every downstream
    anchor-based parser silently skips the whole job unless this runs
    first."""
    lines = list(lines)
    out = []
    i = 0
    while i < len(lines):
        if _INCOMPLETE_DATE_RE.match(lines[i].strip()):
            for j in range(i + 1, min(i + 4, len(lines))):
                if _BARE_YEAR_LINE_RE.match(lines[j].strip()):
                    out.append(f"{lines[i].strip()} {lines[j].strip()}")
                    out.extend(lines[i + 1:j])
                    i = j + 1
                    break
            else:
                out.append(lines[i])
                i += 1
        else:
            out.append(lines[i])
            i += 1
    return out


def _collect_experience_raw(sections):
    """Real fix for a specific, common failure mode: a resume whose
    underlying PDF only has ONE "WORK EXPERIENCE" header covering
    multiple jobs, where an unrelated header (e.g. a sidebar "LANGUAGES"
    list) reappears mid-stream - because the PDF's text order doesn't
    match visual reading order - and steals _find_sections' "current"
    pointer before all the jobs are captured. The remaining job blocks
    then get silently bucketed under whatever section that header opened.

    A genuine languages/skills/contact list should never contain a real
    date range ("Feb 2022 - Mar 2024") - if one shows up in a section
    other than "experience", it's recoverable overflow, not a false
    positive. Dict insertion order matches document order here (each
    section is inserted in the order its header was first seen), so
    concatenating in that order preserves chronological order too."""
    raw = list(sections.get("experience", []))
    for name, lines in sections.items():
        if name in ("experience", "education", "summary"):
            continue
        anchor_idx = next((i for i, l in enumerate(lines) if DATE_RANGE_RE.search(l)), None)
        if anchor_idx is not None:
            # Only take content from that section's OWN first date anchor
            # onward - anything before it is leftover unrelated content
            # (sidebar/contact/language-list drips) that doesn't belong
            # to employment at all, and would otherwise get attributed to
            # the PRECEDING job's bullets (since it lands right before
            # the next job's anchor in the combined raw list).
            raw.extend(lines[anchor_idx:])
    return raw


_INLINE_BULLET_SPLIT_RE = re.compile(r"\s*(?:[\u2022\u25cf\u25aa\u2023]|\+(?=[A-Za-z]))\s*")


def _reconstruct_bullets(raw_lines):
    """Rejoin wrapped PDF lines into the actual logical bullets.

    A raw line list here is just physical PDF text lines, already split
    wherever the page happened to wrap - not at logical bullet
    boundaries. The real bullet separator ("•", or "+" on at least one
    resume's export) shows up INLINE within a line wherever one bullet
    ends and the next begins ("...compliance teams. • Owned and
    prioritized..."), only ever at the very start of a line if that
    bullet happens to start a fresh line. Treating each physical line as
    its own bullet (the previous behavior) both splits one real bullet
    across several lines AND merges the tail of one bullet with the
    start of the next onto a single line. Rejoin everything into one
    string first, then split on the real inline marker."""
    joined = " ".join(l.strip() for l in raw_lines if l.strip())
    parts = _INLINE_BULLET_SPLIT_RE.split(joined)
    return [p.strip() for p in parts if p.strip()]


def _guess_employment(sections):
    """Split the experience section into jobs using date-range lines as
    anchors. Title/company/date are usually one combined row - split that
    row using known job-title words; fall back to the preceding non-bullet
    line(s) if the date's own row has nothing else on it."""
    raw = _collect_experience_raw(sections)
    if not raw:
        return []
    anchors = [i for i, line in enumerate(raw) if DATE_RANGE_RE.search(line)]
    if not anchors:
        return []

    jobs = []
    for idx, start in enumerate(anchors):
        end = anchors[idx + 1] if idx + 1 < len(anchors) else len(raw)
        date_line = raw[start]
        m = DATE_RANGE_RE.search(date_line)
        duration = m.group(0) if m else "NA"
        remainder = DATE_RANGE_RE.sub("", date_line).strip(" |,-")

        if remainder:
            title, company = _split_title_company(remainder)
            fwd_consumed = 0
        else:
            # Try forward first: date-badge-then-title-then-company (date
            # line standalone, title/company right after it, before the
            # first real bullet) is the more constrained, safer guess -
            # it's bounded within this job's own slice (start+1 to the
            # next anchor) and can never reach into a previous job's
            # tail or unrelated overflow content the way scanning
            # backward can. Confirmed necessary on a real resume: a
            # backward scan reached past a short overflow line with no
            # trailing punctuation into unrelated recovered content
            # (leftover summary text + a sidebar language list) and
            # confidently produced a wrong title/company instead of
            # falling through to this.
            fwd, k = [], start + 1
            while k < end and len(fwd) < 2:
                line = raw[k].strip()
                if not line or BULLET_RE.match(raw[k]) or DATE_RANGE_RE.search(line) or len(line) >= 100:
                    break
                fwd.append(line)
                k += 1

            if len(fwd) >= 2:
                title, company, fwd_consumed = fwd[0], fwd[1], 2
            elif len(fwd) == 1:
                title, company, fwd_consumed = fwd[0], "NA", 1
            else:
                fwd_consumed = 0
                lower_bound = anchors[idx - 1] + 1 if idx > 0 else 0
                pre_lines = []
                j = start - 1
                while j >= lower_bound and len(pre_lines) < 6:
                    line = raw[j].strip()
                    # a real title/company line is short and doesn't end mid-sentence;
                    # a wrapped bullet continuation (no bullet char of its own) does -
                    # stop there so we don't pull prose into the title/company guess.
                    if not line:
                        j -= 1
                        continue
                    if BULLET_RE.match(raw[j]) or line.rstrip().endswith((".", ",", ":", ";")):
                        break
                    if len(line) < 100:
                        pre_lines.insert(0, line)
                    j -= 1
                company, title = "NA", "NA"
                sep_idx = next((k for k in range(len(pre_lines) - 1, -1, -1) if "," in pre_lines[k] or "|" in pre_lines[k]), None)
                if sep_idx is not None:
                    company = re.split(r"[,|]", pre_lines[sep_idx])[0].strip() or "NA"
                    title = pre_lines[sep_idx - 1] if sep_idx - 1 >= 0 else "NA"
                elif pre_lines:
                    company = pre_lines[-1]
                    title = pre_lines[-2] if len(pre_lines) >= 2 else "NA"

        bullet_lines = []
        for l in raw[start + 1 + fwd_consumed:end]:
            cl = _clean_bullet(l)
            if cl and not DATE_RANGE_RE.search(cl) and not PAGE_NUM_RE.match(cl):
                bullet_lines.append(cl)
        bullets = _reconstruct_bullets(bullet_lines)
        jobs.append({
            "duration": duration,
            "company": company,
            "title": title,
            "technology": "NA",
            "role": bullets[:8] if bullets else ["NA"],
        })
    return jobs


_PROJECT_DURATION_RE = re.compile(r"^\d+\s*(Days?|Months?|Years?|Weeks?)$", re.IGNORECASE)


def _guess_projects(sections):
    """Best-effort single-project extraction: drop a bare duration badge
    ("122 Days"), take the next short non-sentence-ending line(s) as the
    title, and the remainder as the description. Doesn't attempt to
    split multiple distinct projects apart - a real but lower-priority
    gap, left as a manual-review item rather than guessed at."""
    raw = [_clean_bullet(l) for l in sections.get("projects", []) if _clean_bullet(l)]
    raw = [l for l in raw if not _PROJECT_DURATION_RE.match(l.strip())]
    if not raw:
        return []
    title_lines, i = [], 0
    while i < len(raw) and len(title_lines) < 3:
        line = raw[i]
        if line.rstrip().endswith((".", ",", ":", ";")) or len(line) > 80:
            break
        title_lines.append(line)
        i += 1
    if not title_lines:
        title_lines, i = [raw[0]], 1
    title = " ".join(title_lines).strip()
    description = " ".join(raw[i:]).strip()
    return [{"title": title, "description": description}] if title else []


def _truncate_before_overflow(lines):
    """A section other than "experience" that contains a real date range
    is the same overflow _collect_experience_raw recovers into
    employment (see there for why) - it was copied out for employment,
    but never removed from its original section, so anything using that
    section directly (like the appended resume's own Languages line)
    would otherwise still show it verbatim. Cut the section off at the
    first such date range, and also drop any line long enough that it's
    clearly prose rather than a language name (a couple of trailing
    summary-paragraph words can land just before the date line too, as
    part of the same overflow, without containing a date themselves)."""
    out = []
    for l in lines:
        if DATE_RANGE_RE.search(l):
            break
        if len(l) > 25:
            continue
        out.append(l)
    return out


def extract_candidate_from_resume(resume_text):
    """
    Best-effort local extraction of the candidate.json schema from raw
    resume text (no LLM). Returns a dict following field-schema.md, with
    unresolved fields left as "NA"/"TBC". Always review before generating.
    """
    lines = _repair_split_date_ranges(_lines(resume_text))
    sections = _strip_pii(_find_sections(lines))
    first_name, last_name = _split_name(lines)

    bachelor, master = _guess_education(sections)
    master_duration_m = YEAR_RANGE_RE.search(master) if master != "NA" else None
    education_master_duration = master_duration_m.group(0) if master_duration_m else "NA"
    employment = _guess_employment(sections)
    top_skills, certifications, skill_pool = _parse_skills_and_certs(sections)
    # skills beyond the top 3, as a literal (not invented) "other qualifications" line
    extra_skills = skill_pool[3:6]
    other_qualifications = ", ".join(extra_skills) if extra_skills else "NA"

    resume_block = {
        "name": f"{first_name} {last_name}".strip(),
        "summary": [_clean_bullet(l) for l in sections.get("summary", []) if _clean_bullet(l)],
        "work_experience": [
            {"title": j["title"], "company": j["company"], "duration": j["duration"], "bullets": j["role"]}
            for j in employment
        ],
        "education": [_clean_bullet(l) for l in sections.get("education", []) if _clean_bullet(l)],
        "skills": top_skills,
        "projects": _guess_projects(sections),
        "certificates": certifications,
        "languages": [
            _clean_bullet(l) for l in _truncate_before_overflow(sections.get("languages", []))
            if _clean_bullet(l)
        ],
    }

    candidate = {
        "first_name": first_name,
        "last_name": last_name,
        "notice_period": "TBC",
        "internal_external": "TBC",
        "interview_availability": "TBC",
        "start_availability": "TBC",
        "worked_ford_direct": "No",
        "worked_ford_agency": "No",
        "cdsid": "NA",
        "supervision_duration": "NA",
        "exit_reason": "NA",
        "overall_it_experience": _guess_overall_experience(sections, resume_text),
        # only filled when every job title shares a word (proxy for "one
        # consistent specialization") and the summary states years+domain;
        # otherwise left NA per field-schema.md rather than guessing.
        "experience_core_skill": _guess_core_skill_experience(sections, resume_text, employment),
        "other_qualifications": other_qualifications,
        "hackerrank_score": "NA",
        "certifications": certifications,
        "top_skills": top_skills,
        "education_bachelor": bachelor,
        "education_master": master,
        "education_master_duration": education_master_duration,
        "education_bachelor_detail": bachelor,
        "employment": employment,
        "resume": resume_block,
    }
    return candidate
