#!/usr/bin/env python3
"""
cv_builder.py - Fill the Ford/Magnit "CV Template" (fixed-layout .docx) from
structured candidate data, and optionally append the candidate's resume as
editable Word text after it (contact details, photo, and logos excluded —
see build_resume_section_xml() below).

This is the standalone version of the script used inside the Claude Skill
of the same name — adapted so it has no dependency on that sandboxed
environment (no /mnt/skills path, no external merge_runs.py call) and can
run anywhere: a plain CLI, or imported by streamlit_app.py.

The bundled assets/Candidate_Template.docx ships pre-merged (its runs
already consolidated) precisely so this module doesn't need Word's
run-fragmentation cleaned up at runtime — if you ever replace that asset
with a fresh export from Word, run it through a run-merging pass first
(any DOCX "merge adjacent runs" utility) or the text anchors below may
stop matching.

This module is anchor-based: it relies on the exact paragraph/text
structure of assets/Candidate_Template.docx. If Ford/Magnit ships a new
template revision (relabeled fields, added rows, reordered sections), the
anchors will report `[MISSING]` in build_docx()'s log — diff the new
template against the bundled one (render both to PDF and compare) before
assuming the anchors still hold.

CLI usage:
    python cv_builder.py --data candidate.json --output output.docx
                          [--template /path/to/Candidate_Template.docx]
                          [--photo /path/to/photo.jpg]

Library usage (see streamlit_app.py):
    from cv_builder import build_docx
    docx_bytes = build_docx(candidate_dict, photo_bytes=None, template_path=None)

See field-schema.md for the candidate.json schema (including the "resume"
block) and the NA/TBC conventions, and example_candidate.json for a worked
example.
"""
import argparse
import io
import json
import re
import zipfile
from pathlib import Path

from PIL import Image

RUN_PLAIN = (
    '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
    '<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{}</w:t></w:r>'
)

NAVY = "1F487C"       # label / accent blue used throughout the template
BANNER_NAVY = "17365D"  # banner fill used by "CV TEMPLATE" / "EDUCATION HISTORY" bars


def esc(s: str) -> str:
    """Escape text for safe insertion into w:t nodes."""
    if s is None:
        s = ""
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def replace_all(data: str, old: str, new: str, label: str, expected=2) -> str:
    count = data.count(old)
    if count == 0:
        print(f"  [MISSING] {label}")
        return data
    if count != expected:
        print(f"  [WARN] {label}: expected {expected} occurrences, found {count}")
    return data.replace(old, new)


# ---------------------------------------------------------------------------
# Section 1: simple label -> value fields (anchored on exact template text)
# ---------------------------------------------------------------------------

def fill_simple_fields(data: str, c: dict, warnings: list | None = None) -> str:
    v = lambda k, default="": esc(c.get(k, default))

    # First Name
    data = replace_all(
        data,
        '<w:t>Name:</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:b/>'
        '<w:color w:val="1F487C"/><w:spacing w:val="1"/><w:sz w:val="24"/></w:rPr>'
        '<w:t xml:space="preserve"> </w:t></w:r></w:p><w:p w14:paraId="1D2EFB46"',
        '<w:t>Name:</w:t></w:r>' + RUN_PLAIN.format(" " + v("first_name")) +
        '</w:p><w:p w14:paraId="1D2EFB46"',
        "First Name",
    )

    # Candidate Last Name -> put value right after the label, then linebreak,
    # then "(As per Passport/Aadhar)" on its own line (per Rajkumar's Sep-2026 correction)
    data = replace_all(
        data,
        '<w:t>Name</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="24"/></w:rPr>'
        '<w:t>:</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:b/><w:color w:val="1F487C"/>'
        '<w:sz w:val="24"/></w:rPr><w:t>(As per Passport/Aadhar)</w:t></w:r></w:p>',
        '<w:t>Name</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="24"/></w:rPr>'
        '<w:t xml:space="preserve">: ' + v("last_name") + '</w:t></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:b/><w:color w:val="1F487C"/>'
        '<w:sz w:val="24"/></w:rPr><w:br/><w:t>(As per Passport/Aadhar)</w:t></w:r></w:p>',
        "Candidate Last Name",
    )

    # Notice Period / Internal-External / Interview Date / Start Date
    # NOTE: keep these values SHORT (e.g. "TBC", "Immediate", "7 Days") --
    # the text box has a right-indent reserved for the candidate photo, so
    # long values wrap onto a second line. "TBC" is the safe default.
    data = replace_all(
        data,
        '<w:t xml:space="preserve">Notice Period of the Candidate: </w:t></w:r></w:p>',
        '<w:t xml:space="preserve">Notice Period of the Candidate: </w:t></w:r>' +
        RUN_PLAIN.format(v("notice_period", "TBC")) + '</w:p>',
        "Notice Period",
    )
    data = replace_all(
        data,
        '<w:t xml:space="preserve"> </w:t></w:r></w:p><w:p w14:paraId="2AD715B9"',
        '<w:t xml:space="preserve"> </w:t></w:r>' +
        RUN_PLAIN.format(v("internal_external", "TBC")) + '</w:p><w:p w14:paraId="2AD715B9"',
        "Internal or External Candidate",
    )
    data = replace_all(
        data,
        '<w:t xml:space="preserve"> </w:t></w:r></w:p><w:p w14:paraId="5BC4C4DB"',
        '<w:t xml:space="preserve"> </w:t></w:r>' +
        RUN_PLAIN.format(v("interview_availability", "TBC")) + '</w:p><w:p w14:paraId="5BC4C4DB"',
        "Interview Availability Date",
    )
    data = replace_all(
        data,
        '<w:t xml:space="preserve">Start Availability Date: </w:t></w:r></w:p>',
        '<w:t xml:space="preserve">Start Availability Date: </w:t></w:r>' +
        RUN_PLAIN.format(v("start_availability", "TBC")) + '</w:p>',
        "Start Availability Date",
    )

    # Has worked directly for Ford before? / as Agency worker?
    data = replace_all(
        data,
        '<w:t>(Yes/No) \u2013</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        '<w:b/><w:color w:val="1F487C"/><w:spacing w:val="-1"/><w:sz w:val="24"/></w:rPr>'
        '<w:t xml:space="preserve"> </w:t></w:r></w:p><w:p w14:paraId="0327E602"',
        '<w:t>(Yes/No) \u2013</w:t></w:r>' + RUN_PLAIN.format(" " + v("worked_ford_direct", "No")) +
        '</w:p><w:p w14:paraId="0327E602"',
        "Worked directly for Ford before",
    )
    data = replace_all(
        data,
        '<w:t>(Yes/No) \u2013</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        '<w:b/><w:color w:val="1F487C"/><w:spacing w:val="-3"/><w:sz w:val="24"/></w:rPr>'
        '<w:t xml:space="preserve"> </w:t></w:r></w:p><w:p w14:paraId="33C5FAFA"',
        '<w:t>(Yes/No) \u2013</w:t></w:r>' + RUN_PLAIN.format(" " + v("worked_ford_agency", "No")) +
        '</w:p><w:p w14:paraId="33C5FAFA"',
        "Worked for Ford as Agency worker before",
    )

    # CDSID / Supervision-Duration / Exit reason
    data = replace_all(
        data,
        '<w:t>CDSID:</w:t></w:r></w:p>',
        '<w:t>CDSID:</w:t></w:r>' + RUN_PLAIN.format(" " + v("cdsid", "NA")) + '</w:p>',
        "CDSID",
    )
    data = replace_all(
        data,
        '<w:t xml:space="preserve"> Project:</w:t></w:r></w:p><w:p w14:paraId="55E47A98"',
        '<w:t xml:space="preserve"> Project:</w:t></w:r>' +
        RUN_PLAIN.format(" " + v("supervision_duration", "NA")) + '</w:p><w:p w14:paraId="55E47A98"',
        "Supervision / Duration of Project",
    )
    data = replace_all(
        data,
        '<w:t>reason:</w:t></w:r></w:p>',
        '<w:t>reason:</w:t></w:r>' + RUN_PLAIN.format(" " + v("exit_reason", "NA")) + '</w:p>',
        "Exit reason",
    )

    # Overall IT Experience / Experience in Core Skill
    data = replace_all(
        data,
        '<w:t xml:space="preserve">Overall IT Experience: </w:t></w:r></w:p>',
        '<w:t xml:space="preserve">Overall IT Experience: </w:t></w:r>' +
        RUN_PLAIN.format(v("overall_it_experience", "NA")) + '</w:p>',
        "Overall IT Experience",
    )
    data = replace_all(
        data,
        '<w:t xml:space="preserve"> </w:t></w:r></w:p><w:p w14:paraId="3D8C29C2"',
        '<w:t xml:space="preserve"> </w:t></w:r>' +
        RUN_PLAIN.format(v("experience_core_skill", "NA")) + '</w:p><w:p w14:paraId="3D8C29C2"',
        "Experience in Core Skill",
    )

    # Other Qualifications/Certifications + Hacker Rank score (same paragraph)
    # NOTE: the template's own label spells this "HackerRank" (one word) as
    # of the Sep-2026 template revision — was "Hacker Rank" (two words) in
    # the original template.
    data = replace_all(
        data,
        '<w:t xml:space="preserve">any: HackerRank score: </w:t></w:r></w:p>',
        '<w:t xml:space="preserve">any: ' + v("other_qualifications", "NA") +
        '. HackerRank score: </w:t></w:r>' + RUN_PLAIN.format(v("hackerrank_score", "NA")) + '</w:p>',
        "Other Qualifications / HackerRank score",
    )

    # Certification list (template ships ONE empty bullet paragraph; we
    # replace it with N numbered paragraphs — one per certification —
    # using the decimal numbering definition (numId=3) added to
    # assets/Candidate_Template.docx specifically for this. Accepts either
    # a list of certifications or a single string (semicolon/comma-split
    # as a fallback so older candidate.json files still work).
    certs = c.get("certifications", "NA")
    if isinstance(certs, str):
        if certs.strip().upper() in ("NA", ""):
            certs = ["NA"]
        else:
            certs = [x.strip() for x in re.split(r"[;\n]", certs) if x.strip()]
    # The Certification box on page 1 is a fixed-height text box (see
    # assets/Candidate_Template.docx's "Textbox 10" - <a:noAutofit/>, no
    # dynamic resize), shared with several other fields above it. In
    # practice it only has safe room for ~3 certification lines before
    # text visually overflows past the drawn border. Cap what goes in the
    # box; the full list is unaffected elsewhere (candidate["resume"]["certificates"]
    # still carries all of them into the appended resume section).
    MAX_CERTS_IN_BOX = 3
    if len(certs) > MAX_CERTS_IN_BOX and certs != ["NA"]:
        if warnings is not None:
            warnings.append(
                f"{len(certs)} certifications were supplied but the template's "
                f"Certification box only has safe room for {MAX_CERTS_IN_BOX} lines at "
                "normal size — using a smaller font to fit more instead, since the "
                "appended resume section (the usual fallback for the full list) isn't "
                "always included."
            )
    # A smaller font (9pt vs the box's normal 12pt) buys roughly 30% more
    # lines before text overflows the box's fixed height - used whenever
    # there are more than MAX_CERTS_IN_BOX certifications, since at 12pt
    # a 4th line is already guaranteed to overflow.
    cert_sz = "18" if len(certs) > MAX_CERTS_IN_BOX else "24"
    CERT_ITEM = (
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr>'
        '<w:tabs><w:tab w:val="left" w:pos="719"/></w:tabs><w:spacing w:before="1"/>'
        f'<w:ind w:hanging="360"/><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="{cert_sz}"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="{cert_sz}"/></w:rPr><w:t xml:space="preserve">{{}}</w:t></w:r></w:p>'
    )
    cert_xml = "".join(CERT_ITEM.format(esc(cert)) for cert in certs)
    data = replace_all(
        data,
        '<w:p w14:paraId="1CA6BCE7" w14:textId="3EC77E0E" w:rsidR="002D7672" w:rsidRDefault="002D7672">'
        '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:tabs>'
        '<w:tab w:val="left" w:pos="719"/></w:tabs><w:spacing w:before="1"/><w:ind w:hanging="360"/>'
        '<w:rPr><w:rFonts w:ascii="Calibri"/><w:b/><w:sz w:val="24"/></w:rPr></w:pPr></w:p>',
        cert_xml,
        "Certification list",
    )

    # Top 3 relevant skills
    top = c.get("top_skills", ["NA", "NA", "NA"])
    top = (top + ["NA", "NA", "NA"])[:3]
    for i, label in enumerate(["1.", "2.", "3."]):
        data = replace_all(
            data,
            f'<w:t>{label}</w:t></w:r></w:p>',
            f'<w:t>{label}</w:t></w:r>' + RUN_PLAIN.format(" " + esc(top[i])) + '</w:p>',
            f"Top skill #{i+1}",
        )

    return data


# ---------------------------------------------------------------------------
# Section 2: Education History table
# ---------------------------------------------------------------------------

def fill_education(data: str, c: dict) -> str:
    bachelor = esc(c.get("education_bachelor", "NA"))
    master = esc(c.get("education_master", "NA"))
    duration = esc(c.get("education_master_duration", "NA"))
    final_detail = esc(c.get("education_bachelor_detail", bachelor))

    data = replace_all(
        data,
        '<w:p w14:paraId="1496C6CD" w14:textId="6A14F32E" w:rsidR="002D7672" w:rsidRDefault="002D7672">'
        '<w:pPr><w:pStyle w:val="TableParagraph"/><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:pPr></w:p>',
        '<w:p w14:paraId="1496C6CD" w14:textId="6A14F32E" w:rsidR="002D7672" w:rsidRDefault="002D7672">'
        '<w:pPr><w:pStyle w:val="TableParagraph"/><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>{bachelor}</w:t></w:r></w:p>',
        "Education: Bachelor",
        expected=1,
    )
    data = replace_all(
        data,
        '<w:t>Master</w:t></w:r></w:p>',
        f'<w:t>Master</w:t></w:r><w:r><w:rPr><w:sz w:val="24"/></w:rPr>'
        f'<w:t xml:space="preserve">: {master}</w:t></w:r></w:p>',
        "Education: Master",
        expected=1,
    )
    data = replace_all(
        data,
        '<w:t>MM.YYYY)</w:t></w:r><w:r><w:rPr><w:b/><w:color w:val="1F487C"/><w:spacing w:val="-2"/>'
        '<w:sz w:val="24"/></w:rPr><w:t>\u2013</w:t></w:r></w:p></w:tc></w:tr>',
        f'<w:t>MM.YYYY)</w:t></w:r><w:r><w:rPr><w:sz w:val="24"/></w:rPr>'
        f'<w:t xml:space="preserve"> \u2013 {duration}</w:t></w:r></w:p></w:tc></w:tr>',
        "Education: Duration of study",
        expected=1,
    )
    data = replace_all(
        data,
        '<w:r><w:rPr><w:b/><w:color w:val="1F487C"/><w:spacing w:val="-10"/><w:sz w:val="24"/></w:rPr>'
        '<w:t>\u2013</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
        f'<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve"> {final_detail}</w:t></w:r>'
        '</w:p></w:tc></w:tr></w:tbl>',
        "Education: final degree/university name",
        expected=1,
    )
    return data


# ---------------------------------------------------------------------------
# Section 3: Employment History (dynamic N entries)
# ---------------------------------------------------------------------------

def employment_block_xml(entry: dict, is_first: bool) -> str:
    dur = esc(entry.get("duration", "NA"))
    company = esc(entry.get("company", "NA"))
    title = esc(entry.get("title", "NA"))
    tech = esc(entry.get("technology", "NA"))
    role = entry.get("role", "NA")

    lead_spacing = '<w:spacing w:before="223" w:line="259" w:lineRule="auto"/>' if is_first else \
                   '<w:spacing w:before="223" w:line="259" w:lineRule="auto"/>'

    p = lambda label, value, extra_ppr="" : (
        f'<w:p><w:pPr>{extra_ppr}<w:ind w:left="514"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        f'<w:sz w:val="24"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="{NAVY}"/>'
        f'<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{label} \u2013 </w:t></w:r>'
        + RUN_PLAIN.format(value) + '</w:p>'
    )

    out = []
    out.append(
        f'<w:p><w:pPr>{lead_spacing}<w:ind w:left="514" w:right="6124"/>'
        f'<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:color w:val="{NAVY}"/>'
        f'<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Start Date/End Date/Duration \u2013 </w:t></w:r>'
        + RUN_PLAIN.format(dur) + '</w:p>'
    )
    out.append(p("Company name", company))
    out.append(p("Job title", title))
    out.append(
        '<w:p><w:pPr><w:spacing w:before="24" w:line="256" w:lineRule="auto"/>'
        '<w:ind w:left="514" w:right="186"/><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:b/><w:color w:val="{NAVY}"/><w:sz w:val="24"/></w:rPr>'
        '<w:t xml:space="preserve">Technology handled in the Project: </w:t></w:r>'
        + RUN_PLAIN.format(tech) + '</w:p>'
    )

    # "Role of the candidate in the Project" — accepts either a plain
    # string (one dense paragraph) or a list of short bullet points.
    # Prefer a list for anything non-trivial: recruiters scan this table,
    # and a wall of semicolon-joined text is much slower to read than
    # 4-6 short bullets. See references/field-schema.md.
    out.append(
        '<w:p><w:pPr><w:ind w:left="514"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:b/><w:color w:val="{NAVY}"/><w:sz w:val="24"/></w:rPr>'
        '<w:t xml:space="preserve">Role of the candidate in the Project \u2013</w:t></w:r></w:p>'
    )
    if isinstance(role, list):
        for line in role:
            out.append(
                '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
                '<w:ind w:left="874" w:hanging="360"/><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{esc(line)}</w:t></w:r></w:p>'
            )
    else:
        out.append(
            '<w:p><w:pPr><w:ind w:left="514"/></w:pPr>' + RUN_PLAIN.format(esc(role)) + '</w:p>'
        )
    return "".join(out)


def rebuild_employment_section(data: str, employment: list) -> str:
    # Anchor on the known paraId of the template's FIRST employment block
    # (see assets/Candidate_Template.docx) rather than a generic backward
    # tag search — an unrelated bare "<w:p>" spacer paragraph elsewhere in
    # the section can otherwise cause the boundary to land in the wrong
    # place and leave earlier blank placeholder blocks un-replaced.
    paraid_anchor = 'w14:paraId="36314C86"'
    anchor_idx = data.find(paraid_anchor)
    if anchor_idx == -1:
        print("  [MISSING] Employment History section start anchor (paraId 36314C86)")
        return data
    p_start = data.rfind("<w:p", 0, anchor_idx)
    if p_start == -1:
        print("  [MISSING] Could not locate start of first employment paragraph")
        return data

    end_idx = data.find("<w:sectPr", p_start)
    if end_idx == -1:
        print("  [MISSING] trailing sectPr after Employment History section")
        return data

    new_blocks = "".join(
        employment_block_xml(e, i == 0) for i, e in enumerate(employment)
    )
    return data[:p_start] + new_blocks + data[end_idx:]


# ---------------------------------------------------------------------------
# Section 4: Append the resume as EDITABLE Word text (not an image)
# ---------------------------------------------------------------------------
#
# Standing instruction from Rajkumar (Sep 2026, supersedes the earlier
# "insert as-is / don't reformat" note): the appended resume must be
# editable text, not a picture of the original file, and must NOT include:
#   - the candidate's contact details (email, mobile number, LinkedIn,
#     GitHub, or any other personal link) — TEKsystems/Magnit manage
#     candidate contact, so this doesn't go to the client
#   - the candidate's photo
#   - any organisation/company logo
# The candidate's NAME is fine to keep (it's not a contact method), and
# the resume's own content (roles, dates, bullets, education, skills) is
# reproduced faithfully — condensed lightly if needed, but not rewritten
# in meaning, same as the CV-template fields above.
#
# Formatting is deliberately plain (bold black headers, thin rule, no
# colour fill) rather than the earlier Ford-navy-banner treatment — the
# goal here is a clean, editable, re-flowing document, not a rebrand.

def section_heading_xml(text: str) -> str:
    return (
        '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>'
        '<w:spacing w:before="240" w:after="80"/>'
        '<w:rPr><w:rFonts w:ascii="Calibri"/><w:b/><w:sz w:val="24"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:b/><w:sz w:val="24"/></w:rPr><w:t>{esc(text)}</w:t></w:r></w:p>'
    )


def resume_bullet_xml(text: str) -> str:
    return (
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
        '<w:ind w:left="360" w:hanging="360"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr><w:t>{esc(text)}</w:t></w:r></w:p>'
    )


def resume_plain_xml(text: str, bold=False) -> str:
    b = "<w:b/>" if bold else ""
    return (
        f'<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Calibri"/>{b}<w:sz w:val="22"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Calibri"/>{b}<w:sz w:val="22"/></w:rPr><w:t>{esc(text)}</w:t></w:r></w:p>'
    )


def build_resume_section_xml(resume: dict) -> str:
    """resume schema (see references/field-schema.md): name, summary,
    work_experience, education, skills, projects, certificates, languages.
    Deliberately NO "contact"/"email"/"phone"/"linkedin"/"github" key —
    never render one even if a caller's JSON includes it, since this
    section must not carry the candidate's personal contact details or
    any photo/logo."""
    out = [section_heading_xml("ORIGINAL RESUME")]

    if resume.get("name"):
        out.append(resume_plain_xml(resume["name"], bold=True))

    if resume.get("summary"):
        out.append(section_heading_xml("CAREER SUMMARY"))
        summary = resume["summary"]
        if isinstance(summary, str):
            out.append(resume_plain_xml(summary))
        else:
            for line in summary:
                out.append(resume_bullet_xml(line))

    if resume.get("work_experience"):
        out.append(section_heading_xml("WORK EXPERIENCE"))
        for job in resume["work_experience"]:
            title_line = " / ".join(x for x in [job.get("title"), job.get("company")] if x)
            out.append(resume_plain_xml(title_line, bold=True))
            if job.get("duration"):
                out.append(resume_plain_xml(job["duration"]))
            for bullet in job.get("bullets", []):
                out.append(resume_bullet_xml(bullet))

    if resume.get("education"):
        out.append(section_heading_xml("EDUCATION"))
        for edu in resume["education"]:
            out.append(resume_plain_xml(edu, bold=True))

    if resume.get("skills"):
        out.append(section_heading_xml("SKILLS"))
        out.append(resume_plain_xml(", ".join(resume["skills"]) + "."))

    if resume.get("projects"):
        out.append(section_heading_xml("PROJECTS"))
        for proj in resume["projects"]:
            out.append(resume_plain_xml(proj.get("title", ""), bold=True))
            if proj.get("description"):
                out.append(resume_plain_xml(proj["description"]))

    if resume.get("certificates"):
        out.append(section_heading_xml("CERTIFICATES & COURSES"))
        for cert in resume["certificates"]:
            out.append(resume_plain_xml(cert))

    if resume.get("languages"):
        out.append(section_heading_xml("LANGUAGES"))
        out.append(resume_plain_xml(", ".join(resume["languages"]) + "."))

    return "".join(out)


def append_resume_text(data: str, resume: dict) -> str:
    end_body = data.rfind("</w:body>")
    sect_start = data.rfind("<w:sectPr", 0, end_body)
    sect_end = data.find("</w:sectPr>", sect_start) + len("</w:sectPr>")
    final_sectPr = data[sect_start:sect_end]

    resume_xml = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + build_resume_section_xml(resume)
    return data[:sect_start] + resume_xml + final_sectPr + data[sect_end:]


# ---------------------------------------------------------------------------
# Section 5: photo insertion (CV TEMPLATE's own photo field only — see note)
# ---------------------------------------------------------------------------
#
# This fills the template's own required candidate-photo field (page 1).
# It is UNRELATED to the resume section above, which never includes a
# photo. Only call this if Rajkumar has explicitly supplied a candidate
# photo for that field — don't source one from the resume automatically
# (per his Sep-2026 instruction: no candidate photo in the submission
# unless he's explicitly provided one for the template's photo field).
#
# As of the Sep-2026 template revision, the photo field is an EMPTY white
# rectangle (Ford/Magnit removed the old sample placeholder photo) — there
# is no existing picture to overwrite. This inserts a new picture into the
# document on top of that rectangle instead. The anchor below is the exact
# XML of that rectangle shape in assets/Candidate_Template.docx; if a
# future template revision changes the photo field's shape/position again,
# this anchor will fail to match — see PHOTO_BOX_ANCHOR's [MISSING] warning
# and re-derive it from the new template the same way (search the
# unpacked document.xml for the shape whose <a:off>/<a:ext> matches the
# empty box seen in the rendered PDF).

PHOTO_BOX_ANCHOR = (
    '<wps:wsp><wps:cNvPr id="7" name="Graphic 7"/><wps:cNvSpPr/><wps:spPr>'
    '<a:xfrm><a:off x="4724273" y="321818"/><a:ext cx="1428750" cy="1689100"/></a:xfrm>'
    '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>'
    '<a:pathLst><a:path w="1428750" h="1689100"><a:moveTo><a:pt x="1428750" y="0"/></a:moveTo>'
    '<a:lnTo><a:pt x="0" y="0"/></a:lnTo><a:lnTo><a:pt x="0" y="1689100"/></a:lnTo>'
    '<a:lnTo><a:pt x="1428750" y="1689100"/></a:lnTo><a:lnTo><a:pt x="1428750" y="0"/></a:lnTo>'
    '<a:close/></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
    '</wps:spPr><wps:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0">'
    '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:noAutofit/></wps:bodyPr></wps:wsp>'
)


def insert_photo(data: str, files: dict, photo_bytes: bytes) -> str:
    """files is a {archive_path: bytes} dict for the whole docx (mutated
    in place to add the photo image + updated relationships)."""
    if PHOTO_BOX_ANCHOR not in data:
        print("  [MISSING] Photo box anchor — template's photo field shape may have "
              "changed again; see the comment above insert_photo() in cv_builder.py")
        return data

    # Crop to the photo box's own aspect ratio (not an arbitrary constant —
    # recompute from cx/cy above if the template's box size ever changes)
    # so the image fills it without visible stretching.
    target_ratio = 1428750 / 1689100
    im = Image.open(io.BytesIO(photo_bytes)).convert("RGB")
    w, h = im.size
    cur_ratio = w / h
    if cur_ratio > target_ratio:
        new_w = int(h * target_ratio)
        left = (w - new_w) // 2
        im = im.crop((left, 0, left + new_w, h))
    else:
        new_h = int(w / target_ratio)
        top = (h - new_h) // 2
        im = im.crop((0, top, w, top + new_h))

    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=90)
    fname = "candidate_photo.jpeg"
    files[f"word/media/{fname}"] = buf.getvalue()

    rid = "rIdCandidatePhoto"
    rels_path = "word/_rels/document.xml.rels"
    rels_data = files[rels_path].decode("utf-8")
    if rid not in rels_data:
        rels_data = rels_data.replace(
            "</Relationships>",
            f'<Relationship Id="{rid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="media/{fname}"/></Relationships>',
        )
        files[rels_path] = rels_data.encode("utf-8")

    pic_xml = (
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="9001" name="CandidatePhoto"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="4724273" y="321818"/><a:ext cx="1428750" cy="1689100"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
    )
    print("  Inserted candidate photo into the CV Template's photo field")
    return data.replace(PHOTO_BOX_ANCHOR, PHOTO_BOX_ANCHOR + pic_xml, 1)


# ---------------------------------------------------------------------------
# Library entry point (used by streamlit_app.py)
# ---------------------------------------------------------------------------

DEFAULT_TEMPLATE = Path(__file__).resolve().parent / "assets" / "Candidate_Template.docx"


def build_docx(candidate: dict, photo_bytes: bytes | None = None,
                template_bytes: bytes | None = None) -> tuple[bytes, list[str]]:
    """Build the filled CV Template docx entirely in memory.

    candidate: dict matching field-schema.md (first_name, last_name,
        employment[], optional resume{}, etc.)
    photo_bytes: optional image bytes for the CV Template's own photo
        field (page 1). Never sourced from the resume automatically —
        only pass this when the caller explicitly supplied a photo.
    template_bytes: optional raw bytes of a Candidate_Template.docx to use
        instead of the bundled assets/Candidate_Template.docx.

    Returns (docx_bytes, warnings) — warnings is a list of human-readable
    strings describing anything the builder had to adjust or couldn't be
    sure of (e.g. certifications trimmed to fit a fixed-size box). Empty
    when there's nothing to flag.
    """
    template_bytes = template_bytes or DEFAULT_TEMPLATE.read_bytes()
    warnings: list[str] = []

    files = {}
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zin:
        for name in zin.namelist():
            files[name] = zin.read(name)

    data = files["word/document.xml"].decode("utf-8")

    data = fill_simple_fields(data, candidate, warnings)
    data = fill_education(data, candidate)
    data = rebuild_employment_section(data, candidate.get("employment", []))
    if candidate.get("resume"):
        data = append_resume_text(data, candidate["resume"])
    if photo_bytes:
        data = insert_photo(data, files, photo_bytes)

    files["word/document.xml"] = data.encode("utf-8")

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in files.items():
            zout.writestr(name, content)
    return out_buf.getvalue(), warnings


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True,
                    help="candidate JSON file. Include a top-level \"resume\" key "
                         "(see field-schema.md) to append the resume as editable "
                         "text — contact details, photo, and logos excluded.")
    ap.add_argument("--template", default=None, help="path to Candidate_Template.docx (defaults to bundled asset)")
    ap.add_argument("--output", required=True, help="output .docx path")
    ap.add_argument("--photo", default=None,
                    help="optional candidate photo for the CV TEMPLATE's own photo field. "
                         "Only pass this if explicitly supplying a photo for that field — "
                         "do not source one from the resume.")
    args = ap.parse_args()

    with open(args.data, encoding="utf-8") as f:
        candidate = json.load(f)

    template_bytes = Path(args.template).read_bytes() if args.template else None
    photo_bytes = Path(args.photo).read_bytes() if args.photo else None

    docx_bytes, warnings = build_docx(candidate, photo_bytes=photo_bytes, template_bytes=template_bytes)

    out_path = Path(args.output)
    out_path.write_bytes(docx_bytes)
    print(f"Wrote {out_path}")
    for w in warnings:
        print(f"WARNING: {w}")


if __name__ == "__main__":
    main()
