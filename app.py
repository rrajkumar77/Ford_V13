"""
Ford/Magnit CV Template Builder — Streamlit app.

Upload a candidate resume (PDF), extract the structured fields locally
(regex/heuristics, no LLM or API key), review/edit the result, and
download the filled Ford/Magnit CV Template as a .docx — with the resume
appended as editable text (no contact details, no photo, no logos).

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Deploy: push this folder to a GitHub repo, then on
https://share.streamlit.io point a new app at app.py in that repo.
No API keys or secrets needed.
"""

import json
from pathlib import Path

import streamlit as st

from cv_builder import build_docx
from extract_candidate import extract_candidate_from_resume
from pdf_layout import extract_layout_aware_text

APP_DIR = Path(__file__).resolve().parent

st.set_page_config(page_title="Ford/Magnit CV Builder", page_icon="📄", layout="wide")
st.title("📄 Ford / Magnit CV Template Builder")
st.caption(
    "Fill the Ford/Magnit CV Template from a resume, with the candidate's "
    "contact details, photo, and any logos stripped from the appended "
    "resume section."
)

# ---------------------------------------------------------------------------
# Sidebar: API key + template override
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("Settings")
    custom_template = st.file_uploader(
        "Replace the CV Template (optional)",
        type=["docx"],
        help="Only use this if Ford/Magnit has issued a newer template revision. "
             "If the layout changed materially, fields may not fill correctly — "
             "check the output carefully.",
    )
    st.divider()
    with open(APP_DIR / "example_candidate.json", encoding="utf-8") as f:
        example_json = f.read()
    st.download_button(
        "Download example candidate.json",
        data=example_json,
        file_name="example_candidate.json",
        mime="application/json",
        help="A worked example showing the full schema.",
    )
    with open(APP_DIR / "field-schema.md", encoding="utf-8") as f:
        schema_doc = f.read()
    st.download_button(
        "Download field-schema.md",
        data=schema_doc,
        file_name="field-schema.md",
        mime="text/markdown",
    )

# ---------------------------------------------------------------------------
# Step 1: resume + optional photo
# ---------------------------------------------------------------------------

st.header("1. Upload the resume")
resume_file = st.file_uploader("Candidate resume (PDF)", type=["pdf"])

resume_text = ""
if resume_file is not None:
    try:
        resume_text = extract_layout_aware_text(resume_file)
        with st.expander("Extracted resume text (verify this looks right)"):
            st.text_area("", resume_text, height=250, label_visibility="collapsed")
    except Exception as e:
        st.error(f"Couldn't read that PDF: {e}")

photo_file = st.file_uploader(
    "Candidate photo for the CV Template's photo field (optional)",
    type=["jpg", "jpeg", "png"],
    help="Only upload this if you have an explicit photo to use for the template's "
         "own photo box. It is never pulled from the resume automatically, and the "
         "appended resume text never includes a photo either way.",
)

# ---------------------------------------------------------------------------
# Step 2: extract (optional) or paste JSON directly
# ---------------------------------------------------------------------------

st.header("2. Build the candidate data")

if "candidate_json_text" not in st.session_state:
    st.session_state.candidate_json_text = ""

col1, col2 = st.columns([1, 2])
with col1:
    extract_clicked = st.button(
        "🔎 Extract from resume",
        disabled=not resume_text,
        help="Requires a resume uploaded above." if not resume_text else None,
    )
with col2:
    st.caption(
        "Local extraction (regex/heuristics), no API key or network call. "
        "Review the JSON below and fix anything it missed — see the schema "
        "download in the sidebar. You can also paste your own JSON directly."
    )

if extract_clicked:
    try:
        candidate = extract_candidate_from_resume(resume_text)
        st.session_state.candidate_json_text = json.dumps(candidate, indent=2, ensure_ascii=False)
        st.success("Extracted locally — review the JSON below, it will miss things on unusual layouts.")
    except Exception as e:
        st.error(f"Extraction failed: {e}")

candidate_json_text = st.text_area(
    "candidate.json (edit as needed)",
    height=400,
    key="candidate_json_text",
    placeholder='Paste candidate JSON here, or use "Extract with Claude" / '
                '"Load example" above. See field-schema.md in the sidebar '
                "for the full shape.",
)

# ---------------------------------------------------------------------------
# Step 3: generate
# ---------------------------------------------------------------------------

st.header("3. Generate the CV")

if st.button("🏗️ Generate CV Template docx", type="primary"):
    if not candidate_json_text.strip():
        st.error("The candidate JSON box is empty — extract from a resume, load the example, or paste JSON in first.")
    else:
        try:
            candidate = json.loads(candidate_json_text)
        except json.JSONDecodeError as e:
            st.error(f"That JSON doesn't parse: {e}")
        else:
            try:
                docx_bytes, build_warnings = build_docx(
                    candidate,
                    photo_bytes=photo_file.getvalue() if photo_file else None,
                    template_bytes=custom_template.getvalue() if custom_template else None,
                )
                first = candidate.get("first_name", "Candidate")
                last = candidate.get("last_name", "")
                last = "" if str(last).upper() == "NA" else last
                fname = f"{first}_{last}_Ford_Magnit_CV.docx".replace("__", "_")
                st.success("Generated. Download below, then open it and check the render before sending it on.")
                for w in build_warnings:
                    st.warning(w)
                st.download_button(
                    "⬇️ Download filled CV Template",
                    data=docx_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(f"Couldn't build the docx: {e}")

st.divider()
st.caption(
    "This app fills fields anchored to the exact text of the bundled "
    "Candidate_Template.docx. If Ford/Magnit issues a new template "
    "revision, upload it via the sidebar — if fields don't fill, the "
    "template's wording/structure likely changed and cv_builder.py's "
    "anchors need updating (see the comments at the top of that file)."
)
