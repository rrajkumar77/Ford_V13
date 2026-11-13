"""
pdf_layout.py - Layout-aware PDF text extraction.

pdfplumber's default page.extract_text() reads strictly top-to-bottom
across the full page width, so on a two-column resume it interleaves the
two columns line-by-line into nonsense. This reconstructs proper reading
order: read the whole left column top-to-bottom, then the whole right
column top-to-bottom, with any genuinely full-width lines (name, title,
summary paragraph, section headers that span the page) kept in their
natural position.

Falls back to plain single-stream extraction when there's no real
two-column evidence, so single-column resumes are unaffected.
"""
import io
import re

import pdfplumber

PAGE_NUM_RE = re.compile(r"^page\s+\d+\s+of\s+\d+$", re.IGNORECASE)
SUBGROUP_GAP = 12  # pt gap within a row that separates two distinct text clusters


def _opener(pdf_source):
    """Return a zero-arg callable that (re)opens the PDF. Reading a
    file-like object (e.g. Streamlit's UploadedFile) once and buffering
    the bytes lets us reopen it more than once, since pdfplumber can't
    re-read an exhausted stream."""
    if hasattr(pdf_source, "read"):
        data = pdf_source.read()
        return lambda: pdfplumber.open(io.BytesIO(data))
    return lambda: pdfplumber.open(pdf_source)


def _row_subgroups(row_words):
    row_words = sorted(row_words, key=lambda w: w["x0"])
    groups = [[row_words[0]]]
    for w in row_words[1:]:
        if w["x0"] - groups[-1][-1]["x1"] > SUBGROUP_GAP:
            groups.append([w])
        else:
            groups[-1].append(w)
    return groups


def _global_gap_boundary(all_rows, min_gap=25):
    """Find the true sidebar/main-content column gutter by looking at
    which x-ranges are NEVER touched by any text anywhere on the page,
    rather than sampling gaps from individual rows.

    The previous approach calibrated the boundary from rows with exactly
    2 subgroups - but on a resume with a timeline layout (date column +
    description column, both already right of the true sidebar), most
    2-subgroup rows are that inner date/description split, not the
    sidebar/main split. That pulls the calibrated boundary in from the
    true gutter to sit between the date and description columns instead,
    so every timeline date gets classified as sidebar content.

    Instead: project every subgroup's horizontal extent onto the x-axis
    and merge overlapping/touching ones. A real column gutter is a gap
    no text ever crosses, on any row - whereas a narrow date/description
    gap on timeline rows gets bridged by other rows that flow across it
    (a header or paragraph line spanning the full content width). The
    largest surviving gap after merging is the real boundary.
    """
    intervals = []
    for pidx, top, groups in all_rows:
        if pidx == 0 and top < 190:
            continue
        for g in groups:
            intervals.append((g[0]["x0"], g[-1]["x1"]))
    if not intervals:
        return None
    intervals.sort()
    merged = [list(intervals[0])]
    for x0, x1 in intervals[1:]:
        if x0 <= merged[-1][1] + 3:  # overlap or touch (+3pt coordinate noise)
            merged[-1][1] = max(merged[-1][1], x1)
        else:
            merged.append([x0, x1])
    if len(merged) < 2:
        return None  # nothing ever leaves a gap - single-column page
    best_gap, boundary = 0, None
    for i in range(1, len(merged)):
        gap = merged[i][0] - merged[i - 1][1]
        if gap > best_gap:
            best_gap = gap
            boundary = (merged[i - 1][1] + merged[i][0]) / 2
    return boundary if best_gap >= min_gap else None


def _plain_text(open_pdf):
    with open_pdf() as pdf:
        return "\n\n".join(p.extract_text() or "" for p in pdf.pages)


def extract_layout_aware_text(pdf_source):
    """pdf_source: a file path, or a file-like object (must support .read())."""
    open_pdf = _opener(pdf_source)

    all_rows = []  # (page_idx, top, [subgroups])
    with open_pdf() as pdf:
        for pidx, page in enumerate(pdf.pages):
            words = page.extract_words()
            if not words:
                continue
            rows = {}
            for w in words:
                key = round(w["top"] / 3) * 3
                rows.setdefault(key, []).append(w)
            for key in sorted(rows.keys()):
                all_rows.append((pidx, key, _row_subgroups(rows[key])))

    # Calibrate the column boundary using the global-gap method (see
    # _global_gap_boundary) - it isn't fooled by a timeline layout's
    # inner date/description split, because a real column gutter is a
    # gap NO row on the page ever crosses, whereas an inner split like
    # that gets bridged by other rows (a header or paragraph spanning
    # the full content width).
    #
    # There used to be a fallback here to an older per-row-sample method
    # when this returns None. That turned out actively harmful: on a
    # genuinely single-column resume (a plain bullet list with a
    # hanging indent - the single most common resume layout there is),
    # this correctly returns None, but the old fallback would then mix
    # two unrelated small gaps (bullet-marker-to-text, and
    # title-to-right-aligned-date) from a handful of coincidentally
    # 2-subgroup rows into one meaningless "boundary", and confidently
    # treat the whole single-column document as two-column - scrambling
    # every section into nonsense. None here is a real, trustworthy
    # signal that there's no reliable page-wide column gap to find, so
    # go straight to plain single-stream extraction instead of
    # second-guessing it.
    boundary = _global_gap_boundary(all_rows)
    if boundary is None:
        return _plain_text(open_pdf)

    # The "header zone" (name/title sitting above the two-column layout)
    # isn't always cleanly above a fixed pixel row - a long profile-
    # summary paragraph can run for a while before the sidebar visually
    # starts. Find where the sidebar actually begins on page 0 (the
    # first row with any content left of the boundary) and use that as
    # the header-zone cutoff, instead of a fixed guess that can slice
    # through the middle of that paragraph.
    header_zone_end = 190
    sidebar_tops_p0 = [
        top for pidx, top, groups in all_rows
        if pidx == 0 and any(g[-1]["x1"] <= boundary for g in groups)
    ]
    if sidebar_tops_p0:
        header_zone_end = min(sidebar_tops_p0)

    lines_out = []
    left_buf, right_buf = [], []

    def flush():
        lines_out.extend(left_buf)
        lines_out.extend(right_buf)
        left_buf.clear()
        right_buf.clear()

    current_page = None
    for pidx, top, groups in all_rows:
        # Flush at each page boundary rather than only once at the very
        # end. Without this, an entire multi-page sidebar column (e.g. a
        # "Key Skills" list that continues onto page 2) gets concatenated
        # into one giant contiguous run with content from unrelated
        # sections on other pages sitting right next to it - which is how
        # a stray line from one page's column ends up looking, to the
        # downstream section parser, like it belongs to a header that's
        # actually from a different page.
        if current_page is not None and pidx != current_page:
            flush()
        current_page = pidx

        texts = [" ".join(w["text"] for w in g) for g in groups]
        full_text = " ".join(texts).strip()
        if PAGE_NUM_RE.match(full_text):
            continue
        is_header_zone = pidx == 0 and top < header_zone_end
        if is_header_zone:
            # The header zone (name/title, right at the top of page 1) is
            # a simple top-to-bottom block, not a two-column layout - so
            # it should never be routed through left_buf/right_buf at
            # all. Buffering it conditionally on the same boundary used
            # for the body let a recognized sidebar section header that
            # happens to also fall in this zone jump ahead of the name in
            # the final output (or the reverse), which is how a section
            # header like "PERSONAL INFORMATION" - not the candidate's
            # actual name - ended up being read as the name downstream.
            # Flushing first and appending directly preserves true visual
            # reading order for this zone regardless of the body's column
            # layout.
            flush()
            for t in texts:
                lines_out.append(t)
            continue

        # Body rows: never flush mid-stream (a stray wide-gap row here
        # must not scatter later column content into the wrong section) -
        # route every subgroup to a side using the calibrated boundary.
        for g, t in zip(groups, texts):
            x0, x1 = g[0]["x0"], g[-1]["x1"]
            if x1 <= boundary + 3:
                left_buf.append(t)
            elif x0 >= boundary - 3:
                right_buf.append(t)
            elif x0 < boundary:
                left_buf.append(t)
            else:
                right_buf.append(t)
    flush()
    return "\n".join(lines_out)


if __name__ == "__main__":
    import sys
    print(extract_layout_aware_text(sys.argv[1]))
