"""
Spelling Shed Lesson Generator — Streamlit App
Two-stage flow: suggest words → review/edit → generate
"""

import streamlit as st
import json
import io
import zipfile
from wfa_shared.logo import logo_html
from wfa_shared.streamlit_css import inject_wfa_css
from content_generator import generate_lesson_json, suggest_words
from worksheets_builder import build_worksheets
from slides_builder import build_slides

st.set_page_config(
    page_title="Spelling Lesson Generator",
    page_icon="📝",
    layout="centered"
)

inject_wfa_css(buttons=True, inputs=True, download=True)

st.markdown(f"""
<div style="border-bottom:1px solid #e5e5e5;margin-bottom:16px;padding-bottom:12px">
    {logo_html("Spelling Lesson Generator")}
    <p style="margin:-4px 0 0 0;padding:0;color:#555;font-size:1rem">
        Wallscourt Farm Academy
    </p>
</div>
""", unsafe_allow_html=True)

# ── Persistent state ──────────────────────────────────────────────────────────

if "word_list_input" not in st.session_state:
    st.session_state["word_list_input"] = ""

# ── Year group and rule ───────────────────────────────────────────────────────

col1, col2 = st.columns([1, 2])

with col1:
    year_group = st.selectbox(
        "Which year group?",
        options=["Y2", "Y3", "Y4", "Y5", "Y6"],
        index=2,
    )

with col2:
    spelling_rule = st.text_input(
        "Spelling rule",
        placeholder="e.g. Silent letter — K / Words ending in -tion / Doubling the consonant",
    )

# rule_origin always matches the teaching year group
rule_origin = year_group

# ── Suggest words button ──────────────────────────────────────────────────────

if st.button("Suggest words"):
    if not spelling_rule.strip():
        st.error("Please enter a spelling rule first.")
    else:
        with st.spinner("Suggesting words…"):
            try:
                suggested = suggest_words(spelling_rule.strip(), year_group)
                st.session_state["word_list_input"] = "\n".join(suggested)
            except Exception as ex:
                st.error(f"Word suggestion failed: {ex}")

# ── Word list ─────────────────────────────────────────────────────────────────

word_list_raw = st.text_area(
    "Word list (one per line, 8–10 words)",
    height=200,
    help="Edit the suggested words or type your own.",
    key="word_list_input"
)

key_spelling = st.text_input(
    "Key spelling word (optional)",
    placeholder="e.g. because",
    help="A word from your class key spelling list. Used on the Quick Write slide. Leave blank to skip that slide."
)

# ── Generate ──────────────────────────────────────────────────────────────────

if st.button("Generate lesson", type="primary", use_container_width=True):
    words = [w.strip() for w in word_list_raw.strip().splitlines() if w.strip()]

    errors = []
    if not spelling_rule.strip():
        errors.append("Please enter a spelling rule.")
    if len(words) < 6:
        errors.append("Please enter at least 6 words.")
    if len(words) > 12:
        errors.append("Please enter no more than 12 words.")
    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    with st.status("Generating lesson… this takes about 30–40 seconds", expanded=True) as status:
        st.write("Calling AI to generate lesson data (this is the slow part)…")
        try:
            lesson = generate_lesson_json(
                rule=spelling_rule.strip(),
                words=words,
                year_group=year_group,
                rule_origin=rule_origin,
                key_spelling_word=key_spelling.strip() or None,
            )
        except Exception as ex:
            status.update(label="Content generation failed.", state="error")
            st.error(f"AI generation error: {ex}")
            st.stop()

        st.write("Building teaching slides…")
        try:
            slides_bytes = build_slides(lesson)
        except Exception as ex:
            status.update(label="Slides build failed.", state="error")
            st.error(f"Slides build error: {ex}")
            st.stop()

        st.write("Building worksheets…")
        try:
            ws_bytes = build_worksheets(lesson)
        except Exception as ex:
            status.update(label="Worksheets build failed.", state="error")
            st.error(f"Worksheets build error: {ex}")
            st.stop()

        status.update(label="Done!", state="complete")

    code = lesson.get("code", "XX")
    st.success(f"Lesson generated: **{spelling_rule}** ({year_group}, {len(words)} words)")

    col_a, col_b, col_c, col_d = st.columns(4)

    with col_a:
        st.download_button(
            label="⬇ Teaching slides",
            data=slides_bytes,
            file_name=f"spelling_slides_{code}_{year_group}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    with col_b:
        st.download_button(
            label="⬇ Worksheets",
            data=ws_bytes,
            file_name=f"spelling_worksheets_{code}_{year_group}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    with col_c:
        st.download_button(
            label="⬇ Lesson data",
            data=json.dumps(lesson, indent=2, ensure_ascii=False),
            file_name=f"lesson_{code}_{year_group}.json",
            mime="application/json",
            use_container_width=True,
        )

    with col_d:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"spelling_slides_{code}_{year_group}.pptx", slides_bytes)
            zf.writestr(f"spelling_worksheets_{code}_{year_group}.pptx", ws_bytes)
            zf.writestr(f"lesson_{code}_{year_group}.json",
                        json.dumps(lesson, indent=2, ensure_ascii=False))
        zip_buf.seek(0)
        st.download_button(
            label="⬇ All files (ZIP)",
            data=zip_buf.getvalue(),
            file_name=f"spelling_lesson_{code}_{year_group}.zip",
            mime="application/zip",
            use_container_width=True,
        )

    with st.expander("Preview generated content"):
        st.write(f"**Rule:** {lesson.get('rule')}")
        st.write(f"**Code:** {lesson.get('code')}")
        st.write(f"**Words:** {', '.join(lesson.get('words', []))}")
        if lesson.get("keySpellingWord"):
            st.write(f"**Key spelling word:** {lesson.get('keySpellingWord')}")
        st.write("**Definitions:**")
        for w, d in lesson.get("defs", {}).items():
            st.write(f"- *{w}*: {d}")
        st.write("**Starter question:**", lesson.get("starter", {}).get("question", ""))
