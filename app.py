"""
Spelling Shed Lesson Generator — Streamlit App
Generates teaching slides and worksheets from a spelling rule and word list.
Uses the Anthropic API to generate lesson content, then builds PPTX files via python-pptx.
"""

import streamlit as st
import json
import io
import zipfile

from content_generator import generate_lesson_json, suggest_words
from worksheets_builder import build_worksheets

st.set_page_config(
    page_title="Spelling Shed Lesson Generator",
    page_icon="📝",
    layout="centered"
)

st.title("Spelling Shed Lesson Generator")
st.caption("Wallscourt Farm Academy — EdShed-style lesson resources")

# ── Step 1: Rule and year group ───────────────────────────────────────────────

with st.form("rule_form"):
    col1, col2 = st.columns(2)
    with col1:
        year_group = st.selectbox(
            "Teaching year group",
            options=["Y2", "Y3", "Y4", "Y5", "Y6"],
            index=2,
            help="The year group being taught. If reteaching a lower-year rule, set this to the actual class."
        )
    with col2:
        rule_origin = st.selectbox(
            "Rule level",
            options=["Y2", "Y3", "Y4", "Y5", "Y6"],
            index=2,
            help="The Spelling Shed stage the rule comes from. Usually matches the teaching year group."
        )

    spelling_rule = st.text_input(
        "Spelling rule",
        placeholder="e.g. Silent letter — K  /  Words ending in -tion  /  Doubling the consonant",
        help="Describe the rule clearly. This becomes the lesson title."
    )

    word_list_raw = st.text_area(
        "Word list (one per line, 8–10 words)",
        height=200,
        placeholder="Leave blank to get word suggestions, or enter your own words here.",
        help="Enter 8–10 words, or leave blank and click Suggest words first."
    )

    key_spelling = st.text_input(
        "Key spelling word (optional)",
        placeholder="e.g. because",
        help="A word from your class key spelling list. Leave blank if not needed."
    )

    col_a, col_b = st.columns(2)
    with col_a:
        suggest_btn = st.form_submit_button("Suggest words", use_container_width=True)
    with col_b:
        generate_btn = st.form_submit_button("Generate lesson", type="primary", use_container_width=True)


# ── Word suggestion ───────────────────────────────────────────────────────────

if suggest_btn:
    if not spelling_rule.strip():
        st.error("Please enter a spelling rule first.")
        st.stop()

    with st.spinner("Suggesting words…"):
        try:
            suggested = suggest_words(spelling_rule.strip(), year_group)
            st.session_state["suggested_words"] = suggested
        except Exception as ex:
            st.error(f"Word suggestion failed: {ex}")
            st.stop()

if "suggested_words" in st.session_state and not generate_btn:
    st.subheader("Suggested words")
    st.caption("Copy these into the word list above, edit as needed, then click Generate lesson.")
    st.code("\n".join(st.session_state["suggested_words"]), language=None)


# ── Generate ──────────────────────────────────────────────────────────────────

if generate_btn:
    words = [w.strip() for w in word_list_raw.strip().splitlines() if w.strip()]

    errors = []
    if not spelling_rule.strip():
        errors.append("Please enter a spelling rule.")
    if len(words) < 6:
        errors.append("Please enter at least 6 words (or click Suggest words first).")
    if len(words) > 12:
        errors.append("Please enter no more than 12 words.")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    with st.status("Generating lesson…", expanded=True) as status:

        st.write("Calling AI to generate lesson data…")
        try:
            lesson = generate_lesson_json(
                rule=spelling_rule.strip(),
                words=words,
                year_group=year_group,
                rule_origin=rule_origin,
            )
        except Exception as ex:
            status.update(label="Content generation failed.", state="error")
            st.error(f"AI generation error: {ex}")
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

    col_a, col_b, col_c = st.columns(3)

    with col_a:
        st.download_button(
            label="⬇ Worksheets (PPTX)",
            data=ws_bytes,
            file_name=f"spelling_worksheets_{code}_{year_group}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    with col_b:
        st.download_button(
            label="⬇ Lesson data (JSON)",
            data=json.dumps(lesson, indent=2, ensure_ascii=False),
            file_name=f"lesson_{code}_{year_group}.json",
            mime="application/json",
            use_container_width=True,
        )

    with col_c:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
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
        st.write("**Definitions:**")
        for w, d in lesson.get("defs", {}).items():
            st.write(f"- *{w}*: {d}")
        st.write("**Starter question:**", lesson.get("starter", {}).get("question", ""))
