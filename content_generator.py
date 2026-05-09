"""
content_generator.py
Calls the Anthropic API to generate a complete lesson.json for the Spelling Shed app.
Reads year_group_guidance.md and phoneme_rules.md from the same directory.
"""

import json
import os
import re
from pathlib import Path
import anthropic

# ── Load reference documents ──────────────────────────────────────────────────

_HERE = Path(__file__).parent

def _read(filename):
    p = _HERE / filename
    if p.exists():
        return p.read_text(encoding="utf-8")
    return ""

YEAR_GROUP_GUIDANCE = _read("year_group_guidance.md")
PHONEME_RULES       = _read("phoneme_rules.md")

# ── Prompt ────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are a specialist primary school spelling resource generator for the EdShed Spelling Shed programme, used in UK primary schools.

Your job is to generate a complete lesson.json object for a given spelling rule, word list and teaching year group.

You must follow the JSON schema exactly and generate every field. Return ONLY valid JSON — no preamble, no explanation, no markdown fences. The response must parse directly with json.loads().

## JSON Schema

{
  "rule":   string,        // Full rule name
  "stage":  string,        // e.g. "Stage 4"
  "code":   string,        // 2-3 letter code from the rule (ILY, DC, SL, SLK, TRE etc.)

  "words": [string],       // Exactly the words provided — do not add or remove any

  "phonemes": {
    "<word>": [ {"l":"<letters>", "t":"dot|line"} | {"l":"<letters>", "t":"dot", "sid":N} ]
    // One entry per word. See phoneme rules below.
  },

  "defs": { "<word>": string },       // One definition per word

  "sentences": { "<word>": string },  // One sentence per word, word appears exactly once

  "wordMaps": {
    "words": [string],     // 6 of the words — pick those that best illustrate the rule
    "syllables": { "<word>": string }  // Syllable breaks for each wordMaps word e.g. "hap | pi | ly"
  },

  "syllableCounts": { "<word>": number },   // All words

  "syllableBreaks": { "<word>": string },   // All words, " | " separator

  "starter": {
    "question": string,
    "words": [string],       // Exactly 6 base forms
    "answers": [string],     // Exactly 6 answers matching starter.words
    "answerLabel": string,   // e.g. "adverb:" or "word:"
    "perPairNote": string,   // Empty string if not needed
    "ruleBox": string,
    "ruleText": string
  },

  "thisWeeksWordsQ": string,
  "thisWeeksWordsPrompt": string,
  "thisWeeksWordsExplanation": string,

  "wordSortQ": string,
  "wordSort": {
    "box1label": string,
    "box1sub":   string,
    "box2label": string,
    "box2sub":   string,
    "hint":      string,
    "answerNote":  string,
    "exampleLine": string,
    "verbOnly":  [string],
    "verbNoun":  [ {"word": string, "eg": string} ]
    // verbOnly + all verbNoun[].word must equal all words exactly
  },

  "spellData": [
    {"opts": [string, string, string], "correct": 0|1|2}
    // Exactly one row per word, in words[] order
  ],

  "clozeOrder": [string],  // All words shuffled — no word in same position as words[]

  "etymology": {
    "word":     string,
    "baseForm": string,
    "clicks": [ {"label": string, "body": string} ]
    // Exactly 6 stages
  },

  "wordShed": {
    "baseWord":   string,
    "def":        string,
    "sentence":   string,
    "rhymes":     string,
    "morphology": string
  },

  "morphMatrix": {
    "baseWord": string,       // Different from wordShed.baseWord
    "def":      string,
    "suffixes": [string],     // Exactly 6
    "answers":  [string]      // Exactly 6, same order
  }
}

## Phoneme Rules
""" + PHONEME_RULES + """

## Year Group Content Guidance
""" + YEAR_GROUP_GUIDANCE + """

## Critical rules
- Return ONLY valid JSON. No markdown, no explanation.
- Generate every field completely — no placeholders, no nulls.
- All content must be freshly generated for this specific rule and word list.
- Apply the year group content guidance strictly for the TEACHING year group.
- Phoneme data must be validated: total letters in all `l` values must equal word length.
- clozeOrder must contain all words with no word in the same array position as in words[].
- wordSort.verbOnly + all wordSort.verbNoun[].word must equal all words exactly (no word missing or duplicated).
- morphMatrix.baseWord must differ from wordShed.baseWord.
- starter.answers must have exactly 6 entries.
- etymology.clicks must have exactly 6 entries.
- morphMatrix.suffixes and morphMatrix.answers must each have exactly 6 entries.
- spellData must have exactly one row per word in words[] order.
"""

def _user_prompt(rule, words, year_group, rule_origin):
    cross_year_note = ""
    if year_group != rule_origin:
        cross_year_note = (
            f"\n\nIMPORTANT: This is a {rule_origin} spelling rule being taught to a {year_group} class "
            f"for consolidation. The words and rule come from {rule_origin} level, but ALL generated content "
            f"(sentences, definitions, morphology, cloze, etymology depth) must be pitched at {year_group} level."
        )

    stage_map = {"Y2": "Stage 2", "Y3": "Stage 3", "Y4": "Stage 4", "Y5": "Stage 5", "Y6": "Stage 6"}
    stage = stage_map.get(rule_origin, "Stage 4")

    return f"""Generate a complete lesson.json for the following spelling lesson.

Teaching year group: {year_group}
Rule origin level: {rule_origin} ({stage})
Spelling rule: {rule}
Words ({len(words)}): {", ".join(words)}{cross_year_note}

Pitch ALL generated content (definitions, sentences, morphology, word sort categories, etymology depth, cloze complexity) according to the {year_group} content guidance.

Return only the JSON object."""


# ── Generator ─────────────────────────────────────────────────────────────────

def generate_lesson_json(rule: str, words: list, year_group: str, rule_origin: str) -> dict:
    """
    Call the Anthropic API and return a validated lesson dict.
    Raises ValueError if the response cannot be parsed or fails basic validation.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY environment variable not set.")

    client = anthropic.Anthropic(api_key=api_key)

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[
            {"role": "user", "content": _user_prompt(rule, words, year_group, rule_origin)}
        ]
    )

    raw = response.content[0].text.strip()

    # Strip any accidental markdown fences
    raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.IGNORECASE)
    raw = re.sub(r'\s*```$', '', raw)

    try:
        lesson = json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(f"API returned invalid JSON: {e}\n\nRaw response (first 500 chars):\n{raw[:500]}")

    # Basic validation
    _validate(lesson, words)

    return lesson


def _validate(lesson: dict, words: list):
    """Raise ValueError if the lesson fails critical checks."""
    required_keys = [
        "rule", "stage", "code", "words", "phonemes", "defs", "sentences",
        "wordMaps", "syllableCounts", "syllableBreaks", "starter",
        "thisWeeksWordsQ", "wordSortQ", "wordSort", "spellData",
        "clozeOrder", "etymology", "wordShed", "morphMatrix"
    ]
    missing = [k for k in required_keys if k not in lesson]
    if missing:
        raise ValueError(f"lesson.json missing keys: {missing}")

    # Words match
    if sorted(lesson["words"]) != sorted(words):
        # Allow minor mismatch — just warn in production; here we patch silently
        lesson["words"] = words

    # spell data count
    if len(lesson.get("spellData", [])) != len(words):
        raise ValueError(
            f"spellData has {len(lesson['spellData'])} rows but {len(words)} words expected."
        )

    # clozeOrder
    if sorted(lesson.get("clozeOrder", [])) != sorted(words):
        # Regenerate a simple shuffle if broken
        import random
        shuffled = words[:]
        while any(shuffled[i] == words[i] for i in range(len(words))):
            random.shuffle(shuffled)
        lesson["clozeOrder"] = shuffled

    # etymology clicks
    if len(lesson.get("etymology", {}).get("clicks", [])) != 6:
        raise ValueError("etymology.clicks must have exactly 6 entries.")

    # morphMatrix
    mm = lesson.get("morphMatrix", {})
    if len(mm.get("suffixes", [])) != 6 or len(mm.get("answers", [])) != 6:
        raise ValueError("morphMatrix.suffixes and .answers must each have 6 entries.")

    # wordSort completeness
    ws = lesson.get("wordSort", {})
    sort_words = ws.get("verbOnly", []) + [e["word"] for e in ws.get("verbNoun", [])]
    if sorted(sort_words) != sorted(words):
        raise ValueError(
            f"wordSort words {sorted(sort_words)} do not match lesson words {sorted(words)}."
        )
