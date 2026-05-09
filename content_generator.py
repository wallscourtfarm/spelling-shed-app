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

  "sentences": { "<word>": string },  // One COMPLETE sentence per word. The target word appears exactly once as a real word — NEVER use blanks or underscores. The builder inserts blanks automatically. Good: "The brave knight rode into battle." Bad: "The brave _____ rode into battle."

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

## CRITICAL: Magic-e (split digraph) + suffix rules

When the user-provided word list contains words built from a magic-e base word with -ed, -ing, -er or -est added, you MUST apply the correct spelling rules. Many earlier outputs have produced wrong forms like "bakeing", "driveing", "makeing", "smileing". These are spelling errors and must never appear in any field of the output (words, starter.answers, sentences, defs, spellData.opts, clozeOrder, wordShed, morphMatrix, etymology, syllableBreaks).

The rules:

1. Magic-e base + ING → DROP the e, then add ing.
   - bake + ing → baking (NOT "bakeing")
   - drive + ing → driving (NOT "driveing")
   - make + ing → making (NOT "makeing")
   - smile + ing → smiling (NOT "smileing")
   - shine + ing → shining (NOT "shineing")
   - race + ing → racing (NOT "raceing")
   - hope + ing → hoping (NOT "hopeing")
   - write + ing → writing (NOT "writeing")
   - skate + ing → skating (NOT "skateing")
   - hike + ing → hiking (NOT "hikeing")

2. Magic-e base + ED → DROP the e, then add ed (the final spelling ends in -ed not -eed).
   - bake + ed → baked (NOT "bakeed")
   - hope + ed → hoped (NOT "hopeed")
   - move + ed → moved (NOT "moveed")
   - skate + ed → skated (NOT "skateed")
   - race + ed → raced (NOT "raceed")

3. Magic-e base + ER / EST → DROP the e, then add er/est.
   - large + er → larger (NOT "largeer")
   - nice + est → nicest (NOT "niceest")

4. The starter.answers field MUST contain the correctly-spelt forms. If the starter shows base words like "bake", "drive", "smile" with "+ing" or "+ed", the answers must be "baking", "driving", "smiling" — never the wrong forms.

5. The spellData distractor "opts" can include the wrong form (e.g. "bakeing") AS A DELIBERATE WRONG OPTION, but the correct entry (correct=N) must point to the right form.

Validate every word in your output against these rules before returning. Reject any word that has "magic-e base + ing/ed/er/est" with the e still present (e.g. anything ending in "eing", "eed" except past-tense regular -ed verbs that don't have a magic e in their base, "eer", "eest").

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
- Magic-e + suffix rules above must be applied to every word, anywhere in the output.
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

If any of the words above are derived from magic-e base words plus a suffix (-ed, -ing, -er, -est), make sure every reference to those words across the entire output uses the correct dropped-e spelling. Never write "bakeing", "driveing", "makeing", "smileing" or any similar form.

Return only the JSON object."""


# ── Generator ─────────────────────────────────────────────────────────────────

def generate_lesson_json(rule: str, words: list, year_group: str, rule_origin: str,
                         key_spelling_word: str = None) -> dict:
    """
    Call the Anthropic API and return a validated lesson dict.
    Raises ValueError if the response cannot be parsed or fails basic validation.

    key_spelling_word is an optional class key-spelling word that gets attached
    to the returned lesson as keySpellingWord. The slides builder uses this to
    add the Quick Write Key Spelling Practice slide.
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

    # Attach key spelling word for the slides builder
    if key_spelling_word:
        lesson["keySpellingWord"] = key_spelling_word

    return lesson


# Words ending in these patterns are almost always misspelt magic-e + suffix
# forms. We check the user-supplied words list (not distractors in spellData,
# which are deliberately wrong) and any place the AI generates new word forms.
_MAGIC_E_SUFFIX_BAD_PATTERNS = re.compile(
    r"(?<!^)(eing|eed|eer|eest)$",  # cake + ing → "cakeing" etc.
    re.IGNORECASE,
)

# A small allow-list of legitimate -eer / -eed / -eing endings that aren't
# magic-e + suffix mistakes. Add words here if false positives appear.
_MAGIC_E_SUFFIX_ALLOW = {
    # -eed: real /iː/ + d words
    "agreed", "freed", "treed", "decreed", "guaranteed", "indeed", "kneed",
    "speed", "feed", "need", "seed", "weed", "breed", "creed", "deed", "freed",
    "greed", "reed", "steed", "tweed", "exceed", "succeed", "proceed",
    # -eer: /ɪər/ words
    "deer", "beer", "cheer", "jeer", "leer", "peer", "queer", "seer", "veer",
    "career", "engineer", "pioneer", "volunteer", "auctioneer", "puppeteer",
    "racketeer", "mountaineer", "musketeer", "profiteer", "sneer", "steer",
    # -eing: legitimate words where the e stays
    "being", "freeing", "seeing", "agreeing", "fleeing", "guaranteeing",
    "decreeing", "skiing",  # not -eing but skiing has unusual -iing
    # -eest: rare but possible
    "freest",
}


def _looks_like_magic_e_suffix_error(word: str) -> bool:
    """Return True if word looks like a missed magic-e drop (e.g. 'bakeing')."""
    w = word.lower().strip()
    if w in _MAGIC_E_SUFFIX_ALLOW:
        return False
    return bool(_MAGIC_E_SUFFIX_BAD_PATTERNS.search(w))


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

    # Magic-e + suffix check — applies to fields that should contain CORRECT
    # spellings. spellData distractor opts are deliberately wrong so excluded.
    bad_words = []
    # Main word list
    for w in lesson.get("words", []):
        if _looks_like_magic_e_suffix_error(w):
            bad_words.append(("words", w))
    # Starter answers
    for w in lesson.get("starter", {}).get("answers", []):
        if _looks_like_magic_e_suffix_error(w):
            bad_words.append(("starter.answers", w))
    # spellData correct option only (not distractors)
    for i, row in enumerate(lesson.get("spellData", [])):
        opts = row.get("opts", [])
        correct_idx = row.get("correct", 0)
        if 0 <= correct_idx < len(opts):
            w = opts[correct_idx]
            if _looks_like_magic_e_suffix_error(w):
                bad_words.append((f"spellData[{i}].correct", w))
    # clozeOrder
    for w in lesson.get("clozeOrder", []):
        if _looks_like_magic_e_suffix_error(w):
            bad_words.append(("clozeOrder", w))
    # wordShed
    ws_morph = lesson.get("wordShed", {}).get("morphology", "")
    for w in re.split(r"[,\s]+", ws_morph):
        if w and _looks_like_magic_e_suffix_error(w):
            bad_words.append(("wordShed.morphology", w))
    # morphMatrix.answers
    for w in lesson.get("morphMatrix", {}).get("answers", []):
        if _looks_like_magic_e_suffix_error(w):
            bad_words.append(("morphMatrix.answers", w))

    if bad_words:
        details = ", ".join(f"{loc}={w!r}" for loc, w in bad_words)
        raise ValueError(
            f"Magic-e suffix errors detected (e dropped wrong): {details}. "
            f"Regenerate the lesson."
        )

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


def suggest_words(rule: str, year_group: str, n: int = 10) -> list:
    """
    Ask the API to suggest n spelling words for a given rule and year group.
    Returns a list of word strings.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY environment variable not set.")

    client = anthropic.Anthropic(api_key=api_key)

    stage_map = {"Y2": "Stage 2", "Y3": "Stage 3", "Y4": "Stage 4", "Y5": "Stage 5", "Y6": "Stage 6"}
    stage = stage_map.get(year_group, "Stage 4")
    age_map = {"Y2": "6-7", "Y3": "7-8", "Y4": "8-9", "Y5": "9-10", "Y6": "10-11"}
    age = age_map.get(year_group, "8-9")

    prompt = f"""Suggest exactly {n} spelling words for a UK primary school Spelling Shed lesson.

Spelling rule: {rule}
Year group: {year_group} ({stage}, ages {age})

Requirements:
- All words must clearly illustrate the spelling rule
- Words must be appropriate in difficulty for {year_group} pupils aged {age}
- Words should be varied — different lengths, different contexts
- All words must be real, common English words pupils will encounter in reading
- All words must be CORRECTLY spelt — apply standard spelling rules. For example, when adding -ing or -ed to a magic-e base word, drop the e (bake + ing = baking, not "bakeing"; drive + ed = drove or driven for past tense, not "driveing").
- Return ONLY a JSON array of {n} lowercase strings, nothing else
- Example format: ["word1", "word2", "word3"]"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=200,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = response.content[0].text.strip()
    raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.IGNORECASE)
    raw = re.sub(r'\s*```$', '', raw)

    words = json.loads(raw)
    if not isinstance(words, list):
        raise ValueError("Word suggestion did not return a list.")
    return [w.strip().lower() for w in words if isinstance(w, str)]
