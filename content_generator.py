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
    "box1label": string,    // Label for first sort category (e.g. "a_e pattern", "verb only", "1 syllable")
    "box1sub":   string,    // Optional sub-text under box1label
    "box2label": string,    // Label for second sort category
    "box2sub":   string,    // Optional sub-text under box2label
    "hint":      string,
    "answerNote":  string,
    "exampleLine": string,
    "verbOnly":  [string],   // Words belonging in BOX 1 (despite the legacy name).
                              // The contents must match box1label's sort criterion,
                              // NOT literally "verb-only" words. The field name is a
                              // historical label and means "the first sort group".
    "verbNoun":  [ {"word": string, "eg": string} ]
                              // Words belonging in BOX 2 (despite the legacy name).
                              // The contents must match box2label's sort criterion.
                              // "eg" is a short example/note shown alongside each word
                              // (can be a sample sentence, alternative form, or note).
    // verbOnly + all verbNoun[].word must equal all words exactly
  },

  "spellData": [
    {"opts": [string, string, string], "correct": 0|1|2}
    // Exactly one row per word, in words[] order
  ],

  "clozeOrder": [string],  // All words shuffled — no word in same position as words[]

  "etymology": {
    "word":     string,        // The target word the children should guess
    "baseForm": string,        // Base form of the target word (e.g. "race" for "racing")
    "clicks": [ {"label": string, "body": string} ]
    // Exactly 6 stages, ordered OLDEST origin → MOST RECENT.
    // Card 1 = the most ancient or distant root (Proto-Indo-European, Latin, Greek, Old Norse, Celtic etc.)
    // Card 2 = the next stage of the journey
    // Card 3 = the next stage
    // Cards 1, 2 and 3 are the only cards visible to the class as the guessing
    // clues. They MUST NOT mention the target word, baseForm, or any close
    // derivative (so for "racing"/"race", clues 1-3 must avoid the strings
    // "race", "racing", "racer", "raced", "racecourse" etc.).
    // Clues 1-3 should describe meaning, origin, related concepts and language
    // history WITHOUT naming the answer.
    // Cards 4, 5 and 6 appear after the answer is revealed, so THESE may freely
    // mention the target word — they are the "now you know" follow-up notes.
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

## CRITICAL: never use the phrase "magic-e" or "magic e"

The phrase "magic-e" must never appear in any field of the output. The school does not teach this phrase. Always refer to the pattern as "split digraph" (e.g. a_e, i_e, o_e, u_e, e_e). If you find yourself wanting to write "magic e" or "magic-e", substitute "split digraph" or describe the specific pattern (a_e, i_e etc.).

## CRITICAL: split digraph + suffix rules

When the user-provided word list contains words built from a split digraph base word (one ending in vowel + consonant + e where the e makes the vowel long, e.g. bake, drive, hope) with -ed, -ing, -er or -est added, you MUST apply the correct spelling rules. Many earlier outputs have produced wrong forms like "bakeing", "driveing", "makeing", "smileing". These are spelling errors and must never appear in any field of the output (words, starter.answers, sentences, defs, spellData.opts, clozeOrder, wordShed, morphMatrix, etymology, syllableBreaks).

The rules:

1. Split digraph base + ING → DROP the e, then add ing.
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

2. Split digraph base + ED → DROP the e, then add ed (the final spelling ends in -ed not -eed).
   - bake + ed → baked (NOT "bakeed")
   - hope + ed → hoped (NOT "hopeed")
   - move + ed → moved (NOT "moveed")
   - skate + ed → skated (NOT "skateed")
   - race + ed → raced (NOT "raceed")

3. Split digraph base + ER / EST → DROP the e, then add er/est.
   - large + er → larger (NOT "largeer")
   - nice + est → nicest (NOT "niceest")

4. The starter.answers field MUST contain the correctly-spelt forms. If the starter shows base words like "bake", "drive", "smile" with "+ing" or "+ed", the answers must be "baking", "driving", "smiling" — never the wrong forms.

5. The spellData distractor "opts" can include the wrong form (e.g. "bakeing") AS A DELIBERATE WRONG OPTION, but the correct entry (correct=N) must point to the right form.

Validate every word in your output against these rules before returning. Reject any word that has "split digraph base + ing/ed/er/est" with the e still present (e.g. anything ending in "eing", "eed" except past-tense regular -ed verbs that don't have a split digraph in their base, "eer", "eest").

## CRITICAL: wordSort field semantics

The wordSort fields `verbOnly` and `verbNoun` are LEGACY NAMES from when the only sort was verb-only vs verb/noun. They DO NOT mean the contents must be sorted by part of speech. They are simply "the array of words for box 1" and "the array of words for box 2".

You must populate `verbOnly` and `verbNoun` based on the SORT CRITERION you defined in `box1label` and `box2label`. These should reflect the spelling rule being taught, not part of speech.

Examples of correct sorts for different lessons:

* Split digraph patterns (a_e vs i_e/o_e):
  - box1label="a_e pattern", verbOnly=["danced","baked","saved","shared","placed","glared"]
  - box2label="i_e or o_e pattern", verbNoun=[{"word":"smiled","eg":""},{"word":"hoped","eg":""},{"word":"prized","eg":""}]

* Syllable count (1 vs 2):
  - box1label="1 syllable", verbOnly=["tried","cried","dried"]
  - box2label="2 syllables", verbNoun=[{"word":"carried","eg":""},...]

* Suffix type (-ed vs -ing):
  - box1label="-ed words", verbOnly=["hoped","baked","skated","moved"]
  - box2label="-ing words", verbNoun=[{"word":"hiking","eg":""},...]

* Word class (verb only vs verb that can be noun) — only when the rule is grammatical:
  - box1label="verb only", verbOnly=["tried","dried","cried"]
  - box2label="verb or noun", verbNoun=[{"word":"carried","eg":"She carried the books (verb)"},...]

ALWAYS choose a sort that makes sense for THIS lesson's rule. Never default to part-of-speech sorting unless the lesson is explicitly about word classes.

The contents of `verbOnly` plus all the `word` values in `verbNoun` must together equal every word in the lesson, with no duplicates and nothing missing.

The `eg` field on each verbNoun entry is optional context shown alongside the word in the answers slide. For grammatical sorts use it for example sentences; for spelling-pattern sorts you can leave it empty ("").

## CRITICAL: etymology card ordering and word-hiding

The etymology slide is a guessing activity. Cards 1, 2 and 3 (the first three entries in `etymology.clicks`) are revealed one at a time as clues. Children must guess the target word from these three clues alone before the answer is revealed. Cards 4, 5 and 6 are follow-up information shown after the answer.

Two strict rules:

1. ORDER: cards 1-3 must go OLDEST origin → MOST RECENT. The first card is the most ancient root (e.g. Proto-Indo-European, Old Norse, Latin, Greek, Celtic, Sanskrit, Old English depending on the word). The second card is the next stage of the journey. The third card is the most recent or specific stage that's still a clue. Never lead with "Modern English" or "Today we use...". Modern usage belongs in the LATER cards (4-6) which appear after the reveal.

2. NO ANSWER REVEAL: cards 1, 2 and 3 must NEVER contain the target word, the baseForm, or any close derivative. For example, if the target word is "racing" and baseForm is "race":
   - FORBIDDEN strings in clues 1-3: "race", "racing", "racer", "raced", "races", "racecourse", "raceway"
   - The clue text must describe the meaning, origin and related concepts WITHOUT naming the answer.
   - Use phrases like "the word", "this word", "the verb", "an action meaning…" instead of repeating the answer.
   - Cards 4, 5 and 6 may freely use the target word — they are the post-reveal explanation.

Example of GOOD etymology for target "racing" (baseForm "race"):
  click 1: "Old Norse — comes from a word meaning 'a running, a rush'"
  click 2: "Old French — borrowed into French as a word meaning 'family, generation, breed'"
  click 3: "Modern usage — now describes a contest of speed between competitors"
  click 4: "The verb form 'race' became common in English from the 14th century."
  click 5: "Adding -ing turns the verb into the present participle 'racing'."
  click 6: "Did you know? The same Old Norse root gave us 'rush' too."

Example of BAD etymology (DO NOT do this):
  click 1: "Modern English — the word 'racing' comes from the verb 'race'..."  ← reveals answer in card 1
  click 2: "Old French — race meaning..."  ← reveals answer
  click 3: "Old Norse — ..."  ← order is reversed (most recent first)

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
- etymology.clicks[0..2] (the first 3 cards, shown to the class as guessing clues) MUST NOT contain the target word, baseForm or any close derivative. Order them oldest origin first, never modern usage first.
- morphMatrix.suffixes and morphMatrix.answers must each have exactly 6 entries.
- spellData must have exactly one row per word in words[] order.
- Split digraph + suffix rules above must be applied to every word, anywhere in the output.
- Never use the phrase "magic-e" or "magic e" in any field. Use "split digraph" or the specific pattern (a_e, i_e, o_e, u_e, e_e).
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

If any of the words above are derived from split digraph base words plus a suffix (-ed, -ing, -er, -est), make sure every reference to those words across the entire output uses the correct dropped-e spelling. Never write "bakeing", "driveing", "makeing", "smileing" or any similar form.

Never use the phrase "magic-e" or "magic e" anywhere in the output. Always use "split digraph" or the specific pattern name (a_e, i_e, o_e, u_e, e_e).

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

    # "magic-e" / "magic e" check — the school uses "split digraph" instead.
    # Scan every string value anywhere in the lesson and reject if found.
    forbidden_patterns = re.compile(r"\bmagic[\s\-]e\b", re.IGNORECASE)

    def _scan_for_forbidden(node, path=""):
        hits = []
        if isinstance(node, str):
            if forbidden_patterns.search(node):
                hits.append((path or "<root>", node[:80]))
        elif isinstance(node, dict):
            for k, v in node.items():
                hits.extend(_scan_for_forbidden(v, f"{path}.{k}" if path else k))
        elif isinstance(node, list):
            for i, item in enumerate(node):
                hits.extend(_scan_for_forbidden(item, f"{path}[{i}]"))
        return hits

    forbidden_hits = _scan_for_forbidden(lesson)
    if forbidden_hits:
        details = "; ".join(f"{loc}: {snippet!r}" for loc, snippet in forbidden_hits[:5])
        raise ValueError(
            f"Forbidden phrase 'magic-e' / 'magic e' found in generated lesson "
            f"(school uses 'split digraph' instead). Locations: {details}. Regenerate the lesson."
        )

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
            f"Split digraph + suffix spelling errors detected "
            f"(e was not dropped before adding -ing/-ed/-er/-est): {details}. "
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
    etym = lesson.get("etymology", {})
    if len(etym.get("clicks", [])) != 6:
        raise ValueError("etymology.clicks must have exactly 6 entries.")

    # etymology clues (first 3 cards) must not give the answer away.
    # We forbid the target word, the baseForm, and a few simple morphological
    # variants — anything that pupils would immediately recognise as the answer.
    target_word = (etym.get("word") or "").strip().lower()
    base_form   = (etym.get("baseForm") or "").strip().lower()

    def _morph_variants(w):
        """Return a small set of close spelling derivatives of w."""
        if not w or len(w) < 3:
            return set()
        v = {w}
        # Drop trailing -e, -d, -s if present
        for suf in ("ed", "ing", "es", "s", "er", "ers", "est", "y", "ly"):
            if w.endswith(suf):
                stem = w[:-len(suf)]
                if len(stem) >= 3:
                    v.add(stem)
        # Add common derivatives
        stems_to_extend = list(v)
        for stem in stems_to_extend:
            # Drop final -e then add suffix (split-digraph friendly)
            stem_no_e = stem[:-1] if stem.endswith("e") else stem
            for suf in ("", "s", "es", "ed", "ing", "er", "ers", "est"):
                cand = stem_no_e + suf
                if len(cand) >= 3:
                    v.add(cand)
                cand2 = stem + suf
                if len(cand2) >= 3:
                    v.add(cand2)
        return v

    forbidden_in_clues = set()
    for w in (target_word, base_form):
        forbidden_in_clues |= _morph_variants(w)
    # Don't accidentally forbid generic short fragments
    forbidden_in_clues = {w for w in forbidden_in_clues if len(w) >= 4}

    bad_clues = []
    for i in range(3):  # only check cards 1, 2, 3 — the guessing clues
        card = etym["clicks"][i]
        text = (card.get("label", "") + " " + card.get("body", "")).lower()
        for forbidden in forbidden_in_clues:
            # Word-boundary match so "race" doesn't match inside "embrace"
            if re.search(r"\b" + re.escape(forbidden) + r"\b", text):
                bad_clues.append((i + 1, forbidden, card.get("body", "")[:80]))
                break

    if bad_clues:
        details = "; ".join(
            f"card {n} contains '{w}': {snippet!r}"
            for n, w, snippet in bad_clues
        )
        raise ValueError(
            f"Etymology guessing clues (cards 1-3) reveal the answer word: "
            f"{details}. Cards 1-3 must not contain '{target_word}', '{base_form}' "
            f"or close variants. Regenerate the lesson."
        )

    # First clue should NOT start with "Modern" (etymology cards must go oldest first).
    first_label = (etym["clicks"][0].get("label") or "").strip().lower()
    if first_label.startswith("modern"):
        raise ValueError(
            f"Etymology card 1 starts with 'Modern' — cards must run oldest origin "
            f"first, modern usage last. Got label: {etym['clicks'][0].get('label')!r}. "
            f"Regenerate the lesson."
        )

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
- All words must be CORRECTLY spelt — apply standard spelling rules. For example, when adding -ing or -ed to a split digraph base word (one ending in vowel + consonant + e where the e makes the vowel long, like bake, drive, hope), drop the e (bake + ing = baking, not "bakeing"; drive + ed = drove or driven for past tense, not "driveing").
- Never use the phrase "magic-e" or "magic e" — always say "split digraph".
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
