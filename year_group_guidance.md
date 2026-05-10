# Year Group Content Guidance — Spelling Shed Lesson Generator

This document instructs the AI on how to pitch generated `lesson.json` content for each year group. It is derived from analysis of actual EdShed Spelling Shed materials across Y2–Y6.

The core JSON schema fields are shared across all year groups. For Y4–Y6 the slide layout, schema and activity types are fixed and identical. For Y2 and Y3, additional schema fields are generated alongside the standard ones (`y2Starter`, `spellingPattern`, `sentencesAndSynonyms`, `wordsInAction`, `wordSpotter` and `y2IncludeMorphMatrix` for Y2; `y3Starter`, `spellingPattern` and `wordMatch` for Y3), and the slide deck uses a different structure suited to those year groups. Across all year groups, only the **content** — vocabulary, sentence complexity, definitions, morphology scope, word sort categories, cloze sentences and etymology depth — should vary according to the guidance below. Apply these guidelines when generating every field of `lesson.json`.

---

## Important: Cross-Year Teaching

A teacher may select a lower year group's rule to teach in a higher year group (e.g. Y2 rule taught to Y4 for consolidation/reteaching). When this happens:

- Use the **teaching year group** (the actual class being taught) to pitch all generated content, not the rule's origin year.
- The rule and words come from the selected year level; the sentences, definitions, morphology and complexity must be calibrated to the **teaching year group**.
- A Y2 rule taught in Y4 should have Y4-appropriate sentences and morphology, not Y2-level content.

---

## Year 2 (ages 6–7, Stage 2)

**Spelling rules:** Consonant digraphs making unexpected sounds (ge=/j/, dge=/j/), adding suffixes (-ed, -ing, -s/-es) to single-syllable words with doubling, basic split digraphs, common digraphs (ai, oi, ay, oy, etc.)

**Words:** Short, high-frequency, one or two syllables only. Entirely concrete nouns and common verbs. No abstract vocabulary.

**Definitions:** One sentence, maximum. Very simple syntax. Avoid subordinate clauses.

- Good: "A grey bird that lives in towns."
- Good: "To cover something tightly with paper or fabric."
- Avoid: "When he, she or it..." format — use plain noun/verb description instead.

**Sentences (cloze):** 6–8 words maximum. Simple subject-verb-object. No complex punctuation. One idea per sentence. The target word must be the most obvious completion.

- Good: "She ___ the parcel carefully."
- Avoid: "Despite the rain, she had carefully ___ the present before hiding it under the stairs."

**Morphology matrix:** 3–4 suffixes only. Common suffixes: -s, -ed, -ing, -er. One simple prefix if applicable (un-). Show the doubling rule in action where relevant to the week's pattern. No Latin or Greek roots.

**Word sort categories:** Concrete, phonics-based sorting only. Examples from the materials:

- Sort by number of phonemes (3, 4, 5, 6)
- Sort by number of syllables (1 syllable / 2 syllables)
- Sort by suffix sound (-ed making /d/ vs /t/)
- Sort by letter pattern (words containing 'a' / not containing 'a')
- Never sort by abstract grammatical concepts (noun/verb distinction is fine; adverb/adjective is not)

**Word sort box labels:** Short, concrete, phonics-grounded. "1 syllable" / "2 syllables". "Makes a /d/ sound" / "Makes a /t/ sound". Not "Verb only" / "Noun and verb".

**Etymology:** Do not include etymology for Y2. The etymology slide should contain a simple "Did you know?" fact about one of the words — one or two sentences, concrete and engaging. No Latin/French root chains.

**Word Shed:** Use the most familiar word from the list. Definition one sentence. Two very simple example sentences. Rhymes: common single-syllable rhyming words. Morphology derivatives: -s, -ed, -ing only.

**Starter question:** Phonics-based. "Can you write the word that matches the picture?" / "Can you sound out and write the words?" / "Which of last week's words are being described?"

**thisWeeksWordsQ:** Simple noticing question. "What do all our words have in common?" / "Where does the pattern appear in each word?"

**y2Starter clues:** Written clues must use Y2 reading vocabulary — short sentences, concrete descriptions, no abstract language. Each clue describes one of last week's words without naming it.

**sentencesAndSynonyms:** Use common, everyday synonyms that Y2 pupils will know (e.g. "unusual" for "strange", "attack" for "charge"). Keep original sentences short (6–8 words). The replacement word should be one of this week's words.

**wordsInAction:** The picture prompt should describe a simple, familiar scene. Required words must be from this week's list and appropriate for Y2 writing. Sample answer should be 1–2 short sentences.

**wordSpotter distractors:** Use words from a plausible previous lesson's rule or related phoneme pattern. All should be real, correctly-spelt common words.

**Word count:** 10 words is fine for Y2; however a teacher may choose to use 8 if the rule is especially complex.

---

## Year 3 (ages 7–8, Stage 3)

**Spelling rules:** Words ending in -ture, -tion, -sion; adding -ing/-ed/-en to multisyllabic words with consonant doubling; prefixes (re-, pre-, un-, mis-); silent letters; homophones and near-homophones.

**Words:** Mostly two syllables, some three. Mix of concrete and abstract nouns. Some less common verbs. Words should be in pupils' reading vocabulary even if not always in their written vocabulary.

**Definitions:** One or two sentences. Can include word class. Simple subordinate clauses acceptable.

- Good: "A living creature or animal."
- Good: "To catch someone or something and keep them. It can also be used as a noun."
- Can introduce "It can also mean..." constructions.

**Sentences (cloze):** 8–12 words. One or two clauses. Can include time phrases (yesterday, last week, after). Punctuation can include commas in lists and before conjunctions. Target word should be clear from context but not completely obvious — some inference needed.

**Morphology matrix:** 4–5 suffixes/prefixes. Can include: -tion, -ing, -ed, -er, -ment, -ful, re-, un-, mis-. Where the base word takes spelling changes (drop -e, double consonant), note this. Brief notes on meaning where needed.

**Word sort categories:** Can introduce grammatical sorting if concrete.

- "Words where '-ture' follows a vowel" / "Words where '-ture' follows a consonant"
- "'-ing' words" / "'-ed' words" / "'-en' words"
- Syllable counts (1/2/3)
- Noun and verb / verb only

**Word sort box labels:** Can be slightly more technical than Y2. "Noun and verb" / "Verb only" is appropriate. Grammatical terms pupils have been taught (noun, verb, adjective) are fine.

**Etymology:** Include etymology. Two or three click-reveal stages. Focus on one clear root and how it connects to the modern word. Avoid chains of three+ languages. Keep accessible.

- Good: root language → root word meaning → link to today's word → a simple fact
- Keep body text to one sentence per stage

**Word Shed:** Can include synonyms. Definition up to two sentences. Example sentences can be more complex. Rhymes: two-syllable rhymes acceptable. Morphology derivatives: up to 4, can include prefixed forms.

**Starter question:** Can be pattern-based. "Can you sort these words by their suffix?" / "Can you climb the word ladder by changing one sound each time?" / "Which of last week's words matches each definition?"

**thisWeeksWordsQ:** Pattern noticing. "What do all our words have in common?" / "Can you spot the three different suffixes?"

**y3Starter categories:** Choose categories that reflect a plausible previous lesson. If the current lesson is about -ture words, a previous lesson might have covered -ing/-ed/-en suffixes. Categories must be 2–3 distinct labels and each of the 10 last-week words must fit exactly one category.

**spellingPattern:** The title should match or paraphrase the lesson rule. The body should explain the pattern in 2–3 Y3-appropriate sentences. Include a rule_note about exceptions or related patterns. Examples should show base → result transformation (e.g. "sculpt" → "sculpture").

**wordMatch descriptions:** Concrete and visual — describe what you might see, hear or do in connection with the word. Avoid abstract definitions. 15–20 words maximum per description.

---

## Year 4 (ages 8–9, Stage 4) — EXISTING SKILL — REFERENCE ONLY

**This is the baseline the tool was built from.** All schema field formats and content patterns described in `lesson-data-schema.md` are calibrated for Y4. When teaching year is Y4, follow the schema as written without modification.

Key Y4 markers:

- Definitions: "When he, she or it [verb phrase]; also, [noun use if applicable]."
- Sentences: 10–15 words, one or two complex clauses, varied punctuation
- Morphology: 4–6 suffixes/prefixes including Latin-origin forms where appropriate
- Word sort: grammatical categories (verb only / verb and noun) with bracketed examples
- Etymology: full 6-stage click-reveal chain

---

## Year 5 (ages 9–10, Stage 5)

**Spelling rules:** Words ending in -cial/-tial/-sial, -ough (multiple sounds), -ant/-ance/-ancy/-ent/-ence/-ency, silent letters in less common positions, words with Latin and Greek roots, prefixes (sub-, inter-, super-, anti-), homophones of less common words.

**Words:** Two to four syllables. Mix of abstract and concrete. Technical and subject-specific vocabulary is appropriate (financial, artificial, controversial). Pupils may not know all meanings in advance — that's expected.

**Definitions:** Dictionary-style. Full word class (adjective, noun, verb, adverb). Can include antonyms or usage notes. Two sentences acceptable.

- Good: "Relating to money and financial systems. (adjective)"
- Good: "Made or produced by humans rather than occurring naturally. (adjective) The opposite of natural."

**Sentences (cloze):** 12–18 words. Complex sentences with subordinate clauses, relative clauses, parenthesis. Varied punctuation including colons, semicolons, brackets. The target word should require genuine understanding of its meaning to complete correctly — not guessable from grammar alone.

- Good: "Despite the rain, the official ceremony went ahead as planned, with hundreds of people lining the streets."

**Morphology matrix:** 5–6 entries. Include prefixes with specific meanings (sub- = under, inter- = between, super- = above). Can explore adverbial forms (-ly). Discuss whether adding certain suffixes changes word class. Note any spelling changes explicitly.

**Word sort categories:** More abstract and grammatical.

- Sort by number of syllables (2/3/4/5)
- Sort by suffix pattern (-cial after vowel / -tial after consonant)
- Sort by sound the grapheme makes
- Noun, adjective, adverb categories
- Carroll diagrams with two simultaneous criteria

**Word sort box labels:** Full grammatical terminology. "Adjective" / "Adverb" / "Both". "Regular pattern" / "Irregular". Two-dimensional Carroll diagram sorting is appropriate.

**Etymology:** Full 6-stage chain. Multiple language hops acceptable (Old French → Latin → modern). Can explore semantic shift (how meaning changed over time). Can connect to related words in other languages or word families. Include "Did you know?" stage with a surprising or culturally relevant fact.

**Word Shed:** Full treatment. Formal dictionary definition. Two varied example sentences (different contexts). Synonyms and antonyms both expected. Morphology derivatives: up to 5, including cross-word-class forms.

**Starter question:** Language-analytical. "Can you correct the mistakes in last week's words and explain what the error was?" / "Which two words from last week are homophones? Write a sentence using each correctly."

**thisWeeksWordsQ:** Can be analytical. "Does '-ough' make the same sound in all of this week's words?" / "What is the rule for choosing '-cial' or '-tial'?"

---

## Year 6 (ages 10–11, Stage 6)

**Spelling rules:** Words with -cial after a vowel, -ough patterns consolidated, words with silent letters in unexpected positions, etymology-driven spelling (words from Latin/Greek/French), homophones of sophisticated words, words that are frequently confused (e.g. affect/effect), spelling patterns for schwa endings (-er/-or/-ar).

**Words:** Two to five syllables. Largely abstract or technical. Pupils are expected to have met most of these in reading even if not secure in spelling. Words may have significant morphological complexity (antisocial, superficial, controversially).

**Definitions:** Full dictionary format. Always include word class. Include usage notes, connotation, or register where relevant. Can explicitly compare to synonyms or near-synonyms to highlight nuance.

- Good: "Relating to or affecting a particular race of people. (adjective) Often used in discussions about equality and discrimination."
- Good: "Existing at or near the surface; not deep or thorough. (adjective) The opposite of profound. Often used to criticise shallow thinking."

**Sentences (cloze):** 15–25 words. Varied, sophisticated sentence structures. Multiple clauses. Full range of punctuation. Sentences should reflect the register of the word — formal words in formal contexts, not awkwardly placed in simple sentences. Two words from the list may appear in one sentence.

- Good: "It is crucial that everyone attends their facial appointments on time, as the salon is fully booked for the day."

**Morphology matrix:** 6 entries. Explore prefix + root + suffix combinations. Include forms that change word class (social → socialist → socially → antisocial → unsociable). Discuss whether combinations are real words. Note register differences between derivatives (e.g. 'computerise' vs 'computation').

**Word sort categories:** Sophisticated, multi-criteria.

- Sort by phoneme count (6 or fewer / 7 or more)
- Sort by word class of base form
- Carroll diagrams with abstract criteria (has a prefix meaning 'against' / does not; has an /oo/ sound / does not)
- Sort by spelling pattern for schwa ending (-er/-or/-ar)

**Word sort box labels:** Fully technical. Phoneme counts, prefix meanings, grammatical classes. Pupils at Y6 are expected to handle this metalanguage comfortably.

**Etymology:** Richest treatment. Can trace words through three or more languages. Explore semantic narrowing/broadening/shift. Connect to contemporary usage, news contexts, or subject-specific uses (e.g. computer in the context of Turing; superior in the context of Lake Superior). The "Did you know?" stage should be genuinely surprising or culturally rich.

**Word Shed:** The word shed at Y6 may use a more challenging word from the list rather than the most familiar. Synonyms and antonyms must be genuine and precise, not just vaguely related. Morphology derivatives section at Y6 can include "words associated with [topic]" or "examples of [category]" rather than just suffix forms.

**Starter question:** Demands active recall and analysis. "Can you correct these words from last week and explain the mistake?" / "How many syllables are in each of last week's words?" / "Which words have a soft 'c' sound? Sort them."

**thisWeeksWordsQ:** Analytical and rule-seeking. "What letter comes before '-cial' in each word? What pattern do you notice?" / "Does '-ough' make the same sound as last week?"

---

## Summary Table

| Feature | Y2 | Y3 | Y4 | Y5 | Y6 |
| --- | --- | --- | --- | --- | --- |
| Word syllables | 1–2 | 2–3 | 2–3 | 2–4 | 2–5 |
| Sentence length | 6–8 words | 8–12 words | 10–15 words | 12–18 words | 15–25 words |
| Definition style | Plain noun/verb description | Simple, with word class | "When he/she/it..." formula | Dictionary-style with word class | Full dictionary with connotation |
| Morphology scope | -s, -ed, -ing, -er, un- | Add -tion, -ment, re-, mis- | 4–6 incl. some Latin forms | 5–6, explore word class change | 6, full prefix+root+suffix chains |
| Word sort basis | Phonics/phoneme-based | Phonics + simple grammar | Grammatical (noun/verb) | Multi-criteria, Carroll diagrams | Abstract, multi-criteria |
| Etymology | Fun fact only | 2–3 stages, one root | Full 6-stage chain | Full chain, semantic shift | Richest: 3+ languages, cultural depth |
| Starter type | Picture/definition match | Pattern sort / word ladder | Rule application | Error correction + explanation | Syllable count / analytical sort |
| Cloze difficulty | Obvious from context | Clear but needs some inference | Requires word knowledge | Requires genuine understanding | Register-sensitive, may use two words |
| Additional fields | y2Starter, spellingPattern, sentencesAndSynonyms, wordsInAction, wordSpotter | y3Starter, spellingPattern, wordMatch | — | — | — |
| Slide count | 16–18 | 18 | 22 | 22 | 22 |
