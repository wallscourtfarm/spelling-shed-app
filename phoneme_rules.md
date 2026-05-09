# Phoneme Rules — Sound Button System

Rules confirmed by Innes. Apply exactly.

---

## Symbol types

| Symbol | Code | Rule |
|--------|------|------|
| Filled dot `●` | `{t:'dot'}` | One letter making one sound |
| Filled bar `—` | `{t:'line'}` | Two+ letters making one sound (digraph, double consonant) |
| Bezier arc | `{t:'dot', sid:N}` | Split digraph — arc connects two letters; **no dot under either letter** |

---

## Digraphs and double consonants → `t:'line'`

Single phoneme spelled with two or more letters gets a bar:

| Example | Phoneme entry |
|---------|---------------|
| `sh`, `ch`, `th`, `ph` | `{l:'sh', t:'line'}` |
| `ai`, `ea`, `oa`, `ie`, `ue` | `{l:'ie', t:'line'}` |
| `rr` (as in *carries*) | `{l:'rr', t:'line'}` |
| `pp` (as in *supplies*) | `{l:'pp', t:'line'}` |
| `ll`, `tt`, `ss` etc. | same pattern |
| `ck`, `dge`, `tch` | `{l:'ck', t:'line'}` etc. |

**Double consonants** — even though two letters, they make ONE sound, so always `t:'line'`.

---

## Split digraphs → paired `sid` entries, no dot

### How to identify a split digraph

A split digraph is present when a word follows the pattern:  
**vowel + one or more consonants + silent 'e' at the end of the word**

The five split digraph patterns and their sounds:

| Pattern | Sound | Examples |
|---------|-------|---------|
| `a_e` | long /eɪ/ | make, cake, came, plane, shake |
| `e_e` | long /iː/ | these, theme, complete |
| `i_e` | long /aɪ/ | time, bike, write, knife |
| `o_e` | long /əʊ/ | home, note, gnome, globe, phone |
| `u_e` | long /juː/ | cube, tune, flute, rule |

### How to apply the notation

1. Identify the vowel and the final silent 'e' — these form the split digraph pair
2. Give **both** the vowel AND the 'e' `{t:'dot', sid:1}` (same `sid` number)
3. All consonants between them keep their normal notation (dot or line as usual)
4. **Never** give the vowel or 'e' a regular dot — the arc replaces it
5. If a word has **two** split digraphs, use `sid:1` for the first pair and `sid:2` for the second

```javascript
// make — a_e split digraph
[{l:'m',t:'dot'}, {l:'a',t:'dot',sid:1}, {l:'k',t:'dot'}, {l:'e',t:'dot',sid:1}]

// time — i_e split digraph
[{l:'t',t:'dot'}, {l:'i',t:'dot',sid:1}, {l:'m',t:'dot'}, {l:'e',t:'dot',sid:1}]

// gnome — gn digraph + o_e split digraph (two patterns in one word)
[{l:'gn',t:'line'}, {l:'o',t:'dot',sid:1}, {l:'m',t:'dot'}, {l:'e',t:'dot',sid:1}]

// these — e_e split digraph
[{l:'th',t:'line'}, {l:'e',t:'dot',sid:1}, {l:'s',t:'dot'}, {l:'e',t:'dot',sid:1}]

// phone — ph digraph + o_e split digraph
[{l:'ph',t:'line'}, {l:'o',t:'dot',sid:1}, {l:'n',t:'dot'}, {l:'e',t:'dot',sid:1}]
```

### Words that look like split digraphs but are NOT

Some words end in consonant + e without a split digraph:

| Word | Why NOT a split digraph | Correct notation |
|------|------------------------|-----------------|
| `some` | short /ʌ/ vowel — not a long vowel sound | `[{l:'s',t:'dot'},{l:'o',t:'dot'},{l:'m',t:'dot'},{l:'e',t:'dot'}]` |
| `come` | short /ʌ/ vowel | same — no sid |
| `have` | short /æ/ vowel | no sid |
| `love` | short /ʌ/ vowel | no sid |
| `give` | short /ɪ/ vowel | no sid |
| `live` (verb) | short /ɪ/ vowel | no sid |

**Rule: only apply `sid` when the vowel makes its LONG sound because of the magic e.**

### Validation

The total letters covered by all `l` values must equal the full word length.  
Count as a check before finalising: e.g. `gnome` → gn(2)+o(1)+m(1)+e(1) = 5 = word length ✓  
If the count does not match, you have missed a letter — find it and add it.

The arc infrastructure is built into the template and will handle any sid pairs automatically.

---

## Special cases confirmed

| Case | Rule |
|------|------|
| `or` in *worries* | Two separate dots: `{l:'o',t:'dot'}, {l:'r',t:'dot'}` — NOT an `or` digraph here |
| `ur` in *hurries* | Two separate dots: `{l:'u',t:'dot'}, {l:'r',t:'dot'}` — NOT an `ur` digraph here |
| Final `s` pronounced /z/ | Still a single dot — letter identity matters, not sound value |
| `ie` in Y→IES words | Always `{l:'ie', t:'line'}` — digraph bar |

**Reasoning for or/ur:** In *hurries* and *worries*, the `u`/`o` and `r` are in separate syllable positions (hur-ries, wor-ries) so they do not form a vowel digraph — they are independent phonemes.

---

## Worked examples — Y→IES word set

```javascript
carries:  [{l:'c',t:'dot'},{l:'a',t:'dot'},{l:'rr',t:'line'},{l:'ie',t:'line'},{l:'s',t:'dot'}]
tries:    [{l:'t',t:'dot'},{l:'r',t:'dot'},{l:'ie',t:'line'},{l:'s',t:'dot'}]
flies:    [{l:'f',t:'dot'},{l:'l',t:'dot'},{l:'ie',t:'line'},{l:'s',t:'dot'}]
hurries:  [{l:'h',t:'dot'},{l:'u',t:'dot'},{l:'rr',t:'line'},{l:'ie',t:'line'},{l:'s',t:'dot'}]
worries:  [{l:'w',t:'dot'},{l:'o',t:'dot'},{l:'rr',t:'line'},{l:'ie',t:'line'},{l:'s',t:'dot'}]
copies:   [{l:'c',t:'dot'},{l:'o',t:'dot'},{l:'p',t:'dot'},{l:'ie',t:'line'},{l:'s',t:'dot'}]
studies:  [{l:'s',t:'dot'},{l:'t',t:'dot'},{l:'u',t:'dot'},{l:'d',t:'dot'},{l:'ie',t:'line'},{l:'s',t:'dot'}]
replies:  [{l:'r',t:'dot'},{l:'e',t:'dot'},{l:'p',t:'dot'},{l:'l',t:'dot'},{l:'ie',t:'line'},{l:'s',t:'dot'}]
denies:   [{l:'d',t:'dot'},{l:'e',t:'dot'},{l:'n',t:'dot'},{l:'ie',t:'line'},{l:'s',t:'dot'}]
supplies: [{l:'s',t:'dot'},{l:'u',t:'dot'},{l:'pp',t:'line'},{l:'l',t:'dot'},{l:'ie',t:'line'},{l:'s',t:'dot'}]
```

---

## Silent letter words — worked examples

Silent-letter digraphs (`kn`, `gn`, `wr`) are treated as a **single sound** → `t:'line'`.  
Words ending in magic-e also carry a **split digraph** (`o_e`, `a_e`, `i_e`, `u_e`, `e_e`) → use paired `sid` entries.

```javascript
// kn words — silent k, 'kn' treated as one digraph sound → line
knight:  [{l:'kn',t:'line'},{l:'igh',t:'line'},{l:'t',t:'dot'}]
kneel:   [{l:'kn',t:'line'},{l:'ee',t:'line'},{l:'l',t:'dot'}]
knot:    [{l:'kn',t:'line'},{l:'o',t:'dot'},{l:'t',t:'dot'}]
knuckle: [{l:'kn',t:'line'},{l:'u',t:'dot'},{l:'ck',t:'line'},{l:'le',t:'line'}]
knife:   [{l:'kn',t:'line'},{l:'i',t:'dot',sid:1},{l:'f',t:'dot'},{l:'e',t:'dot',sid:1}]
knee:    [{l:'kn',t:'line'},{l:'ee',t:'line'}]

// gn words — silent g, 'gn' treated as one digraph sound → line
// gnome has a split digraph: o_e (long o across the m)
gnome:   [{l:'gn',t:'line'},{l:'o',t:'dot',sid:1},{l:'m',t:'dot'},{l:'e',t:'dot',sid:1}]
gnaw:    [{l:'gn',t:'line'},{l:'aw',t:'line'}]
gnash:   [{l:'gn',t:'line'},{l:'a',t:'dot'},{l:'sh',t:'line'}]
sign:    [{l:'s',t:'dot'},{l:'igh',t:'line'},{l:'n',t:'dot'}]

// wr words — silent w, 'wr' treated as one digraph sound → line
wrist:   [{l:'wr',t:'line'},{l:'i',t:'dot'},{l:'st',t:'line'}]
wreck:   [{l:'wr',t:'line'},{l:'e',t:'dot'},{l:'ck',t:'line'}]
wren:    [{l:'wr',t:'line'},{l:'e',t:'dot'},{l:'n',t:'dot'}]
wrap:    [{l:'wr',t:'line'},{l:'a',t:'dot'},{l:'p',t:'dot'}]
write:   [{l:'wr',t:'line'},{l:'i',t:'dot',sid:1},{l:'t',t:'dot'},{l:'e',t:'dot',sid:1}]

// CRITICAL RULE: letter coverage must equal the full word length.
// Count the letters in all `l` values and verify they sum to the word.
// gnome = gn(2) + o(1) + m(1) + e(1) = 5 ✓
// WRONG: [{l:'gn',t:'line'},{l:'m',t:'dot'}] → only 3 letters covered, 'o' and 'e' missing
