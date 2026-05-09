"""
slides_builder.py
python-pptx port of slides-template.js — Y4/Y5/Y6 teaching deck.
21 slides. No click animations in this version (added separately).
Returns raw PPTX bytes.
"""

import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree


def build_slides(lesson: dict) -> bytes:

    WORDS       = lesson["words"]
    DEFS        = lesson["defs"]
    SENTENCES   = lesson["sentences"]
    CLOZE_ORDER = lesson["clozeOrder"]
    CODE        = lesson["code"]

    def rgb(h):
        return RGBColor.from_string(h)

    C = dict(
        BLUE    = "1E5FA3",
        PINK    = "E91E8C",
        PURPLE  = "7B1FA2",
        GREEN_S = "388E3C",
        CYAN    = "4DD0E1",
        GREEN_B = "57A657",
        BLACK   = "1A1A1A",
        GREY    = "757575",
        WHITE   = "FFFFFF",
        YELLOW  = "F4C430",
        RED     = "C62828",
    )

    SW, SH   = 10.0, 5.625
    HDR_H    = 1.12
    CYAN_H   = 0.06
    CONT_Y   = HDR_H + CYAN_H
    CONT_H   = 4.22
    BAR_Y    = CONT_Y + CONT_H
    FONT     = "Calibri"

    prs = Presentation()
    prs.slide_width  = Inches(SW)
    prs.slide_height = Inches(SH)
    BLANK = prs.slide_layouts[6]

    # ── Core drawing helpers ──────────────────────────────────────────────────

    def _bodyPr(txBody, valign="middle", margin_in=0.05):
        bp = txBody.find(qn('a:bodyPr'))
        if bp is None:
            bp = etree.SubElement(txBody, qn('a:bodyPr'))
        bp.set('anchor', {"top": "t", "middle": "ctr", "bottom": "b"}.get(valign, "ctr"))
        m = str(Inches(margin_in))
        for a in ('lIns', 'rIns', 'tIns', 'bIns'):
            bp.set(a, m)

    def rect(slide, x, y, w, h, fill, line=None, lpt=0, radius=None):
        shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill)
        if line and lpt > 0:
            shape.line.color.rgb = rgb(line)
            shape.line.width = Pt(lpt)
        else:
            shape.line.fill.background()
        if radius is not None:
            sp = shape._element
            spPr = sp.find(qn('p:spPr'))
            pg = spPr.find(qn('a:prstGeom'))
            if pg is None:
                pg = etree.SubElement(spPr, qn('a:prstGeom'))
            pg.set('prst', 'roundRect')
            al = pg.find(qn('a:avLst'))
            if al is None:
                al = etree.SubElement(pg, qn('a:avLst'))
            for av in list(al):
                al.remove(av)
            adj = int(min(radius / min(w, h), 0.5) * 100000)
            gd = etree.SubElement(al, qn('a:gd'))
            gd.set('name', 'adj')
            gd.set('fmla', f'val {adj}')
        return shape

    def triangle(slide, x, y, w, h, fill, line=None, lpt=1):
        shape = slide.shapes.add_shape(5, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill)
        if line and lpt > 0:
            shape.line.color.rgb = rgb(line)
            shape.line.width = Pt(lpt)
        else:
            shape.line.fill.background()
        return shape

    def txt(slide, text, x, y, w, h, *,
            font=FONT, size=14, bold=False, italic=False,
            color="1A1A1A", align="left", valign="middle",
            wrap=True, margin=0.04, shadow=False):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        _bodyPr(tf._txBody, valign=valign, margin_in=margin)
        p = tf.paragraphs[0]
        p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                       "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)
        run = p.add_run()
        run.text = text
        rf = run.font
        rf.name = font
        rf.size = Pt(size)
        rf.bold = bold
        rf.italic = italic
        rf.color.rgb = rgb(color)
        return tb

    def rich_txt(slide, runs, x, y, w, h, *,
                 font=FONT, align="left", valign="middle", margin=0.04):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        _bodyPr(tf._txBody, valign=valign, margin_in=margin)
        p = tf.paragraphs[0]
        p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER}.get(align, PP_ALIGN.LEFT)
        for text, opts in runs:
            run = p.add_run()
            run.text = text
            rf = run.font
            rf.name = font
            rf.size = Pt(opts.get("size", 14))
            rf.bold = opts.get("bold", False)
            rf.italic = opts.get("italic", False)
            rf.color.rgb = rgb(opts.get("color", "1A1A1A"))
        return tb

    def table(slide, rows, x, y, w, h, col_widths, row_heights,
              border_hex="444444", border_pt=1.5, font=FONT):
        nr, nc = len(rows), len(rows[0])
        ts = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(h))
        tbl = ts.table
        for ci, cw in enumerate(col_widths):
            tbl.columns[ci].width = Inches(cw)
        rhl = row_heights if isinstance(row_heights, list) else [row_heights] * nr
        for ri, rh in enumerate(rhl):
            tbl.rows[ri].height = Inches(rh)
        for ri, row in enumerate(rows):
            for ci, cd in enumerate(row):
                cell = tbl.cell(ri, ci)
                text   = cd.get("text", "")
                bold   = cd.get("bold", False)
                italic = cd.get("italic", False)
                color  = cd.get("color", "1A1A1A")
                fill   = cd.get("fill", "FFFFFF")
                align  = cd.get("align", "left")
                size   = cd.get("size", 14)
                valign = cd.get("valign", "middle")
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(fill)
                tf = cell.text_frame
                tf.word_wrap = True
                _bodyPr(tf._txBody, valign=valign, margin_in=0.04)
                p = tf.paragraphs[0]
                p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                               "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)
                if text:
                    run = p.add_run()
                    run.text = text
                    rf = run.font
                    rf.name = font
                    rf.size = Pt(size)
                    rf.bold = bold
                    rf.italic = italic
                    rf.color.rgb = rgb(color)
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                for edge in ('lnL', 'lnR', 'lnT', 'lnB'):
                    ln = tcPr.find(qn(f'a:{edge}'))
                    if ln is None:
                        ln = etree.SubElement(tcPr, qn(f'a:{edge}'))
                    ln.set('w', str(int(Pt(border_pt))))
                    sf = ln.find(qn('a:solidFill'))
                    if sf is None:
                        sf = etree.SubElement(ln, qn('a:solidFill'))
                    sc = sf.find(qn('a:srgbClr'))
                    if sc is None:
                        sc = etree.SubElement(sf, qn('a:srgbClr'))
                    sc.set('val', border_hex)
        return ts

    # ── Frame (standard header used by most slides) ───────────────────────────

    def add_frame(slide, activity_type, activity_label, question, slide_num):
        ind_col = "F4C430" if activity_type == "Independent" else "66BB6A"
        rect(slide, 0.1, 0.1, 1.35, 0.92, ind_col)
        txt(slide, activity_type, 0.1, 0.1, 1.35, 0.92,
            size=12, bold=True, color=C["WHITE"], align="center", margin=0)
        txt(slide, activity_label, 1.55, 0.06, 7.5, 0.44,
            size=19, color=C["BLACK"], align="center")
        txt(slide, question, 1.55, 0.5, 7.5, 0.6,
            size=22, bold=True, color=C["BLUE"], align="center")
        txt(slide, slide_num, 8.85, 0.78, 1.05, 0.28,
            size=14, color=C["GREY"], align="right")
        rect(slide, 0, HDR_H, SW, CYAN_H, C["CYAN"])
        rect(slide, 0, BAR_Y, SW, SH - BAR_Y, C["GREEN_B"])

    def fit_font(text, box_w, max_pt=36, min_pt=14):
        n = len(text.replace(" ", "")) or len(text)
        req = (box_w * 72) / (n * 0.54)
        return round(min(max_pt, max(min_pt, req)))

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 1 — Title
    # ════════════════════════════════════════════════════════════════════════

    def slide1():
        s = prs.slides.add_slide(BLANK)
        rect(s, 0, 0, SW, 3.6, "87CEEB")
        rect(s, 0, 3.3, SW, 2.325, C["WHITE"])
        rect(s, 0, BAR_Y, SW, SH - BAR_Y, C["GREEN_B"])
        txt(s, "Spelling Shed", 1.5, 0.15, 7.0, 1.4,
            size=64, bold=True, color=C["YELLOW"], align="center")

        def badge(slide, text, x, y, w, h):
            rect(slide, x, y, w, h, C["WHITE"], "BBBBBB", 1.5, radius=0.15)
            txt(slide, text, x, y, w, h, size=20, bold=True,
                color=C["BLACK"], align="center", margin=0)

        badge(s, "Stage: " + lesson["stage"], 3.2, 1.75, 1.6, 0.75)
        badge(s, "Lesson: " + CODE,           5.2, 1.75, 1.6, 0.75)

        obj = f"To spell words: {lesson['rule']}"
        txt(s, obj, 0.5, 3.42, 9.0, 0.42,
            size=18, bold=True, color=C["BLACK"], align="center")
        txt(s, "This week's words:  " + ", ".join(WORDS),
            0.5, 4.35, 9.0, 0.55, size=17, color=C["BLACK"], align="center")

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 2 — Starter (blank)
    # ════════════════════════════════════════════════════════════════════════

    def slide2():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Whole Group", "Starter",
                  lesson["starter"]["question"], f"{CODE}.1")
        txt(s, "The root or base word may need to change.",
            1.0, CONT_Y + 0.15, 8.0, 0.4,
            size=18, color=C["BLACK"], align="center")

        starters = lesson["starter"]["words"]
        cell_w, cell_h = 9.0 / 3, 1.5
        sx, sy = 0.5, CONT_Y + 0.65
        label = lesson["starter"]["answerLabel"]

        for i, word in enumerate(starters):
            col, row = i % 3, i // 3
            x = sx + col * cell_w
            y = sy + row * cell_h
            txt(s, word, x, y, cell_w, 0.8,
                size=44, color=C["BLACK"], align="center")
            txt(s, f"{label} ______________________",
                x + 0.15, y + 1.1, cell_w - 0.3, 0.3,
                size=14, color=C["BLACK"], align="left")

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 3 — Starter (answers)
    # ════════════════════════════════════════════════════════════════════════

    def slide3():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Whole Group", "Starter",
                  "How did the base word change?", f"{CODE}.2")

        starters = lesson["starter"]["words"]
        answers  = lesson["starter"]["answers"]
        cell_w   = 9.0 / 3
        sx, sy   = 0.5, CONT_Y + 0.15
        BASE_H, ANS_H, RULE_H = 0.52, 0.50, 0.28
        PAIR_H   = BASE_H + ANS_H + RULE_H
        row_gap  = 0.20

        fs = max(14, min(28, int((cell_w * 72) / (max(len(w) for w in answers) * 0.54))))

        for i, (word, ans) in enumerate(zip(starters, answers)):
            col, row = i % 3, i // 3
            x = sx + col * cell_w
            y = sy + row * (PAIR_H + row_gap)
            txt(s, word + "  →", x, y, cell_w, BASE_H,
                size=fs, color=C["BLACK"], align="center")
            txt(s, ans, x, y + BASE_H, cell_w, ANS_H,
                size=fs, bold=True, color=C["PINK"], align="center")
            note = lesson["starter"].get("perPairNote", "")
            txt(s, note, x, y + BASE_H + ANS_H, cell_w, RULE_H,
                size=11, italic=True, color=C["GREY"], align="center")

        rule_y = sy + 2 * (PAIR_H + row_gap) + 0.14
        rect(s, 0.5, rule_y, 9.0, 0.45, "FFF9C4", C["YELLOW"], 1.5)
        txt(s, lesson["starter"]["ruleBox"], 0.5, rule_y, 9.0, 0.45,
            size=16, bold=True, color=C["BLACK"], align="center", margin=0)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 4 — This Week's Words
    # ════════════════════════════════════════════════════════════════════════

    def slide4():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Whole Group", "This Week's Words",
                  lesson["thisWeeksWordsQ"], f"{CODE}.3")

        cell_w = 9.0 / 5
        fs = min(fit_font(w, cell_w, 32, 16) for w in WORDS)

        for ri, row_words in enumerate([WORDS[:5], WORDS[5:]]):
            for ci, w in enumerate(row_words):
                x = 0.5 + ci * cell_w
                y = CONT_Y + 0.3 + ri * 1.35
                txt(s, w, x, y, cell_w, 0.85,
                    size=fs, color=C["BLACK"], align="center", valign="bottom")
                rect(s, x + 0.3, y + 0.88, cell_w - 0.6, 0.07, C["PINK"])

        txt(s, lesson["thisWeeksWordsPrompt"],
            0.5, CONT_Y + 3.05, 4.3, 0.4,
            size=15, bold=True, color=C["BLUE"], align="left")
        txt(s, lesson["thisWeeksWordsExplanation"],
            4.8, CONT_Y + 3.05, 4.8, 0.55,
            size=14, color=C["BLACK"], align="left")

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 5 — Etymology
    # ════════════════════════════════════════════════════════════════════════

    def slide5():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Whole Group", "Etymology",
                  "Which of this week's words is this?", f"{CODE}.4")

        etym = lesson["etymology"]
        txt(s, etym["word"], 1.0, CONT_Y + 0.1, 8.0, 0.85,
            size=60, color=C["PINK"], align="center")

        bw = 2.85
        for i, click in enumerate(etym["clicks"][:3]):
            x = 0.5 + i * (bw + 0.15)
            rect(s, x, CONT_Y + 1.1, bw, 1.35, C["WHITE"], C["PINK"], 2)
            txt(s, click["label"] + " — " + click["body"],
                x + 0.12, CONT_Y + 1.15, bw - 0.24, 1.25,
                size=13, color=C["BLACK"], align="center")

        txt(s, etym["baseForm"], 1.25, CONT_Y + 2.55, 2.5, 0.40,
            size=26, color=C["BLUE"], align="center")
        txt(s, "↙                          ↘",
            1.25, CONT_Y + 2.95, 2.5, 0.22,
            size=14, color=C["BLACK"], align="center")

        if len(etym["clicks"]) > 3:
            txt(s, etym["clicks"][3]["label"] + ":\n" + etym["clicks"][3]["body"],
                0.3, CONT_Y + 3.19, 2.1, 0.84,
                size=10, color=C["GREY"], align="center")
        if len(etym["clicks"]) > 4:
            txt(s, etym["clicks"][4]["label"] + ":\n" + etym["clicks"][4]["body"],
                2.6, CONT_Y + 3.19, 2.4, 0.84,
                size=10, color=C["GREY"], align="center")
        if len(etym["clicks"]) > 5:
            txt(s, etym["clicks"][5]["label"] + ":",
                5.2, CONT_Y + 2.55, 4.3, 0.38,
                size=14, color=C["BLACK"], align="left")
            txt(s, etym["clicks"][5]["body"],
                5.2, CONT_Y + 2.95, 4.3, 1.0,
                size=11, color=C["BLACK"], align="left", valign="top")

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 6 — Syllable Count
    # ════════════════════════════════════════════════════════════════════════

    def slide6():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Whole Group", "Syllable Count",
                  "How many syllables are in this week's words?", f"{CODE}.5")

        cell_w, cell_h = 9.0 / 5, 1.6
        sx, sy = 0.5, CONT_Y + 0.2
        sc = lesson["syllableCounts"]
        fs = min(fit_font(w, cell_w, 28, 14) for w in WORDS)

        for i, w in enumerate(WORDS):
            col, row = i % 5, i // 5
            x = sx + col * cell_w
            y = sy + row * cell_h
            n = sc.get(w, 1)
            txt(s, w, x, y, cell_w, 0.8,
                size=fs, color=C["BLACK"], align="center")
            num_col = C["RED"] if n == 1 else "388E3C"
            txt(s, str(n), x, y + 0.82, cell_w, 0.42,
                size=26, bold=True, color=num_col, align="center")
            txt(s, "syllable" if n == 1 else "syllables",
                x, y + 1.24, cell_w, 0.26,
                size=11, color=C["GREY"], align="center")

        txt(s, "Does the length of the word affect the number of syllables it has?",
            0.5, BAR_Y - 0.48, 9.0, 0.35,
            size=13, italic=True, color=C["GREY"], align="center")

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 7 — Word Sort (blank)
    # ════════════════════════════════════════════════════════════════════════

    def slide7():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Whole Group", "Word Sort",
                  lesson["wordSortQ"], f"{CODE}.6")

        for ri, row_words in enumerate([WORDS[:5], WORDS[5:]]):
            txt(s, "          ".join(row_words),
                0.5, CONT_Y + 0.15 + ri * 0.5, 9.0, 0.48,
                size=22, color=C["BLACK"], align="center")

        BOX_Y, BOX_H = CONT_Y + 1.3, 2.05
        ws = lesson["wordSort"]

        rect(s, 0.35, BOX_Y, 4.45, BOX_H, C["WHITE"], C["PURPLE"], 3, radius=0.15)
        txt(s, ws["box1label"] + "\n" + ws["box1sub"],
            0.45, BOX_Y + 0.08, 4.25, 0.75,
            size=16, bold=True, color=C["PURPLE"], align="center")

        rect(s, 5.2, BOX_Y, 4.45, BOX_H, C["WHITE"], C["GREEN_S"], 3, radius=0.15)
        txt(s, ws["box2label"] + "\n" + ws["box2sub"],
            5.3, BOX_Y + 0.08, 4.25, 0.75,
            size=16, bold=True, color=C["GREEN_S"], align="center")

        txt(s, ws["hint"], 0.5, BOX_Y + BOX_H + 0.12, 9.0, 0.32,
            size=12, italic=True, color=C["GREY"], align="center")

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 8 — Word Sort (answers)
    # ════════════════════════════════════════════════════════════════════════

    def slide8():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Whole Group", "Word Sort", "Answers", f"{CODE}.7")

        ws = lesson["wordSort"]
        txt(s, ws["answerNote"], 0.5, CONT_Y + 0.08, 9.0, 0.38,
            size=14, color=C["BLACK"], align="center")

        BOX_Y, BOX_H = CONT_Y + 0.55, 3.05
        verb_only = ws["verbOnly"]
        verb_noun = ws["verbNoun"]

        rect(s, 0.35, BOX_Y, 4.45, BOX_H, C["WHITE"], C["PURPLE"], 3, radius=0.15)
        txt(s, ws["box1label"] + "  " + ws["box1sub"],
            0.45, BOX_Y + 0.08, 4.25, 0.35,
            size=14, bold=True, color=C["PURPLE"], align="center")
        for i, w in enumerate(verb_only):
            txt(s, w, 0.55, BOX_Y + 0.50 + i * 0.52, 4.15, 0.40,
                size=13, bold=True, color=C["PINK"], align="center")

        rect(s, 5.2, BOX_Y, 4.45, BOX_H, C["WHITE"], C["GREEN_S"], 3, radius=0.15)
        txt(s, ws["box2label"], 5.3, BOX_Y + 0.08, 4.25, 0.35,
            size=14, bold=True, color=C["GREEN_S"], align="center")
        for i, item in enumerate(verb_noun):
            rich_txt(s, [
                (item["word"], {"size": 12, "bold": True,  "color": C["PINK"]}),
                ("  " + item["eg"], {"size": 10, "italic": True, "color": C["GREY"]}),
            ], 5.4, BOX_Y + 0.47 + i * 0.362, 4.15, 0.34, valign="middle")

        txt(s, ws["exampleLine"], 0.5, BOX_Y + BOX_H + 0.12, 9.0, 0.30,
            size=12, color=C["BLACK"], align="center")

    # ── Sound button drawing ──────────────────────────────────────────────────

    def draw_sound_buttons(slide, word, cx, wy, font_size):
        """
        Draw sound buttons centred at cx, top at wy.
        Letters in a no-border table; dots under single phonemes; bars under digraphs.
        """
        phonemes = lesson["phonemes"].get(word)
        if not phonemes:
            return

        CELL_W = font_size / 28 * 0.40
        ROW_H  = font_size * 1.4 / 72
        DOT    = max(0.040, font_size * 0.084 / 28)
        LINE_H = max(0.018, font_size * 0.046 / 28)
        PAD    = CELL_W * 0.07
        GAP    = 0.04

        n_cells = sum(len(g["l"]) for g in phonemes)
        total_w = n_cells * CELL_W
        table_x = cx - total_w / 2
        sym_y   = wy + ROW_H + GAP
        line_y  = sym_y + (DOT - LINE_H) / 2

        # Letter table — one cell per letter, no borders
        ts = slide.shapes.add_table(
            1, n_cells,
            Inches(table_x), Inches(wy),
            Inches(total_w), Inches(ROW_H)
        )
        tbl = ts.table
        for ci in range(n_cells):
            tbl.columns[ci].width = Inches(CELL_W)
        tbl.rows[0].height = Inches(ROW_H)

        ci = 0
        for g in phonemes:
            for ch in g["l"]:
                cell = tbl.cell(0, ci)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(C["WHITE"])
                tf = cell.text_frame
                _bodyPr(tf._txBody, valign="middle", margin_in=0)
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                run = p.add_run()
                run.text = ch
                run.font.name = FONT
                run.font.size = Pt(int(font_size))
                run.font.bold = False
                run.font.color.rgb = rgb(C["BLACK"])
                # Remove all borders
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                for edge in ('lnL', 'lnR', 'lnT', 'lnB'):
                    ln = tcPr.find(qn(f'a:{edge}'))
                    if ln is None:
                        ln = etree.SubElement(tcPr, qn(f'a:{edge}'))
                    for sf in ln.findall(qn('a:solidFill')):
                        ln.remove(sf)
                    if ln.find(qn('a:noFill')) is None:
                        etree.SubElement(ln, qn('a:noFill'))
                ci += 1

        # Per-group geometry
        col = 0
        gd = []
        for g in phonemes:
            n  = len(g["l"])
            gx = table_x + col * CELL_W
            gw = n * CELL_W
            gd.append({**g, "gx": gx, "gw": gw, "gc": gx + gw / 2})
            col += n

        # Digraph bars
        for g in gd:
            if g["t"] == "line":
                rect(slide, g["gx"] + PAD, line_y,
                     g["gw"] - 2 * PAD, LINE_H, C["BLACK"])

        # Dots under single phonemes (not split digraphs)
        for g in gd:
            if g["t"] == "dot" and "sid" not in g:
                # oval shape (type 9)
                shape = slide.shapes.add_shape(
                    9,  # MSO_SHAPE_TYPE.OVAL
                    Inches(g["gc"] - DOT / 2), Inches(sym_y),
                    Inches(DOT), Inches(DOT)
                )
                shape.fill.solid()
                shape.fill.fore_color.rgb = rgb(C["BLACK"])
                shape.line.fill.background()

        # Split digraphs — draw a simple arc using a curved connector approximation
        splits = {}
        for g in gd:
            if "sid" in g:
                sid = g["sid"]
                if sid not in splits:
                    splits[sid] = []
                splits[sid].append(g)

        for sid, pair in splits.items():
            if len(pair) < 2:
                continue
            pair.sort(key=lambda g: g["gx"])
            x1, x2 = pair[0]["gc"], pair[-1]["gc"]
            arc_w = x2 - x1
            arc_h = 0.18
            arc_x = x1
            arc_y = sym_y - arc_h
            # Draw as a thin rectangle arc approximation
            shape = slide.shapes.add_shape(
                1,  # rectangle
                Inches(arc_x), Inches(arc_y),
                Inches(arc_w), Inches(LINE_H)
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = rgb(C["BLACK"])
            shape.line.fill.background()

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 9 — Syllable & Phoneme Map (worked examples)
    # ════════════════════════════════════════════════════════════════════════

    def slide9():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Whole Group", "Syllable and Phoneme Map",
                  "How to mark the syllable breaks and sound buttons", f"{CODE}.8")

        map_words = lesson["wordMaps"]["words"][:3]
        syllables = lesson["wordMaps"]["syllables"]
        sc        = lesson["syllableCounts"]

        cw, sx, sy = 2.9, 0.5, CONT_Y + 0.25
        fs = min(fit_font(w, cw, 36, 18) for w in map_words)

        for i, w in enumerate(map_words):
            x  = sx + i * (cw + 0.25)
            cx = x + cw / 2
            brk = syllables.get(w, w)
            n   = sc.get(w, 1)

            txt(s, w, x, sy, cw, 0.7,
                size=fs, color=C["PINK"], align="center")

            txt(s, "Syllable Breaks", x, sy + 0.75, cw, 0.3,
                size=12, bold=True, color=C["BLUE"], align="center")
            rect(s, x, sy + 1.08, cw, 0.5, C["WHITE"], C["BLUE"], 1.5)
            txt(s, brk, x, sy + 1.08, cw, 0.5,
                size=20, color=C["BLACK"], align="center", margin=0)
            txt(s, f"{n} syllable{'s' if n > 1 else ''}",
                x, sy + 1.62, cw, 0.28,
                size=12, italic=True, color=C["GREY"], align="center")

            txt(s, "Sound Buttons", x, sy + 1.95, cw, 0.3,
                size=12, bold=True, color=C["PINK"], align="center")
            draw_sound_buttons(s, w, cx, sy + 2.32, 22)
            txt(s, "● = single sound   — = letters making one sound",
                x, sy + 2.88, cw, 0.22,
                size=9, italic=True, color=C["GREY"], align="center")

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 10 — Word Maps (blank table)
    # ════════════════════════════════════════════════════════════════════════

    def slide10():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Independent", "Word Maps",
                  "Read the word. Write syllable breaks and sound buttons. Rewrite the word.",
                  f"{CODE}.9")

        map_words = lesson["wordMaps"]["words"]
        hdr = [
            {"text": "Words",           "bold": True, "fill": C["WHITE"], "align": "center"},
            {"text": "Syllable Breaks", "bold": True, "fill": C["WHITE"], "align": "center"},
            {"text": "Sound Buttons",   "bold": True, "fill": C["WHITE"], "align": "center"},
            {"text": "My Word",         "bold": True, "fill": C["WHITE"], "align": "center"},
        ]
        rows = [
            [{"text": w, "fill": C["WHITE"], "align": "center", "size": 18},
             {"text": "", "fill": C["WHITE"]},
             {"text": "", "fill": C["WHITE"]},
             {"text": "", "fill": C["WHITE"]}]
            for w in map_words
        ]
        rh = [0.5] + [0.57] * len(rows)
        table(s, [hdr] + rows,
              0.35, CONT_Y + 0.1, 9.3, CONT_H - 0.15,
              [2.0, 2.5, 2.9, 1.9], rh)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 11 — Word Maps (answers)
    # ════════════════════════════════════════════════════════════════════════

    def slide11():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Whole Group", "Word Maps", "Answers", f"{CODE}.10")

        map_words = lesson["wordMaps"]["words"]
        syllables = lesson["wordMaps"]["syllables"]
        hdr = [
            {"text": "Words",           "bold": True, "fill": C["WHITE"], "align": "center"},
            {"text": "Syllable Breaks", "bold": True, "fill": C["WHITE"], "align": "center"},
            {"text": "Sound Buttons",   "bold": True, "fill": C["WHITE"], "align": "center"},
            {"text": "My Word",         "bold": True, "fill": C["WHITE"], "align": "center"},
        ]
        rows = [
            [{"text": w,                   "fill": C["WHITE"], "align": "center", "size": 18},
             {"text": syllables.get(w, w), "fill": C["WHITE"], "align": "center", "size": 16},
             {"text": "",                  "fill": C["WHITE"]},
             {"text": w,                   "fill": C["WHITE"], "align": "center", "size": 18, "color": C["PINK"]}]
            for w in map_words
        ]
        rh = [0.5] + [0.57] * len(rows)
        ts = table(s, [hdr] + rows,
                   0.35, CONT_Y + 0.1, 9.3, CONT_H - 0.15,
                   [2.0, 2.5, 2.9, 1.9], rh)

        # Draw sound buttons over the Sound Buttons column (col index 2)
        TABLE_X   = 0.35
        SB_COL_X  = TABLE_X + 2.0 + 2.5   # 4.85"
        SB_COL_W  = 2.9
        HDR_H_IN  = 0.5
        ROW_H_IN  = 0.57
        SB_FS     = 13

        for ri, w in enumerate(map_words):
            cell_top_y = CONT_Y + 0.1 + HDR_H_IN + ri * ROW_H_IN
            cell_cx    = SB_COL_X + SB_COL_W / 2

            # Estimate height of sound button display to centre it
            row_h_sb   = SB_FS * 1.4 / 72
            dot_h      = max(0.040, SB_FS * 0.084 / 28)
            total_sb_h = row_h_sb + 0.04 + dot_h
            word_top_y = cell_top_y + (ROW_H_IN - total_sb_h) / 2

            draw_sound_buttons(s, w, cell_cx, word_top_y, SB_FS)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 12 — Definitions and In a Sentence (blank)
    # ════════════════════════════════════════════════════════════════════════

    def slide12():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Independent", "Definitions and In a Sentence",
                  "Read the word and its definition, then write your own sentence for each word.",
                  f"{CODE}.11")

        hdr = [
            {"text": "Word",          "bold": True, "fill": C["WHITE"], "align": "center"},
            {"text": "Definition",    "bold": True, "fill": C["WHITE"], "align": "center"},
            {"text": "In a Sentence", "bold": True, "fill": C["WHITE"], "align": "center"},
        ]
        rows = [
            [{"text": w,       "fill": C["WHITE"], "align": "center", "size": 13},
             {"text": DEFS[w], "fill": C["WHITE"], "align": "left",   "size": 9},
             {"text": "",      "fill": C["WHITE"]}]
            for w in WORDS
        ]
        rh = [0.4] + [0.36] * len(rows)
        table(s, [hdr] + rows,
              0.35, CONT_Y + 0.05, 9.3, CONT_H - 0.1,
              [1.55, 3.55, 4.2], rh)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 13 — Definitions and In a Sentence (answers)
    # ════════════════════════════════════════════════════════════════════════

    def slide13():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Independent", "Definitions and In a Sentence",
                  "Answers", f"{CODE}.12")

        hdr = [
            {"text": "Word",          "bold": True, "fill": C["WHITE"], "align": "center"},
            {"text": "Definition",    "bold": True, "fill": C["WHITE"], "align": "center"},
            {"text": "In a Sentence", "bold": True, "fill": C["WHITE"], "align": "center"},
        ]
        rows = [
            [{"text": w,            "fill": C["WHITE"], "align": "center", "size": 13},
             {"text": DEFS[w],      "fill": C["WHITE"], "align": "left",   "size": 9},
             {"text": SENTENCES[w], "fill": C["WHITE"], "align": "left",   "size": 9,
              "color": C["PINK"]}]
            for w in WORDS
        ]
        rh = [0.4] + [0.36] * len(rows)
        table(s, [hdr] + rows,
              0.35, CONT_Y + 0.05, 9.3, CONT_H - 0.1,
              [1.55, 3.4, 4.35], rh)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 14 — Spell Check (blank)
    # ════════════════════════════════════════════════════════════════════════

    def slide14():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Independent", "Spell Check",
                  "Circle the correct spelling of each word.", f"{CODE}.13")

        rows = [
            [{"text": opt, "fill": C["WHITE"], "align": "center", "size": 16}
             for opt in row["opts"]]
            for row in lesson["spellData"]
        ]
        table(s, rows,
              0.35, CONT_Y + 0.05, 9.3, CONT_H - 0.05,
              [3.1, 3.1, 3.1], 0.4)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 15 — Spell Check (answers)
    # ════════════════════════════════════════════════════════════════════════

    def slide15():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Independent", "Spell Check", "Answers", f"{CODE}.14")

        rows = [
            [{"text": opt,
              "fill": C["WHITE"],
              "align": "center",
              "size": 16,
              "bold": j == row["correct"],
              "color": C["PINK"] if j == row["correct"] else C["BLACK"]}
             for j, opt in enumerate(row["opts"])]
            for row in lesson["spellData"]
        ]
        table(s, rows,
              0.35, CONT_Y + 0.05, 9.3, CONT_H - 0.05,
              [3.1, 3.1, 3.1], 0.4)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 16 — Cloze (blank)
    # ════════════════════════════════════════════════════════════════════════

    def slide16():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Independent", "Cloze",
                  "Use the words from the word bank to complete each sentence.",
                  f"{CODE}.15")

        row_h = (CONT_H - 0.10) / len(WORDS)
        wb_rows = [[{"text": w, "fill": C["WHITE"], "align": "center",
                     "size": 12, "valign": "middle"}]
                   for w in WORDS]
        table(s, wb_rows,
              0.30, CONT_Y + 0.05, 1.50, CONT_H - 0.10,
              [1.50], row_h)

        BLANK_STR = "_______________"
        for i, w in enumerate(CLOZE_ORDER):
            sentence = SENTENCES[w]
            idx = sentence.find(w)
            before = sentence[:idx] if idx != -1 else sentence
            after  = sentence[idx + len(w):] if idx != -1 else ""
            rich_txt(s, [
                (before,     {"size": 12, "color": C["BLACK"]}),
                (BLANK_STR,  {"size": 12, "color": C["BLACK"]}),
                (after,      {"size": 12, "color": C["BLACK"]}),
            ], 1.95, CONT_Y + 0.05 + i * row_h, 7.75, row_h,
               align="left", valign="middle", margin=0.03)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 17 — Cloze (answers)
    # ════════════════════════════════════════════════════════════════════════

    def slide17():
        s = prs.slides.add_slide(BLANK)
        add_frame(s, "Independent", "Cloze", "Answers", f"{CODE}.16")

        row_h = (CONT_H - 0.10) / len(WORDS)
        wb_rows = [[{"text": w, "fill": C["WHITE"], "align": "center",
                     "size": 12, "valign": "middle"}]
                   for w in WORDS]
        table(s, wb_rows,
              0.30, CONT_Y + 0.05, 1.50, CONT_H - 0.10,
              [1.50], row_h)

        for i, w in enumerate(CLOZE_ORDER):
            sentence = SENTENCES[w]
            idx = sentence.find(w)
            before = sentence[:idx] if idx != -1 else sentence
            after  = sentence[idx + len(w):] if idx != -1 else ""
            rich_txt(s, [
                (before, {"size": 12, "color": C["BLACK"]}),
                (w,      {"size": 12, "bold": True, "color": C["PINK"]}),
                (after,  {"size": 12, "color": C["BLACK"]}),
            ], 1.95, CONT_Y + 0.05 + i * row_h, 7.75, row_h,
               align="left", valign="middle", margin=0.03)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 18 — Word Shed (blank)
    # ════════════════════════════════════════════════════════════════════════

    def _word_shed_structure(s, show_answers=False):
        """Shared structure for slides 18 and 19."""
        rect(s, 0, BAR_Y, SW, SH - BAR_Y, C["GREEN_B"])
        rect(s, 0.1, 0.05, 1.3, 0.75, C["YELLOW"])
        txt(s, "Independent", 0.1, 0.05, 1.3, 0.75,
            size=11, bold=True, color=C["WHITE"], align="center", margin=0)
        slide_num = f"{CODE}.{'18' if show_answers else '17'}"
        txt(s, slide_num, 8.8, 0.6, 1.1, 0.28,
            size=14, color=C["GREY"], align="right")

        triangle(s, 0.6, 0.02, 8.8, 0.95, "A0A0A0", "808080", 1)
        rect(s, 0.6, 0.88, 8.8, 4.45, "C4956A", "8B5E3C", 2)
        rect(s, 0.88, 1.08, 8.24, 4.12, C["WHITE"], "8B5E3C", 1)
        txt(s, "Word Shed", 2.5, 0.08, 5.0, 0.7,
            size=30, bold=True, color=C["YELLOW"], align="center")

        mid_x, mid_y = 5.0, 3.14
        rect(s, mid_x - 0.02, 1.08, 0.04, 4.12, "AAAAAA")
        rect(s, 0.88, mid_y - 0.02, 8.24, 0.04, "AAAAAA")
        rect(s, mid_x - 1.1, mid_y - 0.3, 2.2, 0.6, C["WHITE"], C["BLACK"], 1.5)
        txt(s, lesson["wordShed"]["baseWord"],
            mid_x - 1.1, mid_y - 0.3, 2.2, 0.6,
            size=18, bold=True, color=C["BLACK"], align="center", margin=0)

        ws = lesson["wordShed"]
        sections = [
            ("Definition",     0.95, 1.12, ws["def"]      if show_answers else ""),
            ("In a Sentence",  mid_x + 0.1, 1.12, ws["sentence"]  if show_answers else ""),
            ("Rhymes With...", 0.95, mid_y + 0.42, ws["rhymes"]   if show_answers else ""),
            ("Add Prefixes or Suffixes", mid_x + 0.1, mid_y + 0.42, ws["morphology"] if show_answers else ""),
        ]
        body_coords = [
            (0.95, 1.5, 3.85, 1.5),
            (mid_x + 0.1, 1.5, 3.85, 1.5),
            (0.95, mid_y + 0.8, 3.85, 1.55),
            (mid_x + 0.1, mid_y + 0.8, 3.85, 1.55),
        ]
        for (label, lx, ly, body_text), (bx, by, bw, bh) in zip(sections, body_coords):
            txt(s, label, lx, ly, 3.85, 0.35,
                size=14, bold=True, color=C["BLACK"], valign="middle")
            if body_text:
                txt(s, body_text, bx, by, bw, bh,
                    size=13, color=C["PINK"], valign="top")

    def slide18():
        s = prs.slides.add_slide(BLANK)
        _word_shed_structure(s, show_answers=False)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 19 — Word Shed (answers)
    # ════════════════════════════════════════════════════════════════════════

    def slide19():
        s = prs.slides.add_slide(BLANK)
        _word_shed_structure(s, show_answers=True)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 20 — Morphology Matrix (blank)
    # ════════════════════════════════════════════════════════════════════════

    def _morph_matrix(s, show_answers=False):
        add_frame(s, "Whole Group", "Morphology Matrix",
                  "How many new words can you create by adding suffix(es)?",
                  f"{CODE}.{'20' if show_answers else '19'}")

        mm = lesson["morphMatrix"]

        txt(s, "Base Word", 0.4, CONT_Y + 0.15, 4.2, 0.38,
            size=20, bold=True, color=C["GREEN_S"], align="center")
        rect(s, 0.4, CONT_Y + 0.55, 4.2, 3.45, C["WHITE"], C["GREEN_S"], 2.5)
        txt(s, mm["baseWord"], 0.4, CONT_Y + 0.85, 4.2, 1.5,
            size=64, color=C["BLACK"], align="center", valign="middle")
        txt(s, mm["def"], 0.6, CONT_Y + 2.55, 3.8, 0.6,
            size=13, italic=True, color=C["GREY"], align="center", valign="middle")

        txt(s, "Suffix", 5.0, CONT_Y + 0.15, 4.6, 0.38,
            size=20, bold=True, color=C["PURPLE"], align="center")

        gcw, grh = 2.3, 0.84
        for i, sf in enumerate(mm["suffixes"]):
            col, row = i % 2, i // 2
            x = 5.0 + col * gcw
            y = CONT_Y + 0.55 + row * grh
            rect(s, x, y, gcw, grh, C["WHITE"], C["PURPLE"], 2)
            if show_answers:
                ans = mm["answers"][i] if i < len(mm["answers"]) else ""
                rich_txt(s, [
                    (sf + "\n",  {"size": 14, "color": C["GREY"]}),
                    (ans,        {"size": 22, "bold": True, "color": C["PINK"]}),
                ], x, y, gcw, grh, align="center", valign="middle", margin=0)
            else:
                txt(s, sf, x, y, gcw, grh,
                    size=26, color=C["BLACK"], align="center", valign="middle", margin=0)

    def slide20():
        s = prs.slides.add_slide(BLANK)
        _morph_matrix(s, show_answers=False)

    # ════════════════════════════════════════════════════════════════════════
    # SLIDE 21 — Morphology Matrix (answers)
    # ════════════════════════════════════════════════════════════════════════

    def slide21():
        s = prs.slides.add_slide(BLANK)
        _morph_matrix(s, show_answers=True)

    # ── Build all slides ──────────────────────────────────────────────────────

    slide1();  slide2();  slide3();  slide4();  slide5();  slide6();  slide7()
    slide8();  slide9();  slide10(); slide11(); slide12(); slide13(); slide14()
    slide15(); slide16(); slide17(); slide18(); slide19(); slide20(); slide21()

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()
