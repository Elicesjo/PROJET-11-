#!/usr/bin/env python3
"""
Génère Elices-Diez_Josy_3_Support_presentation_062026.pptx
"""

from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt

BLUE_DARK = RGBColor(0x1A, 0x52, 0x76)
BLUE_MED = RGBColor(0x2E, 0x86, 0xC1)
BLUE_LIGHT = RGBColor(0xD6, 0xEA, 0xF8)
GRAY_BG = RGBColor(0xF2, 0xF2, 0xF2)
GRAY_TEXT = RGBColor(0x88, 0x88, 0x88)
GREEN = RGBColor(0x1E, 0x8A, 0x44)
RED = RGBColor(0xC0, 0x39, 0x2B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1C, 0x1C, 0x1C)
FONT = "Aptos Display"

SLIDE_W = 33.87
SLIDE_H = 19.05
MARGIN = 2.0
CONTENT_W = SLIDE_W - 2 * MARGIN

GAP = 0.4
CHART_W = CONTENT_W * 0.6 - GAP / 2 + 2.0
TEXT_W = CONTENT_W * 0.4 - GAP / 2 - 2.0
TEXT_L = MARGIN + CHART_W + GAP

TITLE_TOP = 0.6
TITLE_H1 = 1.7
TITLE_H2 = 3.1
LISERET_H = 0.08
BEFORE_LIS = 0.8
AFTER_LIS = 1.0
BAND_H = 2.0

FIGURES = Path("figures")
OUT = Path("Elices-Diez_Josy_3_Support_presentation_062026.pptx")


def _no_border(shape):
    sp = shape._element
    spPr = sp.find(qn("p:spPr"))
    if spPr is None:
        return
    ln = spPr.find(qn("a:ln"))
    if ln is None:
        ln = etree.SubElement(spPr, qn("a:ln"))
    for child in list(ln):
        ln.remove(child)
    etree.SubElement(ln, qn("a:noFill"))


def _rect(slide, left, top, width, height, fill_color):
    shape = slide.shapes.add_shape(1, Cm(left), Cm(top), Cm(width), Cm(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    _no_border(shape)
    return shape


def _set_run(run, text, size, bold, color):
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def _textbox(slide, text, left, top, width, height,
             size=18, bold=False, color=DARK, align=PP_ALIGN.LEFT):
    txb = slide.shapes.add_textbox(Cm(left), Cm(top), Cm(width), Cm(height))
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    _set_run(p.add_run(), text, size, bold, color)
    return txb


def _multiline(slide, lines, left, top, width, height, size=15, space_after=5):
    txb = slide.shapes.add_textbox(Cm(left), Cm(top), Cm(width), Cm(height))
    tf = txb.text_frame
    tf.word_wrap = True
    for i, (text, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        _set_run(p.add_run(), text, size, bold, color if color else DARK)
        if text.strip():
            p.space_after = Pt(space_after)
    return txb


def _picture_fit(slide, path, left, top, max_w, max_h):
    img = Image.open(path)
    iw, ih = img.size
    ratio = iw / ih
    if ratio > max_w / max_h:
        w, h = max_w, max_w / ratio
    else:
        h, w = max_h, max_h * ratio
    slide.shapes.add_picture(
        str(path),
        Cm(left + (max_w - w) / 2),
        Cm(top + (max_h - h) / 2),
        Cm(w),
        Cm(h),
    )


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _add_page_number(slide, n, is_cover=False):
    color = WHITE if is_cover else GRAY_TEXT
    _textbox(slide, str(n), SLIDE_W - 2.2, SLIDE_H - 1.1, 1.8, 0.8,
             size=12, color=color, align=PP_ALIGN.RIGHT)


def _add_title(slide, text):
    title_h = TITLE_H2 if len(text) > 46 else TITLE_H1
    lis_top = TITLE_TOP + title_h + BEFORE_LIS
    ct = lis_top + LISERET_H + AFTER_LIS
    ch = SLIDE_H - ct - 1.0
    txb = slide.shapes.add_textbox(Cm(MARGIN), Cm(TITLE_TOP), Cm(CONTENT_W), Cm(title_h))
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    _set_run(p.add_run(), text, 36, True, DARK)
    _rect(slide, MARGIN, lis_top, CONTENT_W, LISERET_H, BLUE_DARK)
    return ct, ch


def add_kpi_box(slide, value, label, left, top, width, height, value_color=BLUE_MED):
    _rect(slide, left, top, width, height, GRAY_BG)
    pad = 0.4
    _textbox(slide, value, left + pad, top + height * 0.1, width - 2 * pad, height * 0.5,
             size=36, bold=True, color=value_color, align=PP_ALIGN.CENTER)
    _textbox(slide, label, left + pad, top + height * 0.62, width - 2 * pad, height * 0.35,
             size=13, color=GRAY_TEXT, align=PP_ALIGN.CENTER)


def add_card(slide, title, desc, left, top, width, height):
    _rect(slide, left, top, width, height, GRAY_BG)
    _rect(slide, left, top, width, BAND_H, BLUE_DARK)
    pad = 0.4
    txb = slide.shapes.add_textbox(
        Cm(left + pad), Cm(top + (BAND_H - 1.4) / 2), Cm(width - 2 * pad), Cm(1.4)
    )
    _set_run(txb.text_frame.paragraphs[0].add_run(), title, 17, True, WHITE)
    _textbox(slide, desc, left + pad, top + BAND_H + 0.4,
             width - 2 * pad, height - BAND_H - 0.6, size=15, color=DARK)


def add_section_divider(prs, title, subtitle, page_num):
    slide = _blank(prs)
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, BLUE_DARK)
    _textbox(slide, title, 3.0, 6.5, SLIDE_W - 6.0, 3.5,
             size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if subtitle:
        _textbox(slide, subtitle, 3.0, 10.5, SLIDE_W - 6.0, 2.0,
                 size=24, color=BLUE_LIGHT, align=PP_ALIGN.CENTER)
    _add_page_number(slide, page_num, is_cover=True)


def content_slide(prs, n, title, image=None, kpis=None):
    slide = _blank(prs)
    ct, ch = _add_title(slide, title)
    if image:
        _picture_fit(slide, image, MARGIN, ct, CHART_W, ch)
    if kpis:
        gap_v = 0.7
        card_h = (ch - gap_v * (len(kpis) - 1)) / len(kpis)
        for i, (val, lbl, col) in enumerate(kpis):
            add_kpi_box(slide, val, lbl, TEXT_L, ct + i * (card_h + gap_v), TEXT_W, card_h, col)
    _add_page_number(slide, n)


def cards_slide(prs, n, title, cards):
    slide = _blank(prs)
    ct, ch = _add_title(slide, title)
    nb, gap = len(cards), 0.5
    card_w = (CONTENT_W - gap * (nb - 1)) / nb
    for i, (head, desc) in enumerate(cards):
        add_card(slide, head, desc, MARGIN + i * (card_w + gap), ct, card_w, ch)
    _add_page_number(slide, n)


def focus_slide(prs, n, title, stats, note=None):
    slide = _blank(prs)
    ct, ch = _add_title(slide, title)
    nb, gap = len(stats), 0.6
    card_w = (CONTENT_W - gap * (nb - 1)) / nb
    note_h = 1.6 if note else 0.0
    card_h = SLIDE_H - 1.5 - note_h - ct
    for i, (val, lbl, col) in enumerate(stats):
        add_kpi_box(slide, val, lbl, MARGIN + i * (card_w + gap), ct, card_w, card_h, col)
    if note:
        _textbox(slide, note, MARGIN, SLIDE_H - 1.5 - note_h + 0.4, CONTENT_W, note_h,
                 size=14, color=GRAY_TEXT, align=PP_ALIGN.CENTER)
    _add_page_number(slide, n)


def cover_slide(prs, title, subtitle, footer):
    slide = _blank(prs)
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, BLUE_DARK)
    _textbox(slide, title, MARGIN, 2.8, SLIDE_W - 2 * MARGIN, 3.5,
             size=50, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    _textbox(slide, subtitle, MARGIN, 6.8, SLIDE_W - 2 * MARGIN, 2.5,
             size=36, color=BLUE_LIGHT, align=PP_ALIGN.LEFT)
    _rect(slide, MARGIN, 10.2, CONTENT_W, 0.06, BLUE_MED)
    _textbox(slide, footer, MARGIN, SLIDE_H - 1.8, SLIDE_W - 2 * MARGIN, 1.0,
             size=14, color=BLUE_LIGHT, align=PP_ALIGN.LEFT)
    _add_page_number(slide, 1, is_cover=True)


def table_slide(prs, n, title, headers, rows, col_weights=None):
    slide = _blank(prs)
    ct, ch = _add_title(slide, title)
    nb_cols = len(headers)
    if col_weights is None:
        col_weights = [1.0 / nb_cols] * nb_cols
    col_ws = [CONTENT_W * w for w in col_weights]
    row_h = ch / (len(rows) + 1)
    x = MARGIN
    for hdr, cw in zip(headers, col_ws):
        _rect(slide, x, ct, cw, row_h, BLUE_DARK)
        _textbox(slide, hdr, x + 0.2, ct + 0.15, cw - 0.4, row_h - 0.2,
                 size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        x += cw
    for i, row in enumerate(rows):
        y = ct + (i + 1) * row_h
        x = MARGIN
        bg = GRAY_BG if i % 2 == 0 else WHITE
        for cell, cw in zip(row, col_ws):
            _rect(slide, x, y, cw, row_h, bg)
            _textbox(slide, str(cell), x + 0.2, y + 0.1, cw - 0.4, row_h - 0.15,
                     size=13, color=DARK, align=PP_ALIGN.CENTER)
            x += cw
    _add_page_number(slide, n)


def code_slide(prs, n, title, left_lines, right_lines):
    slide = _blank(prs)
    ct, ch = _add_title(slide, title)
    half_w = (CONTENT_W - GAP) / 2
    _rect(slide, MARGIN, ct, half_w, ch, GRAY_BG)
    _multiline(slide, left_lines, MARGIN + 0.35, ct + 0.35, half_w - 0.7, ch - 0.5)
    _rect(slide, MARGIN + half_w + GAP, ct, half_w, ch, GRAY_BG)
    _multiline(slide, right_lines, MARGIN + half_w + GAP + 0.35, ct + 0.35, half_w - 0.7, ch - 0.5)
    _add_page_number(slide, n)


def build():
    prs = Presentation()
    prs.slide_width = Cm(SLIDE_W)
    prs.slide_height = Cm(SLIDE_H)

    cover_slide(
        prs,
        title="Détection de faux billets",
        subtitle="Classification par caractéristiques physiques — ONCFM",
        footer="Josy Elices-Diez  |  Juin 2026",
    )

    cards_slide(prs, 2, "Sommaire", [
        ("Les données", "1 500 billets, 6 mesures physiques, ratio 2:1 vrais / faux"),
        ("Preprocessing", "Encodage, imputation, standardisation, split stratifié 75/25"),
        ("Modélisation", "4 algorithmes : K-means, Rég. logistique, KNN, Random Forest"),
        ("Application", "Script Python — 2 modes d'entrée, sortie VRAI / FAUX par billet"),
    ])

    focus_slide(prs, 3, "Contexte : identifier automatiquement les faux billets", [
        ("1 500", "billets dans le jeu de données", BLUE_MED),
        ("6", "mesures physiques par billet", BLUE_MED),
        ("20", "valeurs manquantes (margin_low)", BLUE_MED),
        ("4", "algorithmes comparés", BLUE_MED),
    ], note="Données ONCFM — scans physiques : diagonal, hauteurs gauche/droite, marges basse/haute, longueur")

    add_section_divider(prs, "Les données", "Exploration du jeu d'entraînement", 4)

    content_slide(prs, 5,
        "Les faux billets se distinguent nettement sur margin_low et length",
        image=FIGURES / "distributions.png",
        kpis=[
            ("1 000", "billets vrais  (67 %)", GREEN),
            ("500",   "faux billets  (33 %)", RED),
            ("2:1",   "ratio vrais / faux", BLUE_MED),
        ],
    )

    content_slide(prs, 6,
        "margin_low : 20 valeurs manquantes comblées par régression linéaire",
        image=FIGURES / "correlation.png",
        kpis=[
            ("20", "valeurs manquantes dans margin_low", BLUE_MED),
            ("Rég.\nlin.", "imputation sur les 5 autres features", BLUE_MED),
        ],
    )

    add_section_divider(prs, "Preprocessing", "Préparer les données avant la modélisation", 7)

    cards_slide(prs, 8, "4 étapes de preprocessing avant la modélisation", [
        ("Encodage cible", "is_genuine : True/False\n→ 1 (vrai) / 0 (faux)"),
        ("Imputation", "margin_low : 20 lignes manquantes comblées par régression linéaire sur les 5 autres features"),
        ("Standardisation", "StandardScaler\n(moyenne=0, écart-type=1)\nObligatoire pour KNN,\nK-means, rég. log."),
        ("Split stratifié", "75 % train — 1 125 billets\n25 % test  — 375 billets\nStratifié : préserve\nle ratio 2:1"),
    ])

    add_section_divider(prs, "Modélisation", "Comparaison de 4 algorithmes", 9)

    content_slide(prs, 10,
        "K-means : prédiction par distance aux centroïdes (non supervisé)",
        image=FIGURES / "cm_kmeans.png",
        kpis=[
            ("98.4 %", "Recall — faux billets", BLUE_MED),
            ("98.7 %", "Accuracy globale", BLUE_MED),
            ("—", "AUC-ROC\n(pas de proba)", GRAY_TEXT),
        ],
    )

    content_slide(prs, 11,
        "Régression logistique : margin_low et length sont les features décisives",
        image=FIGURES / "coef_logreg.png",
        kpis=[
            ("98.4 %", "Recall — faux billets", GREEN),
            ("99.2 %", "Accuracy globale", BLUE_MED),
            ("0.9996", "AUC-ROC", BLUE_MED),
        ],
    )

    content_slide(prs, 12,
        "KNN : k optimal trouvé par GridSearchCV — meilleur AUC-ROC",
        image=FIGURES / "knn_k_optimization.png",
        kpis=[
            ("98.4 %", "Recall — faux billets", GREEN),
            ("99.2 %", "Accuracy globale", BLUE_MED),
            ("0.9997", "AUC-ROC  meilleur", GREEN),
        ],
    )

    content_slide(prs, 13,
        "Random Forest : margin_low domine les feature importances",
        image=FIGURES / "rf_summary.png",
        kpis=[
            ("98.4 %", "Recall — faux billets", GREEN),
            ("99.2 %", "Accuracy globale", BLUE_MED),
            ("0.9994", "AUC-ROC", BLUE_MED),
        ],
    )

    content_slide(prs, 14,
        "Les 3 modèles supervisés atteignent 99.2 % d'accuracy",
        image=FIGURES / "comparaison_modeles.png",
        kpis=[
            ("KNN", "modèle retenu", GREEN),
            ("98.4 %", "Recall(Faux)", GREEN),
            ("0.9997", "AUC-ROC", GREEN),
        ],
    )

    cards_slide(prs, 15, "KNN retenu : meilleure séparation probabiliste des classes", [
        ("Critère 1 — Recall", "Recall(Faux) = 0.984\nPriorité absolue : un faux billet classé vrai est un risque opérationnel majeur"),
        ("Égalité sur 3 modèles", "Rég. log., KNN et Random Forest atteignent tous Recall = 0.984 et F1 = 0.988"),
        ("Tie-break AUC-ROC", "KNN : 0.9997\nRég. log. : 0.9996\nRandom Forest : 0.9994\nKNN discrimine mieux en probabilité"),
        ("K-means écarté", "Algorithme non supervisé : l'alignement cluster/classe est une inférence post-hoc, moins robuste sur données inconnues"),
    ])

    add_section_divider(prs, "Application", "Script Python prêt pour la production", 16)

    code_slide(prs, 17,
        "2 modes d'utilisation — une ligne de sortie par billet",
        left_lines=[
            ("Mode CSV  (multi-billets)", True, BLUE_DARK),
            ("", False, DARK),
            ("python script.py \\", False, DARK),
            ("  --csv billets_production.csv", False, DARK),
            ("", False, DARK),
            ("VRAI", False, GREEN),
            ("VRAI", False, GREEN),
            ("FAUX", False, RED),
            ("VRAI", False, GREEN),
            ("", False, DARK),
            ("Une ligne par billet dans l'ordre du CSV", False, GRAY_TEXT),
        ],
        right_lines=[
            ("Mode valeurs  (1 billet)", True, BLUE_DARK),
            ("", False, DARK),
            ("python script.py --values \\", False, DARK),
            ("  172.28 103.95 103.91 \\", False, DARK),
            ("  4.78 3.31 111.40", False, DARK),
            ("", False, DARK),
            ("FAUX", False, RED),
            ("", False, DARK),
            ("Ordre des valeurs :", False, GRAY_TEXT),
            ("diagonal  height_left  height_right", False, GRAY_TEXT),
            ("margin_low  margin_up  length", False, GRAY_TEXT),
        ],
    )

    focus_slide(prs, 18, "Résultats — modèle KNN validé sur 375 billets", [
        ("99.2 %", "Accuracy globale", GREEN),
        ("98.4 %", "Recall — faux billets", GREEN),
        ("0.9997", "AUC-ROC", GREEN),
        ("2 modes", "CSV ou valeurs directes", BLUE_MED),
    ], note="Entraîné sur 1 125 billets, validé sur 375 — données ONCFM, juin 2026")

    prs.save(str(OUT))
    print(f"OK : {OUT}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build()
