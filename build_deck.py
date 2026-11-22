# -*- coding: utf-8 -*-
"""第1回セミナー資料の生成スクリプト。

同じ内容定義から、スライド原稿（マークダウン）と PowerPoint の両方を出力する。
  python build_deck.py
内容を変えるときは、このファイルの各 slide_XX 関数を編集して再実行する。
"""
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

HERE = Path(__file__).parent
PPTX_OUT = HERE / "01_第1回_固定資産仕訳の検証ポイント.pptx"
MD_OUT = HERE / "01_第1回_スライド原稿.md"

FONT = "Meiryo UI"
SLIDE_W, SLIDE_H = 13.333, 7.5
LEFT = 0.6
CONTENT_W = SLIDE_W - LEFT * 2


def rgb(h):
    return RGBColor.from_string(h)


INK = rgb("1E292E")
MUTED = rgb("6B7378")
LINE = rgb("D5DADC")
WHITE = rgb("FFFFFF")
PAPER = rgb("F4F6F6")
DARK = rgb("16343A")
# 層の色：① 固定資産固有＝ティール、② 仕訳に応じた確認ポイント＝オレンジ
L1 = rgb("0F6E64")
L1_LT = rgb("E1F0ED")
L2 = rgb("B85C14")
L2_LT = rgb("FBEADB")
NEUTRAL = rgb("4A5A61")
NEUTRAL_LT = rgb("EBEEEF")

DASH = "－"

# ---------------------------------------------------------------- helpers


def style_run(run, size, color=INK, bold=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    f.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", FONT)


def set_bullet(p, indent=0.22):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Emu(Inches(indent)))))
    pPr.set("indent", str(-int(Emu(Inches(indent)))))
    bu = etree.SubElement(pPr, qn("a:buChar"))
    bu.set("char", "•")


def fill_tf(tf, paras, size=14, color=INK, bold=False, align=PP_ALIGN.LEFT,
            space_after=4, line_spacing=None):
    """paras: list of str | dict(text|runs, size, color, bold, bullet, align, space_after)."""
    tf.word_wrap = True
    for i, para in enumerate(paras):
        if isinstance(para, str):
            para = {"text": para}
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = para.get("align", align)
        p.space_after = Pt(para.get("space_after", space_after))
        if line_spacing:
            p.line_spacing = line_spacing
        runs = para.get("runs") or [(para.get("text", ""), {})]
        for text, opt in runs:
            r = p.add_run()
            r.text = text
            style_run(r, opt.get("size", para.get("size", size)),
                      opt.get("color", para.get("color", color)),
                      opt.get("bold", para.get("bold", bold)))
        if para.get("bullet"):
            set_bullet(p)


def text(slide, x, y, w, h, paras, size=14, color=INK, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, margin=0.0, name=None, **kw):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(margin))
    fill_tf(tf, paras if isinstance(paras, list) else [paras], size, color, bold, align, **kw)
    return tb


def box(slide, x, y, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=None,
        radius=0.06, name=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        s.name = name
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    return s


def boxed_text(slide, x, y, w, h, fill, paras, size=14, color=INK, bold=False,
               align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, pad=0.18, line=None,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06, name=None, **kw):
    s = box(slide, x, y, w, h, fill, shape=shape, line=line, radius=radius, name=name)
    tf = s.text_frame
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(pad))
    fill_tf(tf, paras if isinstance(paras, list) else [paras], size, color, bold, align, **kw)
    return s


def cell_border(cell, color="D5DADC", width=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    for i, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
        ln = etree.Element(qn(tag), w=str(width), cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=color)
        etree.SubElement(ln, qn("a:prstDash"), val="solid")
        tcPr.insert(i, ln)


def table(slide, x, y, col_w, rows, size=12, row_h=0.38, header_fill=NEUTRAL,
          header_color=WHITE, row_fills=None, cell_fills=None, bold_cols=(),
          center_cols=(), name=None):
    """rows[0] is the header. row_fills: {row_index: color}; cell_fills: {(r,c): color}."""
    nr, nc = len(rows), len(rows[0])
    shape = slide.shapes.add_table(nr, nc, Inches(x), Inches(y),
                                   Inches(sum(col_w)), Inches(row_h * nr))
    if name:
        shape.name = name
    tbl = shape.table
    tbl.first_row = True
    tbl.horz_banding = False
    for c, w in enumerate(col_w):
        tbl.columns[c].width = Inches(w)
    for r in range(nr):
        tbl.rows[r].height = Inches(row_h)
        for c in range(nc):
            cell = tbl.cell(r, c)
            val = rows[r][c]
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            if r == 0:
                fill, color, bold = header_fill, header_color, True
            else:
                fill = (cell_fills or {}).get((r, c), (row_fills or {}).get(r, WHITE))
                color = MUTED if val == DASH else INK
                bold = c in bold_cols
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if (c in center_cols or val in (DASH, "●", "○")) else PP_ALIGN.LEFT
            run = p.add_run()
            run.text = val
            style_run(run, size, color, bold)
            cell_border(cell)
    return shape


def arrow(slide, x, y, w, h, color=MUTED, shape=MSO_SHAPE.RIGHT_ARROW):
    a = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    a.shadow.inherit = False
    a.fill.solid()
    a.fill.fore_color.rgb = color
    a.line.fill.background()
    return a


def md_table(rows):
    out = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * len(rows[0])]
    out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(out)


# ---------------------------------------------------------------- deck scaffolding

prs = Presentation()
prs.slide_width = Inches(SLIDE_W)
prs.slide_height = Inches(SLIDE_H)
BLANK = prs.slide_layouts[6]
MD = []
SLIDE_NO = [0]
LAYER = {
    "intro": (NEUTRAL, NEUTRAL_LT),
    "l1": (L1, L1_LT),
    "l2": (L2, L2_LT),
}


def new_slide(notes):
    SLIDE_NO[0] += 1
    s = prs.slides.add_slide(BLANK)
    s.notes_slide.notes_text_frame.text = notes
    return s


def content_slide(layer, label, title, message, notes, md_body):
    s = new_slide(notes)
    color, tint = LAYER[layer]
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = WHITE
    w = len(label) * 0.165 + 0.4
    boxed_text(s, LEFT, 0.38, w, 0.34, color, label, size=11, color=WHITE, bold=True,
               anchor=MSO_ANCHOR.MIDDLE, pad=0.12, radius=0.5, name="Section label")
    text(s, LEFT, 0.82, CONTENT_W, 0.6, title, size=26, bold=True, name="Title")
    text(s, LEFT, 1.45, CONTENT_W, 0.5, message, size=15, color=color, bold=True, name="Key message")
    text(s, LEFT, 7.0, 6, 0.3, "固定資産仕訳セミナー 第1回", size=10, color=MUTED, name="Footer")
    text(s, SLIDE_W - LEFT - 1, 7.0, 1, 0.3, str(SLIDE_NO[0]), size=10, color=MUTED,
         align=PP_ALIGN.RIGHT, name="Slide number")
    MD.append(f"## {SLIDE_NO[0]}. {title}\n\n*{label}*\n\n**{message}**\n\n{md_body.strip()}\n\n"
              f"> ノート：{notes}\n")
    return s, color, tint


def points_card(slide, x, y, w, h, tint, color, items, title="検証ポイント", size=14, pos=None):
    paras = [{"text": title, "size": 15, "bold": True, "color": color, "space_after": 8}]
    paras += [{"text": t, "bullet": True, "space_after": 7} for t in items]
    card = boxed_text(slide, x, y, w, h, tint, paras, size=size, pad=0.22, name="Check points")
    if pos is not None:
        quadrant_map(slide, x + 0.22, y + h - 1.13, pos, color)
    return card


QUAD = {(0, 0): "取得", (0, 1): "売却", (1, 0): "振替", (1, 1): "除却・減損ほか"}


def quadrant_map(slide, x, y, pos, color):
    """2×2 の分類表のミニ版。pos=(row, col) のマスを強調する。"""
    rh, cw, ch, gap = 0.85, 1.4, 0.37, 0.05
    for c, t in enumerate(["損益なし", "損益あり"]):
        text(slide, x + rh + c * (cw + gap), y, cw, 0.22, t, size=10, color=MUTED, align=PP_ALIGN.CENTER)
    for r, t in enumerate(["現金あり", "現金なし"]):
        cy = y + 0.25 + r * (ch + 0.04)
        text(slide, x, cy, rh - 0.05, ch, t, size=10, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
        for c in range(2):
            on = (r, c) == pos
            boxed_text(slide, x + rh + c * (cw + gap), cy, cw, ch, color if on else WHITE, QUAD[(r, c)],
                       size=11, color=WHITE if on else MUTED, bold=on, align=PP_ALIGN.CENTER,
                       anchor=MSO_ANCHOR.MIDDLE, pad=0.02, radius=0.15, name="Quadrant map")


def bullets_md(items, title="検証ポイント"):
    return f"**{title}**\n\n" + "\n".join(f"- {t}" for t in items)


JHEAD = ["貸借", "勘定科目", "事業", "部門・PJ", "取引先", "税区分"]
JCOLS = [0.6, 2.45, 0.85, 1.5, 0.95, 1.15]  # 7.5


def journal(slide, x, y, rows, l1_rows=(), l2_rows=(), row_h=0.42, size=12, name="Journal"):
    """l1_rows: rows of the 3 fixed-asset accounts (tinted teal); l2_rows: P/L rows (tinted orange)."""
    data = [JHEAD] + rows
    fills = {}
    for i in l1_rows:
        fills[i] = L1_LT
    for i in l2_rows:
        fills[i] = L2_LT
    return table(slide, x, y, JCOLS, data, size=size, row_h=row_h, row_fills=fills,
                 center_cols=(0,), name=name)


def journal_legend(slide, x, y):
    box(slide, x, y + 0.04, 0.2, 0.2, L1_LT, shape=MSO_SHAPE.RECTANGLE, line=LINE)
    text(slide, x + 0.28, y, 3.2, 0.32, "固定資産勘定の行（台帳と元帳の一致の対象）", size=11, color=MUTED)
    box(slide, x + 3.75, y + 0.04, 0.2, 0.2, L2_LT, shape=MSO_SHAPE.RECTANGLE, line=LINE)
    text(slide, x + 4.03, y, 2.5, 0.32, "損益勘定の行", size=11, color=MUTED)


# ---------------------------------------------------------------- slides


def slide_01():
    s = new_slide("表紙。ERPパッケージの開発者向けに、固定資産の仕訳ファイル作成機能を検証するときの観点を全4回で説明する。第1回は仕訳の種類と特性。")
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = DARK
    text(s, 0.9, 1.6, 11.5, 0.5, "ERPパッケージ開発者向けセミナー　第1回", size=18, color=rgb("9CC9C1"), bold=True)
    text(s, 0.9, 2.3, 11.5, 1.9, ["固定資産の仕訳ファイル作成機能", "検証ポイント"], size=40, color=WHITE, bold=True,
         space_after=6)
    text(s, 0.9, 4.45, 11.5, 0.6, "仕訳の種類と特性", size=24, color=WHITE)
    # 2層のモチーフ
    boxed_text(s, 0.9, 5.6, 3.6, 0.5, L1, "① 固定資産固有の観点", size=14, color=WHITE, bold=True,
               anchor=MSO_ANCHOR.MIDDLE, pad=0.15, radius=0.5)
    boxed_text(s, 4.7, 5.6, 3.9, 0.5, L2, "② 仕訳に応じた確認ポイント", size=14, color=WHITE, bold=True,
               anchor=MSO_ANCHOR.MIDDLE, pad=0.15, radius=0.5)
    MD.append("# 固定資産の仕訳ファイル作成機能 検証ポイント\n\n"
              "## 1. 表紙\n\n- ERPパッケージ開発者向けセミナー 第1回\n- 仕訳の種類と特性\n")


def slide_02():
    items = ["第1回（本日）：仕訳の種類と特性", "第2回〜第4回：本番事例"]
    s, color, tint = content_slide(
        "intro", "導入", "本セミナーの目的",
        "仕訳ファイル作成機能を検証するときの観点を、全体像として持ち帰っていただく",
        "全4回の中での第1回の位置づけを説明する。本日は観点を2層に分けて整理する。金額の計算は扱わない。",
        bullets_md(items, "全4回の構成") + "\n\n**本日の観点（2層）**\n\n"
        "- ① 固定資産固有の観点\n- ② 仕訳に応じた確認ポイント\n\n**本日扱わないこと**：金額の計算")
    # 全4回のタイムライン
    for i, (no, label, on) in enumerate([("第1回", "仕訳の種類と特性", True), ("第2回", "本番事例", False),
                                          ("第3回", "本番事例", False), ("第4回", "本番事例", False)]):
        x = LEFT + i * 3.08
        boxed_text(s, x, 2.2, 2.88, 1.15, NEUTRAL if on else NEUTRAL_LT,
                   [{"text": no, "size": 13, "bold": True, "color": WHITE if on else MUTED},
                    {"text": label + ("（本日）" if on else ""), "size": 15, "bold": True,
                     "color": WHITE if on else INK}],
                   anchor=MSO_ANCHOR.MIDDLE, pad=0.2)
    text(s, LEFT, 3.75, 6, 0.4, "本日の観点は2層に分けて整理します", size=15, bold=True)
    boxed_text(s, LEFT, 4.25, 5.95, 1.5, L1_LT,
               [{"text": "① 固定資産固有の観点", "size": 16, "bold": True, "color": L1},
                "固定資産勘定・減価償却累計額・減損損失累計額と、資産の種類ごとの勘定科目テーブル"],
               size=14, pad=0.22)
    boxed_text(s, LEFT + 6.18, 4.25, 5.95, 1.5, L2_LT,
               [{"text": "② 仕訳に応じた確認ポイント", "size": 16, "bold": True, "color": L2},
                "それ以外の勘定科目。現金関連、償却費勘定のテーブル、B/S・P/Lセグメント"],
               size=14, pad=0.22)
    text(s, LEFT, 6.0, CONTENT_W, 0.4,
         [{"runs": [("本日扱わないこと：", {"bold": True}), ("金額の計算（端数、按分など）", {})]}],
         size=13, color=MUTED)


def slide_03():
    kinds = ["取得", "売却", "振替", "除却", "減損", "圧縮記帳", "減価償却費", "過年度の訂正"]
    s, color, tint = content_slide(
        "l1", "① 固定資産固有の観点", "台帳と元帳の一致",
        "固定資産勘定・減価償却累計額・減損損失累計額は、台帳と元帳の一致が最終的な検証基準",
        "仕訳ファイルは総勘定元帳に転記される。固定資産の3つの勘定は、台帳の残高と元帳の残高が一致していなければならない。一致させるには、どの仕訳ファイルでも3勘定の行の項目がそろっている必要がある。1種類のファイルでも項目が欠けると、その単位で一致しなくなる。",
        "固定資産台帳 → 仕訳ファイル（" + "・".join(kinds) + "）→ 総勘定元帳\n\n"
        "- 一致を確認する対象：固定資産勘定、減価償却累計額、減損損失累計額\n"
        "- 一致の単位：勘定科目 × 事業（B/Sセグメント）\n"
        "- 条件：すべての仕訳ファイルで、3勘定の行の項目（勘定科目・事業・資産番号）がそろっていること")
    y, h = 2.2, 2.55
    boxed_text(s, LEFT, y, 2.9, h, L1_LT,
               [{"text": "固定資産台帳", "size": 18, "bold": True, "color": L1},
                {"text": "資産ごとに", "size": 13, "color": MUTED},
                {"text": "資産の種類・事業・", "size": 13, "color": MUTED, "space_after": 0},
                {"text": "部門・プロジェクトを保持", "size": 13, "color": MUTED}],
               anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, pad=0.2)
    arrow(s, LEFT + 3.0, y + h / 2 - 0.25, 0.55, 0.5)
    mx = LEFT + 3.65
    box(s, mx, y, 5.4, h, PAPER)
    text(s, mx + 0.2, y + 0.15, 5, 0.4, "仕訳ファイル", size=18, bold=True)
    for i, k in enumerate(kinds):
        cx = mx + 0.2 + (i % 4) * 1.27
        cy = y + 0.75 + (i // 4) * 0.82
        boxed_text(s, cx, cy, 1.17, 0.62, WHITE, k, size=13, bold=True, align=PP_ALIGN.CENTER,
                   anchor=MSO_ANCHOR.MIDDLE, pad=0.04, line=LINE)
    arrow(s, mx + 5.5, y + h / 2 - 0.25, 0.55, 0.5)
    boxed_text(s, mx + 6.15, y, 2.33, h, L1_LT,
               [{"text": "総勘定元帳", "size": 18, "bold": True, "color": L1},
                {"text": "勘定科目 × セグメントで残高を集計", "size": 13, "color": MUTED}],
               anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, pad=0.2)
    boxed_text(s, LEFT, 5.05, CONTENT_W, 1.65, WHITE,
               [{"runs": [("一致を確認する対象　", {"bold": True, "color": L1}),
                          ("固定資産勘定 ／ 減価償却累計額 ／ 減損損失累計額", {"bold": True})], "size": 16},
                {"runs": [("一致の単位　", {"bold": True, "color": L1}),
                          ("勘定科目 × 事業（B/Sセグメント）", {})], "size": 15},
                {"runs": [("条件　", {"bold": True, "color": L1}),
                          ("すべての仕訳ファイルで、3勘定の行の項目（勘定科目・事業・資産番号）がそろっていること", {})],
                 "size": 15}],
               pad=0.25, line=L1, anchor=MSO_ANCHOR.MIDDLE, space_after=6)


ACCT_TABLE = [
    ["資産の種類", "取得価額", "減価償却累計額", "減損損失累計額"],
    ["建物", "建物", "建物減価償却累計額", "建物減損損失累計額"],
    ["機械装置", "機械装置", "機械装置減価償却累計額", "機械装置減損損失累計額"],
    ["ソフトウェア", "ソフトウェア", "なし（直接控除）", "なし（直接控除）"],
    ["土地", "土地", "なし（償却しない）", "土地減損損失累計額"],
]


def slide_04():
    pts = ["すべての資産の種類・分類に、科目の設定があるか",
           "設定がない組み合わせ（累計額がない資産、償却しない資産）で、行が作られていないか",
           "設定が漏れているとき、エラーで止まるか。科目が空欄のまま仕訳が出ていないか",
           "資産の種類が変わったとき、変更後の科目で振り替えられるか（振替で詳しく扱う）"]
    s, color, tint = content_slide(
        "l1", "① 固定資産固有の観点", "資産の種類・分類ごとの勘定科目テーブル",
        "資産の種類・分類によって使われる勘定科目が異なる。元帳転記時の勘定科目テーブルを考慮する",
        "固定資産固有の観点。同じ取得や償却の仕訳でも、資産の種類によって勘定科目が変わる。ソフトウェアのように累計額を使わない資産、土地のように償却しない資産もあるため、テーブルには『設定がない組み合わせ』も含まれる。",
        "**勘定科目テーブルの例**\n\n" + md_table(ACCT_TABLE) + "\n\n" + bullets_md(pts))
    table(s, LEFT, 2.2, [1.6, 1.5, 2.3, 2.3], ACCT_TABLE, size=13, row_h=0.6, header_fill=L1,
          bold_cols=(0,), cell_fills={(3, 2): PAPER, (3, 3): PAPER, (4, 2): PAPER})
    text(s, LEFT, 5.35, 7.7, 0.6, "グレーのセル＝科目の設定がない組み合わせ。資産の種類ごとに使う勘定が変わる",
         size=11, color=MUTED)
    points_card(s, 8.55, 2.2, 4.18, 4.55, tint, color, pts)


def slide_05():
    cards = [
        ("1", "現金に関連する仕訳", "相手勘定に現金・債権債務が出ると、取引先と税区分が仕訳に入る", "取得、売却"),
        ("2", "B/S・P/Lセグメント", "損益勘定は、資産・負債より細かいセグメントで把握する", "損益勘定がある仕訳"),
        ("3", "償却費勘定のテーブル", "償却費の科目は、資産の種類 × 利用セグメントで決まる", "減価償却費"),
    ]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント", "3つの確認ポイント",
        "仕訳には3勘定以外の勘定科目も出てくる。仕訳の種類に応じて3つのポイントを確認する",
        "ここからは②の層。固定資産勘定以外の勘定科目は、仕訳の種類によって出てくるものが違うので、確認すべきポイントも変わる。1と2は次のスライドの2つの軸に対応し、3は減価償却費のスライドで詳しく扱う。",
        "\n".join(f"{n}. **{t}**：{d}（対象：{tg}）" for n, t, d, tg in cards))
    w = (CONTENT_W - 0.6) / 3
    for i, (n, t, d, tg) in enumerate(cards):
        x = LEFT + i * (w + 0.3)
        box(s, x, 2.2, w, 4.45, tint)
        circ = box(s, x + 0.3, 2.5, 0.7, 0.7, L2, shape=MSO_SHAPE.OVAL)
        tf = circ.text_frame
        fill_tf(tf, [n], size=22, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        text(s, x + 0.3, 3.45, w - 0.6, 0.9, t, size=19, bold=True)
        text(s, x + 0.3, 4.35, w - 0.6, 1.3, d, size=15)
        boxed_text(s, x + 0.3, 5.75, w - 0.6, 0.6, WHITE,
                   [{"runs": [("対象：", {"bold": True, "color": L2}), (tg, {"bold": True})]}],
                   size=14, anchor=MSO_ANCHOR.MIDDLE, pad=0.15)


def slide_06():
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント", "仕訳の種類と2つの軸",
        "現金の有無と損益勘定の有無によって、仕訳の行構成と必要な項目が決まる",
        "現金の軸は取引先・税区分の有無、損益の軸はセグメントの細かさに対応する。売却はすべての要素が出てくる最も複雑な仕訳、振替は固定資産勘定だけで完結する最もシンプルな仕訳。損益ありのグループが最も多い。",
        md_table([["", "損益勘定なし", "損益勘定あり"],
                  ["現金あり", "取得", "売却"],
                  ["現金なし", "振替", "除却、減損、圧縮記帳、減価償却費、過年度の訂正"]]) +
        "\n\n- 現金の有無 → 取引先・税区分が入るか\n- 損益勘定の有無 → 細かいセグメント（部門・プロジェクト）が必要か")
    hx, cw, gap = LEFT + 1.7, 5.1, 0.2
    for i, t in enumerate(["損益勘定なし", "損益勘定あり"]):
        boxed_text(s, hx + i * (cw + gap), 2.15, cw, 0.5, NEUTRAL, t, size=15, color=WHITE, bold=True,
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, pad=0.05)
    rows = [("現金あり", 2.8, 1.75), ("現金なし", 4.7, 2.05)]
    for t, y, h in rows:
        boxed_text(s, LEFT, y, 1.55, h, NEUTRAL, t, size=15, color=WHITE, bold=True,
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, pad=0.05)
    cells = [
        (0, 0, "取得", "取引先・税区分が入る。損益の行はない", NEUTRAL_LT),
        (1, 0, "売却", "取引先・税区分と、細かいセグメントの両方が出てくる最も複雑な仕訳", L2_LT),
        (0, 1, "振替", "固定資産勘定の行だけで完結する最もシンプルな仕訳", L1_LT),
        (1, 1, "除却 ／ 減損 ／ 圧縮記帳\n減価償却費 ／ 過年度の訂正", "損益の行に、細かいセグメントが必要", L2_LT),
    ]
    for c, r, name, desc, fill in cells:
        x = hx + c * (cw + gap)
        y, h = rows[r][1], rows[r][2]
        names = name.split("\n")
        paras = [{"text": n, "size": 20, "bold": True, "space_after": 2} for n in names]
        paras.append({"text": desc, "size": 14, "color": MUTED})
        boxed_text(s, x, y, cw, h, fill, paras, pad=0.25, anchor=MSO_ANCHOR.MIDDLE)


SEG_TABLE = [
    ["", "資産・負債の行", "損益の行"],
    ["事業（B/Sセグメント）", "●", "●"],
    ["部門（P/Lセグメント）", DASH, "●"],
    ["プロジェクト（P/Lセグメント）", DASH, "●"],
]


def slide_07():
    pts = ["損益の行で、部門・プロジェクトが空欄になっていないか",
           "部門・プロジェクトと事業の組み合わせが矛盾していないか",
           "資産・負債の行に、不要な部門・プロジェクトが入っていないか。入っていると元帳の集計単位がずれる"]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント", "B/SセグメントとP/Lセグメント",
        "損益勘定は、資産・負債より細かいセグメントで把握する",
        "セグメントは勘定科目以外の帳票項目のこと。本資料ではB/Sセグメント＝事業、P/Lセグメント＝部門・プロジェクトとして説明する。実際の項目構成は導入先によって異なる。損益の行は、資産・負債の行より細かい単位で編集される。",
        md_table(SEG_TABLE) + "\n\n※本資料ではB/Sセグメント＝事業、P/Lセグメント＝部門・プロジェクトとして説明します。実際の項目構成は導入先によって異なります。\n\n" + bullets_md(pts))
    table(s, LEFT, 2.2, [3.3, 2.05, 2.05], SEG_TABLE, size=14, row_h=0.62,
          bold_cols=(0,), cell_fills={(1, 1): L1_LT, (1, 2): L2_LT, (2, 2): L2_LT, (3, 2): L2_LT})
    arrow(s, LEFT + 3.3, 4.85, 4.1, 0.45, color=L2)
    text(s, LEFT + 3.3, 5.33, 4.1, 0.35, "損益の行ほど細かい単位で把握する", size=12, color=L2, bold=True,
         align=PP_ALIGN.CENTER)
    boxed_text(s, LEFT, 5.85, 7.4, 0.85, PAPER,
               "※本資料では B/Sセグメント＝事業、P/Lセグメント＝部門・プロジェクト として説明します。実際の項目構成は導入先によって異なります。",
               size=12, color=MUTED, pad=0.18, anchor=MSO_ANCHOR.MIDDLE)
    points_card(s, 8.3, 2.2, 4.43, 4.5, tint, color, pts)


ROWS_TABLE = [
    ["行の種類", "勘定科目", "事業", "部門・PJ", "資産番号", "取引先", "税区分"],
    ["固定資産勘定の行（取得価額・累計額）", "●", "●", DASH, "●", DASH, DASH],
    ["損益勘定の行", "●", "●", "●", "●", DASH, DASH],
    ["現金・債権債務の行", "●", "●", DASH, DASH, "●", DASH],
    ["消費税の行", "●", "●", DASH, DASH, "●", "●"],
]


def slide_08():
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント", "仕訳ファイルの行構成と項目",
        "行の種類によって、編集される項目が決まっている",
        "仕訳ファイルの1行ごとに、どの項目が入るかを整理した表。固定資産勘定の行は①の層、損益・現金・消費税の行は②の層の確認ポイントになる。損益の行や債権債務の行にどこまで項目を持たせるかは、パッケージの仕様に合わせて調整する。",
        md_table(ROWS_TABLE) + "\n\n※損益の行や債権債務の行にどこまで項目を持たせるかは、パッケージの仕様に合わせて調整してください。")
    table(s, LEFT, 2.25, [4.03, 1.35, 1.35, 1.35, 1.35, 1.35, 1.35], ROWS_TABLE, size=14, row_h=0.72,
          bold_cols=(0,), row_fills={1: L1_LT, 2: L2_LT})
    boxes = [(L1, "固定資産勘定の行 → ①の層（台帳と元帳の一致）"),
             (L2, "損益の行 → 部門・PJ（セグメント） ／ 現金・消費税の行 → 取引先・税区分")]
    for i, (c, t) in enumerate(boxes):
        text(s, LEFT, 6.0 + i * 0.4, CONTENT_W, 0.35, [{"runs": [("■ ", {"color": c}), (t, {})]}],
             size=13, color=MUTED)


def slide_09():
    rows = [["借", "機械装置", "事業X", DASH, DASH, DASH],
            ["借", "仮払消費税", "事業X", DASH, "取引先A", "課税仕入"],
            ["貸", "未払金", "事業X", DASH, "取引先A", DASH]]
    pts = ["資産の行が、台帳の資産の種類・事業で編集されているか（①）",
           "未払金・消費税の行に、取引先・税区分が入っているか",
           "買掛・支払モジュール側でも同じ取引を計上していないか（二重計上）"]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント ｜ 現金あり", "取得",
        "相手勘定に債権債務が出るため、取引先と税区分が仕訳に入る",
        "取得は現金ありのグループ。損益の行はないが、未払金と仮払消費税の行に取引先・税区分が入る。購買や支払のモジュールでも債務を計上する場合、どちらで仕訳を作るかの仕様を確認し、二重計上がないかを検証する。",
        md_table([JHEAD] + rows) + "\n\n" + bullets_md(pts))
    journal(s, LEFT, 2.3, rows, l1_rows=(1,), row_h=0.5)
    journal_legend(s, LEFT, 4.45)
    points_card(s, 8.4, 2.3, 4.33, 4.4, tint, color, pts, pos=(0, 0))


def slide_10():
    rows = [["借", "未収入金", "事業X", DASH, "取引先B", DASH],
            ["借", "機械装置減価償却累計額", "事業X", DASH, DASH, DASH],
            ["借/貸", "固定資産売却損益", "事業X", "製造1課・PJ-a", DASH, DASH],
            ["貸", "機械装置", "事業X", DASH, DASH, DASH],
            ["貸", "仮受消費税", "事業X", DASH, "取引先B", "課税売上"]]
    pts = ["課税区分の判定（建物は課税、土地は非課税）",
           "売却日までの当期の償却費が、売却の仕訳より先に計上されているか",
           "売却損益の科目とセグメントが、科目ごとの設定どおりか（償却費とは別の観点）",
           "取得価額・累計額の行が漏れなく消し込まれているか（①）"]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント ｜ 現金あり", "売却",
        "取引先・税区分と、細かいセグメントの両方が出てくる最も複雑な仕訳",
        "売却は2つの軸の両方に当てはまる。債権と消費税の行に取引先・税区分が入り、売却損益の行には部門・プロジェクトが入る。売却損益の科目の細かさは、償却費と同じ場合も固定の科目の場合もあるので、設定ごとに確認する。",
        md_table([JHEAD] + rows) + "\n\n" + bullets_md(pts))
    journal(s, LEFT, 2.3, rows, l1_rows=(2, 4), l2_rows=(3,), row_h=0.5)
    journal_legend(s, LEFT, 5.45)
    points_card(s, 8.4, 2.3, 4.33, 4.4, tint, color, pts, pos=(0, 1))


def slide_11():
    change = [["台帳の変更", "資産勘定の振替仕訳"],
              ["資産の種類（例：機械装置 → 工具器具備品）", "作られる（勘定科目が変わる）"],
              ["事業（B/Sセグメント）", "作られる（事業が変わる）"],
              ["部門・プロジェクト（P/Lセグメント）", "作られない"]]
    rows = [["借", "機械装置", "事業Y", DASH, DASH, DASH],
            ["貸", "機械装置", "事業X", DASH, DASH, DASH],
            ["借", "機械装置減価償却累計額", "事業X", DASH, DASH, DASH],
            ["貸", "機械装置減価償却累計額", "事業Y", DASH, DASH, DASH]]
    pts = ["取得価額と累計額（減損損失累計額を含む）の両方が振り替えられているか",
           "変更前の属性の残高を消し、変更後の属性で計上しているか",
           "変更後の科目が、勘定科目テーブルから正しく決まっているか（①）",
           "部門・プロジェクトだけの変更で、資産勘定の振替仕訳が作られていないか"]
    s, color, tint = content_slide(
        "l1", "① 固定資産固有の観点 ｜ 損益なし", "振替",
        "B/Sの項目（資産の種類・事業）が変わったときだけ、資産勘定の振替仕訳が作られる",
        "振替は損益の行がなく、固定資産勘定の行だけで完結する。資産の種類が変わると勘定科目が、事業が変わるとB/Sセグメントが変わるので振替仕訳が作られる。部門・プロジェクトはP/Lセグメントなので、資産勘定は振り替わらない。設例は事業XからYへの変更。",
        md_table(change) + "\n\n**設例：機械Aの事業を事業X → 事業Yに変更**\n\n" + md_table([JHEAD] + rows) + "\n\n" + bullets_md(pts))
    table(s, LEFT, 2.2, [4.4, 3.1], change, size=13, row_h=0.42, header_fill=L1, bold_cols=(0,),
          cell_fills={(3, 1): PAPER})
    text(s, LEFT, 4.05, 7.5, 0.35, "設例：機械Aの事業を 事業X → 事業Y に変更", size=13, bold=True, color=L1)
    journal(s, LEFT, 4.45, rows, l1_rows=(1, 2, 3, 4), row_h=0.4)
    points_card(s, 8.4, 2.2, 4.33, 4.5, tint, color, pts, pos=(1, 0))


def slide_12():
    rows = [["借", "機械装置減価償却累計額", "事業X", DASH, DASH, DASH],
            ["借", "固定資産除却損", "事業X", "製造1課・PJ-a", DASH, DASH],
            ["貸", "機械装置", "事業X", DASH, DASH, DASH]]
    pts = ["除却損の科目とセグメントが、科目ごとの設定どおりか",
           "除却時点の台帳の属性で編集されているか（償却費と違い、後から振り替わらない）",
           "減損損失累計額がある資産では、その行も消し込まれているか（①）",
           "撤去費用は別の仕訳（債務の計上側）で扱われているか"]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント ｜ 損益あり", "除却",
        "現金は出てこないが、除却損の行に細かいセグメントが必要",
        "ここから損益ありのグループ。除却は対価がないので取引先・税区分は出てこない。除却損は一時点で一度だけ発生し、後からセグメントが変わっても振り替わらないため、除却時点の属性で正しく編集されていることが重要。",
        md_table([JHEAD] + rows) + "\n\n" + bullets_md(pts))
    journal(s, LEFT, 2.3, rows, l1_rows=(1, 3), l2_rows=(2,), row_h=0.5)
    journal_legend(s, LEFT, 4.45)
    points_card(s, 8.4, 2.3, 4.33, 4.4, tint, color, pts, pos=(1, 1))


def slide_13():
    direct = [["借", "減損損失", "事業X", "製造1課・PJ-a", DASH, DASH],
              ["貸", "機械装置", "事業X", DASH, DASH, DASH]]
    accum = [["借", "減損損失", "事業X", "製造1課・PJ-a", DASH, DASH],
             ["貸", "機械装置減損損失累計額", "事業X", DASH, DASH, DASH]]
    pts = ["ユーザーが設定した計上方式どおりの貸方科目か（設定によって正解が変わる）",
           "累計額方式では、減損損失累計額も台帳と元帳の一致の対象になる（①）",
           "減損損失の科目とセグメントが、科目ごとの設定どおりか"]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント ｜ 損益あり", "減損",
        "計上方式の設定によって、貸方の勘定科目が変わる",
        "減損は直接控除方式と累計額方式のどちらを採るかで貸方科目が変わる。累計額方式の場合、減損損失累計額は①の層の一致の対象になる。減損損失の科目の細かさは科目ごとの設定を確認する。",
        "**直接控除方式**\n\n" + md_table([JHEAD] + direct) + "\n\n**累計額方式**\n\n" + md_table([JHEAD] + accum) + "\n\n" + bullets_md(pts))
    text(s, LEFT, 2.2, 7.5, 0.35, "直接控除方式", size=14, bold=True)
    journal(s, LEFT, 2.6, direct, l1_rows=(2,), l2_rows=(1,), row_h=0.4)
    text(s, LEFT, 4.0, 7.5, 0.35, "累計額方式", size=14, bold=True)
    journal(s, LEFT, 4.4, accum, l1_rows=(2,), l2_rows=(1,), row_h=0.4)
    journal_legend(s, LEFT, 5.85)
    points_card(s, 8.4, 2.2, 4.33, 4.5, tint, color, pts, pos=(1, 1))


def slide_14():
    rows = [["借", "固定資産圧縮損", "事業X", "製造1課・PJ-a", DASH, DASH],
            ["貸", "機械装置", "事業X", DASH, DASH, DASH]]
    pts = ["圧縮損の科目とセグメントが、科目ごとの設定どおりか",
           "取得と同時に計上する場合と、後日計上する場合の計上日・会計期間",
           "積立金方式を採るユーザーでは、固定資産側の仕訳が作られないこと（設定によって正解が変わる）"]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント ｜ 損益あり", "圧縮記帳",
        "直接減額方式では、圧縮損の行と固定資産勘定の行が作られる",
        "直接減額方式を中心に説明する。取得価額を直接減らすので、固定資産勘定の行は①の層の一致の対象になる。積立金方式を採るユーザーでは固定資産側の仕訳は作られないため、設定によって正解が変わる項目になる。",
        "**直接減額方式**\n\n" + md_table([JHEAD] + rows) + "\n\n" + bullets_md(pts))
    text(s, LEFT, 2.2, 7.5, 0.35, "直接減額方式", size=14, bold=True)
    journal(s, LEFT, 2.6, rows, l1_rows=(2,), l2_rows=(1,), row_h=0.5)
    journal_legend(s, LEFT, 4.3)
    points_card(s, 8.4, 2.2, 4.33, 4.5, tint, color, pts, pos=(1, 1))


DEP_TABLE = [["資産の種類", "製造部門で利用", "販売・管理部門で利用"],
             ["建物", "製造経費・減価償却費", "販管費・減価償却費"],
             ["機械装置", "製造経費・減価償却費", "販管費・減価償却費"],
             ["ソフトウェア", "製造経費・ソフトウェア償却費", "販管費・ソフトウェア償却費"],
             ["土地", "償却しない", "償却しない"]]


def slide_15():
    rows = [["借", "製造経費・減価償却費", "事業X", "製造1課・PJ-a", DASH, DASH],
            ["貸", "機械装置減価償却累計額", "事業X", DASH, DASH, DASH]]
    pts = ["資産の種類 × 利用セグメントのすべての組み合わせに、設定があるか",
           "勘定科目とP/Lセグメントが矛盾していないか（例：製造経費の科目に販売部門）",
           "累計額がない資産（ソフトウェア）は、貸方が資産勘定になるか"]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント ｜ 損益あり", "減価償却費①：償却費勘定のテーブル",
        "償却費の科目は、資産の種類に加えて、資産を利用するセグメントによっても変わる",
        "固定資産勘定・累計額は資産の種類だけで決まるが、償却費は資産の種類と利用セグメントの組み合わせで決まる。同じ機械でも、製造部門で使えば製造経費、販売部門で使えば販管費になる。",
        "**償却費勘定のテーブルの例**\n\n" + md_table(DEP_TABLE) + "\n\n" + md_table([JHEAD] + rows) + "\n\n" + bullets_md(pts))
    table(s, LEFT, 2.2, [1.7, 2.9, 2.9], DEP_TABLE, size=13, row_h=0.45, header_fill=L2, bold_cols=(0,),
          cell_fills={(4, 1): PAPER, (4, 2): PAPER})
    journal(s, LEFT, 4.75, rows, l1_rows=(2,), l2_rows=(1,), row_h=0.42)
    points_card(s, 8.4, 2.2, 4.33, 4.5, tint, color, pts, pos=(1, 1))


def slide_16():
    detail = [["資産", "利用する部門・PJ"],
              ["機械A", "製造1課・PJ-a"],
              ["機械B", "製造1課・PJ-a"],
              ["機械C", "販売1課・PJ-b"],
              ["機械D", "製造1課と販売1課で負担"]]
    rows = [["借", "製造経費・減価償却費", "事業X", "製造1課・PJ-a", DASH, DASH],
            ["借", "販管費・減価償却費", "事業X", "販売1課・PJ-b", DASH, DASH],
            ["貸", "機械装置減価償却累計額", "事業X", DASH, DASH, DASH]]
    pts = ["集約キーに、事業・部門・プロジェクトのすべてが含まれているか",
           "集約キーから部門・プロジェクトが漏れると、P/Lの情報が失われる",
           "複数の部門で負担する資産（機械D）は、部門ごとの行が漏れなく作られているか",
           "集約の単位を設定で変えられる場合は、設定ごとに確認する"]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント ｜ 損益あり", "減価償却費②：集約と配賦",
        "件数が最も多い仕訳。集約と配賦の仕様で、セグメントの情報が失われやすい",
        "減価償却費は資産の数だけ毎月発生するため、仕訳ファイルでは集約されることが多い。集約のキーにP/Lセグメントが含まれていないと、部門・プロジェクト別の損益が把握できなくなる。複数の部門で負担する資産は、1つの資産から複数の損益の行が作られる。",
        "**資産別の償却明細**\n\n" + md_table(detail) + "\n\n↓ 集約\n\n" + md_table([JHEAD] + rows) + "\n\n" + bullets_md(pts))
    text(s, LEFT, 2.15, 4, 0.35, "資産別の償却明細", size=13, bold=True)
    table(s, LEFT, 2.5, [1.3, 3.2], detail, size=12, row_h=0.34, header_fill=NEUTRAL)
    arrow(s, LEFT + 4.75, 3.05, 0.5, 0.55, color=MUTED, shape=MSO_SHAPE.DOWN_ARROW)
    text(s, LEFT + 5.3, 3.12, 1.6, 0.4, "集約", size=13, bold=True, color=MUTED)
    text(s, LEFT, 4.35, 4, 0.35, "仕訳ファイルの行", size=13, bold=True)
    journal(s, LEFT, 4.7, rows, l1_rows=(3,), l2_rows=(1, 2), row_h=0.4)
    points_card(s, 8.4, 2.2, 4.33, 4.5, tint, color, pts, size=13, pos=(1, 1))


def slide_17():
    change = [["台帳の変更", "償却費の振替"],
              ["資産の種類", "振り替わる"],
              ["事業（B/S）", "振り替わる"],
              ["部門・プロジェクト（P/L）", "振り替わる"],
              ["（参考）償却費以外の損益", "振り替わらない"]]
    rows = [["借", "販管費・減価償却費", "事業X", "販売1課・PJ-a", DASH, DASH],
            ["貸", "製造経費・減価償却費", "事業X", "製造1課・PJ-a", DASH, DASH]]
    policy = ["当期の期首から計上済みの償却費をすべて振り替える",
              "変更月以降の償却費だけを新しいセグメントで計上する",
              "変更日を基準に振り替える"]
    edge = ["変更月の当月分", "期首月の変更", "年度内に2回以上の変更", "変更後の売却・除却"]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント ｜ 損益あり", "減価償却費③：台帳変更時の振替",
        "どの変更でも償却費は振り替わる。振り替える範囲はユーザーごとの方針で決まる",
        "資産勘定は部門・プロジェクトの変更では振り替わらないが、償却費はどの変更でも振り替わる。設例では部門の変更によって、セグメントだけでなく勘定科目も製造経費から販管費に変わる。振り替える範囲はユーザーごとの方針なので、正解は1つではない。設定ごとに期待結果を用意して検証する。",
        md_table(change) + "\n\n**設例：機械Aの部門を製造1課 → 販売1課に変更（資産勘定の振替はなし）**\n\n" +
        md_table([JHEAD] + rows) + "\n\n" + bullets_md(policy, "振替範囲の方針の例（ユーザーごと）") + "\n\n" +
        bullets_md(edge, "境界のケース"))
    table(s, LEFT, 2.15, [3.6, 3.9], change, size=12, row_h=0.36, header_fill=L2, bold_cols=(0,),
          cell_fills={(5, 0): PAPER, (5, 1): PAPER})
    text(s, LEFT, 4.08, 7.6, 0.35, "設例：機械Aの部門を 製造1課 → 販売1課 に変更（資産勘定の振替はなし）",
         size=13, bold=True, color=L2)
    journal(s, LEFT, 4.45, rows, l2_rows=(1, 2), row_h=0.4)
    text(s, LEFT, 5.75, 7.5, 0.6, "→ セグメントだけでなく、勘定科目も変わる", size=13, bold=True)
    paras = [{"text": "振替範囲はユーザーごとの方針", "size": 15, "bold": True, "color": L2, "space_after": 6}]
    paras += [{"text": t, "bullet": True, "size": 13, "space_after": 4} for t in policy]
    paras += [{"text": "境界のケース", "size": 15, "bold": True, "color": L2, "space_after": 6}]
    paras += [{"text": t, "bullet": True, "size": 13, "space_after": 3} for t in edge]
    boxed_text(s, 8.4, 2.15, 4.33, 4.6, tint, paras, pad=0.22)
    quadrant_map(s, 8.62, 2.15 + 4.6 - 1.13, (1, 1), color)


def slide_18():
    under = [["借", "雑損", "事業X", "製造1課・PJ-a", DASH, DASH],
             ["貸", "機械装置減価償却累計額", "事業X", DASH, DASH, DASH]]
    over = [["借", "機械装置減価償却累計額", "事業X", DASH, DASH, DASH],
            ["貸", "雑収益", "事業X", "製造1課・PJ-a", DASH, DASH]]
    pts = ["訂正の方向によって、雑損と雑収益が正しく選ばれているか",
           "雑損・雑収益の科目とセグメントが、科目ごとの設定どおりか",
           "計上日は当期で、締め済みの期間に計上していないか",
           "台帳の過年度データの修正と、作られた仕訳が対応しているか（①）"]
    s, color, tint = content_slide(
        "l2", "② 仕訳に応じた確認ポイント ｜ 損益あり", "過年度の訂正",
        "固定資産勘定の相手勘定は雑損・雑収益。訂正の方向で科目が変わる",
        "過年度の誤りを当期に訂正する仕訳。固定資産勘定の相手勘定は雑損または雑収益になり、損益ありのグループに入る。計上日は当期なので、締め済みの期間に計上されないことも確認する。",
        "**償却不足を訂正する場合**\n\n" + md_table([JHEAD] + under) + "\n\n**償却過大を訂正する場合**\n\n" +
        md_table([JHEAD] + over) + "\n\n" + bullets_md(pts))
    text(s, LEFT, 2.2, 7.5, 0.35, "償却不足を訂正する場合", size=14, bold=True)
    journal(s, LEFT, 2.6, under, l1_rows=(2,), l2_rows=(1,), row_h=0.4)
    text(s, LEFT, 4.0, 7.5, 0.35, "償却過大を訂正する場合", size=14, bold=True)
    journal(s, LEFT, 4.4, over, l1_rows=(1,), l2_rows=(2,), row_h=0.4)
    journal_legend(s, LEFT, 5.85)
    points_card(s, 8.4, 2.2, 4.33, 4.5, tint, color, pts, pos=(1, 1))


CHECK = [
    ["層", "確認項目", "対象の仕訳", "設定で正解が変わる"],
    ["①", "3勘定の行の項目（勘定科目・事業・資産番号）が、全ファイルでそろっている", "すべて", ""],
    ["①", "資産の種類・分類ごとの勘定科目テーブルに漏れがない", "すべて", ""],
    ["①", "B/Sの項目の変更だけで資産勘定の振替が作られる", "振替", ""],
    ["①", "減損の計上方式に応じた貸方科目になっている", "減損", "○"],
    ["② 現金", "債権債務・消費税の行に取引先・税区分が入っている", "取得、売却", ""],
    ["② 現金", "購買・支払側との二重計上がない", "取得", ""],
    ["② セグメント", "損益の行に部門・PJが入り、事業と矛盾しない", "損益ありの全仕訳", ""],
    ["② セグメント", "償却費以外の損益は、発生時点の属性で科目ごとの設定どおり", "売却、除却、減損、圧縮記帳、訂正", "○"],
    ["② 償却費", "資産の種類 × 利用セグメントのテーブルに漏れ・矛盾がない", "減価償却費", ""],
    ["② 償却費", "集約キーに必要なセグメント項目がすべて含まれている", "減価償却費", "○"],
    ["② 償却費", "台帳変更時の償却費の振替範囲が方針どおり", "減価償却費", "○"],
    ["② その他", "圧縮記帳の方式、締め済み期間への計上の扱い", "圧縮記帳、訂正", "○"],
]


def slide_19():
    s, color, tint = content_slide(
        "intro", "まとめ", "検証チェックリスト",
        "①固定資産固有の観点と②仕訳に応じた確認ポイントの2層で確認する",
        "本日の内容をチェックリストにまとめた。『設定で正解が変わる』に○がある項目は、ユーザーの設定ごとに期待結果を用意して検証する。第2回以降の本番事例も、このチェックリストのどの項目に当たるかで整理して紹介する。",
        md_table(CHECK))
    fills = {}
    for r in range(1, len(CHECK)):
        fills[(r, 0)] = L1_LT if CHECK[r][0] == "①" else L2_LT
    table(s, LEFT, 2.0, [1.45, 6.0, 2.9, 1.78], CHECK, size=12, row_h=0.34, header_fill=NEUTRAL,
          cell_fills=fills, bold_cols=(0,), center_cols=(0, 3))


def slide_20():
    s = new_slide("第2回以降は本番事例を紹介する。事例は本日のチェックリストのどの項目に当たるかで整理する。質疑。")
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = DARK
    text(s, 0.9, 1.5, 11.5, 0.8, "次回予告", size=36, color=WHITE, bold=True)
    text(s, 0.9, 2.5, 11.5, 1.2,
         ["第2回〜第4回：本番事例",
          "本日のチェックリストのどの項目に当たるかで整理して紹介します"],
         size=20, color=rgb("D6E6E3"), space_after=10)
    msgs = ["固定資産の3勘定は、台帳と元帳の一致が最終的な検証基準",
            "資産の種類・分類ごとの勘定科目テーブルを考慮する",
            "それ以外の勘定科目は、仕訳に応じて確認する",
            "設定によって正解が変わる部分は、設定ごとに期待結果を用意する"]
    for i, m in enumerate(msgs):
        x = 0.9 + (i % 2) * 5.85
        y = 4.05 + (i // 2) * 1.0
        boxed_text(s, x, y, 5.65, 0.85, rgb("21474E"),
                   [{"runs": [(f"{i + 1}  ", {"bold": True, "color": rgb("9CC9C1")}), (m, {})]}],
                   size=15, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, pad=0.22)
    text(s, 0.9, 6.35, 11.5, 0.5, "質疑応答", size=20, color=WHITE, bold=True)
    MD.append("## 20. 次回予告・質疑\n\n- 第2回〜第4回：本番事例（本日のチェックリストのどの項目に当たるかで整理）\n\n"
              "**本日のメッセージ**\n\n" + "\n".join(f"{i + 1}. {m}" for i, m in enumerate(msgs)) + "\n")


for fn in [slide_01, slide_02, slide_03, slide_04, slide_05, slide_06, slide_07, slide_08, slide_09,
           slide_10, slide_11, slide_12, slide_13, slide_14, slide_15, slide_16, slide_17, slide_18,
           slide_19, slide_20]:
    fn()

prs.save(PPTX_OUT)
header = ("<!-- このファイルは build_deck.py から生成されます。内容の変更は build_deck.py で行ってください。 -->\n\n")
MD_OUT.write_text(header + "\n---\n\n".join(MD), encoding="utf-8")
print("saved", PPTX_OUT.name, MD_OUT.name, SLIDE_NO[0], "slides")
