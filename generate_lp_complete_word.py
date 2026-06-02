#!/usr/bin/env python3
"""
Sanhe Interfashion 日本向けLP完全設計書
WHY → LP根本思想 → 競合分析 → 具体的LP構成（実コピー付き）
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph as DocxPara

# ── Colors ─────────────────────────────────────────────────────────────────────
BLUE_DARK  = RGBColor(0x0f, 0x2d, 0x5e)
BLUE_MID   = RGBColor(0x1a, 0x56, 0xdb)
BLUE_LIGHT = RGBColor(0x3b, 0x82, 0xf6)
BLUE_PALE  = RGBColor(0x93, 0xc5, 0xfd)
BLUE_TINT  = RGBColor(0xdb, 0xea, 0xfe)
BLUE_BG    = RGBColor(0xef, 0xf6, 0xff)
WHITE      = RGBColor(0xff, 0xff, 0xff)
GRAY_DARK  = RGBColor(0x1e, 0x29, 0x3b)
GRAY_MID   = RGBColor(0x64, 0x74, 0x8b)
STRIPE1    = RGBColor(0xf8, 0xfa, 0xff)
GREEN_BG   = RGBColor(0xd1, 0xfa, 0xe5)
GREEN_FG   = RGBColor(0x06, 0x5f, 0x46)
AMBER_BG   = RGBColor(0xfe, 0xf3, 0xc7)
AMBER_FG   = RGBColor(0x78, 0x35, 0x0f)
RED_BG     = RGBColor(0xff, 0xe4, 0xe6)
RED_FG     = RGBColor(0x7f, 0x1d, 0x1d)
GOLD       = RGBColor(0xd4, 0xaf, 0x37)
DARK_BG    = RGBColor(0x0d, 0x1b, 0x38)

def hx(rgb): return '{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])

# ── XML helpers ────────────────────────────────────────────────────────────────
def set_bg(cell, rgb):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hx(rgb))
    tcPr.append(shd)

def set_margins(cell, top=80, bottom=80, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side, val in [('top',top),('bottom',bottom),('left',left),('right',right)]:
        m = OxmlElement(f'w:{side}'); m.set(qn('w:w'), str(val)); m.set(qn('w:type'), 'dxa'); tcMar.append(m)
    tcPr.append(tcMar)

def set_valign(cell, align='center'):
    tcPr = cell._tc.get_or_add_tcPr()
    v = OxmlElement('w:vAlign'); v.set(qn('w:val'), align); tcPr.append(v)

def no_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for s in ['top','left','bottom','right','insideH','insideV']:
        b = OxmlElement(f'w:{s}'); b.set(qn('w:val'), 'none'); borders.append(b)
    tblPr.append(borders)

def fp(cell, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0]; p.alignment = align
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    return p

def acp(cell, align=WD_ALIGN_PARAGRAPH.LEFT):
    new_p = OxmlElement('w:p'); cell._tc.append(new_p)
    para = DocxPara(new_p, cell._tc); para.alignment = align
    para.paragraph_format.space_before = Pt(0); para.paragraph_format.space_after = Pt(0)
    return para

def ar(para, text, bold=False, size=10, color=None, italic=False):
    r = para.add_run(text); r.bold = bold; r.italic = italic; r.font.size = Pt(size); r.font.name = 'Meiryo'
    if r._element.rPr is None: r._element.get_or_add_rPr()
    r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Meiryo')
    if color: r.font.color.rgb = color
    return r

def page_break(doc):
    p = doc.add_paragraph(); r = p.add_run()
    br = OxmlElement('w:br'); br.set(qn('w:type'), 'page'); r._r.append(br)
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)

def gap(doc, pt=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(pt)

# ── Block builders ─────────────────────────────────────────────────────────────
def chapter_banner(doc, num, title, subtitle=''):
    t = doc.add_table(rows=1, cols=1); t.style='Table Grid'; no_borders(t)
    cell = t.rows[0].cells[0]; cell.width = Inches(7.0)
    set_bg(cell, DARK_BG); set_margins(cell, 160, 160, 240, 240)
    p = fp(cell)
    ar(p, f'Chapter {num}  ', bold=True, size=10, color=BLUE_PALE)
    ar(p, title, bold=True, size=16, color=WHITE)
    if subtitle:
        np = acp(cell); np.paragraph_format.space_before = Pt(6)
        ar(np, subtitle, size=9.5, color=BLUE_PALE)
    gap(doc, 8)

def section_bar(doc, label, title, color=BLUE_TINT, text_color=BLUE_DARK):
    t = doc.add_table(rows=1, cols=1); t.style='Table Grid'; no_borders(t)
    cell = t.rows[0].cells[0]; cell.width = Inches(7.0)
    set_bg(cell, color); set_margins(cell, 80, 80, 160, 160)
    p = fp(cell)
    ar(p, f'{label}  ', bold=True, size=10, color=BLUE_MID)
    ar(p, title, bold=True, size=11, color=text_color)
    gap(doc, 5)

def lp_section_bar(doc, num, title, importance=''):
    t = doc.add_table(rows=1, cols=1); t.style='Table Grid'; no_borders(t)
    cell = t.rows[0].cells[0]; cell.width = Inches(7.0)
    if importance == 'CRITICAL':
        set_bg(cell, BLUE_DARK)
        p = fp(cell); set_margins(cell, 90, 90, 160, 160)
        ar(p, f'Section {num}  ', bold=True, size=10, color=BLUE_PALE)
        ar(p, title, bold=True, size=12, color=WHITE)
        ar(p, '  ★ 最重要', bold=True, size=9, color=GOLD)
    else:
        set_bg(cell, BLUE_TINT)
        p = fp(cell); set_margins(cell, 80, 80, 160, 160)
        ar(p, f'Section {num}  ', bold=True, size=10, color=BLUE_MID)
        ar(p, title, bold=True, size=11, color=BLUE_DARK)
        if importance:
            ar(p, f'  {importance}', size=9, color=GRAY_MID)
    gap(doc, 5)

def body(doc, text, bold=False, size=10, color=None, indent=0.0, after=3):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(after)
    if indent: p.paragraph_format.left_indent = Inches(indent)
    ar(p, text, bold=bold, size=size, color=color or GRAY_DARK)

def bullet(doc, text, size=9.5, indent=0.18, color=None, marker='▶ ', marker_color=None):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Inches(indent); p.paragraph_format.first_line_indent = Inches(-indent)
    ar(p, marker, bold=True, size=size, color=marker_color or BLUE_MID)
    ar(p, text, size=size, color=color or GRAY_DARK)

def check(doc, text, size=9.5):
    bullet(doc, text, size=size, marker='✓ ', marker_color=GREEN_FG)

def box(doc, title, lines, bg=None, title_color=None, title_size=11):
    bg = bg or BLUE_BG; title_color = title_color or BLUE_DARK
    t = doc.add_table(rows=1, cols=1); t.style='Table Grid'; no_borders(t)
    cell = t.rows[0].cells[0]; cell.width = Inches(7.0)
    set_bg(cell, bg); set_margins(cell, 120, 120, 200, 200)
    p = fp(cell); ar(p, title, bold=True, size=title_size, color=title_color)
    for line in lines:
        np = acp(cell); np.paragraph_format.left_indent = Inches(0.15); np.paragraph_format.space_before = Pt(4)
        ar(np, '• ' + line, size=9.5, color=GRAY_DARK)
    gap(doc, 8)

def two_col_box(doc, left_title, left_lines, right_title, right_lines,
                left_bg=RED_BG, right_bg=BLUE_BG, lt_c=RED_FG, rt_c=BLUE_DARK):
    t = doc.add_table(rows=1, cols=2); t.style='Table Grid'; no_borders(t)
    for j, (ttl, lines, bg, tc) in enumerate([
        (left_title, left_lines, left_bg, lt_c),
        (right_title, right_lines, right_bg, rt_c),
    ]):
        cell = t.rows[0].cells[j]; cell.width = Inches(3.4)
        set_bg(cell, bg); set_margins(cell, 110, 110, 160, 160)
        p = fp(cell); ar(p, ttl, bold=True, size=10, color=tc)
        for line in lines:
            np = acp(cell); np.paragraph_format.space_before = Pt(5)
            ar(np, '• ' + line, size=9, color=GRAY_DARK)
    gap(doc, 8)

def kv(doc, rows, col_w=(2.2, 4.8)):
    t = doc.add_table(rows=len(rows), cols=2); t.style='Table Grid'
    for i, (k, v) in enumerate(rows):
        kc = t.rows[i].cells[0]; vc = t.rows[i].cells[1]
        kc.width = Inches(col_w[0]); vc.width = Inches(col_w[1])
        set_bg(kc, BLUE_TINT if i%2==0 else BLUE_BG); set_margins(kc, 70,70,130,100); set_valign(kc)
        set_bg(vc, WHITE if i%2==0 else STRIPE1); set_margins(vc, 70,70,130,130); set_valign(vc)
        ar(fp(kc), k, bold=True, size=9.5, color=BLUE_DARK)
        ar(fp(vc), v, size=9.5, color=GRAY_DARK)
    gap(doc, 8)

def mtable(doc, headers, rows, col_w=None, center_from=1):
    cols = len(headers); col_w = col_w or [7.0/cols]*cols
    t = doc.add_table(rows=1+len(rows), cols=cols); t.style='Table Grid'
    for j, h in enumerate(headers):
        hc = t.rows[0].cells[j]; hc.width = Inches(col_w[j])
        set_bg(hc, BLUE_DARK); set_margins(hc, 70,70,110,110)
        ar(fp(hc, WD_ALIGN_PARAGRAPH.CENTER), h, bold=True, size=9, color=WHITE)
    for i, row in enumerate(rows):
        bg = WHITE if i%2==0 else STRIPE1
        for j, val in enumerate(row):
            dc = t.rows[i+1].cells[j]; dc.width = Inches(col_w[j])
            set_bg(dc, bg); set_margins(dc, 65,65,110,110)
            align = WD_ALIGN_PARAGRAPH.CENTER if j >= center_from else WD_ALIGN_PARAGRAPH.LEFT
            ar(fp(dc, align), val, size=9, color=GRAY_DARK)
    gap(doc, 8)

def lp_mock_section(doc, num_label, section_name, copy_title, copy_body, cta='', importance=''):
    lp_section_bar(doc, num_label, section_name, importance)
    if copy_title:
        body(doc, f'【実際のコピー案】', bold=True, size=9, color=BLUE_MID, after=1)
        body(doc, f'「{copy_title}」', bold=True, size=11, color=BLUE_DARK, after=2, indent=0.15)
    if copy_body:
        body(doc, copy_body, size=9.5, color=GRAY_DARK, indent=0.15, after=3)
    if cta:
        body(doc, f'CTA: {cta}', bold=True, size=9.5, color=BLUE_MID, indent=0.15, after=2)

# ══════════════════════════════════════════════════════════════════════════════
# BUILD
# ══════════════════════════════════════════════════════════════════════════════
doc = Document()
sec = doc.sections[0]
sec.page_width = Inches(8.5); sec.page_height = Inches(11)
sec.left_margin = sec.right_margin = Inches(0.75)
sec.top_margin  = sec.bottom_margin = Inches(0.75)
doc.styles['Normal'].font.name = 'Meiryo'
doc.styles['Normal'].paragraph_format.space_after  = Pt(0)
doc.styles['Normal'].paragraph_format.space_before = Pt(0)

# ─────────────────────────────────────────────────────────────────────────────
# COVER
# ─────────────────────────────────────────────────────────────────────────────
t = doc.add_table(rows=1, cols=1); t.style='Table Grid'; no_borders(t)
cc = t.rows[0].cells[0]; cc.width = Inches(7.0)
set_bg(cc, DARK_BG); set_margins(cc, 1000, 1000, 400, 400)
p = fp(cc, WD_ALIGN_PARAGRAPH.CENTER)
ar(p, 'SANHE INTERFASHION', bold=True, size=10, color=BLUE_PALE)
for txt, sz, clr, bd, sp in [
    ('日本向けLP  完全設計書', 20, WHITE, True, 14),
    ('', 6, WHITE, False, 6),
    ('WHY から始まる、完全具体化ガイド', 10, BLUE_PALE, False, 4),
    ('根本思想 → 競合分析 → 実コピー付きLP構成 → 実装チェックリスト', 9, BLUE_PALE, False, 3),
    ('', 6, WHITE, False, 12),
    ('Mywalitが15年選んだ工場の、日本市場戦略', 10, RGBColor(0xd4,0xaf,0x37), True, 4),
    ('', 6, WHITE, False, 8),
    ('2025年版', 9, GRAY_MID, False, 0),
]:
    np = acp(cc, WD_ALIGN_PARAGRAPH.CENTER); np.paragraph_format.space_before = Pt(sp)
    if txt: ar(np, txt, bold=bd, size=sz, color=clr)

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# Chapter 0 — WHY このLPを作るのか
# ─────────────────────────────────────────────────────────────────────────────
chapter_banner(doc, 0, 'WHY — このLPを作る本当の理由',
    'すべての設計判断はここから逆算する')

section_bar(doc, '0-1', 'Sanheが今、直面している問題')
body(doc, 'Sanheの実力は数字が証明している。しかし日本市場では「存在しない工場」と同じ状態だ。', bold=True, size=10, color=BLUE_DARK, after=5)

two_col_box(doc,
    '現状（問題）',
    ['日本語のウェブサイトが存在しない',
     '日本のバイヤーがGoogle・展示会以外でSanheを見つける経路がない',
     '英語LPを見ても、日本語バイヤーは信頼判断ができない',
     '商談機会が「紹介」か「展示会ブース」だけに依存している',
     'Mywalit・Roots・Jigsaw等の実績が日本では全く知られていない'],
    '失っている機会',
    ['月に何件の日本バイヤーが韓国革製品OEMを探しているか（推定: 数十〜数百件/月）',
     'その検索者がSanheに辿り着けていない',
     '辿り着いても英語しかなく離脱している',
     '競合（中国系OEM工場・国内エージェント）がその機会を取っている'],
    left_bg=RED_BG, right_bg=AMBER_BG, lt_c=RED_FG, rt_c=AMBER_FG,
)

section_bar(doc, '0-2', 'このLPが解くべき「1つの問い」')
body(doc, 'LPが答えるべき問いは1つだけ。訪問者の頭にある、たった1つの疑問をクリアすることがすべての設計の軸になる。', size=9.5, color=GRAY_MID, after=6)

t = doc.add_table(rows=1, cols=1); t.style='Table Grid'; no_borders(t)
cell = t.rows[0].cells[0]; set_bg(cell, BLUE_DARK); set_margins(cell, 160, 160, 280, 280)
p = fp(cell, WD_ALIGN_PARAGRAPH.CENTER)
ar(p, '「このアジアの工場に発注して、本当に大丈夫か？」', bold=True, size=13, color=WHITE)
np = acp(cell, WD_ALIGN_PARAGRAPH.CENTER); np.paragraph_format.space_before = Pt(10)
ar(np, 'LPのすべてのセクションはこの1つの不安を「証拠で」解消するために存在する', size=9.5, color=BLUE_PALE)
gap(doc, 8)

section_bar(doc, '0-3', 'LPの成功定義（具体的に何が起きれば成功か）')
mtable(doc,
    ['KPI', '初期目標', '意味'],
    [['月間問い合わせ数', '3〜5件/月', '欧米実績 → 日本展開なら十分達成可能な数字'],
     ['サンプル請求率', '問い合わせの50%以上', 'サンプルを見た担当者は高確率で上長に提案する'],
     ['LP→問い合わせ CVR', '2〜5%', '良質なBtoB LPの業界標準値'],
     ['スペックシートDL数', '月10件以上', '検討フェーズのバイヤーをリード化する指標']],
    col_w=[2.0, 1.8, 3.2], center_from=1,
)

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# Chapter 1 — LPの根本思想
# ─────────────────────────────────────────────────────────────────────────────
chapter_banner(doc, 1, 'LPの根本思想（深掘り）',
    '「なぜこの順番か」「なぜこの言葉か」を理解しないと改善できない')

section_bar(doc, '1-1', '5秒ルール — 訪問者がページに与える時間')
body(doc, '訪問者の約57%はページを開いてから5〜8秒以内に「読む/離脱」を決める。最初の画面（ファーストビュー）がすべてを決定する。', size=9.5, after=5)
box(doc, '▶ ファーストビューで5秒以内に伝えるべき3要素',
    ['① 何ができる会社か（What） → 「革製品OEMメーカー、業歴23年」',
     '② なぜ信頼できるか（Why Trust） → 「Mywalit・Roots等欧米5ブランドが15年以上継続取引」',
     '③ 次に何をすればよいか（CTA） → 「サンプルを依頼する」ボタン'],
    bg=GREEN_BG, title_color=GREEN_FG)
box(doc, '❌ Sanheのファーストビューにあってはいけないもの',
    ['"高品質な製品をご提供します" ← 証拠がなく全社が言っている',
     '"OEM・ODMに対応しております" ← ほぼ全ての競合も言っている',
     '工場の外観写真 ← バイヤーが知りたいのは製品のクオリティ、工場の建物ではない',
     '会社概要からスタートする構成 ← 自己紹介より先に相手の課題に共感する'],
    bg=RED_BG, title_color=RED_FG)

section_bar(doc, '1-2', 'BtoBとBtoC、LPの設計は根本的に違う')
mtable(doc,
    ['観点', 'BtoC（一般消費者向け）', 'BtoB（日本バイヤー向け）= Sanheのケース'],
    [['意思決定者',  '1人',        '担当バイヤー ＋ MD/上長の複数人（稟議が必要）'],
     ['決定期間',   '即日〜数日',  '数週間〜数ヶ月（初回発注は特に長い）'],
     ['動機',       '欲求・感情',  '課題解決・リスク回避・上長への説明責任'],
     ['CTAの設計',  '「今すぐ購入」が効く', '「まず話を聞く」「サンプルだけ見てみる」が効く'],
     ['ページの役割','購入完結',   '稟議に使える「証拠書類」を提供することが目的'],
     ['最も重要な要素','写真・価格・レビュー', '実績数字・顧客名・QCプロセス・スペック表']],
    col_w=[1.5, 1.8, 3.7], center_from=99,
)

section_bar(doc, '1-3', '信頼の階段（Trust Ladder）— 各セクションの役割')
body(doc, '1ページのLPは「バイヤーの頭の中の対話」に沿って設計される。セクションは気分で並べるのではなく、信頼を段階的に積み上げる順番で並ぶ。', size=9.5, color=GRAY_MID, after=5)
mtable(doc,
    ['階段', 'バイヤーの内心', '対応するLPセクション', 'Sanheの答え'],
    [['第1段 関連性', '自分に関係ある工場か？', 'Hero（製品カテゴリ・MOQ）', '革製品OEM・300個〜・日本語対応'],
     ['第2段 実力', 'ちゃんとできる工場か？', '信頼数字・スペック表', '50万個・3,000スタイル・95%納期遵守'],
     ['第3段 信頼性', '飛ばれない工場か？', 'クライアント実績・継続年数', 'Mywalit等5ブランド・15年以上継続'],
     ['第4段 安全性', '品質ミスが起きないか？', '4段階QCプロセス・監査対応', '不良率1%未満・Factory Audit可'],
     ['第5段 実現性', '自分のブランドに合うか？', '製品写真・カスタマイズ範囲', 'OEM/ODM・社内デザイナー在籍'],
     ['第6段 行動', '次に何をすべきか？', 'CTA・FAQ', 'サンプル請求・スペックシートDL']],
    col_w=[1.3, 1.9, 1.9, 1.9], center_from=99,
)

section_bar(doc, '1-4', '「具体性の法則」— Sanheにとって最も重要なルール')
body(doc, '根拠のない主張は競合と区別がつかない。Sanheは具体的な数字・名前・年数を持っているのに、それを使わないのは最大の機会損失だ。', size=9.5, color=GRAY_MID, after=5)
mtable(doc,
    ['NG表現（抽象）', '→', 'OK表現（具体・Sanheの実データ）'],
    [['「高品質な製品を提供します」',        '→', '「不良品発生率1%未満。4段階品質管理プロセスで継続維持」'],
     ['「豊富な実績があります」',            '→', '「累計50万個・3,000以上のスタイルを製造」'],
     ['「一流ブランドとの取引実績あり」',    '→', '「Mywalit（イタリア）・Roots Canada・Jigsaw UKと15年以上継続取引」'],
     ['「迅速に対応します」',               '→', '「納期遵守率95%以上。青島→日本 海運3〜7日」'],
     ['「小ロットから対応可能」',           '→', '「MOQ 300個〜。現在30〜40%の余剰生産キャパシティあり」'],
     ['「品質管理を徹底しています」',       '→', '「原材料・生産中・最終・出荷前の4段階検品。Factory Audit受け入れ可」']],
    col_w=[2.5, 0.3, 4.2], center_from=99,
)

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# Chapter 2 — 競合LP分析
# ─────────────────────────────────────────────────────────────────────────────
chapter_banner(doc, 2, '競合LP分析 — Sanheが参考にすべき要素',
    '真似すべき構造・使える要素・差別化できるポイント')

section_bar(doc, '2-1', '同類OEM工場LPの共通構造（業界標準）')
body(doc, '革製品OEM工場の日本向けLPを構造的に分析すると、以下の要素が「あるか/ないか」で信頼感が大きく変わることがわかる。', size=9.5, color=GRAY_MID, after=5)
mtable(doc,
    ['LP要素', '競合に多い？', 'なぜ重要か', 'Sanheの現状'],
    [['数字バー（実績の可視化）',       '中国系:30% / 欧米系:80%', '証拠なき主張より数字1つが100倍説得力ある', '✅ データあり→使っていない'],
     ['クライアントロゴ帯',             '中国系:10% / 欧米系:90%', 'バイヤーが知ってる名前が1つあれば信頼が跳ぶ', '✅ 5ブランドあり→LPに未掲載'],
     ['製品の高品質写真（白背景）',      '中国系:20% / 欧米系:95%', '写真の質がそのまま工場の品質イメージになる', '❌ 要撮影（最重要準備）'],
     ['4〜6ステップのQCプロセス図',    '中国系:40% / 欧米系:70%', 'プロセスの可視化が「安心感」を生む', '✅ 4段階データあり'],
     ['FAQ（最低5Q&A）',              '中国系:50% / 欧米系:60%', 'バイヤーの残り不安を事前につぶす', '✅ データあり→構成化が必要'],
     ['日本語LINEでの問い合わせ導線',   '中国系:60% / 欧米系:5%', '日本バイヤーへのフリクション最小化に必須', '✅ 対応可→LPに組み込む'],
     ['サンプル請求フォーム',           '中国系:30% / 欧米系:70%', 'CTAの最高到達点。あるだけで問い合わせ率倍増', '未整備'],
     ['工場内部・スタッフ写真',         '中国系:70% / 欧米系:40%', '清潔感・規模感・人の顔が見える安心感', '未整備']],
    col_w=[2.0, 1.5, 2.2, 1.3], center_from=99,
)

section_bar(doc, '2-2', 'Sanheが「絶対に真似してはいけない」競合の失敗パターン')
two_col_box(doc,
    '中国系OEM工場LPの典型的失敗',
    ['Heroが「工場の外観写真」+ "Welcome to our factory"',
     '価格だけを前面に出して品質の証拠がない',
     '顧客名は「機密のため非公開」で実績がゼロに見える',
     '「高品質・丁寧・迅速対応」の3点セット（証拠なし）',
     '問い合わせ先がメールアドレス1行だけ'],
    '欧米系OEM工場LPの典型的失敗',
    ['日本語対応がなく、機械翻訳の日本語が逆に信頼を下げる',
     'Heritage（創業ストーリー）が長すぎてバイヤーが求める情報に辿りつけない',
     'MOQ・リードタイム・FOB価格が「要問い合わせ」のみで比較できない',
     'スマホ未対応で日本のバイヤーが閲覧できない'],
    lt_c=RED_FG, rt_c=AMBER_FG, left_bg=RED_BG, right_bg=AMBER_BG,
)

section_bar(doc, '2-3', 'Sanheが真似すべき「良いLP」3つの要素')

body(doc, '① 数字の可視化バー（参考:製造業・人材系BtoB LPに多い）', bold=True, size=10, color=BLUE_DARK, after=2)
box(doc, 'なぜ効くか',
    ['数字はスキャン読みされても伝わる。「50万個」は読み飛ばされない',
     '「品質が高い」という主張 vs「不良率1%未満」という数字では、後者が100倍信頼される',
     'Sanheの場合: 23年・50万個・3,000スタイル・95%・1%・80% の6指標をHero直下に横並び表示'],
)

body(doc, '② Mywalitストーリー活用（参考: IKEAのSUPPLIERS紹介、BtoB製造業の"顧客インタビュー"形式）', bold=True, size=10, color=BLUE_DARK, after=2)
box(doc, 'なぜ効くか',
    ['ロゴを並べるだけよりも、「なぜ選ばれたか」のストーリーが読まれる',
     '「イタリアブランドが選んだ理由」は日本バイヤーへの強力な代理説得になる',
     'Mywalitはレザーグッズ専門・イタリア拠点 → 日本人が最も品質を信頼する産地の一つ',
     '15年継続という事実が「一時的なラッキーではない」ことを示す'],
)

body(doc, '③ 3段階CTA設計（参考: SaaS系BtoB LPの標準設計）', bold=True, size=10, color=BLUE_DARK, after=2)
box(doc, 'なぜ効くか',
    ['「今すぐ発注」は日本BtoBでは機能しない。入口を3段階に分けて、ハードルの低い順に並べる',
     '高意向: 「サンプルを依頼する」（決め手になる）',
     '中意向: 「スペックシート・会社概要PDFをDL」（稟議資料として使える）',
     '低意向: 「LINEで気軽に相談」（言語バリアをなくす最も低摩擦の入口）',
     'Sanheは3つ全部対応できる → CTAをこの3つに絞る'],
)

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# Chapter 3 — Sanhe 日本語LP 完全構成（実コピー付き）
# ─────────────────────────────────────────────────────────────────────────────
chapter_banner(doc, 3, 'Sanhe 日本語LP 完全構成（実コピー付き）',
    'このまま制作に渡せる構成 × コピー × 理由のセット')

body(doc, '以下はLP制作者（デザイナー/コーダー）にそのまま渡せる構成書です。', size=9.5, color=GRAY_MID, after=2)
body(doc, '各セクションに「なぜこの順番か」「何を見せるか」「実際のコピー案」の3点セットを記載しています。', size=9.5, color=GRAY_MID, after=8)

# Section 1: Hero
lp_section_bar(doc, '①', 'HERO — ファーストビュー', 'CRITICAL')
body(doc, '【なぜここに置くか】 最初の5秒。この画面だけで「読む価値があるか」が決まる。Sanheの最強カードであるMywalit実績を冒頭に出す。', size=9.5, color=GRAY_MID, after=5)

box(doc, '✅ 推奨コピー（案A）— 実績・信頼先行型',
    ['メインヘッドライン（大文字・最上部）:',
     '　「Mywalitが15年選び続けた工場が、今日本のブランドと組みます。」',
     '',
     'サブヘッドライン（ヘッドの直下）:',
     '　欧米5ブランド・業歴23年・累計50万個・不良率1%未満。',
     '　その品質と生産体制が、今あなたのブランドに使えます。',
     '',
     'CTA（ボタン2つ）:',
     '　[サンプルを依頼する →]   [スペックシートをDL →]',
     '',
     'サブテキスト（CTAの下・小文字）:',
     '　日本語担当チームあり ｜ MOQ 300個〜 ｜ 青島→日本 最短3日'],
    bg=GREEN_BG, title_color=GREEN_FG)

box(doc, '✅ 推奨コピー（案B）— 問題提起型（バイヤーの悩みから入る）',
    ['メインヘッドライン:',
     '　「なぜ、革製品の本場イタリアのブランドが韓国・中国工場を選んだのか。」',
     '',
     'サブヘッドライン:',
     '　答えは品質にある。Sanheは欧米5ブランドに23年、品質を届けてきた。',
     '',
     'CTA: [まずサンプルを見てみる →]   [詳細スペックを確認する →]'],
    bg=BLUE_BG, title_color=BLUE_DARK)

kv(doc, [
    ('ビジュアル',       '革財布・バッグの高品質プロ撮影写真（白背景）。複数アイテムを並べた俯瞰ショット推奨'),
    ('Trust badge',     '右上または Hero 内に小さく: 「日本語対応 ✓」「Factory Audit 対応 ✓」「MOQ 300個〜」'),
    ('NG例',            '工場外観写真・"Welcome to Sanhe"の文章・会社概要テキストのベタ書き'),
], col_w=(1.8, 5.2))

# Section 2: 数字バー
lp_section_bar(doc, '②', '信頼数字バー — 5秒で実績を刻む', '★ 高重要度')
body(doc, '【なぜここに置くか】 Heroの直下。スクロール1回目で見える位置に実績数字を横並びで表示。スキャン読みするバイヤーの目に刺さる。', size=9.5, color=GRAY_MID, after=5)

mtable(doc,
    ['表示数字', 'ラベル', '補足テキスト（小さめ）'],
    [['23年以上',  '業歴',         '韓国での創業から継続'],
     ['50万個超',  '累計総生産',   'バッグ・財布・ベルト合計'],
     ['3,000超',   'スタイル実績', 'OEM・ODM含む'],
     ['95%以上',   '納期遵守率',   '継続的に高水準を維持'],
     ['1%未満',    '不良品発生率', '4段階QCプロセスによる'],
     ['80%以上',   'リピート率',   '長期継続取引の割合']],
    col_w=[1.1, 1.3, 4.6], center_from=99,
)

# Section 3: クライアント実績
lp_section_bar(doc, '③', 'クライアント実績 — 15年継続の意味', 'CRITICAL')
body(doc, '【なぜここに置くか】 「知っているブランドが使っている」という事実は、どんな言葉よりも強い。ここでMywalitを筆頭に5ブランドを列挙する。', size=9.5, color=GRAY_MID, after=5)

box(doc, '【実際のコピー案】 セクションタイトル',
    ['"1回発注したら終わり"ではありません。',
     'イギリス、カナダ、イタリア——欧米5ブランドが、15年以上Sanheを使い続けています。',
     '',
     '（ブランド名横並び表示）',
     'Jigsaw  ｜  Roots Canada  ｜  Mywalit  ｜  Want Les Essentiels  ｜  Quartz Co.',
     '',
     '（サブテキスト）',
     'すべて15年以上の継続取引。一時的な価格競争ではなく、品質への信頼が関係を支えています。'],
    bg=BLUE_BG, title_color=BLUE_DARK)

# Section 4: Mywalit Story（最重要コンテンツ）
lp_section_bar(doc, '④', 'Mywalitストーリー — LP最強の信頼証拠', 'CRITICAL ★★★')
body(doc,
    '【なぜここに置くか】 Mywalitはイタリアのプレミアムレザーブランド。「革製品の本場」イタリアのブランドがアジア工場を15年使い続けているという事実は、'
    '日本バイヤーへの最強の代理証言になる。このセクションはSanheにしか書けない。',
    size=9.5, color=GRAY_MID, after=5)

t = doc.add_table(rows=1, cols=1); t.style='Table Grid'; no_borders(t)
cell = t.rows[0].cells[0]; set_bg(cell, DARK_BG); set_margins(cell, 160, 160, 240, 240)
p = fp(cell)
ar(p, '【実際のコピー案】 Mywalitセクション全文', bold=True, size=10, color=GOLD)
np = acp(cell); np.paragraph_format.space_before = Pt(10)
ar(np, 'なぜ、革製品の本場イタリアのブランドが、アジアの工場を15年選び続けるのか。', bold=True, size=12, color=WHITE)
np2 = acp(cell); np2.paragraph_format.space_before = Pt(10)
ar(np2,
    'Mywalitはイギリス・マルタに拠点を持つプレミアムレザーグッズブランドです。\n'
    '素材へのこだわりと品質基準の高さで知られるMywalitが、15年以上にわたりSanheを\n'
    '製造パートナーとして選び続けています。',
    size=9.5, color=BLUE_PALE)
np3 = acp(cell); np3.paragraph_format.space_before = Pt(10)
ar(np3,
    'その理由は一つ。Sanheが「欧米品質基準を、アジアのコスト効率で実現できる工場」だからです。\n'
    'イタリアブランドが認めた品質は、あなたのブランドにも使えます。',
    size=9.5, color=WHITE)
np4 = acp(cell); np4.paragraph_format.space_before = Pt(12)
ar(np4, '→ 「イタリアが認めた品質 × 日本語対応チーム」 = あなたのブランドへ', bold=True, size=10, color=GOLD)
gap(doc, 8)

body(doc, '【設計メモ】 このセクションはロゴ1個＋15年という数字＋1段落のコピーで完結させる。長くしない。', size=9, color=GRAY_MID, indent=0.15, after=8)

# Section 5: 製品・価格
lp_section_bar(doc, '⑤', '製品ラインアップ・価格 — 具体数字で比較できる状態に')
body(doc, '【なぜここに置くか】 「自分のブランドに使えるか？」の実務確認。MOQ・価格・リードタイムが明確でないと問い合わせに進めない。', size=9.5, color=GRAY_MID, after=5)

mtable(doc,
    ['製品カテゴリ', 'FOB単価', '円換算目安', 'MOQ', 'リードタイム'],
    [['ハンドバッグ',         'USD 25〜120+', '約4,000〜19,000円', '300個〜', '60〜90日'],
     ['財布・スモールレザーグッズ', 'USD 8〜45', '約1,300〜7,200円', '300個〜', '60〜75日'],
     ['ベルト',               'USD 10〜35',  '約1,600〜5,600円', '300個〜', '60〜75日']],
    col_w=[1.9, 1.2, 1.8, 1.1, 1.0], center_from=1,
)
body(doc, '※ FOB単価。1 USD ≒ 159円（参考値）。実際の価格は仕様・数量・素材により変動。現在30〜40%の余剰生産キャパシティあり。',
    size=8.5, color=GRAY_MID, indent=0.15, after=8)

# Section 6: QC
lp_section_bar(doc, '⑥', '4段階品質管理プロセス — 不良率1%未満の実現方法')
body(doc, '【なぜここに置くか】 「品質が心配」はバイヤー最大の不安。プロセスを見せることで「仕組みとして品質管理している」ことが伝わる。', size=9.5, color=GRAY_MID, after=5)

mtable(doc,
    ['工程', 'タイミング', '内容', '目的'],
    [['① 原材料検査',  '入荷時',      '革・副資材の品質・仕様・数量確認', 'NG素材を工程に入れない'],
     ['② 生産中検査',  '製造工程中',  '縫製・加工の精度を随時チェック', '工程内不良の早期発見・修正'],
     ['③ 最終検査',    '完成品',      '全数検品・規格への適合確認',     '出荷前の最終品質ゲート'],
     ['④ 出荷前検査',  '梱包・出荷直前','梱包・数量・ラベル・輸出仕様確認', '輸送中のリスクを最小化']],
    col_w=[1.4, 1.2, 2.5, 1.9], center_from=99,
)
box(doc, '★ Factory Audit（工場立会検査）受け入れ可',
    ['バイヤーによる工場視察・立会検査を受け入れています',
     '事前予約制（メール/LINEにてご連絡ください）',
     '中国・青島工場 / ベトナム工場 どちらも対応可能',
     '初回取引前の視察を推奨。信頼関係の起点になります'],
    bg=GREEN_BG, title_color=GREEN_FG)

# Section 7: カスタマイズ
lp_section_bar(doc, '⑦', 'カスタマイズ対応範囲 — OEM/ODM両対応')
body(doc, '【なぜここに置くか】 「自分のブランドに合わせてもらえるか」の確認。項目を列挙することで「何でも頼める工場」の印象を与える。', size=9.5, color=GRAY_MID, after=5)

mtable(doc,
    ['カスタマイズ項目', '対応', '補足'],
    [['ハードウェア（金具）素材・形状・色',  '○', '真鍮・亜鉛・ステンレス等から選択可'],
     ['裏地・内装素材の指定',               '○', '顧客指定素材の持ち込み可'],
     ['YKKジッパー',                        '○', '国際品質基準のジッパーを標準採用'],
     ['ロゴ・ブランドタグ',                  '○', '刻印・箔押し・刺繍・プリント対応'],
     ['パッケージ・化粧箱・下げ札',          '○', 'ブランド世界観に合わせた包材設計'],
     ['ステッチカラー',                      '○', 'PantoneカラーベースでのカラーマッチOK'],
     ['素材の産地指定',                      '○', '韓国産・イタリア産・顧客指定素材対応可'],
     ['ODM（デザイン共同開発）',             '○', '社内デザイナー2〜3名在籍']],
    col_w=[2.6, 0.7, 3.7], center_from=1,
)

page_break(doc)

# Section 8: 物流
lp_section_bar(doc, '⑧', '日本向け物流・納品条件 — 具体的な数字で安心を')
body(doc, '【なぜここに置くか】 「どのくらいで届くか」はバイヤーが必ず確認する。具体的な日数と対応インコタームズを明記することで比較検討が容易になる。', size=9.5, color=GRAY_MID, after=5)

mtable(doc,
    ['出荷拠点', '輸送方法', '日本到着目安', 'インコタームズ'],
    [['中国・青島工場', '海運', '3〜7日', 'FOB / CIF / DDP（関税込み納品）対応'],
     ['ベトナム工場',   '海運', '5〜10日', 'FOB / CIF / DDP 対応'],
     ['緊急時',        '航空便', '2〜3日', '要相談（費用は実費）']],
    col_w=[1.5, 1.0, 1.3, 3.2], center_from=99,
)
body(doc, 'DDP（Delivered Duty Paid）対応: 関税・通関手続きをSanheが代行。日本バイヤーは国内受け取りだけで完結。通関サポートあり。',
    size=9.5, color=GRAY_DARK, indent=0.15, after=8)

# Section 9: FAQ
lp_section_bar(doc, '⑨', 'よくある質問（FAQ） — 残りの不安をすべてつぶす')
body(doc, '【なぜここに置くか】 CTAの直前。ここまで読んできたバイヤーの「最後の迷い」を取り除く。問い合わせのハードルを下げる最終関門。', size=9.5, color=GRAY_MID, after=5)

kv(doc, [
    ('Q: 最低発注数（MOQ）はいくつですか？',
     'A: 300個〜です。初回は少量から試したい場合はご相談ください。製品・仕様によって応相談となります。'),
    ('Q: サンプルは作ってもらえますか？',
     'A: 対応可能です。サンプル費用・製作期間は仕様により異なりますので、お気軽にご相談ください。'),
    ('Q: 日本語でやり取りできますか？',
     'A: はい。Sanhe Japan Teamが日本語で対応いたします。Email・LINE・WhatsApp対応可。'),
    ('Q: 工場を見学することはできますか？',
     'A: 可能です（事前予約制）。中国・青島工場とベトナム工場、どちらも受け入れ可能です。'),
    ('Q: 納期はどのくらいかかりますか？',
     'A: 発注確定から60〜90日が目安です（製品・数量により変動）。リードタイム最新情報はお問い合わせください。'),
    ('Q: NDA（秘密保持契約）は締結できますか？',
     'A: はい、対応可能です。デザイン持ち込みの場合は締結を推奨しています。'),
    ('Q: OEMとODMの違いは？どちらに対応していますか？',
     'A: 両方対応しています。OEM（お客様デザイン持ち込み）・ODM（Sanhe社内デザイナーによる共同開発）。'),
    ('Q: 支払い方法と条件は？',
     'A: T/T（電信送金）30%前払い＋残70%出荷前が基本です。L/C（信用状）にも対応しています。'),
], col_w=(3.0, 4.0))

# Section 10: CTA
lp_section_bar(doc, '⑩', 'CTA・お問い合わせ — 3段階の入口設計', 'CRITICAL')
body(doc, '【なぜここに置くか】 最高到達点。CTAは「1つ」ではなく「3段階」で設計する。バイヤーの検討フェーズに合わせた入口を複数用意することで、取りこぼしを防ぐ。', size=9.5, color=GRAY_MID, after=5)

mtable(doc,
    ['CTA', '想定ユーザー', '実際のボタン文言', '獲得するもの'],
    [['① 最重要CTA', '具体的に検討中', '「まずサンプルを依頼してみる」',    '商談接点＋バイヤー情報'],
     ['② 第二CTA',   '比較・検討中',   '「スペックシート・会社概要をDL」',  'メールアドレス（リード化）'],
     ['③ 第三CTA',   '初めて知った',   '「LINEで気軽に相談する（日本語OK）」', '低摩擦の最初の一歩']],
    col_w=[1.0, 1.4, 2.3, 2.3], center_from=99,
)

box(doc, '▶ フォーム設計（サンプル依頼フォームの項目）',
    ['会社名（必須）',
     'お名前（必須）',
     'メールアドレス（必須）',
     '電話番号（任意）',
     '製品カテゴリ（必須: ハンドバッグ / 財布・SLG / ベルト / その他）',
     '発注数量の目安（任意: 〜500個 / 500〜1000個 / 1000個以上）',
     'ご質問・ご要望（任意・自由記述）',
     '→ 以上7項目以内。多すぎる入力は離脱の原因になる'],
    bg=BLUE_BG, title_color=BLUE_DARK)

kv(doc, [
    ('連絡先情報', 'Sanhe Interfashion Japan Team'),
    ('メール',     '________@sanhe.com（LP掲載時に更新）'),
    ('LINE / WhatsApp', '対応可（QRコードをLP上に掲載することで日本バイヤーの入り口を最大化）'),
    ('本社',       '韓国（管理本部）'),
    ('生産工場',   '中国・青島 ／ ベトナム'),
], col_w=(1.8, 5.2))

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# Chapter 4 — 実装チェックリスト
# ─────────────────────────────────────────────────────────────────────────────
chapter_banner(doc, 4, '実装準備チェックリスト',
    'LP制作を開始する前に揃えるべきものリスト')

section_bar(doc, '4-1', '制作開始前に準備するコンテンツ')
body(doc, '以下がすべて揃ってから制作に入る。特に写真と許諾確認を先に動かすこと。', size=9.5, color=GRAY_MID, after=5)

mtable(doc,
    ['準備物', '重要度', '現状', 'アクション'],
    [['製品写真（白背景・高解像度）',      '★★★ 必須', '未撮影', 'プロカメラマン依頼（韓国・青島）。費用: ¥3〜8万'],
     ['工場内部・製造ライン写真',          '★★★ 必須', '未整備', '明るい場所・清潔感のある構図で20〜30枚'],
     ['Mywalitロゴの使用許可',             '★★★ 必須', '未確認', 'Mywalit担当者へ書面で使用許可申請（メール可）'],
     ['他4ブランドのロゴ使用許可',         '★★ 推奨',  '未確認', 'Jigsaw・Roots・WLE・Quartzそれぞれに確認'],
     ['日本語コピー（全10セクション）',     '★★★ 必須', '本書が基準', 'このガイドのコピー案を日本語ネイティブが校正'],
     ['スペックシートPDF',                 '★★ 推奨',  '会社概要あり', '日本語版スペックシート（本書の内容を1枚に）'],
     ['担当者の写真・プロフィール',         '★ あれば', '未整備', '日本語窓口担当の顔写真＋名前掲載で信頼UP'],
     ['QRコード（LINE / WhatsApp）',       '★★ 推奨',  '未確認', 'Japan Team のLINEアカウント or WhatsApp QR生成'],
     ['Google Analytics 4設置',           '★★★ 必須', '未設置', 'LP公開と同時に設置。CVR測定に必須']],
    col_w=[2.2, 0.9, 0.9, 2.0], center_from=99,
)

section_bar(doc, '4-2', 'ツール選定（どこで作るか）')
mtable(doc,
    ['ツール', '難易度', '月額', '日本語LP適性', '推奨タイミング'],
    [['STUDIO',      '易',    '¥0〜3,000', '◎ 最適',  '自作スタート推奨。日本語UIで操作しやすい'],
     ['Webflow',     '中',    '$14〜',     '◎ 最適',  '外注発注時の標準。デザイナーへの依頼が容易'],
     ['Framer',      '易〜中', '$0〜',     '○',       'AI生成機能で英語版LP作成に有利'],
     ['WordPress',   '中〜難', '$5〜',     '○',       '長期的に自社管理したい場合'],
     ['Wix',         '易',    '$17〜',     '△',       'カスタマイズ限界があるため最終手段']],
    col_w=[1.2, 0.7, 0.9, 1.1, 3.1], center_from=1,
)

section_bar(doc, '4-3', '公開後の最優先アクション')
for item in [
    'Google Analytics 4で「問い合わせ完了」のゴールイベントを設定する',
    'Google Search Consoleにサイトマップを送信（検索流入を始める）',
    'LinkedIn・Instagramのプロフィールに日本語LPのURLを追加する',
    '展示会ブース（ライフスタイルWeek等）のパンフに日本語LP URLとQRコードを印刷する',
    '公開後2週間、毎日ヒートマップ（Hotjar等）でスクロール深度を確認する',
    '問い合わせ第1件目が来たら、対応時間・回答品質を記録し、テンプレート化する',
]:
    check(doc, item)

gap(doc, 10)

# FINAL SUMMARY
t = doc.add_table(rows=1, cols=1); t.style='Table Grid'; no_borders(t)
cell = t.rows[0].cells[0]; set_bg(cell, DARK_BG); set_margins(cell, 180, 180, 260, 260)
p = fp(cell, WD_ALIGN_PARAGRAPH.CENTER)
ar(p, '最終サマリー — このLPで伝えること、1文で', bold=True, size=11, color=GOLD)
np = acp(cell, WD_ALIGN_PARAGRAPH.CENTER); np.paragraph_format.space_before = Pt(14)
ar(np, '「イタリアブランドが15年選んだ品質を、日本のあなたのブランドに使えます。', bold=True, size=13, color=WHITE)
np2 = acp(cell, WD_ALIGN_PARAGRAPH.CENTER)
ar(np2, '300個〜、日本語対応、まずサンプルから。」', bold=True, size=13, color=WHITE)
np3 = acp(cell, WD_ALIGN_PARAGRAPH.CENTER); np3.paragraph_format.space_before = Pt(12)
ar(np3, 'Sanhe Interfashion Japan Team', size=9.5, color=BLUE_PALE)

# Save
output = 'Sanhe_LP完全設計書.docx'
doc.save(output)
print(f'✅ Saved: {output}')
