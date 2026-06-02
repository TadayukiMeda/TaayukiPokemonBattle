#!/usr/bin/env python3
"""Sanhe LP Strategy Phase 1-4 Pitch Document Generator"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph as DocxPara

# ── Colors ────────────────────────────────────────────────────────────────────
BLUE_DARK  = RGBColor(0x0f, 0x2d, 0x5e)
BLUE_MID   = RGBColor(0x1a, 0x56, 0xdb)
BLUE_LIGHT = RGBColor(0x3b, 0x82, 0xf6)
BLUE_PALE  = RGBColor(0x93, 0xc5, 0xfd)
BLUE_TINT  = RGBColor(0xdb, 0xea, 0xfe)
BLUE_BG    = RGBColor(0xef, 0xf6, 0xff)
WHITE      = RGBColor(0xff, 0xff, 0xff)
GRAY_DARK  = RGBColor(0x1e, 0x29, 0x3b)
GRAY_MID   = RGBColor(0x64, 0x74, 0x8b)
GRAY_LIGHT = RGBColor(0xf1, 0xf5, 0xf9)
GREEN_BG   = RGBColor(0xd1, 0xfa, 0xe5)
GREEN_FG   = RGBColor(0x06, 0x5f, 0x46)
AMBER_BG   = RGBColor(0xfe, 0xf3, 0xc7)
AMBER_FG   = RGBColor(0x78, 0x35, 0x0f)
RED_BG     = RGBColor(0xff, 0xe4, 0xe6)
RED_FG     = RGBColor(0x7f, 0x1d, 0x1d)
STRIPE1    = RGBColor(0xf8, 0xfa, 0xff)


def hx(rgb):
    return '{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])

# ── Low-level XML helpers ─────────────────────────────────────────────────────

def set_bg(cell, rgb):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hx(rgb))
    tcPr.append(shd)

def set_margins(cell, top=80, bottom=80, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'), str(val))
        m.set(qn('w:type'), 'dxa')
        tcMar.append(m)
    tcPr.append(tcMar)

def set_vAlign(cell, align='center'):
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), align)
    tcPr.append(vAlign)

def no_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'none')
        borders.append(b)
    tblPr.append(borders)

def add_cell_para(cell, align=WD_ALIGN_PARAGRAPH.LEFT):
    new_p = OxmlElement('w:p')
    cell._tc.append(new_p)
    para = DocxPara(new_p, cell._tc)
    para.alignment = align
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    return para

def first_para(cell, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    return p

def add_run(para, text, bold=False, size=10, color=None, italic=False):
    r = para.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.name = 'Meiryo'
    if r._element.rPr is None:
        r._element.get_or_add_rPr()
    r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Meiryo')
    if color:
        r.font.color.rgb = color
    return r

def page_break(doc):
    p = doc.add_paragraph()
    r = p.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r._r.append(br)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

def gap(doc, pt=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(pt)

# ── Document-level block builders ─────────────────────────────────────────────

def phase_banner(doc, num, title_jp, title_en):
    t = doc.add_table(rows=1, cols=1)
    t.style = 'Table Grid'
    no_table_borders(t)
    cell = t.rows[0].cells[0]
    cell.width = Inches(7.0)
    set_bg(cell, BLUE_DARK)
    set_margins(cell, 140, 140, 220, 220)
    p = first_para(cell)
    add_run(p, f'Phase {num}  ', bold=True, size=10, color=BLUE_PALE)
    add_run(p, title_jp, bold=True, size=15, color=WHITE)
    add_run(p, f'   {title_en}', size=9, color=BLUE_PALE)
    gap(doc, 8)

def section_bar(doc, label, title):
    t = doc.add_table(rows=1, cols=1)
    t.style = 'Table Grid'
    no_table_borders(t)
    cell = t.rows[0].cells[0]
    cell.width = Inches(7.0)
    set_bg(cell, BLUE_TINT)
    set_margins(cell, 80, 80, 160, 160)
    p = first_para(cell)
    add_run(p, f'{label}  ', bold=True, size=10, color=BLUE_MID)
    add_run(p, title, bold=True, size=11, color=BLUE_DARK)
    gap(doc, 6)

def body_p(doc, text, bold=False, size=10, color=None, left=0.0, after=3):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(after)
    if left:
        p.paragraph_format.left_indent = Inches(left)
    add_run(p, text, bold=bold, size=size, color=color or GRAY_DARK)
    return p

def bullet_p(doc, text, level=0, check=False, size=9.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    indent = 0.18 + level * 0.25
    p.paragraph_format.left_indent = Inches(indent)
    p.paragraph_format.first_line_indent = Inches(-0.18)
    marker = '✓ ' if check else '▶ '
    c = GREEN_FG if check else BLUE_MID
    add_run(p, marker, bold=True, size=size, color=c)
    add_run(p, text, size=size, color=GRAY_DARK)

def label_badge(doc, text, bg, fg):
    t = doc.add_table(rows=1, cols=1)
    t.style = 'Table Grid'
    no_table_borders(t)
    cell = t.rows[0].cells[0]
    set_bg(cell, bg)
    set_margins(cell, 60, 60, 160, 160)
    p = first_para(cell, WD_ALIGN_PARAGRAPH.CENTER)
    add_run(p, text, bold=True, size=9, color=fg)
    gap(doc, 4)

def info_box(doc, title, lines, bg=None, title_color=None):
    bg = bg or BLUE_BG
    title_color = title_color or BLUE_DARK
    t = doc.add_table(rows=1, cols=1)
    t.style = 'Table Grid'
    no_table_borders(t)
    cell = t.rows[0].cells[0]
    cell.width = Inches(7.0)
    set_bg(cell, bg)
    set_margins(cell, 120, 120, 200, 200)
    p = first_para(cell)
    add_run(p, title, bold=True, size=11, color=title_color)
    for line in lines:
        np = add_cell_para(cell)
        np.paragraph_format.left_indent = Inches(0.15)
        np.paragraph_format.space_before = Pt(3)
        add_run(np, '• ' + line, size=9.5, color=GRAY_DARK)
    gap(doc, 8)

# ── KV table (key → value rows) ───────────────────────────────────────────────

def kv_table(doc, rows, col_w=(2.2, 4.8)):
    t = doc.add_table(rows=len(rows), cols=2)
    t.style = 'Table Grid'
    for i, (k, v) in enumerate(rows):
        bg_k = BLUE_TINT if i % 2 == 0 else BLUE_BG
        bg_v = WHITE if i % 2 == 0 else STRIPE1
        kc = t.rows[i].cells[0]
        vc = t.rows[i].cells[1]
        kc.width = Inches(col_w[0])
        vc.width = Inches(col_w[1])
        set_bg(kc, bg_k); set_margins(kc, 70, 70, 130, 100)
        set_bg(vc, bg_v); set_margins(vc, 70, 70, 130, 130)
        set_vAlign(kc); set_vAlign(vc)
        add_run(first_para(kc), k, bold=True, size=9.5, color=BLUE_DARK)
        add_run(first_para(vc), v, size=9.5, color=GRAY_DARK)
    gap(doc, 8)

# ── Multi-column table ────────────────────────────────────────────────────────

def multi_table(doc, headers, rows, col_w=None, center_from=1):
    cols = len(headers)
    col_w = col_w or [7.0 / cols] * cols
    t = doc.add_table(rows=1 + len(rows), cols=cols)
    t.style = 'Table Grid'
    for j, h in enumerate(headers):
        hc = t.rows[0].cells[j]
        hc.width = Inches(col_w[j])
        set_bg(hc, BLUE_DARK); set_margins(hc, 70, 70, 110, 110)
        p = first_para(hc, WD_ALIGN_PARAGRAPH.CENTER)
        add_run(p, h, bold=True, size=9, color=WHITE)
    for i, row in enumerate(rows):
        bg = WHITE if i % 2 == 0 else STRIPE1
        for j, val in enumerate(row):
            dc = t.rows[i + 1].cells[j]
            dc.width = Inches(col_w[j])
            set_bg(dc, bg); set_margins(dc, 65, 65, 110, 110)
            align = WD_ALIGN_PARAGRAPH.CENTER if j >= center_from else WD_ALIGN_PARAGRAPH.LEFT
            p = first_para(dc, align)
            add_run(p, val, size=9, color=GRAY_DARK)
    gap(doc, 8)

# ── Pain → Solution two-column table ─────────────────────────────────────────

def pain_solution(doc, pairs):
    t = doc.add_table(rows=1 + len(pairs), cols=2)
    t.style = 'Table Grid'
    for j, (label, bg) in enumerate([
        ('バイヤーの不安・痛み', RED_FG),
        ('Sanheの回答・解決策', BLUE_DARK),
    ]):
        hc = t.rows[0].cells[j]
        set_bg(hc, bg); set_margins(hc, 80, 80, 140, 140)
        p = first_para(hc, WD_ALIGN_PARAGRAPH.CENTER)
        add_run(p, label, bold=True, size=9.5, color=WHITE)
    for i, (pain, sol) in enumerate(pairs):
        bg_p = RED_BG if i % 2 == 0 else RGBColor(0xff, 0xf8, 0xf8)
        bg_s = BLUE_BG if i % 2 == 0 else STRIPE1
        pc = t.rows[i+1].cells[0]; sc = t.rows[i+1].cells[1]
        set_bg(pc, bg_p); set_margins(pc, 75, 75, 130, 130)
        set_bg(sc, bg_s); set_margins(sc, 75, 75, 130, 130)
        add_run(first_para(pc), pain, size=9.5, color=GRAY_DARK)
        add_run(first_para(sc), sol, size=9.5, color=GRAY_DARK)
    gap(doc, 8)

# ── LP wireframe section list ─────────────────────────────────────────────────

def lp_wireframe(doc, sections):
    t = doc.add_table(rows=len(sections), cols=3)
    t.style = 'Table Grid'
    for i, (num, title, detail) in enumerate(sections):
        bg = BLUE_DARK if i == 0 else (BLUE_TINT if i % 2 == 0 else BLUE_BG)
        fg = WHITE if i == 0 else BLUE_DARK
        fg2 = BLUE_PALE if i == 0 else GRAY_DARK

        nc = t.rows[i].cells[0]; tc2 = t.rows[i].cells[1]; dc = t.rows[i].cells[2]
        nc.width = Inches(0.5); tc2.width = Inches(2.2); dc.width = Inches(4.3)
        for c in [nc, tc2, dc]:
            set_bg(c, bg); set_margins(c, 70, 70, 120, 120); set_vAlign(c)
        add_run(first_para(nc, WD_ALIGN_PARAGRAPH.CENTER), num, bold=True, size=9.5, color=fg)
        add_run(first_para(tc2), title, bold=True, size=9.5, color=fg)
        add_run(first_para(dc), detail, size=9, color=fg2)
    gap(doc, 8)

# ═════════════════════════════════════════════════════════════════════════════
# BUILD DOCUMENT
# ═════════════════════════════════════════════════════════════════════════════

doc = Document()
sec = doc.sections[0]
sec.page_width  = Inches(8.5)
sec.page_height = Inches(11)
sec.left_margin = sec.right_margin = Inches(0.75)
sec.top_margin  = sec.bottom_margin = Inches(0.75)

normal = doc.styles['Normal']
normal.font.name = 'Meiryo'
normal.paragraph_format.space_after = Pt(0)
normal.paragraph_format.space_before = Pt(0)

# ─────────────────────────────────────────────────────────────────────────────
# COVER PAGE
# ─────────────────────────────────────────────────────────────────────────────
t = doc.add_table(rows=1, cols=1)
t.style = 'Table Grid'; no_table_borders(t)
cc = t.rows[0].cells[0]; cc.width = Inches(7.0)
set_bg(cc, BLUE_DARK); set_margins(cc, 900, 900, 400, 400)

p = first_para(cc, WD_ALIGN_PARAGRAPH.CENTER)
add_run(p, 'SANHE INTERFASHION', bold=True, size=11, color=BLUE_PALE)

for txt, sz, clr, bd, sp in [
    ('LP 戦略実行計画', 22, WHITE,      True,  16),
    ('Phase 1 〜 Phase 4  完全ガイド',  11, BLUE_PALE, False, 8),
    ('',  6, WHITE, False, 6),
    ('自社棚卸し  ▶  ペルソナ設計  ▶  英語LP  ▶  日本語LP', 10, BLUE_PALE, False, 4),
    ('',  6, WHITE, False, 10),
    ('2025年版  会社ピッチ資料', 9, RGBColor(0x6b,0xa3,0xd6), False, 0),
]:
    np = add_cell_para(cc, WD_ALIGN_PARAGRAPH.CENTER)
    np.paragraph_format.space_before = Pt(sp)
    if txt:
        add_run(np, txt, bold=bd, size=sz, color=clr)

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 1 — 自社棚卸し
# ─────────────────────────────────────────────────────────────────────────────
phase_banner(doc, 1, '自社棚卸し', 'Company Self-Assessment')

body_p(doc,
    '「何を・誰に・どう伝えるか」を決める前に、Sanheが何者で何ができるかを言語化する。'
    'LP制作・商談・展示会出展、すべてのベースになる作業。',
    size=9.5, color=GRAY_MID, after=8)

# ── 1-1 強み・製品ライン ──────────────────────────────────────────────────────
section_bar(doc, '1-1', '強み・製品ラインの言語化')
body_p(doc, '■ Sanheのコア強み', bold=True, size=10, color=BLUE_DARK, after=2)
for item in [
    '韓国国内工場（ソウル近郊）での一貫生産 ── 品質管理を自社コントロール',
    '欧米ラグジュアリーブランド向けOEM実績10年以上（Want Les Essentiels・Roots等）',
    '小ロット対応（MOQ 50個〜） ── 日本のセレクトショップニーズに直結',
    '牛革・豚革・馬革・合成皮革など幅広い素材・仕上げ対応',
    'OEM（型紙から）・ODM（既存型のカスタマイズ）どちらも可',
    '全品検品体制 ── 欧米バイヤーに鍛えられた品質基準',
]:
    bullet_p(doc, item, check=True)

gap(doc, 8)
body_p(doc, '■ 製品ラインアップ', bold=True, size=10, color=BLUE_DARK, after=2)
multi_table(doc,
    ['カテゴリ', '主要品目', 'OEM', 'ODM', '最小発注数'],
    [
        ['ウォレット類',    '二つ折り財布・長財布・カードケース・コインケース', '○', '○', '50個〜'],
        ['バッグ類',        'トートバッグ・ショルダーバッグ・クラッチ・ポーチ',  '○', '○', '30個〜'],
        ['小物・ストラップ','ベルト・キーケース・名刺入れ・ストラップ',          '○', '△', '100個〜'],
        ['カスタムOEM',     '型紙・素材・金具すべてゼロから設計',               '○', '―', '要相談'],
    ],
    col_w=[1.8, 2.9, 0.55, 0.55, 1.2],
    center_from=2,
)

# ── 1-2 既存顧客実績 ─────────────────────────────────────────────────────────
section_bar(doc, '1-2', '既存顧客実績の整理（英語・日本語）')
body_p(doc, 'LPや商談で「どこが使っているか」を示すことが最大の信頼材料になる。', size=9.5, color=GRAY_MID, after=6)

kv_table(doc, [
    ('クライアント名',  'Want Les Essentiels（ウォント・レ・ゼッセンシエル）'),
    ('国・規模',        'カナダ発・北米展開。ラグジュアリーアクセサリーブランド'),
    ('担当製品',        'ウォレット・カードケース・旅行小物シリーズ'),
    ('英語表記（LP用）','OEM Manufacturer for Want Les Essentiels (Canada) — Wallets & Card Holders'),
    ('日本語表記（商談用）', 'カナダ発ラグジュアリーブランド「Want Les Essentiels」へ革小物を継続供給中'),
], col_w=(2.3, 4.7))

kv_table(doc, [
    ('クライアント名',  'Roots（ルーツ）'),
    ('国・規模',        'カナダ最大手ライフスタイルブランド。全世界120店舗以上'),
    ('担当製品',        '革小物シリーズの一部（詳細は守秘義務に準じて開示範囲を決める）'),
    ('英語表記（LP用）','Production Partner for Roots Canada — Leather Accessories'),
    ('日本語表記（商談用）', 'カナダ「Roots」の革小物を製造。北米品質基準をクリアした工場として実績あり'),
], col_w=(2.3, 4.7))

info_box(doc, '▶ 実績を使う際の注意点',
    ['NDA・守秘義務の確認: ロゴ・数量を公開してよいか先方に確認する',
     '「製造パートナー」表記のみ許可されることが多い → "Production Partner for Roots" のような表現に留める',
     '写真・製品画像の二次利用権: 使用許可があればLP掲載が最も効果的',
     '取引証明書（Invoice・PO）: 商談時に提示できるよう整理しておく'],
    bg=AMBER_BG, title_color=AMBER_FG)

# ── 1-3 BtoB必須スペック ─────────────────────────────────────────────────────
section_bar(doc, '1-3', 'BtoB必須スペック一覧')
body_p(doc, '日本のバイヤーが最初に確認する「基本仕様」を一枚で見せられる状態にしておく。', size=9.5, color=GRAY_MID, after=6)

kv_table(doc, [
    ('MOQ（最小発注数）',    '50個〜（品目・仕様により応相談）'),
    ('サンプルリードタイム', '2〜3週間（素材在庫あり） / 4〜5週間（素材調達込み）'),
    ('量産リードタイム',     '4〜6週間（発注確定・素材手配完了から）'),
    ('OEM対応範囲',          '型紙・素材・ハードウェア（金具）・箔押し・刻印 すべて対応'),
    ('ODM対応',              '既存型ベースのカスタマイズ可。型代不要で小ロット対応しやすい'),
    ('対応素材',             '牛革（フルグレイン/コレクテッド）・豚革・馬革・ヌバック・合成皮革'),
    ('対応仕上げ',           'スムース・型押し・ヌメ革・スエード・パンチング・染色各種'),
    ('品質基準',             '全品検品実施。欧米ブランド向け基準適用（サイズ誤差±1mm以内等）'),
    ('言語対応',             '韓国語・英語（日本語対応: 要確認・通訳手配可）'),
    ('支払い条件',           'T/T：前払30% ＋ 出荷前残70%（初回取引）'),
    ('サンプル費用',         '実費請求（量産確定時に一部返還の相談可）'),
    ('出荷港',               '仁川港（ICN）/ 釜山港（PUS）選択可'),
], col_w=(2.3, 4.7))

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 2 — 日本ペルソナ設定
# ─────────────────────────────────────────────────────────────────────────────
phase_banner(doc, 2, '日本ペルソナ設定', 'Japan Buyer Persona')

body_p(doc,
    '「全員に届けようとすると誰にも届かない」。ターゲットを一人に絞ることでLP・商談トークがシャープになる。',
    size=9.5, color=GRAY_MID, after=8)

# ── 2-1 ペルソナプロファイル ─────────────────────────────────────────────────
section_bar(doc, '2-1', 'ペルソナプロファイル（一人に絞る）')

kv_table(doc, [
    ('ペルソナ名（仮）',  '田中 美咲（Misaki Tanaka）28〜38歳'),
    ('職種・役割',        'セレクトショップ勤務 バイヤー  /  革小物ブランド 調達担当'),
    ('会社規模',          '従業員20〜150名。国内展開のライフスタイル・ファッション系企業'),
    ('決裁権',            '〜50万円/発注は自己判断。それ以上は上長（MD・経営者）承認が必要'),
    ('仕入れ経験',        '国内卸・中国工場とは取引経験あり。韓国工場は初めてか、数社経験あり'),
    ('情報収集',          '展示会（ライフスタイルWeek・JFW）・Instagram・業界紙・LinkedIn'),
    ('英語力',            '読める・書ける（TOEIC 600〜750程度）。口頭交渉には通訳が必要な場合も'),
    ('KPI・評価軸',       '仕入れ原価・製品クオリティ・納期遵守率・バイヤー接点数'),
], col_w=(2.2, 4.8))

info_box(doc, '▶ ペルソナの一日（仕入れ担当の日常）',
    ['朝: MDとの売上・在庫会議 → 不足品目の補充計画を立てる',
     '午前: 既存サプライヤーへの納期確認メール・QC確認',
     '午後: 新規工場の展示会ブースを回る or Instagram/LinkedInで海外工場をリサーチ',
     '夕方: サンプル依頼・見積もり依頼のメール送付',
     '判断基準: サンプルの品質と担当者の対応スピードで発注先を決める'],
    bg=GREEN_BG, title_color=GREEN_FG)

# ── 2-2 不安マップ ───────────────────────────────────────────────────────────
section_bar(doc, '2-2', '不安マップ ／ 決め手マップ')
body_p(doc, '韓国の知らない工場に初めて発注するとき、バイヤーが頭の中で考えていること。', size=9.5, color=GRAY_MID, after=6)

pain_solution(doc, [
    ('本当に届くの？納期を守れる工場かどうか分からない',
     'リードタイム明記＋過去納期実績（例: 欧米ブランド向け納期遵守率98%）を提示'),
    ('品質が毎ロット変わる海外工場に怖い経験がある',
     '全品検品体制・欧米基準適用・サンプルと量産品の差異ゼロ保証を訴求'),
    ('100個以上じゃないと受けてもらえないと思っている',
     'MOQ 50個〜を明示。「初回トライアル30個〜対応可」と商談で伝える'),
    ('英語でのやり取りが煩雑で工数がかかる',
     '日本語対応窓口（担当者）or 日本語対応可能なエージェント経由を提示'),
    ('実物を見ないと発注できない（サンプル入手が面倒）',
     'サンプル送付フロー・費用・期間を明記。展示会でサンプル直接確認できる旨を伝える'),
    ('韓国工場と取引した実績・事例がない、稟議を通せるか不安',
     'Want Les Essentiels・Roots等の北米実績を証拠として提示。「北米で通った品質」は稟議の説得材料になる'),
])

# ── 2-3 検索ワード・情報収集経路 ─────────────────────────────────────────────
section_bar(doc, '2-3', '検索ワード・情報収集経路')
body_p(doc, '■ バイヤーが使う日本語検索ワード（SEO・Google広告にも転用可）', bold=True, size=10, color=BLUE_DARK, after=3)

multi_table(doc,
    ['検索意図', 'キーワード例', '競合度（推測）', 'Sanheの刺さる訴求点'],
    [
        ['工場探し',      '韓国 革製品 OEM 工場',          '低〜中', '韓国製品質＋英語対応実績'],
        ['小ロット',      'レザー 小物 OEM 小ロット',       '中',     'MOQ 50個〜と明示'],
        ['具体品目',      '革財布 OEM 製造',                '中',     'ウォレット専門ラインを訴求'],
        ['比較・検討',    '韓国 中国 OEM 品質 違い',        '低',     '韓国製のQCの厳しさをコンテンツ化'],
        ['展示会系',      'ライフスタイルWeek 仕入れ 革製品','低',    '展示会出展で直接接点'],
        ['問い合わせ系',  '革小物 海外工場 見積もり 問い合わせ','低', 'CTAで即見積もり対応を打ち出す'],
    ],
    col_w=[1.4, 2.0, 1.2, 2.4],
    center_from=2,
)

body_p(doc, '■ バイヤーの情報収集チャネル（接触ポイント）', bold=True, size=10, color=BLUE_DARK, after=3)
for ch, detail in [
    ('展示会', 'ライフスタイルWeek（東京ビッグサイト）・JFW・Tokyo Leather Fair → 直接ブースでサンプル確認'),
    ('Instagram', '@工場アカウントで「素材感・縫製クオリティ・作業工程」を検索 → Reels・製造動画が効果的'),
    ('LinkedIn', '調達・購買担当者が使用。英語コンテンツを発信するだけでリーチできる'),
    ('業界紙',   'テキスタイル・アパレル系メディア（SENKEN・WWD Japan）への掲載で信頼度UP'),
    ('紹介・口コミ', '同業バイヤー間での情報共有。一社実績ができると紹介が生まれやすい'),
]:
    bullet_p(doc, f'【{ch}】 {detail}', check=True)

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 3 — LP構成設計（英語本社LP）
# ─────────────────────────────────────────────────────────────────────────────
phase_banner(doc, 3, 'LP構成設計（英語本社LP）', 'English Corporate LP Structure')

body_p(doc,
    '英語LPは「工場の実力証明書」。北米実績を最上部に置き、Webflow/Framerで1ページ完結のランディングページを作る。',
    size=9.5, color=GRAY_MID, after=8)

# ── 3-1 LP構成・ワイヤーフレーム ─────────────────────────────────────────────
section_bar(doc, '3-1', '英語本社LP 必要要素と構成')

lp_wireframe(doc, [
    ('①', 'HERO（最重要）',
     'キャッチコピー + サブヘッド + CTA。「Trusted by North American Brands」を必ず入れる'),
    ('②', 'TRUST BAR',
     'クライアントロゴ横並び（Want Les Essentiels・Roots）。Heroの直下に置く'),
    ('③', 'PRODUCTS',
     '製品カテゴリ写真4〜6枚。クオリティの高い実物写真が必須。照明・背景に注意'),
    ('④', 'CAPABILITIES',
     'OEM / ODM / MOQ / Lead Time をアイコン付き4ブロックで明示'),
    ('⑤', 'PROCESS',
     '「Inquiry → Sample → Production → Delivery」の4ステップ。安心感を与える'),
    ('⑥', 'QUALITY',
     '検品基準・使用素材・認証あれば掲載。工場内部写真があれば最も効果的'),
    ('⑦', 'ABOUT',
     '設立年・工場所在地・従業員数・代表コメント。顔が見える情報は信頼を上げる'),
    ('⑧', 'CTA（最下部）',
     '「Request a Sample」「Download Catalog」「Contact Us」の3択。メールフォーム必須'),
])

# ── 3-2 信頼軸設計 ───────────────────────────────────────────────────────────
section_bar(doc, '3-2', '「韓国製 ＋ 北米実績あり」信頼軸の設計')
body_p(doc, 'LPを開いた瞬間（5秒以内）に伝わらなければ、バイヤーは離脱する。', size=9.5, color=GRAY_MID, after=6)

info_box(doc, '✅ Heroセクションで必ず伝える3要素',
    ['① 何ができる工場か : "Korean Leather Goods OEM Manufacturer"',
     '② 誰の実績があるか : "Trusted by Want Les Essentiels & Roots (Canada)"',
     '③ 次の行動 : "Request a Sample Today → "（CTAボタン）'],
    bg=GREEN_BG, title_color=GREEN_FG)

info_box(doc, '❌ よくある失敗（やってはいけないこと）',
    ['Heroに「高品質」「誠実な対応」などの根拠のない言葉を並べる',
     'クライアントロゴがページ下部にある（スクロールしないと信頼が伝わらない）',
     '問い合わせフォームがなく「メールはこちら」のみ（離脱率が高い）',
     '製品写真がスマホ撮影・背景が工場の床（品質への不信感につながる）'],
    bg=RED_BG, title_color=RED_FG)

# ── 3-3 ツール比較 ───────────────────────────────────────────────────────────
section_bar(doc, '3-3', 'ツール比較 ／ 制作判断フロー')
body_p(doc, '英語LPは「コーディング不要ツール＋プロのデザイン」が最速・最コスパ。', size=9.5, color=GRAY_MID, after=6)

multi_table(doc,
    ['ツール', '制作難易度', '月額料金', '英語LP適性', '特徴・推奨理由'],
    [
        ['Webflow',      '中',    '$14〜/月',  '◎ 最適', 'デザイン自由度最高。フリーランサーへの依頼も容易'],
        ['Framer',       '易〜中', '$0〜/月',  '◎ 最適', 'AI生成機能あり。直感的操作で高品質LP制作可'],
        ['Squarespace',  '易',    '$16〜/月',  '○ 可',   'テンプレート豊富。写真重視ブランドに向く'],
        ['Notion + Super','易',   '$12〜/月',  '△ 代替', 'カタログ・仕様書として使う場合に有効'],
        ['WordPress',    '中〜難', '$5〜/月',  '○ 可',   'カスタム自由度高。初心者は制作コスト大'],
    ],
    col_w=[1.3, 1.1, 1.1, 1.1, 2.4],
    center_from=1,
)

body_p(doc, '■ 制作判断フロー', bold=True, size=10, color=BLUE_DARK, after=3)
for step, desc in [
    ('STEP 1 予算確認',  '¥0〜¥50,000で自作 ／ ¥150,000〜で外注 ／ ¥300,000〜でデザイン会社'),
    ('STEP 2 写真素材', 'プロ撮影の製品写真がなければ、まず写真撮影を先に投資する（最重要）'),
    ('STEP 3 ツール選定', '自作ならFramer（無料）→ Webflow。外注ならWebflow指定で依頼'),
    ('STEP 4 テキスト', '英語コピーはChatGPT + ネイティブチェック（Fiverr等で$50〜）で用意'),
    ('STEP 5 公開・検証', 'Googleアナリティクス設置 → 問い合わせ数を2週間計測して改善'),
]:
    bullet_p(doc, f'【{step}】 {desc}')

page_break(doc)

# ─────────────────────────────────────────────────────────────────────────────
# PHASE 4 — 日本語版LP制作
# ─────────────────────────────────────────────────────────────────────────────
phase_banner(doc, 4, '日本語版LP制作', 'Japanese Landing Page')

body_p(doc,
    '英語版の「直訳」は絶対にNG。日本のBtoBバイヤーの不安を順番に解消していく構成に作り直す。',
    size=9.5, color=GRAY_MID, after=8)

# ── 4-1 構成の組み直し方針 ────────────────────────────────────────────────────
section_bar(doc, '4-1', '日本語LP 構成の組み直し方針')
body_p(doc, 'PASONAの法則：「問題提起 → 共感 → 解決策 → 証拠 → 行動促進」の順に設計する。', size=9.5, color=GRAY_MID, after=6)

lp_wireframe(doc, [
    ('①', 'ファーストビュー（最重要）',
     '「なぜ韓国工場が選ばれているのか」を問いかけ形式で提示。5秒以内にターゲットに刺さるか'),
    ('②', '課題の共感',
     '「こんな悩みありませんか？」中国工場の品質ばらつき・納期遅延など、バイヤーの日常課題を列挙'),
    ('③', 'Sanheの解決策',
     '課題に対する具体的な回答。数字・実績で裏付ける（例: 全品検品・MOQ50個〜・リードタイム4週間）'),
    ('④', '信頼証拠（最重要）',
     'Want Les Essentiels・Rootsのロゴ・取引実績。日本語で「北米ブランドに選ばれた品質」と伝える'),
    ('⑤', '製品ラインアップ',
     '高品質写真4〜6枚。価格ではなく「素材感・縫製精度・デザイン性」を見せる'),
    ('⑥', '製造プロセス',
     '相談〜サンプル〜量産〜納品の流れ。日本語で丁寧に。「初めてでも安心」メッセージ'),
    ('⑦', 'よくある質問（FAQ）',
     '「最低発注数は？」「サンプルはもらえる？」「日本語で話せる？」の3つは必須'),
    ('⑧', 'CTA（行動促進）',
     '「まずサンプルを見てみる」「資料・スペックシートをDL」など低ハードルの入口を複数用意'),
])

# ── 4-2 強調ポイント ─────────────────────────────────────────────────────────
section_bar(doc, '4-2', '日本バイヤー向け 強調すべき4ポイント')

multi_table(doc,
    ['強調ポイント', '伝えるべきメッセージ', 'LPでの見せ方'],
    [
        ['品質保証',
         'ロット間のブレをゼロに近づける全品検品。欧米基準で鍛えられた精度',
         '検品工程の写真・「ロット間誤差±1mm以内」など数値で証明'],
        ['納期の明確さ',
         'サンプル2〜3週間、量産4〜6週間を明記。遅延時の連絡ルールも伝える',
         '工程スケジュール表・「納期遵守率○%」の実績数字'],
        ['サンプル対応',
         '実物確認なしに発注できない日本バイヤーへの配慮。サンプル送付フロー明記',
         '「まずサンプルから」CTAを目立つ場所に複数配置'],
        ['日本語窓口',
         '言語バリアが最大の不安要因。日本語対応の有無・エージェント名を明記',
         '担当者の顔写真・名前・連絡先（日本語対応可）を掲載'],
    ],
    col_w=[1.5, 2.8, 2.7],
    center_from=99,
)

# ── 4-3 信頼構築ビジュアル ────────────────────────────────────────────────────
section_bar(doc, '4-3', '信頼構築ビジュアル要素')
body_p(doc, '日本のBtoBバイヤーは「見えないもの」を信頼しない。テキストより写真・数字・証拠を優先する。', size=9.5, color=GRAY_MID, after=6)

multi_table(doc,
    ['ビジュアル素材', '重要度', '準備方法・コスト感'],
    [
        ['製品実物写真（白背景・高解像度）', '★★★ 必須',
         'プロカメラマン依頼推奨。韓国でのコスト：¥30,000〜¥80,000/半日'],
        ['工場内部・製造工程写真',            '★★★ 必須',
         '自社スマホ撮影可。明るい場所・清潔感が伝わる構図で20〜30枚以上準備'],
        ['クライアントロゴ（許可取得済み）',   '★★★ 必須',
         '先方の許可を書面で取得。ロゴファイル（SVG/PNG透過）をもらう'],
        ['素材証明書・品質検査レポート',       '★★ 推奨',
         '自社発行でも可。「使用素材: ○○産フルグレイン牛革」等の明記が効果的'],
        ['担当者・スタッフの顔写真',           '★★ 推奨',
         '日本語対応窓口の顔が見えると問い合わせ率が上がる。自撮りでも可'],
        ['受賞・認証・メディア掲載',           '★ あれば',
         '業界誌・展示会賞・ISO等。信頼スコアを底上げする'],
    ],
    col_w=[2.5, 1.0, 3.5],
    center_from=1,
)

# ── 4-4 CTA設計 ─────────────────────────────────────────────────────────────
section_bar(doc, '4-4', 'CTA設計（行動促進の設計）')
body_p(doc, 'BtoBは意思決定に時間がかかる。いきなり「発注」を求めず、段階的な接点設計が重要。', size=9.5, color=GRAY_MID, after=6)

multi_table(doc,
    ['CTAの種類', '想定ユーザー', '入手するもの', '優先度'],
    [
        ['サンプルを請求する',          '具体的に検討中のバイヤー',   'バイヤー情報＋商談機会',   '★★★'],
        ['スペックシートをDLする',       '比較・検討中のバイヤー',     'メールアドレス（リード）',  '★★★'],
        ['無料相談・オンライン商談',     '韓国工場初経験の担当者',     '信頼関係の構築',           '★★'],
        ['展示会で会う（ブース番号明記）','展示会参加予定のバイヤー',  '対面接点・サンプル確認',   '★★'],
        ['カタログをメールで受け取る',   '情報収集フェーズのバイヤー', 'メールアドレス（リード）',  '★'],
    ],
    col_w=[2.0, 2.0, 1.9, 1.1],
    center_from=3,
)

info_box(doc, '▶ フォーム設計の鉄則（日本BtoB）',
    ['入力項目は5項目以内に絞る：会社名・名前・メール・電話・問い合わせ内容',
     '「必須」を明確にし、任意項目は最小化する（バイヤーは時間がない）',
     '送信後の自動返信メールに「3営業日以内に返信します」と明記する',
     'プライバシーポリシーへのリンクを必ず設置（日本企業は確認する）',
     'スマホ最適化を必ず確認：日本のバイヤーの50%以上がスマホでアクセス'],
    bg=GREEN_BG, title_color=GREEN_FG)

gap(doc, 10)

# ── 全体まとめ ────────────────────────────────────────────────────────────────
t = doc.add_table(rows=1, cols=1)
t.style = 'Table Grid'; no_table_borders(t)
cell = t.rows[0].cells[0]; cell.width = Inches(7.0)
set_bg(cell, BLUE_DARK); set_margins(cell, 160, 160, 240, 240)

p = first_para(cell, WD_ALIGN_PARAGRAPH.CENTER)
add_run(p, 'Phase 1〜4  実行チェックサマリー', bold=True, size=12, color=WHITE)

for txt, done in [
    ('Phase 1: 強み・実績・スペックの言語化 → 英語・日本語の両バージョン作成', True),
    ('Phase 2: ペルソナ一人に絞る → 不安マップ・決め手マップ完成', True),
    ('Phase 3: 英語LP構成確定 → ツール選定・写真素材準備', False),
    ('Phase 4: 日本語LP構成 → サンプルCTA設計・フォーム設置', False),
]:
    np = add_cell_para(cell)
    np.paragraph_format.left_indent = Inches(0.2)
    np.paragraph_format.space_before = Pt(6)
    marker = '✅ ' if done else '□  '
    col = RGBColor(0x6b, 0xe0, 0xb0) if done else BLUE_PALE
    add_run(np, marker + txt, size=9.5, color=col)

# ── 出力 ─────────────────────────────────────────────────────────────────────
output = 'Sanhe_LP戦略実行計画_Phase1-4.docx'
doc.save(output)
print(f'✅ Saved: {output}')
