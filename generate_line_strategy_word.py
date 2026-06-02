from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ── Page margins ──
for section in doc.sections:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.8)

# ── Styles helper ──
def set_font(run, size, bold=False, color=None, italic=False):
    run.font.name = 'Hiragino Kaku Gothic Pro'
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20 if level == 1 else 14)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    if level == 1:
        set_font(run, 16, bold=True, color=(30, 30, 30))
        # gold left border via shading — use bottom border trick
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        left = OxmlElement('w:left')
        left.set(qn('w:val'), 'single')
        left.set(qn('w:sz'), '24')
        left.set(qn('w:space'), '8')
        left.set(qn('w:color'), 'B5872C')
        pBdr.append(left)
        pPr.append(pBdr)
    elif level == 2:
        set_font(run, 13, bold=True, color=(40, 40, 40))
    elif level == 3:
        set_font(run, 11, bold=True, color=(80, 70, 60))
    return p

def add_body(doc, text, indent=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    if indent:
        p.paragraph_format.left_indent = Cm(0.8)
    run = p.add_run(text)
    set_font(run, 10.5, color=(50, 48, 44))
    return p

def add_bullet(doc, text, highlight=False):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    run_dot = p.add_run('• ')
    set_font(run_dot, 10.5, color=(181, 135, 44))
    run = p.add_run(text)
    if highlight:
        set_font(run, 10.5, bold=True, color=(181, 135, 44))
    else:
        set_font(run, 10.5, color=(50, 48, 44))
    return p

def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def cell_text(cell, text, size=10, bold=False, color=(50, 48, 44), align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.font.name = 'Hiragino Kaku Gothic Pro'
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor(*color)

def add_spacer(doc, size=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run('')
    run.font.size = Pt(size)

# ══════════════════════════════════════
# TITLE PAGE
# ══════════════════════════════════════
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_before = Pt(8)
p.paragraph_format.space_after = Pt(4)
run = p.add_run('SANHE INTERFASHION')
set_font(run, 9, color=(181, 135, 44))
run.font.name = 'Helvetica Neue'

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(6)
run = p.add_run('日本市場向けLP・資料・LINE戦略ガイド')
set_font(run, 20, bold=True, color=(20, 20, 18))

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
run = p.add_run('LPと資料とLINEをどう設計するか')
set_font(run, 12, color=(110, 104, 96))

add_spacer(doc, 12)

# Divider
p = doc.add_paragraph()
run = p.add_run('─' * 48)
set_font(run, 9, color=(200, 190, 178))

add_spacer(doc, 8)

# ══════════════════════════════════════
# SECTION 1: 基本構造
# ══════════════════════════════════════
add_heading(doc, '1｜なぜ「資料」が起点になるのか', 1)
add_body(doc, 'LPもLINEも、結局は資料（PDF）を軸に設計される。資料の中身が決まれば、LPで何を訴求すべきか、LINEで何を話すかが自然に決まる。逆に言えば、渡せる資料がない状態でLP公開しても機能しない。')

add_spacer(doc)

# フロー図
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(6)
p.paragraph_format.left_indent = Cm(0.5)
run = p.add_run('LP（集客）  →  資料DL（リード取得）  →  LINE相談（商談化）')
set_font(run, 11, bold=True, color=(181, 135, 44))

add_spacer(doc)

# ══════════════════════════════════════
# SECTION 2: 考える順番
# ══════════════════════════════════════
add_heading(doc, '2｜考える順番', 1)

# ① 資料の中身
add_heading(doc, '① 資料の中身から逆算する（最初に決めること）', 2)
add_body(doc, '「何を読んだらLINEに連絡したくなるか？」を先に決める。')

add_spacer(doc, 4)
add_heading(doc, '渡せる資料の候補', 3)
add_bullet(doc, '会社概要PDF（sanhe_pitch.html ベース）')
add_bullet(doc, '製品・価格カタログ（FOB価格帯・MOQ・リードタイム）')
add_bullet(doc, '品質管理プロセス詳細資料')
add_bullet(doc, '→ 一番刺さるのは「サンプル依頼の手引き ＋ 価格帯シート」の組み合わせ', highlight=True)

add_spacer(doc, 6)

# ② ダウンロード要件
add_heading(doc, '② ダウンロードに何を要求するか', 2)
add_body(doc, 'B2Bならフォームあり一択。「誰がDLしたか」がわかることでLINEでの初回接触が変わる。')

add_spacer(doc, 4)

tbl = doc.add_table(rows=3, cols=3)
tbl.style = 'Table Grid'
tbl.allow_autofit = True

# Header
header_cells = tbl.rows[0].cells
shade_cell(header_cells[0], '1E1C18')
shade_cell(header_cells[1], '1E1C18')
shade_cell(header_cells[2], '1E1C18')
cell_text(header_cells[0], 'パターン', 10, bold=True, color=(181, 135, 44))
cell_text(header_cells[1], 'メリット', 10, bold=True, color=(181, 135, 44))
cell_text(header_cells[2], 'デメリット', 10, bold=True, color=(181, 135, 44))

# Row 1
r1 = tbl.rows[1].cells
shade_cell(r1[0], 'F5F3EF')
cell_text(r1[0], 'フォームなし（直接DL）', 10)
cell_text(r1[1], 'ダウンロード数多い', 10)
cell_text(r1[2], '誰がDLしたか不明', 10)

# Row 2
r2 = tbl.rows[2].cells
shade_cell(r2[0], 'FDF9F2')
shade_cell(r2[1], 'FDF9F2')
shade_cell(r2[2], 'FDF9F2')
cell_text(r2[0], 'フォームあり（会社名・メール必須）★推奨', 10, bold=True, color=(100, 75, 20))
cell_text(r2[1], 'リードが取れる', 10, bold=True, color=(100, 75, 20))
cell_text(r2[2], 'DL数は減る', 10)

add_spacer(doc, 6)

# ③ LINE運用体制
add_heading(doc, '③ LINE運用体制（一番コケやすい）', 2)
add_bullet(doc, '誰が対応するか（日本語で答えられる人材確認）')
add_bullet(doc, '何時間以内に返信するか（24時間以内を守れるか）')
add_bullet(doc, 'LINE公式アカウント vs 個人LINE　→ 公式アカウント推奨（無料プランあり）')

add_spacer(doc, 8)

# ══════════════════════════════════════
# SECTION 3: 個人LINE vs 公式
# ══════════════════════════════════════
add_heading(doc, '3｜なぜ個人LINEではダメか', 1)

add_heading(doc, '信頼性の問題', 2)
add_body(doc, '相手は企業の調達担当者や購買部門。個人LINEで連絡が来たら「これ、本当にSanheの公式窓口か？」と疑う。稟議を通す際に上長に見せる連絡履歴が「個人のLINEアカウント」だと、それだけで却下されるケースがある。')

add_spacer(doc, 6)
add_heading(doc, '運用上の問題', 2)

tbl2 = doc.add_table(rows=5, cols=3)
tbl2.style = 'Table Grid'

header2 = tbl2.rows[0].cells
shade_cell(header2[0], '1E1C18')
shade_cell(header2[1], '1E1C18')
shade_cell(header2[2], '1E1C18')
cell_text(header2[0], '問題', 10, bold=True, color=(181, 135, 44))
cell_text(header2[1], '個人LINE', 10, bold=True, color=(181, 135, 44))
cell_text(header2[2], 'LINE公式アカウント', 10, bold=True, color=(181, 135, 44))

rows2 = [
    ('担当者が辞めたら', '全会話履歴が消える', 'アカウントは会社資産として残る'),
    ('複数人で対応', '不可能', '複数スタッフで共有対応できる'),
    ('対応時間外', '既読無視に見える', '自動返信メッセージが設定できる'),
    ('ブランド表示', '個人名が出る', '「Sanhe Interfashion」と表示される'),
]

for i, (a, b, c) in enumerate(rows2):
    row = tbl2.rows[i + 1].cells
    bg = 'F5F3EF' if i % 2 == 0 else 'FFFFFF'
    shade_cell(row[0], bg)
    cell_text(row[0], a, 10, bold=True, color=(60, 55, 48))
    cell_text(row[1], b, 10, color=(160, 50, 50))
    cell_text(row[2], c, 10, color=(50, 110, 60))

add_spacer(doc, 6)

add_heading(doc, '一番シンプルな理由', 2)
add_body(doc, '個人LINEだと「担当者個人との関係」になってしまう。担当者が変わったとき、関係がゼロリセットされる。')
add_spacer(doc, 2)
add_body(doc, '公式アカウントなら「Sanheという会社との関係」になる。これはB2Bの長期取引では致命的な差だ。')

add_spacer(doc, 4)
add_heading(doc, '無料プランで十分な理由', 2)
add_body(doc, '月1,000通まで無料で送れる。LP経由のリードが月に何十件も来るフェーズになったら有料プランを検討すればいい。今は無料で始めて、運用に慣れることが先だ。')

add_spacer(doc, 8)

# ══════════════════════════════════════
# SECTION 4: 着手順序
# ══════════════════════════════════════
add_heading(doc, '4｜まず着手すること（優先順）', 1)

steps = [
    ('Step 1', '資料（PDF）を1枚作る', '会社概要 + 価格帯シート + サンプル依頼の手引き。すべての起点になる。'),
    ('Step 2', 'LP上のダウンロードフォームを設計する', '会社名・名前・メールアドレス必須。電話番号・質問は任意。送信後にPDFのDLリンクを表示。'),
    ('Step 3', 'LINE公式アカウントを開設する', '無料プランで開始。自動返信メッセージを設定（例：「24時間以内にご返信します」）。'),
    ('Step 4', 'LP → フォーム → LINE の導線を繋げる', 'LP上のCTAから資料DLフォームへ。DL後にLINEのQRコードを表示する。'),
]

for num, title, desc in steps:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    run_num = p.add_run(f'{num}  ')
    set_font(run_num, 10, bold=True, color=(181, 135, 44))
    run_title = p.add_run(title)
    set_font(run_title, 11, bold=True, color=(20, 20, 18))
    add_body(doc, desc, indent=True)

add_spacer(doc, 10)

# Footer line
p = doc.add_paragraph()
run = p.add_run('─' * 48)
set_font(run, 9, color=(200, 190, 178))

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(6)
run = p.add_run('Sanhe Interfashion Japan Strategy｜社内資料')
set_font(run, 8, color=(160, 150, 140))

# Save
path = '/home/user/TaayukiPokemonBattle/Sanhe_LINE戦略設計ガイド.docx'
doc.save(path)
print(f'Saved: {path}')
