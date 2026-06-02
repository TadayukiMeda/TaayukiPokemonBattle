from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ページ設定
section = doc.sections[0]
section.page_width  = Cm(21)
section.page_height = Cm(29.7)
section.left_margin   = Cm(2)
section.right_margin  = Cm(2)
section.top_margin    = Cm(2)
section.bottom_margin = Cm(2)

NAVY   = RGBColor(0x1a, 0x2a, 0x4a)
ACCENT = RGBColor(0xc8, 0xa9, 0x6e)
BLUE   = RGBColor(0x12, 0x55, 0xa0)
MUTED  = RGBColor(0x66, 0x66, 0x66)
WHITE  = RGBColor(0xff, 0xff, 0xff)
GREEN  = RGBColor(0x1a, 0x6e, 0x30)
RED    = RGBColor(0xb5, 0x2a, 0x2a)
AMBER  = RGBColor(0x7a, 0x52, 0x00)

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top','left','bottom','right'):
        tag = OxmlElement(f'w:{edge}')
        tag.set(qn('w:val'), kwargs.get(edge, 'none'))
        tag.set(qn('w:sz'), '4')
        tag.set(qn('w:space'), '0')
        tag.set(qn('w:color'), kwargs.get(f'{edge}_color', 'auto'))
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def para_fmt(para, size=11, bold=False, color=None, align=None, space_before=0, space_after=6):
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after  = Pt(space_after)
    if align: para.alignment = align
    for run in para.runs:
        run.font.size  = Pt(size)
        run.font.bold  = bold
        if color: run.font.color.rgb = color

def add_heading(text, level=1, color=NAVY, size=16):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(6)
    run = p.add_run(text)
    run.font.size  = Pt(size)
    run.font.bold  = True
    run.font.color.rgb = color
    return p

def add_body(text, size=10.5, color=None, bold=False, indent=0, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(space_after)
    if indent: p.paragraph_format.left_indent = Cm(indent)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color: run.font.color.rgb = color
    return p

def add_separator():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'DDDDDD')
    pBdr.append(bottom)
    pPr.append(pBdr)

# ============================================================
# 表紙ヘッダー
# ============================================================
tbl = doc.add_table(rows=1, cols=1)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
cell = tbl.cell(0, 0)
set_cell_bg(cell, '1a2a4a')
p = cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_before = Pt(8)
r = p.add_run('EXHIBITION STRATEGY — UPDATED')
r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = ACCENT
p2 = cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
r2 = p2.add_run('展示会アプローチ戦略（完全版）')
r2.font.size = Pt(20); r2.font.bold = True; r2.font.color.rgb = WHITE
p3 = cell.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
r3 = p3.add_run('革製品ブランドのオーナー社長と対面できる展示会リスト ｜ 作成日：2026年5月26日')
r3.font.size = Pt(9); r3.font.color.rgb = RGBColor(0xaa, 0xbb, 0xcc)
p4 = cell.add_paragraph()
r4 = p4.add_run('方針転換：メールアプローチ → 展示会対面アプローチに軸を移す。ターゲット企業が出展している場で直接提案する。ペルソナの意思決定トリガー「展示会での直接対話」を正面から活用する。')
r4.font.size = Pt(9); r4.font.color.rgb = RGBColor(0xcc, 0xdd, 0xee)
p4.paragraph_format.space_after = Pt(10)
p4.paragraph_format.left_indent = Cm(0)

doc.add_paragraph()

# ============================================================
# サマリー数値
# ============================================================
tbl2 = doc.add_table(rows=1, cols=4)
tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
labels = [('8','対象展示会（確認済み）'),('4週','最短接触まで（6/24〜）'),('30+','確認済みターゲット企業'),('2','バッグ専門B2B展（最重要）')]
for i,(val,lbl) in enumerate(labels):
    c = tbl2.cell(0, i)
    set_cell_bg(c, 'ffffff')
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(val + '\n')
    r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = NAVY
    r2 = p.add_run(lbl)
    r2.font.size = Pt(8); r2.font.color.rgb = MUTED

doc.add_paragraph()

# ============================================================
# 重要発見ボックス
# ============================================================
tbl3 = doc.add_table(rows=1, cols=1)
c = tbl3.cell(0,0)
set_cell_bg(c, 'fff3cd')
p = c.paragraphs[0]
r = p.add_run('【重要発見】バッグブランド専門のB2B展示会が2つある')
r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = NAVY
p2 = c.add_paragraph()
r2 = p2.add_run('「ハンドバッグかばん見本市（TBL）」と「K・N・O・T collection」は、革バッグ・財布ブランドだけが出展するBtoB商談展。来場者もバッグ業界関係者のみ。FaWやギフトショーより')
r2.font.size = Pt(10)
r3 = p2.add_run('はるかにターゲット精度が高い。')
r3.font.size = Pt(10); r3.font.bold = True
r4 = p2.add_run('\n今すぐ主催者に問い合わせて次回日程を確認することが最優先。')
r4.font.size = Pt(10)
p2.paragraph_format.space_after = Pt(8)

doc.add_paragraph()

# ============================================================
# 展示会カード共通ヘッダー作成関数
# ============================================================
def add_show_card(priority_label, date_str, weeks_str, show_name, show_sub, meta_items,
                  header_color, actions, companies, site_url, company_note=''):

    add_separator()

    # ヘッダー行
    tbl = doc.add_table(rows=1, cols=2)
    tbl.columns[0].width = Cm(4.5)
    tbl.columns[1].width = Cm(12.5)

    # 左：日付ブロック
    lc = tbl.cell(0,0)
    set_cell_bg(lc, header_color)
    p = lc.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(priority_label + '\n')
    r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = ACCENT
    r2 = p.add_run(date_str + '\n')
    r2.font.size = Pt(13); r2.font.bold = True; r2.font.color.rgb = WHITE
    r3 = p.add_run(weeks_str)
    r3.font.size = Pt(8); r3.font.color.rgb = RGBColor(0xaa, 0xbb, 0xcc)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)

    # 右：情報ブロック
    rc = tbl.cell(0,1)
    set_cell_bg(rc, 'ffffff')
    p2 = rc.paragraphs[0]
    r4 = p2.add_run(show_name)
    r4.font.size = Pt(13); r4.font.bold = True; r4.font.color.rgb = NAVY
    p2.paragraph_format.space_before = Pt(6)
    p3 = rc.add_paragraph()
    r5 = p3.add_run(show_sub)
    r5.font.size = Pt(8.5); r5.font.color.rgb = MUTED
    p4 = rc.add_paragraph()
    r6 = p4.add_run('　'.join([f'[{m}]' for m in meta_items]))
    r6.font.size = Pt(8.5); r6.font.color.rgb = BLUE
    p4.paragraph_format.space_after = Pt(6)

    # アクション + 企業の2段テーブル
    tbl2 = doc.add_table(rows=1, cols=2)
    tbl2.columns[0].width = Cm(8.5)
    tbl2.columns[1].width = Cm(8.5)

    ac = tbl2.cell(0,0)
    ec = tbl2.cell(0,1)
    set_cell_bg(ac, 'f8f9fa')
    set_cell_bg(ec, 'f8f9fa')

    ap = ac.paragraphs[0]
    r7 = ap.add_run('▼ アクション手順')
    r7.font.size = Pt(8); r7.font.bold = True; r7.font.color.rgb = MUTED
    ap.paragraph_format.space_before = Pt(4)
    for i, act in enumerate(actions, 1):
        p = ac.add_paragraph()
        r = p.add_run(f'{i}. {act}')
        r.font.size = Pt(9)
        p.paragraph_format.left_indent = Cm(0.3)
        p.paragraph_format.space_after = Pt(3)

    ep = ec.paragraphs[0]
    r8 = ep.add_run('▼ 確認済み出展企業（過去回）')
    r8.font.size = Pt(8); r8.font.bold = True; r8.font.color.rgb = MUTED
    ep.paragraph_format.space_before = Pt(4)
    for name, note in companies:
        p = ec.add_paragraph()
        r = p.add_run(f'・{name}')
        r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = NAVY
        r2 = p.add_run(f'　{note}')
        r2.font.size = Pt(8); r2.font.color.rgb = MUTED
        p.paragraph_format.left_indent = Cm(0.3)
        p.paragraph_format.space_after = Pt(2)
    if company_note:
        p = ec.add_paragraph()
        r = p.add_run(company_note)
        r.font.size = Pt(8); r.font.color.rgb = MUTED
        r.font.italic = True
        p.paragraph_format.space_after = Pt(4)

    su = doc.add_paragraph()
    r = su.add_run(f'公式サイト：{site_url}')
    r.font.size = Pt(8.5); r.font.color.rgb = BLUE
    su.paragraph_format.space_before = Pt(2)
    su.paragraph_format.space_after  = Pt(10)


# ============================================================
# SECTION 1
# ============================================================
add_heading('最優先：今すぐ動ける展示会（2026年6月〜）', size=13)

add_show_card(
    priority_label='★★★ 最優先',
    date_str='6月24〜26日 / 今から約4週間後',
    weeks_str='2026年 ｜ 東京ビッグサイト',
    show_name='ライフスタイルWeek【夏】 国際ファッション雑貨EXPO',
    show_sub='主催：RX Japan ｜ 東京ビッグサイト ｜ BtoB商談展',
    meta_items=['1,250社出展', '来場38,000名', 'バッグ・革小物カテゴリあり', '来場無料登録制（6/19締切）'],
    header_color='1a2a4a',
    actions=[
        '今すぐ lifestyle-expo.jp/summer/ で来場者登録（無料・6/19締切・以降5,000円）',
        '登録後に出展社検索で「バッグ・革小物」カテゴリを絞り込み、ターゲット企業リストを作成',
        '可能なら出展企業にLinkedIn/メールで事前コンタクト（「6/24にお会いしたい」）',
        '6月24〜26日に来場・対面提案。革サンプル＋信頼データシートを持参',
        '展示会後48時間以内に名刺交換した全社へ個別フォローアップメール送信',
    ],
    companies=[
        ('バッグ・革小物', 'カテゴリ確認済み'),
        ('アクセサリー・帽子・手袋', 'カテゴリ確認済み'),
        ('サステナブルファッション', 'カテゴリ確認済み'),
        ('海外ブランド（輸入）', 'カテゴリ確認済み'),
    ],
    site_url='https://www.lifestyle-expo.jp/summer/ja-jp.html',
    company_note='※個別企業名は来場者登録後に出展社検索で確認できる'
)

# ============================================================
# SECTION 2
# ============================================================
add_heading('高優先：バッグブランド専門B2B展（最重要カテゴリ）', size=13)

add_show_card(
    priority_label='★★★ バッグ専門B2B',
    date_str='2026年 11月（予定）',
    weeks_str='今から約24〜26週間後',
    show_name='ハンドバッグかばん見本市（TOKYO BAG AND LUGGAGE FAIR / TBL）',
    show_sub='主催：一般社団法人日本バッグ協会 ｜ 浅草橋ヒューリックホール等 ｜ 業界特化BtoB',
    meta_items=['革バッグ・財布ブランドのみ出展', '年2回（5月・11月）', '次回：11月'],
    header_color='1255a0',
    actions=[
        '今すぐ tbl.gr.jp（日本バッグ協会）に問い合わせて11月回の日程・来場登録方法を確認',
        '出展者も来場者もバッグ・革製品業界関係者のみ → ターゲット精度が全展示会で最高',
        'Dakota・genten・UNISON DEPT等の本命ターゲットと直接接触できる唯一の場',
    ],
    companies=[
        ('Dakota（ダコタ）', '女性向け革バッグ・財布'),
        ('genten（ゲンテン）', 'タンニン革バッグ・財布'),
        ('UNISON DEPT.', '革バッグ・日本橋馬喰町'),
        ('BAGGY PORT', '革・帆布バッグ（神戸）'),
        ('master-piece（MSPC）', '高機能バッグ'),
        ('Kiefer neu', '革バッグ・財布（名古屋）'),
        ('AS2OV', '機能性バッグ'),
        ('F.CLIO / MATAGI / Pid', '各バッグブランド'),
    ],
    site_url='https://www.tbl.gr.jp/mihonichi/'
)

add_show_card(
    priority_label='★★★ バッグ専門B2B',
    date_str='2026年 10月（予定）',
    weeks_str='今から約19〜21週間後',
    show_name='K・N・O・T collection（ケイ・エヌ・オー・ティー コレクション）',
    show_sub='主催：日本バッグ工業連合会 ｜ 浅草橋ヒューリックホール＆カンファレンス（東京）',
    meta_items=['バッグ・財布ブランド専門', '年2回（4月・10月）', '浅草橋開催（ターゲットエリア）'],
    header_color='1255a0',
    actions=[
        '日本バッグ工業連合会に問い合わせて10月回の日程・来場登録方法を確認',
        '浅草橋ヒューリックホールでの開催 ＝ ペルソナ企業の集積エリアに直接行ける',
        '4月回（春）はすでに終了。10月回（秋）が次のターゲット',
    ],
    companies=[
        ('株式会社ナダヤ（mu）', '革小物・バッグ（大阪発・浅草橋SR）'),
        ('大阪バッグ共同組合加盟各社', '革バッグ・財布メーカー多数'),
    ],
    site_url='日本バッグ工業連合会（公式サイト）',
    company_note='※詳細出展企業リストは日本バッグ工業連合会に要問い合わせ'
)

# ============================================================
# SECTION 3
# ============================================================
add_heading('中〜高優先：ブランドに出会える大型展（2026年9〜10月）', size=13)

add_show_card(
    priority_label='優先度：高',
    date_str='9月2〜4日 / 今から約14週間後',
    weeks_str='2026年 ｜ 東京ビッグサイト',
    show_name='東京インターナショナルギフトショー【秋】（第102回）',
    show_sub='主催：ビジネスガイド社 ｜ 東京ビッグサイト',
    meta_items=['2,000〜2,500社出展', '来場130,000名', 'バッグ・革製品ゾーンあり'],
    header_color='2c4a2c',
    actions=[
        'giftshow.co.jp で来場者事前登録（7月以降）',
        '6月展での成果を踏まえ、アプローチトークを改善して臨む',
        '「ライフスタイルゾーン」「ファッション・アパレル雑貨ゾーン」を重点巡回',
    ],
    companies=[
        ('Kanmi.（カンミ）', '浅草発の革小物・バッグ'),
        ('株式会社服部（HATTORI）', '豊岡鞄認定・革バッグ・財布'),
        ('株式会社ナダヤ', '革小物・バッグメーカー'),
    ],
    site_url='https://www.giftshow.co.jp/tigs/'
)

add_show_card(
    priority_label='優先度：高',
    date_str='9〜10月（例年） / 今から約16〜20週間後',
    weeks_str='2026年 ｜ 渋谷ストリームホール等',
    show_name='ジャパンレザーアワード（JAPAN LEATHER AWARD）',
    show_sub='主催：日本皮革産業連合会（JLIA）｜ 渋谷周辺（毎年）｜ 入場無料（一般公開）',
    meta_items=['応募作品281点展示', '革ブランドの作り手が集まる', '入場無料'],
    header_color='2c4a2c',
    actions=[
        'award.jlia.or.jp で2026年の開催情報を定期確認する',
        '受賞者・参加者は「品質と職人性にこだわる」ブランドオーナー ＝ ペルソナに完全一致',
        '来場者も革ブランド・バイヤー業界関係者中心。自然な形で提案できる',
    ],
    companies=[
        ('エース株式会社', 'バッグ部門ベストプロダクト賞（2024年）'),
        ('株式会社吉田（PORTER）', 'バッグ部門受賞（2023年）'),
        ('有限会社清川商店', 'Japan Leather Award受賞'),
        ('株式会社村瀬鞄行', 'アーティスティックデザイン賞'),
        ('土屋鞄製造所', 'ワークショップ協力・関係企業'),
    ],
    site_url='https://award.jlia.or.jp/'
)

add_show_card(
    priority_label='優先度：高（ペルソナ明記）',
    date_str='10月7〜9日 / 今から約19週間後',
    weeks_str='2026年 ｜ 東京ビッグサイト',
    show_name='FaW TOKYO【秋展】（ファッションワールド東京）',
    show_sub='主催：RX Japan ｜ 東京ビッグサイト ｜ ペルソナが情報収集源として明記している展示会',
    meta_items=['約850〜1,050社出展', '来場23,000名', '出展者リスト：7月公開予定'],
    header_color='2c4a2c',
    actions=[
        '出展者リストは2026年7月公開予定 → 7月になったら即アクセスしてリスト取得',
        '「ブランド＆デザイナーEXPO」「日本のファッション輸出EXPO」ゾーンに革バッグ企業が集中',
        '6月・9月展での実績・名刺を持って、より深い商談として臨む',
    ],
    companies=[
        ('ARTPHERE（アートフィアー）', '豊岡鞄認定・革ダレスバッグ'),
        ('倉敷元造', '日本製キャンバス×レザーバッグ'),
        ('株式会社服部（HATTORI）', '豊岡鞄・革バッグ150年の老舗'),
    ],
    site_url='https://www.fashion-tokyo.jp/hub/ja-jp.html',
    company_note='※全リストは2026年7月のサイト公開後に確認可能'
)

# ============================================================
# SECTION 4: 参考
# ============================================================
add_heading('参考：その他の関連展示会', size=13)

add_show_card(
    priority_label='参考',
    date_str='12月3〜4日 / 今から約27週間後',
    weeks_str='2026年 ｜ 都立産業貿易センター台東館',
    show_name='TOKYO LEATHER FAIR（東京レザーフェア）',
    show_sub='主催：協同組合資材連 ｜ 都立産業貿易センター台東館 ｜ 革素材・副資材の専門商談展',
    meta_items=['約150〜180社出展（素材メーカー中心）', '来場者側にブランド企業', '業界参加登録制'],
    header_color='444444',
    actions=[
        '⚠️ 出展者は革素材・タンナー・パーツメーカーが中心',
        '革ブランドのオーナー社長は来場者側として参加している場合が多い',
        '素材を買いに来るブランドバイヤーに直接接触する機会として活用する',
    ],
    companies=[
        ('山陽（Sanyo Leather）', '皮革素材卸'),
        ('小笠原染革所', '染革メーカー'),
        ('久保柳商店', '浅草皮革卸問屋'),
    ],
    site_url='https://tlf.jp/'
)

# ============================================================
# タイムライン
# ============================================================
add_separator()
add_heading('今日から12月までの展示会タイムライン', size=13)

timeline = [
    ('2026年5月26日（今日）', '★ ライフスタイルWeek夏の来場者登録【今すぐ】',
     'lifestyle-expo.jp/summer/ で無料登録。出展社検索で「バッグ・革小物」を絞り込みターゲット企業一覧を取得。', True),
    ('今日〜今週中', '★ TBL（日本バッグ協会）と日本バッグ工業連合会に問い合わせ【最優先】',
     'tbl.gr.jp へ「11月の見本市の来場登録方法を教えてください」とメール。K・N・O・Tの10月日程も確認。', True),
    ('6月1〜10日', '対面提案ツールの準備',
     '日本語パンフレット・Jigsaw UK実績チラシ・革サンプル・名刺を6月24日までに用意する。', False),
    ('6月24〜26日', '【初の対面】ライフスタイルWeek夏に来場・対面提案',
     '東京ビッグサイト。「バッグ・革小物」カテゴリの出展企業を優先ルートで回る。Sanheの革サンプルを手渡しで提案。', True),
    ('6月27日〜7月初旬', 'フォローアップ ＋ FaW TOKYO出展者リスト公開の確認',
     '名刺交換した全社に個別メール送信。fashion-tokyo.jp の出展者リストが7月公開次第即チェック。', False),
    ('9月2〜4日', '【2回目】東京ギフトショー秋に来場',
     '6月の成果・改善を反映。Kanmi.・HATTORI・ナダヤ等を重点ターゲットに。', True),
    ('9〜10月', 'ジャパンレザーアワードに来場',
     '渋谷での開催。受賞者・参加者と自然な形でコネクションを作る。', False),
    ('10月7〜9日', '【最重要】FaW TOKYO 秋展に来場',
     'ペルソナが最も情報収集に使う展示会。3回の対面経験を積んだ上で最も深い商談を狙う。', True),
    ('10月（日程要確認）', '【バッグ専門】K・N・O・T collection 秋展に来場',
     '浅草橋ヒューリックホール。バッグブランドのみの専門展。', True),
    ('11月（日程要確認）', '【バッグ専門】ハンドバッグかばん見本市（TBL）秋展に来場',
     'Dakota・genten・UNISON DEPT等の本命ターゲットと直接接触できる最もターゲット精度が高い展示会。', True),
    ('12月3〜4日', '東京レザーフェアに来場（来場者側のブランドをターゲットに）',
     '都立産業貿易センター台東館。素材を買いに来る革ブランドのバイヤーに直接接触できる機会として活用。', False),
]

for date, title, desc, is_key in timeline:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(0)
    r1 = p.add_run(date + '　')
    r1.font.size = Pt(8.5); r1.font.color.rgb = MUTED; r1.font.bold = True
    r2 = p.add_run(title)
    r2.font.size = Pt(10.5); r2.font.bold = True
    r2.font.color.rgb = ACCENT if is_key else NAVY

    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after  = Pt(4)
    p2.paragraph_format.left_indent  = Cm(0.5)
    r3 = p2.add_run(desc)
    r3.font.size = Pt(9); r3.font.color.rgb = MUTED

    add_separator()

# 保存
out = '/home/user/TaayukiPokemonBattle/Sanhe_展示会戦略.docx'
doc.save(out)
print(f'保存完了: {out}')
