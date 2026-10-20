# -*- coding: utf-8 -*-
"""
workflow-bot 運用マニュアル生成スクリプト
配色: シンプル(黒基調)。警告は赤ハイライトマーカー。絵文字なし。です・ます調。
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ---------- 全体スタイル設定 ----------
style = doc.styles['Normal']
style.font.name = 'Meiryo'
style.font.size = Pt(10.5)
rpr = style.element.get_or_add_rPr()
rFonts = rpr.find(qn('w:rFonts'))
if rFonts is None:
    rFonts = OxmlElement('w:rFonts')
    rpr.append(rFonts)
rFonts.set(qn('w:eastAsia'), 'Meiryo')

# 見出しスタイル
for lvl, sz in [('Heading 1', 18), ('Heading 2', 14), ('Heading 3', 12)]:
    s = doc.styles[lvl]
    s.font.name = 'Meiryo'
    s.font.size = Pt(sz)
    s.font.color.rgb = RGBColor(0, 0, 0)
    s.font.bold = True
    rpr = s.element.get_or_add_rPr()
    rF = rpr.find(qn('w:rFonts'))
    if rF is None:
        rF = OxmlElement('w:rFonts')
        rpr.append(rF)
    rF.set(qn('w:eastAsia'), 'Meiryo')

# ページ余白
for section in doc.sections:
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)


# ---------- ヘルパー ----------
def add_para(text='', bold=False, size=None, align=None, space_after=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    r.bold = bold
    if size:
        r.font.size = Pt(size)
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    return p


def add_bullet(text, level=0):
    p = doc.add_paragraph(text, style='List Bullet')
    p.paragraph_format.left_indent = Cm(0.6 + level * 0.6)
    return p


def add_warning(text):
    """薄いグレー背景＋左に太い縦線の枠。ラベル「【重要】」のみ太字黒。本文は普通の黒文字。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    # ラベル
    r1 = p.add_run('【重要】  ')
    r1.bold = True
    r1.font.color.rgb = RGBColor(0, 0, 0)
    # 本文
    r2 = p.add_run(text)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    # 枠線: 左だけ太く、他は薄く
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    # 左辺: 太い黒縦線
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), '24')
    left.set(qn('w:space'), '8')
    left.set(qn('w:color'), '404040')
    pBdr.append(left)
    # 他3辺: 薄いグレー
    for edge in ('top', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '4')
        b.set(qn('w:color'), 'BFBFBF')
        pBdr.append(b)
    pPr.append(pBdr)
    # 薄いグレー背景
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F5F5F5')
    pPr.append(shd)
    return p


def add_note(text):
    p = doc.add_paragraph()
    r = p.add_run('（補足） ' + text)
    r.italic = True
    r.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    return p


def add_code(text):
    """等幅フォントで枠風の段落"""
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = 'Consolas'
    r.font.size = Pt(9.5)
    # 枠線
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '4')
        b.set(qn('w:color'), '808080')
        pBdr.append(b)
    pPr.append(pBdr)
    # 灰色背景
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F2F2F2')
    pPr.append(shd)
    return p


def add_screenshot_placeholder(desc):
    p = doc.add_paragraph()
    r = p.add_run(f'[スクリーンショット挿入箇所: {desc}]')
    r.italic = True
    r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'dashed')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '6')
        b.set(qn('w:color'), '808080')
        pBdr.append(b)
    pPr.append(pBdr)
    return p


def add_table(headers, rows, col_widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = 'Light Grid Accent 1'
    # ヘッダー
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ''
        p = hdr[i].paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(10)
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = t.rows[r_idx + 1].cells[c_idx]
            cell.text = ''
            para = cell.paragraphs[0]
            run = para.add_run(str(val))
            run.font.size = Pt(9.5)
    if col_widths:
        for row in t.rows:
            for c_idx, w in enumerate(col_widths):
                row.cells[c_idx].width = Cm(w)
    return t


def add_checklist(items):
    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.4)
        r = p.add_run('[ ] ' + item)
        r.font.size = Pt(10.5)


def page_break():
    doc.add_page_break()


# ================================================================
# 表紙
# ================================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('\n\n\n\nワークフローBot')
r.bold = True; r.font.size = Pt(28)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('運用・引継ぎマニュアル')
r.bold = True; r.font.size = Pt(22)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('\n\nVersion 1.0（対象システムバージョン v1.5.0）')
r.font.size = Pt(12)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('作成日: 2026年9月28日')
r.font.size = Pt(11)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('\n\n\n\n\n\n\n\n\n本書の対象読者:\nGoogle Apps Script を少し触ったことがあり、\nGASがどのようなものかは理解しているが、コードを読むのは苦手な方')
r.font.size = Pt(11)

page_break()

# ================================================================
# 目次(手動)
# ================================================================
doc.add_heading('目次', level=1)
toc_items = [
    '第1章  システム概要',
    '第2章  システム構成要素と設置場所',
    '第3章  管理画面の使い方（日常運用）',
    '第4章  コード変更を本番に反映する（デプロイ手順）',
    '第5章  監視とヘルスチェック',
    '第6章  トラブルシューティング（症状別対応）',
    '第7章  バックアップと復元',
    '第8章  セキュリティ運用',
    '第9章  リスクと事故想定FAQ',
    '第10章 エスカレーション先と参考リンク',
    '付録A  用語集',
    '付録B  スプレッドシート列の完全リファレンス',
    '付録C  ScriptProperties 必須キー一覧',
    '付録D  引継ぎ時チェックリスト',
]
for it in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    r = p.add_run(it)
    r.font.size = Pt(11)

page_break()

# ================================================================
# 第1章 システム概要
# ================================================================
doc.add_heading('第1章 システム概要', level=1)

doc.add_heading('1.1 このシステムは何をするものか', level=2)
add_para('ワークフローBotは、Google Chat 上で「申請・承認」や「定型文送信」を行うためのBotです。管理者がワークフロー（例：稟議申請、外出届など）を作成すると、指定した Google Chat スペースに「申請する」ボタンが投稿されます。ユーザーがそのボタンを押すと入力フォームが開き、送信されると、指定された承認者のスペースに承認カードが投稿されます。承認者は「承認」「却下」などのボタンを押すだけで処理が完結します。')
add_para('全ての申請データはスプレッドシートに自動保存され、必要に応じて外部スプレッドシートにもコピーされます。')

doc.add_heading('1.2 全体アーキテクチャ', level=2)
add_para('本システムは以下の3層で構成されます。')
add_para('（1）Google Chat（ユーザーが操作する場所）')
add_bullet('ボタンやフォームの表示、承認者への通知を担当します。')
add_para('（2）Google Apps Script（頭脳）')
add_bullet('Chat からのリクエストを受け取り、スプレッドシートへの読み書き、承認カードの送信、権限判定などを行います。')
add_bullet('本システムはこのGASスクリプトが中心です。')
add_para('（3）Google スプレッドシート（データ保管庫）')
add_bullet('ワークフロー定義、申請データ、管理者リスト、監査ログを保存します。')

add_screenshot_placeholder('Google Chat → GAS → Spreadsheet の関係を示す図（構築者作成推奨）')

doc.add_heading('1.3 主な機能一覧', level=2)
add_table(['機能', '説明'], [
    ['ワークフロー作成', '管理画面から入力項目・承認者・送信先スペースを設定できます'],
    ['申請投稿', 'Chatのボタンを押すとフォームが開き、送信するとカードが投稿されます'],
    ['承認処理', '承認者はカード上のボタンを押すだけで承認・却下ができます'],
    ['ステータス更新', '承認/却下によりカード上のステータスが即座に書き換わります'],
    ['スレッド返信', '承認時にスレッド返信メッセージを自動投稿できます'],
    ['メンション通知', '承認者ユーザーIDを設定すると自動的に @メンションされます'],
    ['外部シート連携', '任意の外部スプレッドシートにデータをコピーできます'],
    ['管理者権限管理', '3段階（オーナー/エディタ/メンバー）で管理者権限を制御できます'],
    ['共同編集者', 'ワークフローごとに編集許可ユーザーを設定できます'],
    ['監査ログ', '全ての管理操作は自動的にログに記録されます'],
    ['自動バックアップ', '毎日、スプレッドシートのバックアップが作成されます'],
], col_widths=[4.5, 11])

doc.add_heading('1.4 バージョン', level=2)
add_para('本マニュアルの対象は v1.5.0（2026年9月28日時点）です。バージョンは管理画面左下に表示されます。')
add_note('システムを更新した場合は、本マニュアル末尾の「変更履歴」に追記してください。')

page_break()

# ================================================================
# 第2章 システム構成要素と設置場所
# ================================================================
doc.add_heading('第2章 システム構成要素と設置場所', level=1)

doc.add_heading('2.1 Apps Script プロジェクト', level=2)
add_para('本システムのコードは、Google Apps Script プロジェクトに配置されています。')
add_para('プロジェクトへのアクセス方法:')
add_bullet('Google ドライブから該当スプレッドシートを開く')
add_bullet('メニューの「拡張機能」→「Apps Script」をクリック')
add_bullet('または、script.google.com にアクセスして一覧から選択')

add_warning('Apps Script プロジェクトの所有者を変更する際は、必ずデプロイURLが変わっていないことを確認してください。所有者変更後はデプロイURLが変わるケースがあり、Chat側の設定も変更が必要になります。')

doc.add_heading('2.2 スクリプトファイル一覧', level=2)
add_para('プロジェクトには以下のファイルがあります。それぞれの役割を把握しておくと、トラブル時にどこを確認すべきか判断しやすくなります。')
add_table(['ファイル名', '役割'], [
    ['Main.js', 'Chat/Web からのリクエストを最初に受け取る入口。バージョン番号もここで定義'],
    ['Chat.js', '申請カードの投稿、承認処理、フォーム表示など Chat 連携の中核'],
    ['Adminapi.js', '管理画面から呼ばれる API 群（ワークフロー保存、削除、管理者管理など）'],
    ['Auth.js', '認証・権限判定・トークン生成・監査ログ・入力バリデーション'],
    ['Data.js', 'スプレッドシート読み書きのヘルパー関数群'],
    ['Schema.js', 'シートの列構造を自動修復するヘルパー'],
    ['Backup.js', '自動バックアップ処理'],
    ['Monitor.js', '定期実行の監視処理'],
    ['admin.html', '管理画面のHTMLテンプレート'],
    ['JavaScript.html', '管理画面のフロントエンドロジック（ブラウザ側）'],
    ['Stylesheet.html', '管理画面のスタイル定義'],
    ['appsscript.json', 'Apps Script プロジェクトの設定ファイル（スコープ等）'],
], col_widths=[4.0, 12.0])

doc.add_heading('2.3 スプレッドシート構成', level=2)
add_para('スプレッドシートには以下のシートがあります。頭にアンダースコアが付いているシートはシステム用です。')
add_para('スプレッドシートIDは ScriptProperties の SPREADSHEET_ID に設定されています（第2.5節参照）。')

add_para('主要システムシート:', bold=True)
add_table(['シート名', '役割'], [
    ['_設定', 'ワークフロー定義本体（1行=1ワークフロー）'],
    ['_項目設定', '各ワークフローのフォーム項目定義'],
    ['_ボタン設定', '各ワークフローのアクションボタン定義'],
    ['_管理者', '管理者ユーザー一覧と権限（オーナー/エディタ/メンバー）'],
    ['_監査ログ', '全ての管理操作の履歴'],
], col_widths=[4.0, 12.0])

add_para('各ワークフローごとに、そのワークフロー名のデータシートが自動で作成されます。例えば「稟議申請」というワークフローを作ると、「稟議申請」という名前のシートが右端に追加されます。')

add_warning('システムシート（_で始まるシート）を手動で削除・列変更しないでください。特に列を削除するとデータ書き込み時にエラーになります。復旧手順は第7章にあります。')

add_para('各シートの列の完全な意味は「付録B スプレッドシート列の完全リファレンス」を参照してください。')

doc.add_heading('2.4 Google Chat アプリ設定', level=2)
add_para('Google Chat 側では、本Botを「Chatアプリ」として登録しています。設定は Google Cloud Console の Chat API 設定画面で確認できます。')
add_bullet('アプリ名: workflow-bot（環境により異なる場合あり）')
add_bullet('アバターURL・説明文: Chat API 設定画面で編集可能')
add_bullet('接続設定: HTTPエンドポイント方式で、FUNCTION_URL（ScriptProperties）に指定されたURLが登録されている')

add_warning('接続URLを変更する場合は、Chat API 側のURL設定と、ScriptProperties の FUNCTION_URL の両方を同時に更新する必要があります。片方だけの更新はカードのボタンが動かなくなる原因になります。')

doc.add_heading('2.5 ScriptProperties（設定値の保管場所）', level=2)
add_para('パスワードやスプレッドシートIDなどの秘密情報は、コードには書かず ScriptProperties（スクリプトプロパティ）に保存されています。')
add_para('確認方法:')
add_bullet('Apps Script エディタを開く')
add_bullet('左サイドバーの歯車アイコン「プロジェクトの設定」をクリック')
add_bullet('下部の「スクリプト プロパティ」欄で確認・編集できる')

add_screenshot_placeholder('Apps Script のプロジェクト設定 → スクリプトプロパティ画面')

add_para('必須キーの一覧は「付録C」を参照してください。')

add_warning('ScriptProperties の値を空にすると、システム全体が動かなくなります。特に SPREADSHEET_ID と SERVICE_ACCOUNT_KEY は誤って削除しないよう注意してください。')

doc.add_heading('2.6 サービスアカウント', level=2)
add_para('Botとして Chat に投稿するために、Google Cloud のサービスアカウントを使用しています。そのキー（JSON）は ScriptProperties の SERVICE_ACCOUNT_KEY に保存されています。')
add_bullet('サービスアカウントは Google Cloud Console で作成・管理されます')
add_bullet('キーは 数年に一度、または退職者対応時にローテーション（再発行）することを推奨します')

add_warning('サービスアカウントキーが漏洩した場合、直ちに Google Cloud Console でキーを無効化し、新しいキーを発行してScriptProperties に上書きしてください。詳細は第8章。')

page_break()

# ================================================================
# 第3章 管理画面の使い方
# ================================================================
doc.add_heading('第3章 管理画面の使い方（日常運用）', level=1)

doc.add_heading('3.1 管理画面へのアクセス', level=2)
add_para('管理画面のURLは、本番デプロイURLに `/exec` が付いたものです。ScriptProperties の FUNCTION_URL がまさにそのURLです。')
add_bullet('URL 例: https://script.google.com/a/macros/cmind-co.jp/s/AKfycbw.../exec')
add_bullet('ブラウザで開くと Google ログインが求められ、管理者として登録されていれば管理画面が表示されます')

add_warning('管理者として _管理者 シートに登録されていないユーザーがアクセスすると「権限がありません」と表示されます。追加は既存のオーナー権限管理者に依頼してください。')

add_para('URLの種類に注意:')
add_bullet('/exec で終わるURL: 本番用（バージョン固定）')
add_bullet('/dev で終わるURL: 開発用（最新コードが即時反映されるが、管理者以外はアクセス不可）')

doc.add_heading('3.2 3段階の権限モデル', level=2)
add_para('本システムには3種類の権限があります。')
add_table(['権限', 'できること', 'できないこと'], [
    ['オーナー', '全ワークフローの閲覧・編集・削除・復元、管理者の追加削除、権限変更', '（制限なし）'],
    ['エディタ', '全ワークフローの閲覧・編集、有効/無効切替', 'ワークフローの完全削除、アーカイブ復元、管理者管理'],
    ['メンバー', '自分が作成した／自分が共同編集者に指定されたワークフローの閲覧・編集', '他人のワークフロー閲覧、管理者管理、アーカイブ復元'],
], col_widths=[2.5, 8.0, 5.5])

add_note('管理者を新規追加する際、既定は「エディタ」です。必要に応じて追加後にプルダウンで変更してください。')

doc.add_heading('3.3 ワークフローの新規作成', level=2)
add_para('手順:')
add_bullet('管理画面 左上の「新規」ボタンを押します')
add_bullet('「基本情報」タブで、ワークフロー名、種類（申請・承認 or 定型文送信）、実行スペース、送信先スペースを選択します')
add_bullet('「フォーム項目」タブで、入力してもらいたい項目を追加します')
add_bullet('「承認・通知」タブで、承認者や返信メッセージなどを設定します')
add_bullet('「保存」ボタンを押します')

add_screenshot_placeholder('管理画面のワークフロー作成タブ')

add_para('保存後、「設置」ボタンを押すと、指定したスペースに「申請する」ボタンが投稿されます。')

add_warning('Bot が対象スペースに参加していないと「設置」してもボタンが表示されません。事前に対象スペースに Bot をメンバー追加してください（「@workflow-bot」のようにメンションして招待）。')

doc.add_heading('3.4 共同編集者の設定', level=2)
add_para('メンバー権限のユーザーでも、特定のワークフローだけは編集させたい場合、「基本情報」タブの「共同編集者」欄にメールアドレスをカンマ区切りで入力してください。')
add_bullet('例: yamada@example.com, sato@example.com')
add_bullet('オーナー・エディタは共同編集者の設定に関係なく全ワークフローを編集可能です')
add_bullet('自分が作成したワークフローは、メンバー権限でも常に編集可能です')

doc.add_heading('3.5 ワークフローの編集・アーカイブ・完全削除', level=2)
add_para('編集:')
add_bullet('サイドバーからワークフローを選択し、内容を変更して「保存」を押します')
add_bullet('項目を削除する変更を行うと、既存データがある場合は確認ダイアログが表示されます')

add_para('アーカイブ（論理削除）:')
add_bullet('ワークフロー編集画面の「削除」ボタンを押します')
add_bullet('サイドバーから消えますが、30日以内なら「アーカイブ」画面から復元できます')
add_bullet('30日経過後に完全削除されます')

add_para('完全削除:')
add_bullet('「アーカイブ」画面を開き、対象の「完全削除」ボタンを押します')
add_bullet('データシートも含めて全て削除されます')

add_warning('完全削除は取り消せません。データシートと監査ログ以外の全ての痕跡が消えます。必要な場合は先にデータシートを別ブックにコピーしてください。')

doc.add_heading('3.6 承認者と共同編集者の違い', level=2)
add_para('この2つは混同しやすいので注意してください。')
add_table(['項目', '承認者', '共同編集者'], [
    ['役割', '申請を承認する人', 'ワークフロー定義を編集できる人'],
    ['設定場所', '「承認・通知」タブ', '「基本情報」タブ'],
    ['管理画面へのアクセス', '不要', '必要（＝管理者に登録済であること）'],
    ['通知', 'メンション通知が飛ぶ', '通知なし'],
], col_widths=[4.0, 6.0, 6.0])

doc.add_heading('3.7 管理者の追加・削除・権限変更', level=2)
add_para('（オーナー権限を持つユーザーのみ実行可能）')
add_bullet('左サイドバー下部の「管理者」ボタンをクリック')
add_bullet('追加: 下部の入力欄にメールアドレスを入力し、権限を選んで「追加」を押す')
add_bullet('権限変更: 一覧の右側プルダウンから変更する（自動で保存対象になる）')
add_bullet('削除: 一覧の「×」ボタンをクリック')
add_bullet('全ての変更後、「保存」ボタンを必ず押す')

add_warning('自分自身を管理者から外すことはできません。また、オーナーを最低1名残す必要があります（システムに阻止されます）。')

page_break()

# ================================================================
# 第4章 デプロイ手順
# ================================================================
doc.add_heading('第4章 コード変更を本番に反映する（デプロイ手順）', level=1)

doc.add_heading('4.1 なぜ「反映作業」が必要なのか', level=2)
add_para('Apps Scriptでコードを保存しても、そのままでは本番URLには反映されません。本Botは「本番デプロイURL（バージョン固定）」を使っているため、意図的にバージョンを差し替える操作が必要です。')

add_warning('この仕組みを理解していないと、「コードを直したはずなのに動きが変わらない」というトラブルが必ず発生します。')

add_para('用語:')
add_bullet('デプロイ: 特定バージョンのコードを外部から呼び出せる状態にすること')
add_bullet('デプロイID: 各デプロイに割り振られる固有ID（AKfycbw... の形式）')
add_bullet('本番デプロイ: Chat のボタンや管理画面URLに紐づいた1つのデプロイ')

doc.add_heading('4.2 デプロイの種類', level=2)
add_para('本システムには通常2つのデプロイが存在します。')
add_table(['デプロイ', 'URL末尾', 'バージョン', '用途'], [
    ['本番用', '/exec', '固定（例: v58）', 'Chatと管理画面で使用中'],
    ['開発用（@HEAD）', '/dev', '常に最新', '動作確認用（オーナーのみアクセス可）'],
], col_widths=[3.5, 2.0, 3.0, 7.0])

doc.add_heading('4.3 コード変更を本番に反映する手順', level=2)
add_para('（前提: コード修正済み、動作確認済み）')

add_para('手順1: Apps Script エディタを開く')
add_bullet('スプレッドシートから「拡張機能」→「Apps Script」を開く')

add_para('手順2: 右上の「デプロイ」ボタンをクリック → 「デプロイを管理」を選択')
add_screenshot_placeholder('Apps Script エディタ右上の「デプロイ」ボタン')

add_para('手順3: 本番デプロイの右にある鉛筆アイコン（編集）をクリック')
add_bullet('「Web app」と書かれている行の右端の鉛筆')

add_screenshot_placeholder('デプロイを管理ダイアログで本番デプロイの編集ボタン')

add_para('手順4: 「バージョン」プルダウンから「新しいバージョン」を選択')

add_para('手順5: 「説明」欄に変更内容を短く記入（例: v1.5.1 承認ボタンの色を変更）')

add_para('手順6: 「デプロイ」ボタンをクリック')
add_bullet('数秒後に完了メッセージが表示されます')
add_bullet('URLは変わりません（同じデプロイIDが上書き更新される）')

add_warning('絶対に「新しいデプロイ」ボタン（＝新規URL発行）を押さないでください。URLが変わってしまうと Chat 側の設定変更が必要になり、既存カードが動かなくなります。')

doc.add_heading('4.4 反映を確認する', level=2)
add_para('反映後の動作確認:')
add_bullet('管理画面をブラウザで再読み込みし、左下のバージョン番号が更新されているか確認')
add_bullet('Chat のボタンを押して申請が正常に投稿されるか確認')

doc.add_heading('4.5 問題が起きた時の巻き戻し（ロールバック）', level=2)
add_para('新しいデプロイで問題が起きた場合、前のバージョンに戻せます。')
add_bullet('「デプロイを管理」ダイアログを開く')
add_bullet('本番デプロイの編集画面で「バージョン」から古いバージョンを選択')
add_bullet('「デプロイ」ボタンで確定')

add_note('過去のバージョンは Apps Script 内に自動保存されており、任意のバージョンに戻せます。')

page_break()

# ================================================================
# 第5章 監視とヘルスチェック
# ================================================================
doc.add_heading('第5章 監視とヘルスチェック', level=1)

doc.add_heading('5.1 ヘルスチェック関数', level=2)
add_para('システム全体が正常かを確認する関数「healthCheck」があります。')
add_para('実行方法:')
add_bullet('Apps Script エディタを開く')
add_bullet('関数名プルダウンから「healthCheck」を選択')
add_bullet('「実行」ボタンをクリック')
add_bullet('「実行ログ」に結果（JSON形式）が表示される')

add_screenshot_placeholder('Apps Script の関数実行画面')

add_para('結果の見方:')
add_code('{\n  "ok": true,\n  "version": "1.5.0",\n  "env": "prod",\n  "sheetOk": true,\n  "serviceAccountOk": true,\n  ...\n}')
add_bullet('ok: 全体が正常なら true')
add_bullet('sheetOk: スプレッドシートに正常アクセスできれば true')
add_bullet('serviceAccountOk: サービスアカウントで Chat API 認証できれば true')

add_warning('ok が false の場合、その他の項目を確認して原因を特定してください。sheetOk が false なら SPREADSHEET_ID を確認、serviceAccountOk が false なら SERVICE_ACCOUNT_KEY を確認してください。')

doc.add_heading('5.2 監査ログの確認', level=2)
add_para('全ての管理操作（ワークフロー作成、削除、管理者変更など）は _監査ログ シートに自動記録されます。')
add_bullet('列: タイムスタンプ、実行者、アクション、対象、詳細')
add_bullet('保管期間: オンライン1年、その後アーカイブ3年')

add_para('よく使う確認シナリオ:')
add_bullet('「誰が何を消したか」を追跡: アクション列に workflow.archive / workflow.purge を含む行を検索')
add_bullet('「管理者を誰が追加/削除したか」: アクション列 admins.save で絞り込み')
add_bullet('「エラー発生履歴」: アクション列に error.data 等を含む行')

doc.add_heading('5.3 直近エラーの確認', level=2)
add_para('システムでエラーが発生した場合、直近のエラーは ScriptProperties の LAST_ERROR キーに自動保存されます。')
add_bullet('Apps Script → プロジェクトの設定 → スクリプトプロパティ で LAST_ERROR の値を参照')
add_bullet('形式: 「発生日時 関数名: エラーメッセージ」')

doc.add_heading('5.4 Apps Script 実行ログ', level=2)
add_para('Apps Script エディタ左サイドバーの「実行数」（時計アイコン）をクリックすると、直近の実行履歴と、それぞれのログが確認できます。')
add_bullet('関数が失敗している場合は、この画面から詳細ログを追える')
add_bullet('保存期間は Apps Script の仕様で最長7日程度')

doc.add_heading('5.5 定期的な健康診断（推奨頻度）', level=2)
add_table(['項目', '頻度', '確認内容'], [
    ['healthCheck 実行', '週1回', 'ok:true であること'],
    ['_監査ログ 目視', '月1回', 'error.* のログが多発していないか'],
    ['管理者一覧の棚卸し', '四半期に1回', '退職者や異動者が残っていないか'],
    ['バックアップ確認', '月1回', 'バックアップドライブに新しいファイルがあるか'],
    ['サービスアカウントキー更新', '年1回', 'ローテーション（推奨）'],
], col_widths=[5.5, 3.0, 8.0])

page_break()

# ================================================================
# 第6章 トラブルシューティング
# ================================================================
doc.add_heading('第6章 トラブルシューティング（症状別対応）', level=1)

def trouble(sym, causes, checks, actions):
    doc.add_heading(sym, level=3)
    add_para('想定原因:', bold=True)
    for c in causes:
        add_bullet(c)
    add_para('確認手順:', bold=True)
    for c in checks:
        add_bullet(c)
    add_para('対処:', bold=True)
    for a in actions:
        add_bullet(a)

doc.add_heading('6.1 管理画面が開かない・「権限がありません」と出る', level=2)
trouble(
    '症状: 管理画面のURLを開くと「権限がありません」または「アクセス権限が無いです」と表示される',
    [
        '_管理者 シートにメールアドレスが登録されていない',
        '登録メールアドレスに全角空白や不可視文字が混入している',
        '古い（削除された）デプロイURLを開いている',
        '個人アカウントで開いている（会社アカウントで開くべき）',
    ],
    [
        'スプレッドシートを開き _管理者 シートに自分のメールが正確に登録されているか確認',
        '管理画面URLが ScriptProperties の FUNCTION_URL と一致しているか確認',
        'ブラウザで会社アカウントに切り替わっているか確認（右上のアカウントアイコン）',
    ],
    [
        '既存オーナー権限管理者に依頼して自分を追加してもらう',
        '自分のメールを手入力ではなくコピー&ペーストで登録し直す（不可視文字対策）',
        'Chat のカードから管理画面を開く場合は最新デプロイURLになるよう「設置」ボタンで再配置',
    ],
)

doc.add_heading('6.2 Chat のボタンを押しても反応しない', level=2)
trouble(
    '症状: 「申請する」ボタンを押してもフォームが開かない、あるいは無反応',
    [
        'Bot がスペースに参加していない',
        'デプロイURLが変わったのに Chat 側の登録URLが古い',
        'Apps Script の実行時間上限（30秒）を超えた',
        'サービスアカウントキーが無効化された',
    ],
    [
        'スペースのメンバー一覧に「workflow-bot」がいるか確認',
        'Chat API 側の登録URLと ScriptProperties の FUNCTION_URL が一致するか確認',
        'Apps Script の「実行数」画面で該当リクエストのログを確認',
        'healthCheck を実行し serviceAccountOk を確認',
    ],
    [
        'Bot をスペースに追加（「@workflow-bot」でメンション招待）',
        'Chat API 側のURLを更新するか、管理画面で「設置」ボタンで再配置',
        'カードを削除して「設置」ボタンで再配置',
    ],
)

doc.add_heading('6.3 承認カードが承認者に届かない', level=2)
trouble(
    '症状: 申請フォームは送信できるが、承認者側にカードが表示されない',
    [
        '送信先スペースに Bot が未参加',
        '承認者に指定したユーザーがスペースメンバーでない',
        '外部シート指定が誤っていて例外が発生している',
    ],
    [
        '_監査ログ シートで対象申請のログを確認',
        'ScriptProperties の LAST_ERROR を確認',
        'ワークフロー編集画面を保存し直して「保存後の警告」が表示されないか確認',
    ],
    [
        '送信先スペースに Bot を追加',
        '外部シートURLを修正するか外部シート連携をオフにする',
    ],
)

doc.add_heading('6.4 メンション通知が飛ばない', level=2)
trouble(
    '症状: 承認者に @メンション通知が届かない',
    [
        '承認者を「メール手入力」で追加し、ユーザーIDが未取得',
        '承認者がスペースの HUMAN メンバーでない',
    ],
    [
        'ワークフロー編集画面で「メンバーから選択」ボタンから承認者を再選択',
        '選択後に緑色で「N名にメンション通知が飛びます」と表示されているか確認',
    ],
    [
        '「メンバーから選択」で承認者をクリックし直して保存',
    ],
)

doc.add_heading('6.5 同時申請でデータが混ざる（過去バージョンで発生）', level=2)
add_para('v1.4.0 以降ではロック機構で対処済みです。もし v1.3.x 以前で運用している場合は、直ちに最新版へアップデートしてください。')

doc.add_heading('6.6 スプレッドシートの列を誤って削除した', level=2)
trouble(
    '症状: システムシートの列を消してしまい、書き込みエラーが発生する',
    [
        '手動でシステムシート（_設定 など）の列を削除した',
    ],
    [
        'どの列が消えたかを確認（「付録B」の列リファレンスと比較）',
    ],
    [
        '第7章「バックアップと復元」を参照し、直前のバックアップから復元',
        '軽微な列削除なら手動で列を追加し、ヘッダー名を「付録B」の通りに戻す',
    ],
)

doc.add_heading('6.7 「他の管理操作が進行中です」と出る', level=2)
trouble(
    '症状: 保存・削除操作時にこのメッセージが表示される',
    [
        '複数の管理者が同時に管理操作を行った',
        '前の処理がまだ完了していない',
    ],
    [
        '10〜30秒待ってから再度操作',
    ],
    [
        '待って再実行。頻発する場合は運用ルールで「管理操作は同時に1名まで」を徹底',
    ],
)

doc.add_heading('6.8 バージョンが古いまま更新されない', level=2)
trouble(
    '症状: コード修正したのに管理画面のバージョン表示が変わらない',
    [
        '本番デプロイが更新されていない（第4章参照）',
    ],
    [
        '「デプロイを管理」で本番デプロイの現在バージョンを確認',
    ],
    [
        '第4.3節の手順で本番デプロイを新しいバージョンに更新',
    ],
)

page_break()

# ================================================================
# 第7章 バックアップと復元
# ================================================================
doc.add_heading('第7章 バックアップと復元', level=1)

doc.add_heading('7.1 自動バックアップの仕組み', level=2)
add_para('本システムは毎日自動でスプレッドシートを Google ドライブの指定フォルダにコピーしています。')
add_bullet('実行時刻: 深夜（Apps Script のトリガーで設定）')
add_bullet('保管世代: 直近30日分')
add_bullet('30日を超えた古いバックアップは自動削除されます')

doc.add_heading('7.2 バックアップフォルダの確認方法', level=2)
add_bullet('Google ドライブを開く')
add_bullet('検索欄で「workflow-bot バックアップ」等と検索するか、事前に共有されたフォルダURLを開く')
add_bullet('ファイル名は「元スプレッドシート名_YYYYMMDD_HHMMSS」形式')

add_screenshot_placeholder('バックアップフォルダの実際の場所とファイル命名例')

doc.add_heading('7.3 スプレッドシートを復元する手順', level=2)
add_para('（誤削除・データ破損時）')
add_bullet('バックアップフォルダから復元したい日付のファイルを開く')
add_bullet('「ファイル」→「コピーを作成」で新しいスプレッドシートを作成')
add_bullet('新しいスプレッドシートのIDをURLから取得（/d/ の後の英数字部分）')
add_bullet('ScriptProperties の SPREADSHEET_ID を新しいIDに変更')
add_bullet('healthCheck を実行して ok:true を確認')

add_warning('復元後は、バックアップ日以降の申請データが失われます。この事実を関係者に周知した上で復元してください。')

doc.add_heading('7.4 バックアップ設定の確認・変更', level=2)
add_para('バックアップの実行は Apps Script のトリガーで動いています。')
add_bullet('Apps Script エディタ左サイドバーの「トリガー」（時計アイコン）をクリック')
add_bullet('「dailyBackup」等のトリガーが登録されているか確認')
add_bullet('停止したい場合はトリガーの右クリックメニューから削除')

doc.add_heading('7.5 スプレッドシートの部分復元', level=2)
add_para('全体を戻すほどではないが、特定シート・特定行だけ戻したい場合:')
add_bullet('バックアップファイルを開き、対象シートを右クリック → 「コピーして別ブックへ」')
add_bullet('本番スプレッドシートに手動で該当シートをコピー')
add_bullet('列構造が現行と一致しているか「付録B」で確認')

page_break()

# ================================================================
# 第8章 セキュリティ運用
# ================================================================
doc.add_heading('第8章 セキュリティ運用', level=1)

doc.add_heading('8.1 権限管理の原則', level=2)
add_bullet('オーナー権限は最小限（推奨2〜3名）に絞る')
add_bullet('エディタ権限も業務上必要な人だけに付与')
add_bullet('メンバー権限は基本的にワークフロー作成者のみ')
add_bullet('退職者・異動者は速やかに削除')

doc.add_heading('8.2 管理者の棚卸し', level=2)
add_para('四半期に1回、以下を実施してください。')
add_checklist([
    '管理画面の「管理者」から一覧を出力',
    '退職者・異動者がいないか確認',
    'オーナー権限が過剰に付与されていないか確認',
    '削除・降格を実施',
    '_監査ログ で admins.save のログが記録されているか確認',
])

doc.add_heading('8.3 サービスアカウントキーの管理', level=2)
add_para('サービスアカウントキー（SERVICE_ACCOUNT_KEY）は Bot の身分証明書に相当します。')
add_bullet('保管場所: ScriptProperties のみ（コードには絶対に書かない）')
add_bullet('更新頻度: 年1回のローテーション推奨')
add_bullet('漏洩時: 直ちに旧キーを無効化し、新キーを発行')

add_warning('サービスアカウントキーが第三者に漏洩すると、そのキーで Chat 上の任意のメッセージ送信や、Bot 権限で行える全ての操作が実行可能になります。GitHub 等に絶対に公開しないでください。')

doc.add_heading('8.4 サービスアカウントキー更新手順', level=2)
add_bullet('Google Cloud Console にログイン')
add_bullet('該当プロジェクト → 「IAM と管理」→ 「サービスアカウント」')
add_bullet('対象サービスアカウントを選択 → 「キー」タブ')
add_bullet('「鍵を追加」→「新しい鍵を作成」→ JSON形式でダウンロード')
add_bullet('JSON全体を ScriptProperties の SERVICE_ACCOUNT_KEY に上書き')
add_bullet('healthCheck を実行し serviceAccountOk:true を確認')
add_bullet('動作確認後、古いキーを Google Cloud Console 側で削除')

doc.add_heading('8.5 プライバシー同意', level=2)
add_para('管理画面に初回アクセスした管理者には、プライバシー通知への同意ダイアログが表示されます。同意日時は _管理者 シートの「同意日時」列に記録されます。')

doc.add_heading('8.6 監査ログの改ざん防止', level=2)
add_para('_監査ログ シートは追記専用が理想です。')
add_bullet('編集者権限を極力減らす（オーナー以外は閲覧のみ推奨）')
add_bullet('定期的に別ストレージへエクスポート')

page_break()

# ================================================================
# 第9章 リスクと事故想定FAQ
# ================================================================
doc.add_heading('第9章 リスクと事故想定FAQ', level=1)
add_para('Chat移行責任者向けに、想定される事故ケースとその発生確率・軽減策をまとめます。')

def risk(q, prob, impact, mitigation):
    doc.add_heading(q, level=3)
    add_para(f'発生確率: {prob}', bold=True)
    add_para(f'影響度: {impact}', bold=True)
    add_para('軽減策:', bold=True)
    for m in mitigation:
        add_bullet(m)

risk(
    'Q1: 同時申請でデータが混ざることは？',
    '極めて低い（v1.4.0以降）',
    '過去バージョンでは中程度',
    ['LockService による排他制御と SpreadsheetApp.flush() による整合性確保で対処済み', '過去バージョンで運用の場合は必ずアップデート'],
)

risk(
    'Q2: Bot がスペース未参加で承認カードが飛ばないことは？',
    '低い',
    '中（承認フローが止まる）',
    ['保存時にBot参加状態を自動チェックし警告表示', '手動でBotをスペース追加する運用ルール', '「設置」ボタン押下時に自動追加を試行'],
)

risk(
    'Q3: 承認者のメールに不可視文字が混入して権限誤判定されることは？',
    '低い',
    '中（本人なのに承認できない）',
    ['保存時にメール文字列を正規化（不可視文字除去＋小文字化）', '承認判定側も同様に正規化して比較'],
)

risk(
    'Q4: 管理者が誤ってスプレッドシート列を削除したら？',
    '低い（教育次第）',
    '大（システム停止級）',
    ['列削除時の自動復旧処理（ensureDataSheetSchema_ 等）', '毎日の自動バックアップ', '運用マニュアルでの周知'],
)

risk(
    'Q5: サービスアカウントキー漏洩',
    '極めて低い',
    '極大',
    ['コードにキーを直書きしない（ScriptProperties のみ）', '年次ローテーション', 'GitHub 公開リポジトリでは扱わない'],
)

risk(
    'Q6: ScriptProperties が空になる（誤削除）',
    '低い',
    '大（全機能停止）',
    ['本マニュアル付録Cで必須キー一覧を明示', 'healthCheck で異常検知', 'ScriptProperties 変更前にスクリーンショット記録推奨'],
)

risk(
    'Q7: デプロイURL変更でカードが動かなくなる',
    '低い',
    '大',
    ['「新しいデプロイ」ではなく「デプロイ管理→編集」で更新するルール徹底', '本マニュアル第4章で明示'],
)

risk(
    'Q8: Apps Script の1日実行回数上限（クオータ）に達する',
    '極めて低い（通常業務量では到達困難）',
    '中',
    ['api_getQuotaStatus で状況確認可能', 'ユーザー数が急増する場合は Google Workspace Enterprise 契約検討'],
)

risk(
    'Q9: メンバー権限のユーザーが他人のワークフローを触ってしまう',
    '極めて低い（v1.5.0以降）',
    '中',
    ['3段階権限＋共同編集者機能で制御', 'API側で二重チェック（フロントエンドの制御だけに頼らない）'],
)

risk(
    'Q10: 大量エラーで運用担当が業務を圧迫される',
    '低い（想定される主要リスクは網羅対処済み）',
    '中',
    ['LAST_ERROR 記録機能で原因特定を高速化', '警告モーダルでサイレント失敗を事前検知', '本マニュアル第6章の症状別対応で対処時間短縮'],
)

page_break()

# ================================================================
# 第10章 エスカレーション先と参考リンク
# ================================================================
doc.add_heading('第10章 エスカレーション先と参考リンク', level=1)

doc.add_heading('10.1 社内連絡先', level=2)
add_para('システムに関する問い合わせ・障害連絡は以下へ。')
add_table(['カテゴリ', '連絡先'], [
    ['開発・技術問い合わせ', 'konosuke.okazaki@cmind-co.jp'],
    ['管理者権限追加依頼', '既存オーナー権限管理者'],
    ['Google Workspace 全般', 'IT 部門（貴社の担当窓口）'],
], col_widths=[6.0, 10.0])

doc.add_heading('10.2 外部参考リンク', level=2)
add_bullet('Google Apps Script 公式: https://developers.google.com/apps-script')
add_bullet('Google Chat API: https://developers.google.com/chat')
add_bullet('Google Cloud Console: https://console.cloud.google.com/')

doc.add_heading('10.3 GitHub リポジトリ', level=2)
add_para('本システムのソースコードは GitHub 上で管理されています（プライベートリポジトリ）。')
add_bullet('リポジトリURL: 開発担当者から共有されます')
add_bullet('変更履歴・コミットログはリポジトリ上で確認可能')

page_break()

# ================================================================
# 付録A 用語集
# ================================================================
doc.add_heading('付録A 用語集', level=1)
add_table(['用語', '意味'], [
    ['GAS', 'Google Apps Script の略。Google製サービスを自動化するスクリプト環境'],
    ['デプロイ', 'コードを外部から呼び出せる状態にすること'],
    ['デプロイID', '各デプロイに割り当てられる固有ID（AKfycbw... 形式）'],
    ['@HEAD', '「最新コード」を指すデプロイモード。/dev URL がこれに該当'],
    ['ScriptProperties', 'GAS が持つキーバリュー型の設定保管領域。秘密情報を保存する'],
    ['サービスアカウント', 'ユーザーではなくプログラムが Google API を叩くための身分証明'],
    ['スペース', 'Google Chat の会話単位。個人DM・グループ・ルームの総称'],
    ['カード', 'Google Chat のリッチな UI 要素（ボタンやフォーム付きメッセージ）'],
    ['ワークフロー', '本システムの基本単位。1つの申請・記録の型'],
    ['アーカイブ', '論理削除。30日以内なら復元可能'],
    ['完全削除', '物理削除。取り消し不可'],
    ['監査ログ', '誰がいつ何をしたかを記録した履歴'],
    ['オーナー/エディタ/メンバー', '3段階の管理者権限'],
    ['共同編集者', 'ワークフローごとに指定できる編集許可ユーザー'],
    ['ヘルスチェック', 'システム全体の正常性を確認する診断関数'],
    ['LockService', '同時実行を排他制御する GAS の仕組み'],
    ['トークン', '承認リンク等に埋め込まれる、なりすまし防止のための署名文字列'],
], col_widths=[4.5, 11.5])

page_break()

# ================================================================
# 付録B シート列リファレンス
# ================================================================
doc.add_heading('付録B スプレッドシート列の完全リファレンス', level=1)

doc.add_heading('B.1 _設定 シート', level=2)
add_table(['列', 'ヘッダー名', '意味'], [
    ['A', 'ID', 'ワークフロー内部ID（自動採番）'],
    ['B', 'ワークフロー名', '表示名（シート名にもなる）'],
    ['C', '種類', '申請・承認 または 定型文送信'],
    ['D', '実行スペースID', 'ボタンを設置するスペースID'],
    ['E', '送信先スペースID', 'カードを投稿するスペースID'],
    ['F', '作成者', '作成者メールアドレス'],
    ['G', '作成日', '作成日時'],
    ['H', 'ステータス', '有効 / 無効 / archived'],
    ['I', 'ボタンカードメッセージ名', '設置ボタンカードの Chat内メッセージ名'],
    ['J', '承認者', '承認者メールアドレス（カンマ区切り）'],
    ['K', '外部スプレッドシートURL', '外部シート連携用（任意）'],
    ['L', 'アーカイブ日時', 'アーカイブされた日時'],
    ['M', '承認者UserID', 'メンション用のユーザーID（カンマ区切り）'],
    ['N', 'メンション文言', '承認依頼メッセージ本文'],
    ['O', '共同編集者', '共同編集者メールアドレス（カンマ区切り、v1.5.0〜）'],
], col_widths=[1.5, 5.0, 9.5])

doc.add_heading('B.2 _項目設定 シート', level=2)
add_table(['列', 'ヘッダー名', '意味'], [
    ['A', 'ワークフローID', '所属するワークフローのID'],
    ['B', '項目名', 'フォーム項目の表示名'],
    ['C', '入力タイプ', 'text/textarea/dropdown/checkbox/date'],
    ['D', '選択肢', 'ドロップダウン/チェックボックスの選択肢（カンマ区切り）'],
    ['E', '必須', 'TRUE/FALSE'],
    ['F', '順番', '表示順'],
    ['G', 'フィールドID', '内部識別子（自動採番、変更不可）'],
    ['H', '削除フラグ', '論理削除フラグ（TRUE=削除済み）'],
    ['I', 'データ列', 'データシート上の対応列番号'],
], col_widths=[1.5, 4.0, 10.5])

doc.add_heading('B.3 _ボタン設定 シート', level=2)
add_table(['列', 'ヘッダー名', '意味'], [
    ['A', 'ワークフローID', '所属するワークフローのID'],
    ['B', 'ボタン名', '表示名（例: 承認、却下）'],
    ['C', 'ステータス値', '押下時に記録されるステータス'],
    ['D', '色', 'green/blue/yellow/orange/red/gray'],
    ['E', '順番', '表示順'],
    ['F', 'スレッド返信', 'TRUE/FALSE'],
    ['G', '返信メッセージ', '返信文本文（変数使用可）'],
    ['H', 'DM通知', '（未使用/将来拡張用）'],
    ['I', 'DMメッセージ', '（未使用/将来拡張用）'],
], col_widths=[1.5, 4.0, 10.5])

doc.add_heading('B.4 _管理者 シート', level=2)
add_table(['列', 'ヘッダー名', '意味'], [
    ['A', 'メールアドレス', '管理者のメール'],
    ['B', '同意日時', 'プライバシー通知への同意日時'],
    ['C', '権限', 'owner / editor / member（v1.5.0〜。空欄は owner 扱い）'],
], col_widths=[1.5, 4.0, 10.5])

doc.add_heading('B.5 _監査ログ シート', level=2)
add_table(['列', 'ヘッダー名', '意味'], [
    ['A', 'タイムスタンプ', '発生日時'],
    ['B', '実行者', '操作したユーザーのメール'],
    ['C', 'アクション', '操作種別（workflow.create 等）'],
    ['D', '対象', '対象ID等'],
    ['E', '詳細', '操作の詳細（最大500文字）'],
], col_widths=[1.5, 4.0, 10.5])

doc.add_heading('B.6 各ワークフローのデータシート', level=2)
add_para('ワークフローごとにそのワークフロー名のシートが作成されます。列構成は以下です。')
add_table(['列', 'ヘッダー名', '意味'], [
    ['A', 'タイムスタンプ', '申請日時'],
    ['B', '送信者', '申請者の名前'],
    ['C', '送信者メール', '申請者のメールアドレス'],
    ['D以降', '各フォーム項目', 'ワークフロー定義したフォーム項目'],
    ['末尾-1', 'ステータス', '申請・承認型のみ（承認/却下等）'],
    ['末尾', 'メッセージ名', 'Chat内メッセージの内部ID'],
], col_widths=[2.5, 4.5, 8.5])

add_warning('デフォルトのシステム列（タイムスタンプ、送信者、送信者メール、ステータス、メッセージ名）を手動で削除しないでください。削除するとデータ書き込み時にエラーになります。')

page_break()

# ================================================================
# 付録C ScriptProperties
# ================================================================
doc.add_heading('付録C ScriptProperties 必須キー一覧', level=1)
add_table(['キー名', '値の例・形式', '説明'], [
    ['SPREADSHEET_ID', '1iHLqk...（英数字44文字）', 'メインスプレッドシートのID'],
    ['FUNCTION_URL', 'https://script.google.com/a/macros/.../exec', '本番デプロイURL'],
    ['CONTACT_EMAIL', 'konosuke.okazaki@cmind-co.jp', '問い合わせ先メール'],
    ['ENV_NAME', 'prod', '環境名（管理画面表示用）'],
    ['SERVICE_ACCOUNT_KEY', '{ "type":"service_account",...} JSON文字列', 'サービスアカウントキー全文'],
    ['LAST_ERROR', '（自動記録、手動編集不要）', '直近エラーの記録'],
], col_widths=[4.5, 5.5, 6.0])

add_warning('LAST_ERROR 以外のキーが空になっているとシステムが動作しません。ScriptProperties を編集する前に、現在の値を必ずスクリーンショットまたはテキスト保存しておいてください。')

doc.add_heading('C.1 ScriptProperties の値をバックアップする方法', level=2)
add_bullet('Apps Script → プロジェクトの設定 → スクリプトプロパティ')
add_bullet('全ての行を目視でメモ、またはスクリーンショット保存')
add_bullet('SERVICE_ACCOUNT_KEY は JSON なのでテキストエディタに保存推奨')

add_warning('SERVICE_ACCOUNT_KEY のバックアップは秘密情報です。個人PCの平文ファイルに保存せず、社内の安全なパスワードマネージャに保管してください。')

page_break()

# ================================================================
# 付録D 引継ぎチェックリスト
# ================================================================
doc.add_heading('付録D 引継ぎ時チェックリスト', level=1)
add_para('新任担当者への引継ぎ時、以下を一緒に確認してください。')

doc.add_heading('D.1 アクセス確認', level=2)
add_checklist([
    'メインスプレッドシートを開けるか',
    'Apps Script エディタを開けるか',
    'ScriptProperties を閲覧できるか',
    '管理画面URLを開いて管理画面が表示されるか',
    'オーナー権限が付与されているか',
    'Google Cloud Console のプロジェクトにアクセスできるか',
    'バックアップフォルダにアクセスできるか',
    'GitHub リポジトリにアクセスできるか',
])

doc.add_heading('D.2 実操作演習', level=2)
add_checklist([
    'テスト用ワークフローを1つ作成できた',
    'テスト用ワークフローを Chat に「設置」できた',
    'Chat からテスト申請を送信し承認カードが表示された',
    'テスト用ワークフローをアーカイブし、復元できた',
    'テスト用ワークフローを完全削除できた',
    '管理者を追加し、権限を変更し、削除できた',
    'healthCheck を実行し ok:true を確認した',
    '監査ログで自分の操作が記録されているのを確認した',
])

doc.add_heading('D.3 知識確認', level=2)
add_checklist([
    '3段階の権限（オーナー/エディタ/メンバー）の違いを説明できる',
    '本番デプロイ更新手順（第4.3節）を説明できる',
    '「権限がありません」エラーの原因候補を3つ以上挙げられる',
    'サービスアカウントキー漏洩時の対応手順を説明できる',
    'バックアップからの復元手順を説明できる',
    '本マニュアルの各章の場所を把握している',
])

doc.add_heading('D.4 情報引継ぎ', level=2)
add_checklist([
    'ScriptProperties 全値のバックアップを受領した',
    'サービスアカウントキーの管理場所（パスワードマネージャ等）を共有された',
    '本マニュアルの最新版を受領した',
    '前任者への緊急連絡手段を確認した（引継ぎ後3ヶ月間）',
])

page_break()

# ================================================================
# 変更履歴
# ================================================================
doc.add_heading('変更履歴', level=1)
add_table(['日付', 'バージョン', '変更内容', '担当'], [
    ['2026-09-28', '1.0', '初版作成（対象システム v1.5.0）', '岡崎'],
    ['', '', '', ''],
    ['', '', '', ''],
    ['', '', '', ''],
], col_widths=[3.0, 2.0, 8.5, 2.5])

# ================================================================
# 保存
# ================================================================
out_path = 'C:\\Users\\user\\Documents\\Claude\\workflow-bot\\workflow-bot_運用マニュアル_v1.0.docx'
doc.save(out_path)
print('生成完了:', out_path)
