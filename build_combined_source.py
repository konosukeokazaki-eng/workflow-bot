# -*- coding: utf-8 -*-
"""
GAS全ソースファイルを1つのテキストファイルにまとめる（NotebookLM 用）
"""
import os

BASE = r'C:\Users\user\Documents\Claude\workflow-bot'
OUT = os.path.join(BASE, 'workflow-bot_全ソース_v1.5.1.txt')

# 論理順で並べる（読み手が上から順に理解しやすい順）
FILES = [
    ('appsscript.json',  '設定',    'Apps Scriptプロジェクト設定（スコープ・実行権限等）'),
    ('Main.js',          'GAS',     'エントリポイント・定数・doGet/doPost・ルーティング'),
    ('Auth.js',          'GAS',     '認証・権限判定・トークン・監査ログ・バリデーション'),
    ('Data.js',          'GAS',     'スプレッドシート読み書きヘルパー'),
    ('Schema.js',        'GAS',     'シート列構造の自動修復ヘルパー'),
    ('Chat.js',          'GAS',     'Chat連携・申請カード投稿・承認処理'),
    ('Adminapi.js',      'GAS',     '管理画面のサーバー側API群'),
    ('Backup.js',        'GAS',     '自動バックアップ処理'),
    ('Monitor.js',       'GAS',     '定期実行監視処理'),
    ('admin.html',       'HTML',    '管理画面のHTMLテンプレート'),
    ('JavaScript.html',  'HTML',    '管理画面フロントエンドJS（ブラウザ側）'),
    ('Stylesheet.html',  'HTML',    '管理画面スタイル定義（CSS）'),
]

def sep(char='=', n=80):
    return char * n

with open(OUT, 'w', encoding='utf-8', newline='\n') as out:
    # 全体ヘッダー
    out.write(sep('=') + '\n')
    out.write('workflow-bot 全ソースコード統合ファイル\n')
    out.write('対象システムバージョン: v1.5.1\n')
    out.write('生成日: 2026-09-29\n')
    out.write('用途: NotebookLM 等の AI ヘルプデスク用ソース\n')
    out.write(sep('=') + '\n\n')

    # ファイル一覧
    out.write('【ファイル一覧】\n')
    for i, (fname, ftype, desc) in enumerate(FILES, 1):
        out.write(f'  {i:2d}. {fname}  ({ftype})  - {desc}\n')
    out.write('\n')
    out.write(sep('=') + '\n\n')

    # 各ファイルの中身
    for i, (fname, ftype, desc) in enumerate(FILES, 1):
        path = os.path.join(BASE, fname)
        if not os.path.exists(path):
            print(f'SKIP (not found): {fname}')
            continue

        out.write('\n')
        out.write(sep('#') + '\n')
        out.write(f'# ファイル {i}/{len(FILES)}: {fname}\n')
        out.write(f'# 種別: {ftype}\n')
        out.write(f'# 役割: {desc}\n')
        out.write(sep('#') + '\n\n')

        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()

        out.write(content)
        if not content.endswith('\n'):
            out.write('\n')
        out.write('\n')
        out.write(f'--- {fname} ここまで ---\n')

    # 末尾
    out.write('\n' + sep('=') + '\n')
    out.write('全ソースファイルここまで\n')
    out.write(sep('=') + '\n')

size = os.path.getsize(OUT)
with open(OUT, 'r', encoding='utf-8') as f:
    lines = sum(1 for _ in f)
print(f'生成完了: {OUT}')
print(f'サイズ: {size:,} バイト ({size/1024:.1f} KB)')
print(f'行数: {lines:,}')
