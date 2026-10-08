# -*- coding: utf-8 -*-
"""
Sheet 5: 05_validate-on-match徹底比較
"""

import openpyxl
from builder.styles import *

def build_sheet_05(wb):
    ws = wb.create_sheet(title="05_validate-on-match徹底比較")
    setup_sheet_view(ws)

    # 列幅定義
    col_widths = {
        "A": 4,    # 余白
        "B": 22,   # 評価項目
        "C": 36,   # background-validation (定期パトロール)
        "D": 36,   # validate-on-match (貸出時直前検査)
        "E": 38,   # トレードオフ・現場の技術的視点
        "F": 16,   # 推奨判定
    }
    set_col_widths(ws, col_widths)

    # 1. タイトルバナー
    render_title_banner(
        ws,
        title_text="background-validation vs validate-on-match 徹底比較マトリクス",
        subtitle_text="定期パトロール型 vs 貸出時直前検査型 の多角的トレードオフ、意思決定フロー、および3大アンチパターン",
        max_col_letter="F"
    )

    row = 5

    # 2. 一騎打ち比較表
    render_section_header(ws, row, "1. 2大検証方式の多角的徹底比較マトリクス", "F")
    row += 1

    compare_headers = ["比較評価項目", "background-validation (定期パトロール)", "validate-on-match (貸出時直前検査)", "詳細なトレードオフ・現場の視点", "判定"]
    render_table_headers(ws, row, compare_headers, "B")
    row += 1

    comparison_data = [
        (
            "検証の実行タイミング",
            "非同期（裏方スレッド）\nbackground-validation-millis の周期ごとにアイドル接続を点検。",
            "同期（リクエスト直前）\nアプリが getConnection() を呼んだ瞬間にプール内で点検。",
            "非同期はリクエスト処理から完全に切り離されるのに対し、同期はリクエスト処理のクリティカルパス上に乗る。",
            "設計思想の違い"
        ),
        (
            "アプリの応答時間\n(Latencyへの影響)",
            "【影響ゼロ (0 ms)】\n検証は裏で行われるため、getConnection() の貸出速度は最速。",
            "【遅延増加 (+1〜50ms)】\n毎回DBへ往復通信(RTT)が発生。クラウドやマルチAZ等でRTTが大きいと致命傷に。",
            "低レイテンシが必須のWebサービス・API基盤では background-validation が圧倒的に有利。",
            "バックグラウンド\n完全勝利"
        ),
        (
            "データベースへの負荷\n(DBクエリ/ping数)",
            "【一定・予測可能】\nプール内の待機接続数と周期（例:3分）だけで決まり、アクセス増でも負荷は不変。",
            "【アクセス数に完全比例】\n毎秒1万リクエストあれば毎秒1万回のpingがDBに殺到し、DB CPUを圧迫。",
            "アクセス急増時にDBがヘルスチェックのパケットでパンクするリスクを完全に排除できる。",
            "バックグラウンド\n完全勝利"
        ),
        (
            "ゾンビ接続の検知精度\n(すり抜けリスク)",
            "【極小のウィンドウリスクあり】\n点検直後に切断され、次の点検が来るまでのわずかな隙間に借りるとエラーの可能性。",
            "【100%安全 (リスクゼロ)】\n貸し出すまさにその瞬間に点検するため、死んだ接続をアプリに渡すことは絶対にない。",
            "1回のエラーも絶対に許されず、リトライもできない特殊なトランザクションでは validate-on-match に軍配。",
            "オンマッチ\n優位"
        ),
        (
            "FW / NATテーブルの\n維持効果 (Keep-Alive)",
            "【非常に高い】\nアクセスが全くない深夜やアイドル時でも定期的にパケットが流れるため、FWの切断を防ぐ。",
            "【効果なし】\nユーザーアクセスがない時間帯は通信が完全にゼロになるため、FWのタイムアウトで切断される。",
            "クラウド（AWS NAT GW等）での「朝イチ切断障害」を防げるのは background-validation のみ。",
            "バックグラウンド\n完全勝利"
        ),
        (
            "DBフェイルオーバー時の\n挙動と復旧速度",
            "設定したミリ秒間隔で順次ゾンビが検出・破棄され、新品接続が補充されて自律的に正常復帰する。",
            "フェイルオーバー直後の最初のリクエスト群で異常が検知され、その場で破棄・再接続を試行する。",
            "バックグラウンド検証の場合、周期を短め（例:15〜30秒）にしておけば、ユーザーが来る前に復旧が完了する。",
            "引き分け\n(設定次第)"
        ),
        (
            "推奨ユースケース",
            "・高スループットWebシステム\n・低遅延が求められるマイクロサービス/API\n・クラウド(AWS/Azure/GCP)上の全システム",
            "・夜間バッチ処理専用データソース\n・アクセス頻度が低く、絶対失敗できない業務\n・オンプレミスでRTTが0.1ms未満の近接環境",
            "現代のエンタープライズWebアーキテクチャでは 95% 以上のケースで background-validation が選定される。",
            "用途別の棲み分け"
        ),
    ]

    for item in comparison_data:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_CENTER
        ws[f"B{row}"].border = BORDER_THIN

        ws[f"C{row}"] = item[1]
        ws[f"C{row}"].font = FONT_BODY
        ws[f"C{row}"].fill = fill
        ws[f"C{row}"].alignment = ALIGN_LEFT
        ws[f"C{row}"].border = BORDER_THIN

        ws[f"D{row}"] = item[2]
        ws[f"D{row}"].font = FONT_BODY
        ws[f"D{row}"].fill = fill
        ws[f"D{row}"].alignment = ALIGN_LEFT
        ws[f"D{row}"].border = BORDER_THIN

        ws[f"E{row}"] = item[3]
        ws[f"E{row}"].font = FONT_BODY
        ws[f"E{row}"].fill = fill
        ws[f"E{row}"].alignment = ALIGN_LEFT
        ws[f"E{row}"].border = BORDER_THIN

        ws[f"F{row}"] = item[4]
        ws[f"F{row}"].font = FONT_BADGE_REC if "完全勝利" in item[4] else FONT_BADGE_INFO
        ws[f"F{row}"].fill = FILL_BADGE_REC if "完全勝利" in item[4] else FILL_BADGE_INFO
        ws[f"F{row}"].alignment = ALIGN_CENTER
        ws[f"F{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    row += 2

    # 3. 意思決定フローチャート
    render_section_header(ws, row, "2. 方式選定の意思決定フローチャート（どちらを選ぶべきか？）", "F")
    row += 1

    flow_headers = ["判定ステップ", "条件分岐（質問）", "判断基準・Yesの場合の推奨", "判断基準・Noの場合の推奨"]
    render_table_headers(ws, row, flow_headers, "B")
    row += 1

    flows = [
        (
            "Step 1: レイテンシ要求",
            "システムのレスポンス速度（数ミリ秒の遅延短縮）や、高いスループット（秒間数百リクエスト以上）が要求されるか？",
            "【Yes】→ 迷わず `background-validation` を採用！\nリクエスト処理に余計なping往復を一切挟まない構成にする。",
            "【No】→ Step 2へ進む。"
        ),
        (
            "Step 2: インフラ環境",
            "システムはAWS, Azure, GCPなどのパブリッククラウド、または途中にステートフルFWが存在する環境か？",
            "【Yes】→ `background-validation` を強く推奨！\n無通信時のNAT/FWサイレントドロップを防ぐKeep-Alive効果が必須となるため。",
            "【No】→ Step 3へ進む。"
        ),
        (
            "Step 3: 処理の性質",
            "夜間に単発で動くバッチ処理等で、アクセス頻度が極めて低く、万一の切断時でも絶対にリトライ不可なクリティカル処理か？",
            "【Yes】→ `validate-on-match` の採用を検討可能。\nリクエストごとのpingオーバーヘッドが許容され、確実性が最優先されるため。",
            "【No】→ 汎用的な標準構成として `background-validation` を採用。"
        ),
    ]

    for item in flows:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_CENTER
        ws[f"B{row}"].border = BORDER_THIN

        ws[f"C{row}"] = item[1]
        ws[f"C{row}"].font = FONT_BODY
        ws[f"C{row}"].fill = fill
        ws[f"C{row}"].alignment = ALIGN_LEFT
        ws[f"C{row}"].border = BORDER_THIN

        ws[f"D{row}"] = item[2]
        ws[f"D{row}"].font = FONT_BODY_BOLD
        ws[f"D{row}"].fill = fill
        ws[f"D{row}"].alignment = ALIGN_LEFT
        ws[f"D{row}"].border = BORDER_THIN

        ws[f"E{row}"] = item[3]
        ws[f"E{row}"].font = FONT_BODY
        ws[f"E{row}"].fill = fill
        ws[f"E{row}"].alignment = ALIGN_LEFT
        ws[f"E{row}"].border = BORDER_THIN

        # F列をマージ
        ws[f"F{row}"] = "判定"
        ws[f"F{row}"].font = FONT_BADGE_INFO
        ws[f"F{row}"].fill = FILL_BADGE_INFO
        ws[f"F{row}"].alignment = ALIGN_CENTER
        ws[f"F{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    row += 2

    # 4. やってはいけない！3大アンチパターン
    render_section_header(ws, row, "3. 設計者が陥りがちな【3大アンチパターン】と重大な副作用", "F")
    row += 1

    anti_headers = ["アンチパターン名", "実際の設定内容", "引き起こされる重大な副作用・障害", "正しい処方箋"]
    render_table_headers(ws, row, anti_headers, "B")
    row += 1

    anti_patterns = [
        (
            "アンチパターン1：\n【両方trueの二重検証】",
            "<validate-on-match>true</validate-on-match>\n<background-validation>true</background-validation>",
            "【最悪の二重オーバーヘッド】\n貸出時に毎回pingが飛ぶためリクエスト遅延が発生する上、裏でもバックグラウンドスレッドが定期巡回してDBにpingを投げ続ける。\n何一つメリットがなく、ただDBとネットワークに無駄な負荷をかけるだけになる。",
            "どちらか一方のみを true にする！\n通常は background-validation=true、validate-on-match=false を徹底。"
        ),
        (
            "アンチパターン2：\n【両方falseの完全放置】",
            "<validate-on-match>false</validate-on-match>\n<background-validation>false</background-validation>",
            "【エンドユーザー直撃爆弾】\n接続の健全性確認が一切行われない。\nFWの無通信切断やDB再起動が発生すると、プール内のすべての接続がゾンビ化し、次にアクセスしたユーザー全員の画面にシステムエラーが噴出する！",
            "本番運用環境では必ず background-validation を有効化し、適切な millis を設定すること！"
        ),
        (
            "アンチパターン3：\n【FWタイムアウトより\n長いmillis設定】",
            "FW切断時間: 15分 (900,000ms)\n設定値: 30分 (1,800,000ms)",
            "【検知の遅延とゾンビ放置】\nFWが15分でセッションを削除したあと、最長15分間もゾンビ接続がプール内に放置される。\nその間にリクエストが来れば確実に通信エラーとなり、Keep-Alive効果も一切得られない。",
            "background-validation-millis は、必ず「FWタイムアウトの半分以下（例：5分＝300,000ms）」に設定すること！"
        ),
    ]

    for item in anti_patterns:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BADGE_DANGER
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_CENTER
        ws[f"B{row}"].border = BORDER_THIN

        ws[f"C{row}"] = item[1]
        ws[f"C{row}"].font = FONT_CODE
        ws[f"C{row}"].fill = fill
        ws[f"C{row}"].alignment = ALIGN_LEFT
        ws[f"C{row}"].border = BORDER_THIN

        ws[f"D{row}"] = item[2]
        ws[f"D{row}"].font = FONT_BODY
        ws[f"D{row}"].fill = fill
        ws[f"D{row}"].alignment = ALIGN_LEFT
        ws[f"D{row}"].border = BORDER_THIN

        ws.merge_cells(f"E{row}:F{row}")
        e_cell = ws[f"E{row}"]
        e_cell.value = item[3]
        e_cell.font = FONT_BODY_BOLD
        e_cell.fill = fill
        e_cell.alignment = ALIGN_LEFT
        e_cell.border = BORDER_THIN
        ws[f"F{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=28)
        row += 1

    return ws
