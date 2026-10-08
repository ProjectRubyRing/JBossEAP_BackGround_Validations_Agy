# -*- coding: utf-8 -*-
"""
Sheet 7: 07_障害対策とベストプラクティス
"""

import openpyxl
from builder.styles import *

def build_sheet_07(wb):
    ws = wb.create_sheet(title="07_障害対策とベストプラクティス")
    setup_sheet_view(ws)

    # 列幅定義
    col_widths = {
        "A": 4,    # 余白
        "B": 24,   # 分類 / 事象名
        "C": 42,   # 発生シナリオ / 症状・エラーログ
        "D": 42,   # 根本原因 (Root Cause)
        "E": 44,   # 具体的対策・チューニング処方箋
    }
    set_col_widths(ws, col_widths)

    # 1. タイトルバナー
    render_title_banner(
        ws,
        title_text="現場で役立つチューニングの黄金律・トラブルシューティング",
        subtitle_text="FWタイムアウト計算式、主要クラウド別推奨値、5大代表的障害事象の処方箋、TRACEログ読解",
        max_col_letter="E"
    )

    row = 5

    # 2. パラメータ設計の黄金律
    render_section_header(ws, row, "1. パラメータ設計の黄金律と主要クラウド機器別 推奨値テーブル", "E")
    row += 1

    rule_headers = ["対象インフラ・ネットワーク機器", "機器側のアイドル切断タイムアウト", "background-validation-millis 推奨値", "設計根拠・マージン考慮"]
    render_table_headers(ws, row, rule_headers, "B")
    row += 1

    rules = [
        (
            "AWS NAT Gateway\n(Amazon VPC)",
            "350 秒 (約5分50秒)\n※変更不可の固定値",
            "120,000 ms (2分)\n〜 150,000 ms (2.5分)",
            "タイムアウト(350s)の半分以下を徹底。\n2分間隔で定期疎通することでConntrackテーブルが常時延命され、NAT切断を100%回避。"
        ),
        (
            "Azure Load Balancer\n(Microsoft Azure)",
            "240 秒 (4分)\n※デフォルト値 (変更可)",
            "90,000 ms (1.5分)\n〜 120,000 ms (2分)",
            "デフォルトの4分切断に対して十分な安全マージンを確保。\nAzure側のTCP Keepalive設定と併用するとさらに堅牢。"
        ),
        (
            "GCP Cloud NAT\n(Google Cloud)",
            "1,200 秒 (20分)\n※デフォルト値",
            "300,000 ms (5分)",
            "GCPは比較的長いが、5分間隔で疎通させておくことで夜間のサイレント削除を確実に防止可能。"
        ),
        (
            "オンプレミス FW / UTM\n(Cisco ASA / Palo Alto / Fortinet)",
            "900 秒 (15分)\n〜 3,600 秒 (60分)",
            "300,000 ms (5分)\n〜 600,000 ms (10分)",
            "企業の標準的FW設定（15分〜30分）に対して、1/3〜1/2の周期で設定。\nDB側のリソース消費を最小化しつつ健全性を維持。"
        ),
        (
            "MySQL / MariaDB\n(wait_timeout)",
            "28,800 秒 (8時間)\n※厳しい環境では 600秒 (10分)",
            "180,000 ms (3分)\n〜 300,000 ms (5分)",
            "DBMS側でアイドル接続を切断する設定が入っている場合、その値よりも短い間隔で設定して切断を防ぐ。"
        ),
    ]

    for item in rules:
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
        ws[f"D{row}"].font = FONT_CODE_BOLD
        ws[f"D{row}"].fill = fill
        ws[f"D{row}"].alignment = ALIGN_CENTER
        ws[f"D{row}"].border = BORDER_THIN

        ws[f"E{row}"] = item[3]
        ws[f"E{row}"].font = FONT_BODY
        ws[f"E{row}"].fill = fill
        ws[f"E{row}"].alignment = ALIGN_LEFT
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    row += 2

    # 3. 現場の5大トラブルシューティング
    render_section_header(ws, row, "2. 現場で遭遇する5大トラブルシューティング マトリクス", "E")
    row += 1

    trouble_headers = ["代表的事象", "発生シナリオ・エラーログの症状", "根本原因 (Root Cause)", "現場での恒久対策・処方箋"]
    render_table_headers(ws, row, trouble_headers, "B")
    row += 1

    troubles = [
        (
            "事象1：月曜朝イチの\n1発目エラー障害",
            "【症状】\n月曜日の始業直後や連休明け、最初のユーザーがアクセスした瞬間に Communications link failure や Connection reset by peer が発生する。\n2回目以降は正常に動作する。",
            "【原因】\n土日にアクセスがなく、FWやクラウドNATのアイドルタイムアウトでセッションがサイレント破棄されていた。\nJBoss側で background-validation が未設定（無効）だったため、死んだ接続を最初のユーザーに渡してしまった。",
            "【処方箋】\n1. background-validation=true に設定。\n2. background-validation-millis を FW切断時間の半分以下（例: 120000ms）に設定。\n3. 定期通信によりセッションテーブルが常時維持され、月曜朝イチエラーが完全消滅する。"
        ),
        (
            "事象2：JBossログに\nIJ000457が大量発生",
            "【症状】\nserver.log に警告ログが定期的に大量出力される：\nWARN [org.jboss.jca.core.connectionmanager.pool.validator.ConnectionValidator] IJ000457: Destroying connection that is not valid",
            "【原因】\nbackground-validation-millis の周期がFWやDBの切断時間よりも長いため、点検のたびに「既に切断された接続」が発見されて破棄されている状態。\n切断自体は防げておらず、点検の間隔が長すぎる。",
            "【処方箋】\n1. background-validation-millis を短縮する（例: 600000ms → 120000ms）。\n2. タイムアウトが来る前に点検させることで、そもそも「接続が無効化する前」に通信を発生させ、破棄ログの発生を抑制する。"
        ),
        (
            "事象3：DB再起動・フェイル\nオーバー後の接続全滅",
            "【症状】\n夜間パッチ当て等でDBを再起動、またはAuroraがマルチAZフェイルオーバーした後、JBossのプール内の接続が全滅し、長時間業務エラーが継続する。",
            "【原因】\nDB再起動により古い接続がすべて切断されたが、JBoss側で use-fast-fail や exception-sorter が未設定のため、死んだ接続を1本ずつタイムアウトまで待ちながら破棄しており復旧が著しく遅延した。",
            "【処方箋】\n1. use-fast-fail=true を設定（無効接続検知時に即座に後続をスキップ）。\n2. 適切な exception-sorter を設定。\n3. flush-strategy を FailingConnectionOnly に設定。\n4. 運用手順に JBoss CLI での flush-idle-connection-in-pool 実行を組み込む。"
        ),
        (
            "事象4：DB側のCPU高騰と\nスロークエリログ溢れ",
            "【症状】\nDB管理部門（DBA）から「JBossから毎秒数千回の SELECT 1 が飛んできており、DBのCPUが逼迫している」とクレームが入る。",
            "【原因】\nデータソースで validate-on-match=true が設定されており、さらに check-valid-connection-sql に生SQLが指定されていたため、全リクエストでSQLが実行されていた。",
            "【処方箋】\n1. validate-on-match=false に変更。\n2. background-validation=true に切り替え。\n3. check-valid-connection-sql を削除し、JDBC4ValidConnectionChecker（または専用Checker）に変更してSQLパース負荷をゼロにする。"
        ),
        (
            "事象5：検証処理ハングによる\nリクエストスレッド枯渇",
            "【症状】\nネットワーク瞬断時、JBossのHTTPリクエスト処理スレッドが次々とブロックされ、新規リクエストを受け付けなくなってサーバーが完全にハングアップする。",
            "【原因】\nblocking-timeout-wait-millis がデフォルトの30秒など長すぎる値になっており、DB不通時にアプリスレッドが接続待ちキューで詰まったカスケード障害。",
            "【処方箋】\n1. blocking-timeout-wait-millis を 5000〜10000ms (5〜10秒) に短縮。\n2. クイックフェイルさせ、アプリ側で速やかにエラーハンドリング（またはサーキットブレーカー発動）を行わせる。"
        ),
    ]

    for item in troubles:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BADGE_DANGER if "事象" in item[0] else FONT_BODY_BOLD
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
        ws[f"E{row}"].font = FONT_BODY_BOLD
        ws[f"E{row}"].fill = fill
        ws[f"E{row}"].alignment = ALIGN_LEFT
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=32)
        row += 1

    row += 2

    # 4. TRACEログ採取手順
    render_section_header(ws, row, "3. トラブル時の詳細TRACEログ採取手順とログ読解ガイド", "E")
    row += 1

    trace_headers = ["手順・フェーズ", "JBoss CLI コマンド / 操作", "確認ポイント・ログ出力の見方"]
    render_table_headers(ws, row, trace_headers, "B")
    row += 1

    traces = [
        (
            "1. TRACEロガーの有効化\n(一時的な調査時)",
            "/subsystem=logging/logger=org.jboss.jca.core.connectionmanager.pool.validator:add(level=TRACE)",
            "ConnectionValidator がどのプールを何ミリ秒間隔で点検しているか、各接続の検査結果（Valid / Invalid）が詳細に出力されるようになります。"
        ),
        (
            "2. 出力ログの確認\n(server.log)",
            "grep \"ConnectionValidator\" server.log\n\n[TRACE] Checking pool: MyDataSource\n[TRACE] Validating connection: ManagedConnection@12345\n[WARN] IJ000457: Destroying connection that is not valid",
            "・定期的に Checking pool が発火しているか\n・検査間隔が想定通りか\n・Destroying connection が出た際の時刻と直前の通信遮断の相関を確認。"
        ),
        (
            "3. 調査完了後のロガー削除\n(元に戻す)",
            "/subsystem=logging/logger=org.jboss.jca.core.connectionmanager.pool.validator:remove",
            "調査終了後は必ずTRACEロガーを削除してください。\n放置すると大量のログ出力によりディスク容量を圧迫し、I/O負荷の原因になります。"
        ),
    ]

    for item in traces:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_LEFT
        ws[f"B{row}"].border = BORDER_THIN

        ws[f"C{row}"] = item[1]
        ws[f"C{row}"].font = FONT_CODE
        ws[f"C{row}"].fill = FILL_CODE_BLOCK
        ws[f"C{row}"].alignment = ALIGN_CODE
        ws[f"C{row}"].border = BORDER_THIN

        # D列とE列をマージ
        ws.merge_cells(f"D{row}:E{row}")
        d_cell = ws[f"D{row}"]
        d_cell.value = item[2]
        d_cell.font = FONT_BODY
        d_cell.fill = fill
        d_cell.alignment = ALIGN_LEFT
        d_cell.border = BORDER_THIN
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    return ws
