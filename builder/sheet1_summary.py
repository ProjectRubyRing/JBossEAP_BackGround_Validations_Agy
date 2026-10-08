# -*- coding: utf-8 -*-
"""
Sheet 1: 01_概要・全体サマリー
"""

import openpyxl
from builder.styles import *

def build_sheet_01(wb):
    ws = wb.create_sheet(title="01_概要・全体サマリー")
    setup_sheet_view(ws)

    # 列幅定義
    col_widths = {
        "A": 4,    # 余白
        "B": 24,   # カテゴリ / 項目名
        "C": 48,   # 主要内容 / 設定値 / 説明
        "D": 44,   # 影響・メリット・現場のポイント
        "E": 22,   # 参照先 / 備考
    }
    set_col_widths(ws, col_widths)

    # 1. タイトルバナー
    render_title_banner(
        ws,
        title_text="JBoss EAP データソース接続検証（background-validation-millis）完全技術解説書",
        subtitle_text="動作原理・歴史的背景・小学生向け直感解説から現場チューニング・障害対策まで網羅 (JBoss EAP 7.x / 8.x 対応)",
        max_col_letter="E"
    )

    row = 5

    # 2. エグゼクティブサマリー
    render_section_header(ws, row, "エグゼクティブサマリー（1分でわかる結論と重要性）", "E")
    row += 1

    summary_headers = ["重要ポイント", "技術的結論", "ビジネス・運用へのインパクト", "重要度"]
    render_table_headers(ws, row, summary_headers, "B")
    row += 1

    summary_data = [
        (
            "1. 解決する本質的課題",
            "ステートフルFWやクラウドNATによる「サイレント接続切断（ゾンビコネクション）」を、エンドユーザーがアクセスする前にバックグラウンドで事前検知・破棄・再生成する。",
            "アクセスが途切れた後の「朝イチ1発目の通信エラー」や「昼休み明けの画面フリーズ」をゼロに撲滅。",
            "★★★★★ (最重要)"
        ),
        (
            "2. 性能への影響 (vs 即時検証)",
            "「validate-on-match（貸出時直前検証）」と異なり、ユーザーリクエスト処理スレッドのレイテンシを0ms阻害しない。検証処理を非同期専用スレッドに完全分離する。",
            "毎秒数千〜数万リクエストの高負荷Webシステムでも、DBへの無駄なSQL ping負荷を発生させず安定稼働を実現。",
            "★★★★★ (極めて高い)"
        ),
        (
            "3. クラウド環境での必須性",
            "AWS NAT Gateway(350秒)やAzure Load Balancer(4分)など、クラウド機器の短いアイドルタイムアウトよりも短い間隔で定期疎通し、NATセッションテーブルを延命（Keep-Alive）する。",
            "クラウド（AWS/Azure/GCP/OpenShift）移行時におけるコネクション切断障害の最大の防壁となる。",
            "★★★★★ (必須級)"
        ),
        (
            "4. 黄金のチューニング設計",
            "「background-validation-millis ≦ ファイアウォール/NATタイムアウト ÷ 2」が設計の鉄則。かつ idle-timeout-minutes との周期衝突を避ける。",
            "設計ミスによるDB過負荷や、検証周期のすり抜けによる障害を確実に防止できる。",
            "★★★★☆ (推奨)"
        ),
    ]

    for item in summary_data:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_LEFT
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
        ws[f"E{row}"].font = FONT_BADGE_REC if "★★★★★" in item[3] else FONT_BADGE_INFO
        ws[f"E{row}"].fill = FILL_BADGE_REC if "★★★★★" in item[3] else FILL_BADGE_INFO
        ws[f"E{row}"].alignment = ALIGN_CENTER
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=24)
        row += 1

    row += 2

    # 3. 主要パラメータ一元クイックリファレンス
    render_section_header(ws, row, "接続検証コアパラメータ クイックリファレンス", "E")
    row += 1

    param_headers = ["パラメータ名", "デフォルト値", "本番環境 推奨設定", "設定の要点・メカニズム"]
    render_table_headers(ws, row, param_headers, "B")
    row += 1

    param_data = [
        (
            "background-validation",
            "false",
            "true (有効化)",
            "バックグラウンド検証スレッド(ConnectionValidator)を起動するかどうか。validate-on-matchとは排他的に使用する。"
        ),
        (
            "background-validation-millis",
            "0 (無効)",
            "120000〜300000 (2〜5分)\n※NWタイムアウトの半分以下",
            "アイドル状態の接続を定期点検する間隔（ミリ秒）。AWS NAT GW(350s)環境なら120000ms(2分)が定番。"
        ),
        (
            "validate-on-match",
            "false",
            "false (原則無効)\n※バッチ系のみ検討可",
            "接続プールから取得する瞬間に毎回検証する。安全だが毎リクエストに遅延が発生しDB負荷が激増するため通常はfalse。"
        ),
        (
            "use-fast-fail",
            "false",
            "true (推奨)",
            "直前の検証で無効と判定された接続があった場合、以降のアイドル接続検証を高速にスキップしてプール全体を迅速クリーンアップ。"
        ),
        (
            "valid-connection-checker\n-class-name",
            "未設定",
            "JDBC4ValidConnectionChecker\nまたは DBMS専用クラス",
            "JDBC 4.0標準の Connection.isValid() を呼び出し、重いSQLパケットではなく軽量RPC/pingでDBと疎通確認を行う。"
        ),
        (
            "check-valid-connection-sql",
            "未設定",
            "未設定 (checker推奨)\n※旧環境のみ SELECT 1 等",
            "検証用の生SQL。JDBC4チェッカーが使えないレガシー環境でのみ使用。パケットサイズ・パース負荷が高いため避ける。"
        ),
        (
            "flush-strategy",
            "FailingConnectionOnly",
            "FailingConnectionOnly\nまたは IdleConnections",
            "無効接続検知時にプールをどうクリアするか。FailingConnectionOnlyは壊れた接続のみ破棄し正常接続を維持。"
        ),
    ]

    for item in param_data:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_CODE_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_LEFT
        ws[f"B{row}"].border = BORDER_THIN

        ws[f"C{row}"] = item[1]
        ws[f"C{row}"].font = FONT_BODY_MUTED
        ws[f"C{row}"].fill = fill
        ws[f"C{row}"].alignment = ALIGN_CENTER
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

        auto_fit_row_height(ws, row, col_widths, min_height=24)
        row += 1

    row += 2

    # 4. 本書の構成とシートナビゲーション
    render_section_header(ws, row, "本書の構成とシート別ナビゲーションガイド", "E")
    row += 1

    nav_headers = ["シート番号・シート名", "主な対象読者", "主な解説内容", "必読トピック"]
    render_table_headers(ws, row, nav_headers, "B")
    row += 1

    nav_data = [
        (
            "02_小学生でもわかる解説",
            "初学者・非エンジニア・\n直感的に概念を掴みたい方",
            "テーマパークのゴーカートと裏方整備士に例えたストーリー。\nなぜ車庫で待っている間にパンクするのか、直前検査と定期パトロールの違い。",
            "・見えないイタズラ小僧（FW）の正体\n・パトロール間隔が短すぎ/長すぎの悲劇"
        ),
        (
            "03_歴史と背景・必要性の経緯",
            "インフラエンジニア・\nアーキテクト・設計者",
            "J2EE黎明期からEAP 8(Jakarta EE 10)までの進化の歴史。\n「SELECT 1」のDB枯渇問題からIronJacamarでのミリ秒制御導入までの経緯。",
            "・DB CPU 50%消費問題の悪夢\n・なぜ「分(minutes)」から「ミリ秒」になったか"
        ),
        (
            "04_動作原理と内部機構",
            "シニアエンジニア・\n性能チューナー・開発者",
            "IronJacamarのConnectionValidator内部スレッド、アイドル接続走査、\nステートフルFWのConntrackサイレント削除とTCPハーフオープンの物理挙動。",
            "・なぜ貸出中接続は検査しないのか\n・TCP RSTが飛ばないゾンビ化の真相"
        ),
        (
            "05_validate-on-match徹底比較",
            "システム設計者・\n性能試験担当者",
            "background-validation と validate-on-match の多角的一騎打ち比較。\n両方有効化の二重負荷、両方無効化の事故などアンチパターンを詳解。",
            "・レイテンシ vs 安全性のトレードオフ\n・採用判断意思決定フローチャート"
        ),
        (
            "06_設定リファレンスとCLI",
            "運用管理者・環境構築担当・\nインフラ担当者",
            "10種類以上の関連パラメータ完全定義、DBMS別推奨XML（Oracle, PostgreSQL,\nMySQL, SQL Server）、JBoss CLIによる設定変更・テスト・フラッシュ手順。",
            "・コピペで使える本番用XML設定例\n・即時反映CLIコマンド集"
        ),
        (
            "07_障害対策とベストプラクティス",
            "運用保守・SRE・トラブル\nシューティング担当者",
            "パラメータ設計の黄金律、クラウド別（AWS/Azure/GCP）タイムアウト値、\n月曜朝イチの切断エラーやDB再起動時の障害対処法、TRACEログ採取手順。",
            "・「FWタイムアウト ÷ 2」の黄金律\n・IJ000457エラー大量発生時の処方箋"
        ),
    ]

    for item in nav_data:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_LEFT
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
        ws[f"E{row}"].font = FONT_BODY_MUTED
        ws[f"E{row}"].fill = fill
        ws[f"E{row}"].alignment = ALIGN_LEFT
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    row += 2

    # 5. 対象環境情報
    render_section_header(ws, row, "対象製品・アーキテクチャ スコープ", "E")
    row += 1

    scope_headers = ["対象項目", "対応バージョン・仕様", "備考・前提条件"]
    render_table_headers(ws, row, scope_headers, "B")
    row += 1

    scope_data = [
        ("対象アプリケーションサーバ", "JBoss EAP 7.0 〜 7.4 / JBoss EAP 8.0 / WildFly 26〜31+", "IronJacamar JCAコンテナ搭載環境"),
        ("Java / Jakarta EE 規格", "Java EE 7 / Java EE 8 / Jakarta EE 10", "EAP 8 は Jakarta EE 10 / Java 17・21 推奨"),
        ("JDBC ドライバ規格", "JDBC 4.0 / 4.1 / 4.2 / 4.3 準拠ドライバ", "Connection.isValid(int timeout) API を利用"),
        ("対象データベース", "Oracle 19c/23ai, PostgreSQL 13〜16, MySQL 8.0, SQL Server 2019/2022 等", "主要エンタープライズDBMS全般"),
        ("インフラ環境", "AWS (EC2/ECS/EKS/RDS), Azure, GCP, Red Hat OpenShift, オンプレミス", "ステートフルFW・クラウドNAT配下環境"),
    ]

    for item in scope_data:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_LEFT
        ws[f"B{row}"].border = BORDER_THIN

        ws[f"C{row}"] = item[1]
        ws[f"C{row}"].font = FONT_BODY
        ws[f"C{row}"].fill = fill
        ws[f"C{row}"].alignment = ALIGN_LEFT
        ws[f"C{row}"].border = BORDER_THIN

        ws[f"D{row}"] = item[2]
        ws[f"D{row}"].font = FONT_BODY_MUTED
        ws[f"D{row}"].fill = fill
        ws[f"D{row}"].alignment = ALIGN_LEFT
        ws[f"D{row}"].border = BORDER_THIN

        # E列はマージまたは余白
        ws[f"E{row}"] = "検証済み"
        ws[f"E{row}"].font = FONT_BADGE_REC
        ws[f"E{row}"].fill = FILL_BADGE_REC
        ws[f"E{row}"].alignment = ALIGN_CENTER
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=20)
        row += 1

    return ws
