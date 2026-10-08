# -*- coding: utf-8 -*-
"""
Sheet 3: 03_歴史と背景・必要性の経緯
"""

import openpyxl
from builder.styles import *

def build_sheet_03(wb):
    ws = wb.create_sheet(title="03_歴史と背景・必要性の経緯")
    setup_sheet_view(ws)

    # 列幅定義
    col_widths = {
        "A": 4,    # 余白
        "B": 18,   # 時代・年代
        "C": 22,   # JBoss/EAPバージョン
        "D": 44,   # 直面したシステム課題・現場の悲鳴
        "E": 44,   # コネクション検証技術の進化と解決策
        "F": 22,   # 技術的エポック・意義
    }
    set_col_widths(ws, col_widths)

    # 1. タイトルバナー
    render_title_banner(
        ws,
        title_text="JBoss EAPにおけるデータソース接続検証の歴史的変遷と必然性",
        subtitle_text="J2EE黎明期のDB CPU枯渇の悲劇から、IronJacamarによるミリ秒精密制御、そして現代クラウドのNAT切断問題まで",
        max_col_letter="F"
    )

    row = 5

    # 2. 歴史年表
    render_section_header(ws, row, "コネクション検証技術の進化年表（黎明期からEAP 8最新アーキテクチャまで）", "F")
    row += 1

    history_headers = ["時代・年代", "JBoss製品・アーキテクチャ", "当時のシステム環境と直面した問題・課題", "コネクション検証技術の進化・解決策", "技術的エポック"]
    render_table_headers(ws, row, history_headers, "B")
    row += 1

    history_data = [
        (
            "1990年代末〜2000年代初頭\n【黎明期】",
            "JBoss 2.x / 3.x\n(J2EE 1.2 / 1.3)",
            "・Webアクセスのたびに新規TCP接続とDB認証を行うため、性能が極めて低速。\n・コネクションプール技術が導入されたが、DB再起動やNW瞬断時にプール内の接続がすべて「死に体（stale）」になり、アプリが全滅する問題が頻発。",
            "・初期のコネクション管理実装。\n・エラーが発生してから初めて破棄するリアクティブ（事後対処）なアプローチが主流。\n・まだ定期的な事前検証という概念は未成熟だった。",
            "コネクションプールの\n標準化と死活監視の\n必要性の認識"
        ),
        (
            "2004年〜2010年\n【エンタープライズ拡大期】",
            "JBoss AS 4.x / 5.x\n(JCA 1.5 / JBossCX)",
            "【DB CPU枯渇の悲劇（SELECT 1の悪夢）】\n・高負荷システムでリクエストごとに `SELECT 1 FROM DUAL` を発行（validate-on-match）した結果、Oracle等のDB側でCPUの30〜50%がヘルスチェックだけで食いつぶされ、共有メモリのラッチ競合でDBが悲鳴を上げた。",
            "・ベンダー専用チェッカー（`OracleValidConnectionChecker`等）が登場し、SQLパースを伴わない軽量ping(TTCプロトコル)へ進化。\n・さらにリクエスト処理と検証を切り離すため、分単位のバックグラウンド検証（`background-validation-minutes`）が考案された。",
            "ベンダー専用Checkerと\nバックグラウンド検証の\n原型誕生"
        ),
        (
            "2011年〜2016年\n【新世代アーキテクチャ期】",
            "JBoss AS 7 / EAP 6\n(JCA 1.6 / IronJacamar)",
            "・JBossアーキテクチャがモジュール化（JBoss Modules）され全面刷新。\n・分単位（minutes）の設定では、秒単位で変化するネットワーク機器のタイムアウトや、短時間の通信断に対応できず、よりきめ細やかな制御が求められた。",
            "・次世代JCAコンテナ「IronJacamar」を新規採用。\n・ミリ秒精度の制御パラメータ `<background-validation-millis>` が正式導入！\n・タイマースケジューラとスレッドプールが刷新され、高負荷環境でも極めて低オーバーヘッドで検証が可能に。",
            "IronJacamar採用と\nミリ秒（millis）制御の\n公式確立"
        ),
        (
            "2017年〜2023年\n【クラウド・コンテナ移行期】",
            "JBoss EAP 7.0 〜 7.4\n(JCA 1.7 / Java EE 7・8)",
            "【クラウドの洗礼：見えない壁（350秒切断問題）】\n・AWSやAzureへの移行に伴い、AWS NAT Gateway(350秒)やAzure LB(4分)による「サイレント切断」が多発。\n・「夜間や昼休みの後、最初のユーザーアクセスで必ず通信が切れる」という障害が世界中で急増。",
            "・`background-validation-millis` を「120000ms（2分）」等に設定し、NATセッションの延命（Keep-Alive）兼ゾンビ接続の事前刈り取りとして運用することが、クラウド設計の必須プラクティスに昇華。",
            "クラウドNAT対策としての\n必須運用ベスト\nプラクティス化"
        ),
        (
            "2024年〜2026年（現在）\n【クラウドネイティブ最適化期】",
            "JBoss EAP 8.0\n(Jakarta EE 10 / Java 17・21)",
            "・マイクロサービス化、Kubernetes/OpenShiftでのPod高速スケール、Aurora等のマネージドDBの高速フェイルオーバー。\n・旧来のプロプライエタリなJDBC実装から、標準仕様への回帰と高速リカバリが要求される。",
            "・JDBC 4.0標準 `Connection.isValid(timeout)` を利用する `JDBC4ValidConnectionChecker` が主流に定着。\n・`flush-strategy`（破損接続のみ破棄）と連携し、マルチAZ切り替え時でもダウンタイムを極小化する自律回復アーキテクチャが完成。",
            "JDBC 4標準化と\n自律回復型クラウド\nネイティブ設計"
        ),
    ]

    for item in history_data:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_CENTER
        ws[f"B{row}"].border = BORDER_THIN

        ws[f"C{row}"] = item[1]
        ws[f"C{row}"].font = FONT_CODE_BOLD
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
        ws[f"F{row}"].font = FONT_BADGE_INFO
        ws[f"F{row}"].fill = FILL_BADGE_INFO
        ws[f"F{row}"].alignment = ALIGN_CENTER
        ws[f"F{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=32)
        row += 1

    row += 2

    # 3. なぜ分単位からミリ秒単位になったのか？
    render_section_header(ws, row, "深掘り：なぜ「分（minutes）」から「ミリ秒（millis）」への進化が必要だったのか？", "F")
    row += 1

    why_millis_headers = ["進化の動機・技術的背景", "分単位（minutes）の限界", "ミリ秒単位（millis）での解決とメリット"]
    render_table_headers(ws, row, why_millis_headers, "B")
    row += 1

    why_millis_data = [
        (
            "1. クラウド機器の端数タイムアウトへの適合",
            "クラウドサービス（例：AWS NAT Gateway）のアイドルタイムアウトは「350秒（5分50秒）」という秒単位の端数で設計されています。\n分単位（5分＝300秒、6分＝360秒）では、6分を選ぶとタイムアウト（350秒）を超過してしまい切断を防げず、5分を選ぶと余裕が少なすぎました。",
            "ミリ秒単位であれば「120,000ms（2分）」や「180,000ms（3分）」など、クラウドインフラの仕様値に対して安全マージンを精密に確保した設定が自在に可能になりました。"
        ),
        (
            "2. DBフェイルオーバー検知の迅速化",
            "オンプレミスの物理DB時代は夜間計画停止が中心でしたが、クラウドアクティブ/スタンバイ（例：Auroraフェイルオーバー）では数十秒での切り替えが求められます。\n分単位の巡回では、切り替え完了後も最大数分間ゾンビ接続が残り続け、業務停止時間が長期化しました。",
            "必要に応じて15,000ms（15秒）や30,000ms（30秒）といった短いインターバルでの設定が可能となり、フェイルオーバー後の自律復帰スピードが劇的に向上しました。"
        ),
        (
            "3. スレッドタイマースケジューラの精度向上",
            "旧来のJBossCXではOSタイマーや低精度のスレッドスリープに依存しており、タイマーのドリフト（時間のずれ）や同期ズレが発生しやすい構造でした。",
            "IronJacamarでは高精度なJava Concurrencyユーティリティ（ScheduledThreadPoolExecutorベース）が組み込まれ、ミリ秒単位で正確かつ低CPU消費なタイマースケジューリングが実現しました。"
        ),
    ]

    for item in why_millis_data:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_LEFT
        ws[f"B{row}"].border = BORDER_THIN

        ws.merge_cells(f"C{row}:D{row}")
        c_cell = ws[f"C{row}"]
        c_cell.value = item[1]
        c_cell.font = FONT_BODY
        c_cell.fill = fill
        c_cell.alignment = ALIGN_LEFT
        c_cell.border = BORDER_THIN
        ws[f"D{row}"].border = BORDER_THIN

        ws.merge_cells(f"E{row}:F{row}")
        e_cell = ws[f"E{row}"]
        e_cell.value = item[2]
        e_cell.font = FONT_BODY
        e_cell.fill = fill
        e_cell.alignment = ALIGN_LEFT
        e_cell.border = BORDER_THIN
        ws[f"F{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    row += 2

    # 4. 歴史から学ぶ教訓まとめ
    render_section_header(ws, row, "歴史から得られたシステム設計の教訓（アーキテクチャの鉄則）", "F")
    row += 1

    lesson_headers = ["教訓の項目", "過去の失敗とアンチパターン", "現代における正解アーキテクチャ"]
    render_table_headers(ws, row, lesson_headers, "B")
    row += 1

    lessons = [
        (
            "教訓1：DBにヘルスチェックで負荷をかけるな",
            "リクエストの都度SQL（SELECT 1）を発行すると、Webアクセス増に伴いDBが死ぬ。DBは業務トランザクション処理に専念させるべき。",
            "・検証処理は非同期（background-validation）で行う。\n・検証方法は重いSQLではなくJDBC4標準 isValid() またはドライバ固有pingを使用する。"
        ),
        (
            "教訓2：ネットワーク機器の無言切断を信用するな",
            "TCPソケットがOS上でESTABLISHEDだからといって、物理的に通信可能とは限らない。FWやNATは無言で経路を遮断する。",
            "・タイムアウト時間よりも短い周期で定期的に疎通させ、セッションテーブルを意図的に延命（Keep-Alive）させる設計が不可欠。"
        ),
        (
            "教訓3：事後対処（リアクティブ）から事前対処（プロアクティブ）へ",
            "ユーザーがアクセスしてエラー（通信例外）になってから接続を作り直す設計は、エンドユーザーに障害を押し付ける最悪の設計。",
            "・バックグラウンドで待機中接続を常時健全に保ち、ユーザーには常に「点検済みの健全な接続」のみを渡すプロアクティブ設計を徹底する。"
        ),
    ]

    for item in lessons:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_LEFT
        ws[f"B{row}"].border = BORDER_THIN

        ws.merge_cells(f"C{row}:D{row}")
        c_cell = ws[f"C{row}"]
        c_cell.value = item[1]
        c_cell.font = FONT_BODY
        c_cell.fill = fill
        c_cell.alignment = ALIGN_LEFT
        c_cell.border = BORDER_THIN
        ws[f"D{row}"].border = BORDER_THIN

        ws.merge_cells(f"E{row}:F{row}")
        e_cell = ws[f"E{row}"]
        e_cell.value = item[2]
        e_cell.font = FONT_BODY_BOLD
        e_cell.fill = fill
        e_cell.alignment = ALIGN_LEFT
        e_cell.border = BORDER_THIN
        ws[f"F{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    return ws
