# -*- coding: utf-8 -*-
"""
Sheet 6: 06_設定リファレンスとCLI
"""

import openpyxl
from builder.styles import *

def build_sheet_06(wb):
    ws = wb.create_sheet(title="06_設定リファレンスとCLI")
    setup_sheet_view(ws)

    # 列幅定義
    col_widths = {
        "A": 4,    # 余白
        "B": 28,   # パラメータ名 / コマンド種別
        "C": 14,   # データ型 / 属性
        "D": 22,   # デフォルト値 / 推奨値
        "E": 44,   # 役割・機能説明
        "F": 38,   # 設定時の注意事項・推奨構成
    }
    set_col_widths(ws, col_widths)

    # 1. タイトルバナー
    render_title_banner(
        ws,
        title_text="JBoss EAP データソース接続検証パラメータ完全リファレンス & CLI集",
        subtitle_text="全パラメータの詳細定義、主要DBMS別（Oracle, PostgreSQL, MySQL, SQL Server）本番XML設定スニペット、管理CLI",
        max_col_letter="F"
    )

    row = 5

    # 2. パラメータ完全リファレンス表
    render_section_header(ws, row, "1. データソース接続検証 関連パラメータ完全定義表", "F")
    row += 1

    param_headers = ["パラメータ名 (XML / CLI)", "データ型", "デフォルト値 / 推奨値", "役割・内部動作の詳細説明", "設計・チューニングの注意事項"]
    render_table_headers(ws, row, param_headers, "B")
    row += 1

    parameters = [
        (
            "background-validation",
            "boolean",
            "デフォルト: false\n推奨: true",
            "バックグラウンド検証スレッド（ConnectionValidator）による定期健全性チェックを有効化する。validate-on-match とは排他。",
            "本番環境では原則 true を推奨。validate-on-match が false であることを確認すること。"
        ),
        (
            "background-validation-millis",
            "long\n(ミリ秒)",
            "デフォルト: 0 (無効)\n推奨: 120000〜300000\n(2〜5分)",
            "アイドル状態の接続を検証する時間間隔（ミリ秒）。この時間が経過するごとにワーカーランナーがアイドル接続を検査する。",
            "経由するFW/NATのタイムアウト時間の半分以下に設定すること。idle-timeout-minutes と競合しない値にする。"
        ),
        (
            "validate-on-match",
            "boolean",
            "デフォルト: false\n推奨: false",
            "アプリケーションがプールから接続を取得する（貸出時）たびに即時検証を行う。取得の都度DB疎通が発生する。",
            "低遅延が求められるWebでは false 必須。true にするとリクエストごとにRTT遅延が加算されDB負荷が高騰する。"
        ),
        (
            "use-fast-fail",
            "boolean",
            "デフォルト: false\n推奨: true",
            "検証処理において、ある接続が無効と判定された場合、後続のアイドル接続の検証を即座にスキップして早期終了する。",
            "DB障害やネットワーク断が発生した際、無駄に全接続のタイムアウトを待たずに高速にプールをクリーンアップできる。"
        ),
        (
            "valid-connection-checker\n-class-name",
            "string\n(クラス名)",
            "デフォルト: 未設定\n推奨: JDBC4またはDBMS専用",
            "接続の健全性を確認するチェッカークラス。JDBC 4.0標準 isValid() を呼ぶクラスや、DBMS専用クラスを指定する。",
            "最新EAP 7.4/8.0では org.jboss.jca.adapters.jdbc.extensions.novendor.JDBC4ValidConnectionChecker が標準推奨。"
        ),
        (
            "check-valid-connection-sql",
            "string\n(SQL文)",
            "デフォルト: 未設定\n推奨: 未設定 (原則不使用)",
            "検証用のSQL文（例: SELECT 1）。チェッカークラスが指定されていない場合に使用されるフォールバック。",
            "SQLパース負荷やパケットサイズが大きいため本番非推奨。valid-connection-checker を優先使用すべき。"
        ),
        (
            "exception-sorter\n-class-name",
            "string\n(クラス名)",
            "デフォルト: 未設定\n推奨: DBMS専用クラス",
            "SQL実行時に例外が発生した際、それが「致命的な通信断」か「単なるSQL構文エラー/制約違反」かを判別するクラス。",
            "通信切断エラーと判別された場合のみ接続をプールから破棄する。誤って正常接続を破棄しないために極めて重要。"
        ),
        (
            "idle-timeout-minutes",
            "integer\n(分)",
            "デフォルト: 15 (分)\n推奨: 10〜15 (分)",
            "接続がアイドル状態でプール内に保持される最大時間（分）。この時間を超えた接続はプールサイズ縮小のためクローズされる。",
            "min-pool-size で指定された接続には適用されない点に注意。background-validation-millis の周期より長く設定する。"
        ),
        (
            "blocking-timeout-wait-millis",
            "long\n(ミリ秒)",
            "デフォルト: 30000 (30秒)\n推奨: 5000〜15000 (5〜15秒)",
            "プール内の全接続が使用中の際、新規接続が空くまでアプリケーションスレッドが待機する最大許容時間（ミリ秒）。",
            "長すぎるとDB過負荷時にJBossのリクエストスレッドが詰まりカスケード障害（全滅）を引き起こすため短めに設定。"
        ),
        (
            "flush-strategy",
            "string\n(列挙型)",
            "デフォルト: FailingConnectionOnly\n推奨: FailingConnectionOnly",
            "エラー検知時やフラッシュ時にプール内の接続をどう破棄するか（FailingConnectionOnly / IdleConnections / EntirePool）。",
            "通常は FailingConnectionOnly（破損した接続のみ破棄）が最も安全。EntirePool は正常接続まで全破棄するため注意。"
        ),
    ]

    for item in parameters:
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

        ws[f"F{row}"] = item[4]
        ws[f"F{row}"].font = FONT_BODY
        ws[f"F{row}"].fill = fill
        ws[f"F{row}"].alignment = ALIGN_LEFT
        ws[f"F{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    row += 2

    # 3. DBMS別 推奨XML設定スニペット
    render_section_header(ws, row, "2. 主要DBMS別 本番推奨XML設定スニペット（standalone.xml）", "F")
    row += 1

    xml_headers = ["データベース種別", "推奨XML設定スニペット (<validation> タグ周辺)", "設定のポイント・解説"]
    render_table_headers(ws, row, xml_headers, "B")
    row += 1

    xml_snippets = [
        (
            "Oracle Database\n(19c / 21c / 23ai)",
            """<validation>
    <valid-connection-checker class-name="org.jboss.jca.adapters.jdbc.extensions.oracle.OracleValidConnectionChecker"/>
    <validate-on-match>false</validate-on-match>
    <background-validation>true</background-validation>
    <background-validation-millis>120000</background-validation-millis>
    <use-fast-fail>true</use-fast-fail>
    <exception-sorter class-name="org.jboss.jca.adapters.jdbc.extensions.oracle.OracleExceptionSorter"/>
</validation>""",
            "Oracle専用のチェッカー（pingDatabase使用）と例外ソータを指定。\nORA-03113やORA-01012等の致命的通信エラー時に的確に接続破棄が行われます。"
        ),
        (
            "PostgreSQL\n(13 / 14 / 15 / 16)",
            """<validation>
    <valid-connection-checker class-name="org.jboss.jca.adapters.jdbc.extensions.novendor.JDBC4ValidConnectionChecker"/>
    <validate-on-match>false</validate-on-match>
    <background-validation>true</background-validation>
    <background-validation-millis>120000</background-validation-millis>
    <use-fast-fail>true</use-fast-fail>
    <exception-sorter class-name="org.jboss.jca.adapters.jdbc.extensions.postgres.PostgreSQLExceptionSorter"/>
</validation>""",
            "JDBC 4.0標準チェッカーを採用。\nPostgreSQL特有の致命的例外（接続クローズ等）を正しくソートしてプールを保護します。"
        ),
        (
            "MySQL / MariaDB\n(8.0 / 10.x)",
            """<validation>
    <valid-connection-checker class-name="org.jboss.jca.adapters.jdbc.extensions.mysql.MySQLValidConnectionChecker"/>
    <validate-on-match>false</validate-on-match>
    <background-validation>true</background-validation>
    <background-validation-millis>120000</background-validation-millis>
    <use-fast-fail>true</use-fast-fail>
    <exception-sorter class-name="org.jboss.jca.adapters.jdbc.extensions.mysql.MySQLExceptionSorter"/>
</validation>""",
            "MySQL専用チェッカー（COM_PINGパケット送信）を指定。\nMySQLの wait_timeout による切断を確実に防止します。"
        ),
        (
            "Microsoft SQL Server\n(2019 / 2022)",
            """<validation>
    <valid-connection-checker class-name="org.jboss.jca.adapters.jdbc.extensions.novendor.JDBC4ValidConnectionChecker"/>
    <validate-on-match>false</validate-on-match>
    <background-validation>true</background-validation>
    <background-validation-millis>120000</background-validation-millis>
    <use-fast-fail>true</use-fast-fail>
    <exception-sorter class-name="org.jboss.jca.adapters.jdbc.extensions.mssql.MSSQLExceptionSorter"/>
</validation>""",
            "公式Microsoft JDBC DriverとJDBC4ValidConnectionCheckerの組み合わせ。\nMSSQLExceptionSorterにより接続切断を的確にトラップします。"
        ),
    ]

    for item in xml_snippets:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_CENTER
        ws[f"B{row}"].border = BORDER_THIN

        ws.merge_cells(f"C{row}:D{row}")
        c_cell = ws[f"C{row}"]
        c_cell.value = item[1]
        c_cell.font = FONT_CODE
        c_cell.fill = FILL_CODE_BLOCK
        c_cell.alignment = ALIGN_CODE
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

        auto_fit_row_height(ws, row, col_widths, min_height=45)
        row += 1

    row += 2

    # 4. JBoss CLI コマンド集
    render_section_header(ws, row, "3. JBoss EAP 管理CLIコマンド集（運用・保守・トラブルシューティング）", "F")
    row += 1

    cli_headers = ["操作内容", "JBoss CLI コマンド構文 (jboss-cli.sh / jboss-cli.bat)", "説明・実行時の挙動"]
    render_table_headers(ws, row, cli_headers, "B")
    row += 1

    cli_commands = [
        (
            "現在の設定値を確認",
            "/subsystem=datasources/data-source=MyDataSource:read-resource(recursive=true,include-runtime=true)",
            "データソースの全設定値およびランタイム統計（現在の接続数、アクティブ数など）をJSON形式で確認します。"
        ),
        (
            "バックグラウンド検証を有効化\n(millisを設定)",
            "/subsystem=datasources/data-source=MyDataSource:write-attribute(name=background-validation,value=true)\n/subsystem=datasources/data-source=MyDataSource:write-attribute(name=background-validation-millis,value=120000L)",
            "稼働中のデータソースに対して background-validation を true にし、検証周期を 120,000ms (2分) に変更します。"
        ),
        (
            "validate-on-match を無効化",
            "/subsystem=datasources/data-source=MyDataSource:write-attribute(name=validate-on-match,value=false)",
            "即時検証を無効化し、二重検証オーバーヘッドを排除します。"
        ),
        (
            "データソースの接続テスト実行",
            "/subsystem=datasources/data-source=MyDataSource:test-connection-in-pool",
            "現在プール内に存在する接続が正常に疎通可能かをCLIから即座にテストし、結果（true/false）を返します。"
        ),
        (
            "コネクションプールの強制フラッシュ\n(ゾンビ接続の一括掃除)",
            "/subsystem=datasources/data-source=MyDataSource:flush-idle-connection-in-pool",
            "プール内で待機しているすべてのアイドル接続を強制的に破棄します。\nDB再起動後などにゾンビを一掃したい場合に極めて有効です。"
        ),
    ]

    for item in cli_commands:
        fill = FILL_ZEBRA_EVEN if (row % 2 == 0) else FILL_ZEBRA_ODD
        ws[f"B{row}"] = item[0]
        ws[f"B{row}"].font = FONT_BODY_BOLD
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = ALIGN_LEFT
        ws[f"B{row}"].border = BORDER_THIN

        ws.merge_cells(f"C{row}:D{row}")
        c_cell = ws[f"C{row}"]
        c_cell.value = item[1]
        c_cell.font = FONT_CODE
        c_cell.fill = FILL_CODE_BLOCK
        c_cell.alignment = ALIGN_CODE
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

    return ws
