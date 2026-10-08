# -*- coding: utf-8 -*-
"""
Sheet 2: 02_小学生でもわかる解説
"""

import openpyxl
from builder.styles import *

def build_sheet_02(wb):
    ws = wb.create_sheet(title="02_小学生でもわかる解説")
    setup_sheet_view(ws)

    # 列幅定義
    col_widths = {
        "A": 4,    # 余白
        "B": 20,   # ステップ / キャラクター
        "C": 48,   # たとえ話ストーリー（ゴーカート屋さん）
        "D": 48,   # 実際のITの世界（システムでの本当の動き）
        "E": 24,   # ポイント・まとめ
    }
    set_col_widths(ws, col_widths)

    # 1. タイトルバナー
    render_title_banner(
        ws,
        title_text="【たとえ話で直感理解】小学生でもわかる！コネクションプールと点検員のお話",
        subtitle_text="むずかしいIT用語を「テーマパークのゴーカート屋さん」に置き換えて、しくみと役割をスッキリ理解！",
        max_col_letter="E"
    )

    row = 5

    # 2. 登場人物対応表
    render_section_header(ws, row, "ステップ1：登場人物の紹介（むずかしい言葉を置き換えよう）", "E")
    row += 1

    char_headers = ["たとえ話のキャラクター", "ゴーカート屋さんでの役割", "実際のITシステムでの正体", "専門用語での意味"]
    render_table_headers(ws, row, char_headers, "B")
    row += 1

    characters = [
        ("お客さん", "テーマパークに遊びに来て「ゴーカートに乗りたい！」という人。", "Webブラウザでボタンを押した一般ユーザーのリクエスト。", "HTTPリクエスト / アプリケーションスレッド"),
        ("レンタルゴーカート", "お客さんを乗せて遠くのお城（目的地）まで走る乗り物。", "アプリとデータベースの間をつなぐ「専用の通信トンネル」。", "データベース・コネクション (Connection)"),
        ("ゴーカート車庫", "いつでもすぐ乗れるように、何台もの車を並べて待機させておく場所。", "あらかじめ通信トンネルを複数作って貯めておく仕組み。", "コネクションプール (Connection Pool)"),
        ("目的地のお城", "ゴーカートに乗ってたどり着く、大切な宝物（データ）がある場所。", "顧客データや注文情報が保管されている大切な保管庫。", "データベース (Database / RDBMS)"),
        ("見えないイタズラ小僧", "車庫で誰も使わずじっとしている車のタイヤの空気を、こっそり抜く妖怪。", "通信がないまま一定時間たつと、勝手に切断するネットワーク機器。", "ファイアウォール / クラウドNAT / ルーター"),
        ("裏方の専属整備士", "車庫の中を見回って、車がパンクしていないか調べる真面目な整備士。", "プールの中の接続が生きているか、裏で定期的に調べるスレッド。", "ConnectionValidator (検証スレッド)"),
        ("整備士のタイマー時計", "「◯◯分おきに見回りに行こう！」とセットする目覚まし時計。", "「何ミリ秒ごとに点検するか」を決める超重要パラメータ！", "background-validation-millis"),
    ]

    for item in characters:
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
        ws[f"E{row}"].font = FONT_CODE_BOLD
        ws[f"E{row}"].fill = fill
        ws[f"E{row}"].alignment = ALIGN_LEFT
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=22)
        row += 1

    row += 2

    # 3. 大事件：なぜ車庫で待っている間に壊れるの？
    render_section_header(ws, row, "ステップ2：事件発生！なぜ誰も使っていないのにパンクするの？", "E")
    row += 1

    incident_headers = ["場面", "ゴーカート屋さんでの出来事", "ITの世界で起きている恐ろしい現実", "なぜ大問題なのか？"]
    render_table_headers(ws, row, incident_headers, "B")
    row += 1

    incidents = [
        (
            "昼間の大忙しな時間",
            "お客さんがひっきりなしに来て、車庫のゴーカートは常に走り回っています。風を切って走っている間は、イタズラ小僧も手を出せません。",
            "アクセスが大量にある時間帯は、常にデータ通信が行われているため、途中のファイアウォールも通信を切断しません。すべて快調です！",
            "トラブルは起きない（正常稼働）"
        ),
        (
            "夜中や昼休みの静かな時間",
            "お客さんがピタッと来なくなりました。車庫には10台のゴーカートがじーっと待機しています。\nすると…見えないイタズラ小僧が現れました！「誰も使ってないなら、邪魔だからタイヤの空気を抜いちゃえ！」とこっそりプシューと空気を抜いて消えました。",
            "夜間や昼休みにアクセスが途切れると、ファイアウォールやルーターは「通信が終わった」と勝手に判断し、通信の通り道を無言で閉鎖（切断）します。\nしかも恐ろしいことに、サーバーには「切断したよ」と教えてくれません（サイレントドロップ）！",
            "見た目は正常なのに中身は死んでいる！\n（ゾンビコネクションの発生）"
        ),
        (
            "翌朝最初のお客さんが来店！",
            "朝イチでやってきた最初のお客さんが「1台貸して！」と言いました。\n店員は「はいどうぞ！」と車庫にあった車を渡しました。\nお客さんがアクセルを踏んだ瞬間…ガタガタッ！「きゃー！パンクしてて動かない！」と大事故になってしまいました！",
            "翌朝、最初のユーザーがアクセスした時、JBossは車庫にあった死んだ接続をそのままアプリに渡してしまいます。\nアプリがDBにアクセスしようとした瞬間…「通信切断エラー（Connection reset）」が発生し、画面にエラーが出て大パニックになります！",
            "朝一番のユーザーが必ずエラーを踏む！\n（朝イチ障害の正体）"
        ),
    ]

    for item in incidents:
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
        ws[f"E{row}"].font = FONT_BADGE_DANGER if "大問題" in item[3] or "エラー" in item[3] else FONT_BADGE_INFO
        ws[f"E{row}"].fill = FILL_BADGE_DANGER if "大問題" in item[3] or "エラー" in item[3] else FILL_BADGE_INFO
        ws[f"E{row}"].alignment = ALIGN_CENTER
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    row += 2

    # 4. 作戦1と作戦2の比較
    render_section_header(ws, row, "ステップ3：パンクを防ぐ2つの作戦（どっちが賢い？）", "E")
    row += 1

    plan_headers = ["作戦名", "ゴーカート屋さんでのやり方", "ITシステムでの動作（メリット・デメリット）", "結果・判定"]
    render_table_headers(ws, row, plan_headers, "B")
    row += 1

    plans = [
        (
            "作戦1：乗る直前に毎回\nタイヤを叩く作戦！\n(validate-on-match)",
            "お客さんが「貸して！」と窓口に来るたびに、整備士がトンカチを持って走ってきて、タイヤをガンガン叩いて空気が入っているか調べます。\nパンクしていたらその場で別の車に取り替えます。",
            "【メリット】100%確実に元気な接続を渡せる。\n【大デメリット】お客さんが1秒間に100人来たら大行列！毎回点検するせいで貸し出しに時間がかかり、整備士も叩きすぎてヘトヘト（DBのCPUがパンク）になります！",
            "安全だけど遅い！\n大混雑でクレームの嵐！\n（アンチパターンになりがち）"
        ),
        (
            "作戦2：裏方で定期的に\nパトロールする作戦！\n(background-validation)",
            "お客さんが待つ窓口とは関係なく、裏方で整備士がタイマー時計（millis）を見ながら、定期的に車庫を見回ります！\nもしパンクした車が見つかったら、裏でこっそり処分して新しいピカピカの車を並べておきます。",
            "【メリット】お客さんは窓口に来たら、点検済みの車にすぐに乗れる！待ち時間はゼロ秒！\n裏で点検するので、お客さんが100人来てもスムーズに貸し出せます！",
            "待ち時間ゼロで超快適！\nプロの現場での大本命！\n（★★★★★ 推奨）"
        ),
    ]

    for item in plans:
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
        ws[f"E{row}"].font = FONT_BADGE_REC if "推奨" in item[3] else FONT_BADGE_WARN
        ws[f"E{row}"].fill = FILL_BADGE_REC if "推奨" in item[3] else FILL_BADGE_WARN
        ws[f"E{row}"].alignment = ALIGN_CENTER
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=30)
        row += 1

    row += 2

    # 5. パトロール間隔の絶妙なバランス
    render_section_header(ws, row, "ステップ4：パトロール時計（millis）の絶妙な合わせ方", "E")
    row += 1

    timer_headers = ["時計の設定", "ゴーカート屋さんでの様子", "ITシステムでの実際の動き", "良し悪しの判定"]
    render_table_headers(ws, row, timer_headers, "B")
    row += 1

    timer_settings = [
        (
            "短すぎる設定\n（例：1秒ごと＝1000ms）",
            "整備士が1秒ごとに車庫中を走り回ってタイヤを叩き続けます。\n整備士は息切れして倒れてしまい、他の仕事が一切できなくなります。",
            "バックグラウンドスレッドが頻繁にDBへ問い合わせるため、ネットワーク帯域とDBのCPUを無駄食いし、システム全体の性能が低下します。",
            "× ダメ！（整備士の過労死）"
        ),
        (
            "長すぎる設定\n（例：1時間ごと＝3600000ms）",
            "イタズラ小僧は15分おきに空気を抜くのに、整備士は1時間後にしか見回りに来ません。\n結果として、車庫の中はパンクした車だらけになり、お客さんに渡してしまいます。",
            "ファイアウォールのタイムアウト（例：15分）よりも点検間隔が長いため、切断されたゾンビ接続を検知できず、ユーザー障害を防げません。",
            "× ダメ！（見回りが遅すぎて手遅れ）"
        ),
        (
            "黄金のちょうどいい設定！\n（例：3分ごと＝180000ms）",
            "イタズラ小僧が来る（15分）よりも前に、余裕を持ってサッと見回ります。\nしかも、見回りでタイヤを触ることで「この車はまだ使ってるよ！」とイタズラ小僧を追い払う効果（Keep-Alive）もあります！",
            "ファイアウォールのタイムアウトが来る前に定期的に通信が発生するため、切断自体を予防しつつ、万一の障害も迅速に検知して新品に交換します！",
            "◎ 大成功！（黄金のベストバランス）"
        ),
    ]

    for item in timer_settings:
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
        ws[f"E{row}"].font = FONT_BADGE_REC if "大成功" in item[3] else FONT_BADGE_DANGER
        ws[f"E{row}"].fill = FILL_BADGE_REC if "大成功" in item[3] else FILL_BADGE_DANGER
        ws[f"E{row}"].alignment = ALIGN_CENTER
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    row += 2

    # 6. 小学生向けまとめクイズ＆Q&A
    render_section_header(ws, row, "ステップ5：小学生でもスッキリ納得！よくある疑問 Q&A", "E")
    row += 1

    qa_headers = ["質問（素朴なギモン）", "わかりやすい回答（結論）", "プロの世界での詳しい理由"]
    render_table_headers(ws, row, qa_headers, "B")
    row += 1

    qas = [
        (
            "Q1. そもそも、なぜ最初から絶対にパンクしない最強のゴーカートを作らないの？",
            "ゴーカート（通信トンネル）自体は頑丈でも、道（ネットワーク）の途中にいる見えないイタズラ小僧（FW）が勝手に道を塞いじゃうからです！",
            "TCP接続は両端（JBossとDB）が正常でも、経由するステートフルFWやNATルータがメモリ節約のためにセッションテーブルを一方的に破棄（Silent Drop）するため、アプリ側だけで切断を防ぐことは物理的に不可能です。"
        ),
        (
            "Q2. お客さんが乗ってコースを走っている最中の車も、整備士が点検しに来るの？",
            "絶対に点検しません！走っている最中に整備士が飛び出してきたら大事故になっちゃうからです！点検するのは車庫で待っている車だけです。",
            "現在トランザクションを実行中の接続（In-Use接続）に対してバックグラウンドで検証パケットを送ると、実行中のクエリと競合してデータ破壊や例外が起きるため、JBossは「アイドル接続（待機中）」のみを厳密に選別して検査します。"
        ),
        (
            "Q3. ミリ秒（millis）って何？なんで「1分」じゃなくて「60000ミリ秒」って書くの？",
            "ミリ秒は「1秒の1000分の1」という超細かい時間の単位です！コンピュータの世界は一瞬で動くので、1秒や1分よりずっと細かく時間を合わせたいからです！",
            "クラウド（AWS/Azure）ではタイムアウトが「350秒」など中途半端な秒数に設定されていることが多く、分単位の設定（minutes）では柔軟に対応できません。そのためミリ秒単位（millis）での精密制御が必要になりました。"
        ),
    ]

    for item in qas:
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

        # D列とE列をマージして広い領域で解説
        ws.merge_cells(f"D{row}:E{row}")
        d_cell = ws[f"D{row}"]
        d_cell.value = item[2]
        d_cell.font = FONT_BODY_MUTED
        d_cell.fill = fill
        d_cell.alignment = ALIGN_LEFT
        d_cell.border = BORDER_THIN
        ws[f"E{row}"].border = BORDER_THIN

        auto_fit_row_height(ws, row, col_widths, min_height=26)
        row += 1

    return ws
