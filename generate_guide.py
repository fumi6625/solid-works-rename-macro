"""
generate_guide.py
マクロ使用方法ガイド.xlsx を生成する。

実行方法:
    pip install openpyxl
    python generate_guide.py
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# カラー定義
# ============================================================
C_HEADER_BG  = "1F4E79"
C_HEADER_FG  = "FFFFFF"
C_TITLE_BG   = "2E75B6"
C_STEP_BG    = "BDD7EE"
C_CAUTION_BG = "FFF2CC"
C_CAUTION_FG = "7F6000"
C_SECTION_BG = "D6E4F0"
C_GOOD_BG    = "E2EFDA"
C_SCREEN_BG  = "F2F2F2"   # 画面操作説明の背景（薄いグレー）
C_ARROW_FG   = "C55A11"   # 矢印・強調色

# ============================================================
# スタイルヘルパー
# ============================================================
def hfont(bold=False, size=10, color="000000", name="游ゴシック"):
    return Font(bold=bold, size=size, color=color, name=name)

def hfill(color):
    return PatternFill("solid", fgColor=color)

def hborder(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def halign(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def cell(ws, row, col, value="", bold=False, size=10, fg="000000",
         bg=None, h="left", v="center", wrap=False, border=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = hfont(bold=bold, size=size, color=fg)
    c.alignment = halign(h=h, v=v, wrap=wrap)
    if bg:
        c.fill = hfill(bg)
    if border:
        c.border = hborder()
    return c

def merge(ws, r1, c1, r2, c2, value="", bold=False, size=10,
          fg="000000", bg=None, h="left", v="center", wrap=False, border=False):
    ws.merge_cells(start_row=r1, start_column=c1,
                   end_row=r2, end_column=c2)
    c = ws.cell(row=r1, column=c1, value=value)
    c.font      = hfont(bold=bold, size=size, color=fg)
    c.alignment = halign(h=h, v=v, wrap=wrap)
    if bg:
        c.fill = hfill(bg)
    if border:
        c.border = hborder()
    return c

def col_w(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def row_h(ws, row, h):
    ws.row_dimensions[row].height = h

# ============================================================
# シート1：概要
# ============================================================
def build_overview(wb):
    ws = wb.active
    ws.title = "概要"
    ws.sheet_view.showGridLines = False
    col_w(ws, [2, 18, 58, 2])

    # タイトル
    merge(ws, 1, 2, 2, 3,
          "SolidWORKS アセンブリ コピー＆リネーム マクロ\n操作説明資料",
          bold=True, size=16, fg=C_HEADER_FG, bg=C_TITLE_BG,
          h="center", v="center", wrap=True)
    row_h(ws, 1, 22); row_h(ws, 2, 22)

    r = 4

    def section(title, r):
        merge(ws, r, 2, r, 3, title, bold=True, size=12,
              fg=C_HEADER_FG, bg=C_HEADER_BG)
        row_h(ws, r, 20)
        return r + 1

    # ---- 概要 ----
    r = section("■ このマクロでできること", r)
    txt = (
        "SolidWORKS の親アセンブリ（.sldasm）と、そのアセンブリ名で始まる子コンポーネント\n"
        "ファイルをまとめて新しい名前でコピーします。\n\n"
        "コピー後の親アセンブリを開くと、コピーされた子ファイルと正しくリンクされた状態\n"
        "になっています。元のファイルは変更・削除されません。"
    )
    merge(ws, r, 2, r, 3, txt, size=10, wrap=True, v="top")
    row_h(ws, r, 72); r += 2

    # ---- 動作の流れ ----
    r = section("■ 動作の流れ", r)
    steps = [
        ("STEP 1", "SolidWORKS で親アセンブリを開く"),
        ("STEP 2", "マクロを起動する"),
        ("STEP 3", "現在のアセンブリ名・子ファイル数を確認して「はい」をクリック"),
        ("STEP 4", "新しいアセンブリ名を入力して「OK」をクリック"),
        ("STEP 5", "保存先フォルダを選択して「OK」をクリック"),
        ("STEP 6", "実行内容を確認して「はい」をクリック"),
        ("STEP 7", "「ファイル名を変更しコピーが完了しました。」で完了！"),
    ]
    for step, desc in steps:
        cell(ws, r, 2, step, bold=True, fg=C_HEADER_FG, bg=C_HEADER_BG,
             h="center", v="center", border=True)
        cell(ws, r, 3, desc, border=True)
        row_h(ws, r, 18); r += 1
    r += 1

    # ---- 子ファイルの検出ルール ----
    r = section("■ 子ファイルの検出ルール（重要）", r)
    rule = (
        "アセンブリ内のコンポーネントのうち、ファイル名の先頭が「親ファイル名」と一致するものが\n"
        "コピー対象になります。それ以外はコピーされません。\n\n"
        "【例】  親ファイル名 = WIDGET-A.sldasm  →  新しい名前 = GADGET-B\n\n"
        "  WIDGET-A-01.sldprt      → ✓ コピー対象  → GADGET-B-01.sldprt\n"
        "  WIDGET-A-02.sldprt      → ✓ コピー対象  → GADGET-B-02.sldprt\n"
        "  WIDGET-A-FRAME.sldprt   → ✓ コピー対象  → GADGET-B-FRAME.sldprt\n"
        "  OTHER-PART.sldprt       → ✗ 対象外（名前が違う・コピーされない）"
    )
    merge(ws, r, 2, r, 3, rule, wrap=True, v="top", bg=C_CAUTION_BG)
    row_h(ws, r, 120); r += 2

    # ---- 注意事項 ----
    r = section("⚠ 実行前の注意事項", r)
    cautions = [
        "□ 実行前に必ず元ファイルのバックアップを取得してください",
        "□ 親アセンブリを SolidWORKS で開いた状態でマクロを実行してください",
        "□ 子ファイルが他のアプリケーションで開かれていないことを確認してください",
        "□ 保存先フォルダへの書き込み権限があることを確認してください",
    ]
    for c_text in cautions:
        merge(ws, r, 2, r, 3, c_text, fg=C_CAUTION_FG, bg=C_CAUTION_BG, border=True)
        row_h(ws, r, 18); r += 1


# ============================================================
# シート2：インストール手順（刷新：1ファイルのみ・UserForm不要）
# ============================================================
def build_install(wb):
    ws = wb.create_sheet("インストール手順")
    ws.sheet_view.showGridLines = False
    col_w(ws, [2, 6, 22, 46, 2])

    merge(ws, 1, 2, 1, 4,
          "マクロのインストール手順",
          bold=True, size=14, fg=C_HEADER_FG, bg=C_HEADER_BG,
          h="center", v="center")
    row_h(ws, 1, 28)

    # ポイント説明
    merge(ws, 2, 2, 2, 4,
          "💡 インポートするファイルは「RenameAndCopy.bas」の 1ファイルのみです。UserForm の手動設定は不要です。",
          bold=True, size=10, fg=C_ARROW_FG, bg=C_GOOD_BG, wrap=True)
    row_h(ws, 2, 20)

    r = 4

    steps = [
        ("1",
         "マクロファイルを用意する",
         "generate_bas.py を実行して RenameAndCopy.bas を生成します。\n\n"
         "① generate_bas.py があるフォルダでコマンドプロンプトを開く\n"
         "② 以下のコマンドを実行する：\n"
         "     python generate_bas.py\n\n"
         "→ 同じフォルダに「RenameAndCopy.bas」が作成されます。\n\n"
         "※ generate_bas.py が既に含まれていればこの手順は不要です。"),

        ("2",
         "SolidWORKS でマクロエディタを開く",
         "SolidWORKS のメニューから操作します：\n\n"
         "  ①  上部メニューバーの「Tools（ツール）」をクリック\n"
         "  ②  「Macro（マクロ）」にマウスを合わせる\n"
         "  ③  「Edit Macro（マクロを編集）」をクリック\n\n"
         "→ Visual Basic for Applications（VBA）エディタが開きます。"),

        ("3",
         "新しいマクロプロジェクトを作成する",
         "VBA エディタが開いたら新しいマクロファイルを作成します：\n\n"
         "  ①  表示された「Edit Macro」ダイアログの「新規（New）」ボタンをクリック\n"
         "  ②  「名前を付けて保存」ダイアログが開く\n"
         "  ③  ファイル名に「RenameAndCopy」と入力\n"
         "  ④  保存場所を確認して「保存」をクリック\n\n"
         "→ 「RenameAndCopy.swp」というマクロファイルが作成されます。\n"
         "   VBA エディタに空のコード画面が表示されます。"),

        ("4",
         "RenameAndCopy.bas をインポートする",
         "VBA エディタのメニューから .bas ファイルを読み込みます：\n\n"
         "  ①  VBA エディタ上部メニューの「File（ファイル）」をクリック\n"
         "  ②  「Import File...（ファイルのインポート）」をクリック\n"
         "  ③  ファイル選択ダイアログで「RenameAndCopy.bas」を選択\n"
         "  ④  「開く」をクリック\n\n"
         "→ 左側のプロジェクトツリーに「RenameAndCopy」モジュールが追加されます。\n\n"
         "⚠ 自動的に作成された「Module1」など空のモジュールがある場合は\n"
         "   右クリック → 「Remove Module1」で削除してください。"),

        ("5",
         "マクロを保存する",
         "VBA エディタで Ctrl+S を押してマクロを保存します。\n\n"
         "→「RenameAndCopy.swp」ファイルにコードが保存されます。\n\n"
         "以上でインストール完了です。"),

        ("6",
         "（確認）動作テストを行う",
         "テスト用のアセンブリファイルで動作を確認します：\n\n"
         "  ①  SolidWORKS でテスト用のアセンブリ（.sldasm）を開く\n"
         "  ②  Tools → Macro → Run Macro で「RenameAndCopy.swp」を選択\n"
         "  ③  マクロ一覧に「StartRenameAndCopy」が表示されることを確認\n"
         "  ④  「実行（Run）」をクリックしてダイアログが表示されれば成功"),
    ]

    for num, title, detail in steps:
        cell(ws, r, 2, num, bold=True, fg=C_HEADER_FG, bg=C_HEADER_BG,
             h="center", v="top", border=True)
        cell(ws, r, 3, title, bold=True, bg=C_STEP_BG, border=True, v="top")
        cell(ws, r, 4, detail, wrap=True, v="top", border=True)
        lines = detail.count("\n") + 1
        row_h(ws, r, max(20, lines * 14 + 6))
        r += 1

    r += 1
    merge(ws, r, 2, r, 4,
          "💡 補足：SolidWORKS 2020 以降で動作確認済みです。",
          fg=C_CAUTION_FG, bg=C_CAUTION_BG, wrap=True)
    row_h(ws, r, 18)


# ============================================================
# シート3：操作手順（詳細な画面説明付き）
# ============================================================
def build_operation(wb):
    ws = wb.create_sheet("操作手順")
    ws.sheet_view.showGridLines = False
    col_w(ws, [2, 6, 22, 46, 2])

    merge(ws, 1, 2, 1, 4,
          "マクロの操作手順（実行のたびに行う手順）",
          bold=True, size=14, fg=C_HEADER_FG, bg=C_HEADER_BG,
          h="center", v="center")
    row_h(ws, 1, 28)

    r = 3

    def section_hdr(ws, r, title):
        merge(ws, r, 2, r, 4, f"【 {title} 】",
              bold=True, size=11, fg=C_HEADER_FG, bg=C_TITLE_BG)
        row_h(ws, r, 20)
        return r + 1

    # ========== 準備 ==========
    r = section_hdr(ws, r, "準備")
    prep_steps = [
        ("準備①",
         "元ファイルのバックアップを取る",
         "対象の親アセンブリと子ファイルが入ったフォルダを別の場所にコピーして\n"
         "バックアップを保管してください。\n"
         "（マクロ実行中にエラーが起きた場合に復元できます）"),
        ("準備②",
         "SolidWORKS で親アセンブリを開く",
         "コピー・リネームしたい親アセンブリ（.sldasm）を SolidWORKS で開きます。\n\n"
         "✓ 正しい状態：タイトルバーに「〇〇〇.sldasm」と表示されている\n"
         "✗ 誤った状態：パーツファイル（.sldprt）が開かれている\n\n"
         "※ パーツファイルが開かれた状態でマクロを実行するとエラーになります。"),
    ]
    for step, title, detail in prep_steps:
        cell(ws, r, 2, step, bold=True, fg=C_HEADER_FG, bg=C_HEADER_BG,
             h="center", v="top", border=True)
        cell(ws, r, 3, title, bold=True, bg=C_STEP_BG, border=True, v="top")
        cell(ws, r, 4, detail, wrap=True, v="top", border=True)
        row_h(ws, r, max(20, detail.count("\n") * 14 + 10))
        r += 1
    r += 1

    # ========== マクロ起動 ==========
    r = section_hdr(ws, r, "マクロの起動")
    run_steps = [
        ("操作①",
         "マクロの実行メニューを開く",
         "SolidWORKS の上部メニューから：\n\n"
         "  「Tools（ツール）」→「Macro（マクロ）」→「Run Macro（マクロを実行）」\n\n"
         "→ ファイル選択ダイアログが開きます。"),
        ("操作②",
         "RenameAndCopy.swp を選択して実行",
         "  ①  ダイアログで「RenameAndCopy.swp」を選択\n"
         "  ②  「開く」をクリック\n"
         "  ③  マクロの関数一覧が表示される\n"
         "  ④  「StartRenameAndCopy」が選択されていることを確認\n"
         "  ⑤  「実行（Run）」ボタンをクリック"),
    ]
    for step, title, detail in run_steps:
        cell(ws, r, 2, step, bold=True, fg=C_HEADER_FG, bg=C_HEADER_BG,
             h="center", v="top", border=True)
        cell(ws, r, 3, title, bold=True, bg=C_STEP_BG, border=True, v="top")
        cell(ws, r, 4, detail, wrap=True, v="top", border=True)
        row_h(ws, r, max(20, detail.count("\n") * 14 + 10))
        r += 1
    r += 1

    # ========== ダイアログ操作 ==========
    r = section_hdr(ws, r, "ダイアログの操作（マクロ起動後に表示される画面）")

    dialog_steps = [
        ("画面①",
         "開始確認ダイアログ",
         "【表示内容】\n"
         "  ・現在のアセンブリ名（例：WIDGET-A.sldasm）\n"
         "  ・対象子ファイル数（例：4件）\n\n"
         "【操作】\n"
         "  ✓ 内容を確認して「はい（Yes）」をクリックすると次に進みます\n"
         "  ✗ 「いいえ（No）」でキャンセルできます\n\n"
         "⚠ 子ファイル数が 0 件の場合は命名規則を確認してください\n"
         "  （子ファイル名が親ファイル名で始まっていない可能性があります）"),

        ("画面②",
         "新しいアセンブリ名の入力",
         "【表示内容】\n"
         "  「新しいアセンブリ名を入力してください。（拡張子 .sldasm は不要です）」\n\n"
         "【操作】\n"
         "  ①  入力欄に新しい名前を入力する\n"
         "      例：WIDGET-A → GADGET-B と変更したい場合は「GADGET-B」と入力\n"
         "  ②  「OK」をクリックする\n\n"
         "⚠ 注意：拡張子（.sldasm）は入力しないでください\n"
         "⚠ 次の記号はファイル名に使用できません：\\ / : * ? \" < > |"),

        ("画面③",
         "保存先フォルダの選択",
         "【表示内容】\n"
         "  フォルダ選択ダイアログ\n"
         "  （デフォルトで現在の親ファイルと同じフォルダが選択されています）\n\n"
         "【操作】\n"
         "  ①  コピーファイルを保存したいフォルダを選択\n"
         "      ・同じフォルダに保存する場合はそのまま「OK」\n"
         "      ・別のフォルダを選ぶ場合はツリーから目的のフォルダをクリック\n"
         "  ②  「OK」をクリックする\n\n"
         "💡 新しいフォルダを作成したい場合は「新しいフォルダーを作成」ボタンを使用"),

        ("画面④",
         "実行確認ダイアログ",
         "【表示内容】\n"
         "  ・現在の親ファイル名\n"
         "  ・新しい親ファイル名\n"
         "  ・保存先フォルダ\n"
         "  ・対象子ファイル数\n\n"
         "【操作】\n"
         "  ①  表示された内容が正しいことを確認する\n"
         "  ②  問題なければ「はい（Yes）」をクリックして実行\n"
         "  ③  内容を変更したい場合は「いいえ（No）」でキャンセルし、最初からやり直す"),

        ("画面⑤",
         "完了メッセージ",
         "【表示内容】\n"
         "  「ファイル名を変更しコピーが完了しました。」\n"
         "  ・保存先フォルダ\n"
         "  ・新しい親ファイル名\n"
         "  ・コピーされた子ファイル数\n\n"
         "【操作】\n"
         "  「OK」をクリックしてマクロを終了する\n\n"
         "【完了後の確認】\n"
         "  ①  保存先フォルダを開いて新しいファイルが作成されていることを確認\n"
         "  ②  新しい親アセンブリ（.sldasm）を SolidWORKS で開く\n"
         "  ③  全てのコンポーネントが正常に表示されていれば成功"),
    ]

    for step, title, detail in dialog_steps:
        cell(ws, r, 2, step, bold=True, fg=C_HEADER_FG, bg=C_TITLE_BG,
             h="center", v="top", border=True)
        cell(ws, r, 3, title, bold=True, bg=C_STEP_BG, border=True, v="top")
        cell(ws, r, 4, detail, wrap=True, v="top", border=True)
        row_h(ws, r, max(20, detail.count("\n") * 14 + 10))
        r += 1
    r += 1

    # ========== 変換例 ==========
    r = section_hdr(ws, r, "ファイル名変換例")

    # ヘッダー
    for col_idx, txt in enumerate(["種別", "変換前（元ファイル）", "変換後（コピーファイル）"], start=2):
        cell(ws, r, col_idx, txt, bold=True, fg=C_HEADER_FG,
             bg=C_HEADER_BG, h="center", border=True)
    row_h(ws, r, 18); r += 1

    examples = [
        ("親アセンブリ",  "WIDGET-A.sldasm",       "GADGET-B.sldasm"),
        ("子パーツ①",    "WIDGET-A-01.sldprt",    "GADGET-B-01.sldprt"),
        ("子パーツ②",    "WIDGET-A-02.sldprt",    "GADGET-B-02.sldprt"),
        ("子パーツ③",    "WIDGET-A-FRAME.sldprt", "GADGET-B-FRAME.sldprt"),
        ("対象外（コピーなし）", "OTHER-PART.sldprt", "（変更なし・コピーされない）"),
    ]
    for kind, before, after in examples:
        cell(ws, r, 2, kind, h="center", border=True)
        cell(ws, r, 3, before, border=True)
        bg = C_GOOD_BG if "コピーなし" not in kind else C_CAUTION_BG
        cell(ws, r, 4, after, bg=bg, border=True)
        row_h(ws, r, 18); r += 1


# ============================================================
# シート4：よくある質問
# ============================================================
def build_faq(wb):
    ws = wb.create_sheet("よくある質問")
    ws.sheet_view.showGridLines = False
    col_w(ws, [2, 6, 58, 2])

    merge(ws, 1, 2, 1, 3,
          "よくある質問・トラブルシューティング",
          bold=True, size=14, fg=C_HEADER_FG, bg=C_HEADER_BG,
          h="center", v="center")
    row_h(ws, 1, 28)

    faqs = [
        ("Q", "「アセンブリファイルを開いてください」と表示される",
         "A", "パーツファイル（.sldprt）が開かれています。\n"
              "親アセンブリ（.sldasm）を SolidWORKS で開いてから再度マクロを実行してください。"),

        ("Q", "子ファイル数が「0件」と表示される",
         "A", "子ファイルのファイル名が親ファイル名で始まっていない可能性があります。\n\n"
              "確認方法：\n"
              "  ・親ファイル名（拡張子なし）と子ファイル名を並べて確認する\n"
              "  ・子ファイル名の先頭が親ファイル名と完全に一致しているか確認する\n\n"
              "例）親が「ABC-001.sldasm」の場合、子は「ABC-001-○○.sldprt」形式である必要があります。"),

        ("Q", "コピー後の親アセンブリを開いたら参照エラーが出る",
         "A", "以下の順で確認してください：\n\n"
              "  ①  保存先フォルダに全ての子ファイルがコピーされているか確認する\n"
              "  ②  コピーが正常に完了していた場合、SolidWORKS の\n"
              "      「外部参照の修復」機能で手動パス指定を行う\n\n"
              "手動修復の手順：\n"
              "  File → Open → コピーした親 .sldasm を開く\n"
              "  → 「参照が見つかりません」ダイアログが表示されたら「参照先の変更」を選択\n"
              "  → コピー先フォルダを指定する"),

        ("Q", "VBA エディタで日本語が文字化けしている",
         "A", "ファイルのエンコーディングが原因です。以下の手順で解決できます：\n\n"
              "  ①  generate_bas.py を Python で実行して RenameAndCopy.bas を再生成する\n"
              "      （このスクリプトが Shift-JIS エンコーディングでファイルを生成します）\n"
              "  ②  VBA エディタで既存モジュールを右クリック → Remove（削除）\n"
              "  ③  File → Import File で再生成した .bas をインポートし直す\n\n"
              "⚠ メモ帳などのテキストエディタで直接 .bas ファイルを編集・保存すると\n"
              "  UTF-8 で保存されて再び文字化けします。編集は必ず VBA エディタ上で行ってください。"),

        ("Q", "「既に同名ファイルが存在します」と表示される",
         "A", "指定した保存先に同じ名前のファイルが既に存在します。\n\n"
              "  ・「はい」→ そのファイルのみ上書きしてコピー続行\n"
              "  ・「いいえ」→ そのファイルはスキップして残りは続行\n"
              "  ・「キャンセル」→ 以降の全ファイルをスキップ"),

        ("Q", "元のファイルが変更されないか心配",
         "A", "このマクロは元ファイルを一切変更・削除しません。コピーのみを行います。\n"
              "安心して実行してください。"),

        ("Q", "サブアセンブリ（.sldasm が入れ子になっている）場合はどうする？",
         "A", "サブアセンブリのファイル名が親ファイル名で始まる場合は、そのファイルも\n"
              "コピー対象になります。\n\n"
              "ただし、サブアセンブリが参照している孫ファイルは自動ではコピーされません。\n"
              "必要な場合は、コピーされたサブアセンブリを開き、同様にマクロを実行してください。"),
    ]

    r = 3
    for q_lbl, q_txt, a_lbl, a_txt in faqs:
        cell(ws, r, 2, q_lbl, bold=True, fg=C_HEADER_FG, bg=C_TITLE_BG,
             h="center", v="top", border=True)
        cell(ws, r, 3, q_txt, bold=True, bg=C_STEP_BG, wrap=True,
             v="top", border=True)
        row_h(ws, r, max(18, q_txt.count("\n") * 14 + 14))
        r += 1

        cell(ws, r, 2, a_lbl, bold=True, fg=C_CAUTION_FG,
             bg=C_CAUTION_BG, h="center", v="top", border=True)
        cell(ws, r, 3, a_txt, wrap=True, v="top", border=True)
        row_h(ws, r, max(18, a_txt.count("\n") * 14 + 14))
        r += 2


# ============================================================
# メイン処理
# ============================================================
def main():
    wb = Workbook()
    build_overview(wb)
    build_install(wb)
    build_operation(wb)
    build_faq(wb)

    out = "マクロ使用方法ガイド.xlsx"
    wb.save(out)
    print(f"Excel ガイドを生成しました：{out}")


if __name__ == "__main__":
    main()
