# SolidWORKS アセンブリ コピー＆リネーム マクロ

SolidWORKS の親アセンブリと子コンポーネントファイルをまとめて新しい名前でコピーする VBA マクロです。  
コピーされた親アセンブリを開くと、コピーされた子ファイルと正しくリンクされた状態になります。

---

## ファイル構成

| ファイル | 説明 |
|---|---|
| `generate_bas.py` | **最初に実行**。`RenameAndCopy.bas` を Shift-JIS エンコーディングで生成するスクリプト |
| `RenameAndCopy.bas` | VBA マクロ本体（`generate_bas.py` で生成、このファイルを SolidWORKS にインポートする） |
| `マクロ使用方法ガイド.xlsx` | Excel 操作説明資料 |
| `generate_guide.py` | Excel 資料を再生成するスクリプト |

---

## インストール手順

### ステップ 1：`RenameAndCopy.bas` を生成する

```bash
python generate_bas.py
```

`RenameAndCopy.bas` が同フォルダに作成されます（Shift-JIS エンコーディング）。

> **なぜ generate_bas.py で生成するのか？**  
> VBA エディタ（Windows）は Shift-JIS（CP932）でファイルを読み込みます。  
> `generate_bas.py` を使うと日本語が文字化けしない .bas ファイルが生成されます。

### ステップ 2：SolidWORKS でマクロエディタを開く

`Tools（ツール）` → `Macro（マクロ）` → `Edit Macro（マクロを編集）`

### ステップ 3：新しいマクロプロジェクトを作成する

- 「新規（New）」をクリック
- ファイル名 `RenameAndCopy` で保存 → `RenameAndCopy.swp` が作成される

### ステップ 4：`RenameAndCopy.bas` をインポートする

VBA エディタのメニュー `File → Import File...` から `ReneAndCopy.bas` を選択してインポート。

> **UserForm の手動設定は不要です。** `RenameAndCopy.bas` の 1 ファイルだけで動作します。

### ステップ 5：保存

`Ctrl+S` でマクロを保存。

---

## 使用方法

1. SolidWORKS で親アセンブリ（`.sldasm`）を開く
2. `Tools → Macro → Run Macro` で `RenameAndCopy.swp` を選択
3. `StartRenameAndCopy` を選択して「実行（Run）」
4. ダイアログに従って操作：
   - **画面①**：現在のアセンブリ名・子ファイル数を確認 → 「はい」
   - **画面②**：新しいアセンブリ名を入力 → 「OK」（拡張子不要）
   - **画面③**：保存先フォルダを選択 → 「OK」
   - **画面④**：実行内容を確認 → 「はい」で実行
5. 「ファイル名を変更しコピーが完了しました。」で完了

---

## 子ファイルの検出ルール

ファイル名の先頭が親ファイル名と一致するコンポーネントがコピー対象になります。

```
親: WIDGET-A.sldasm  → 新しい名前: GADGET-B

WIDGET-A-01.sldprt      → ✓ コピー  → GADGET-B-01.sldprt
WIDGET-A-FRAME.sldprt   → ✓ コピー  → GADGET-B-FRAME.sldprt
OTHER-PART.sldprt       → ✗ 対象外（コピーされない）
```

---

## 注意事項

- **実行前に必ずバックアップを取得してください**
- 元のファイルは変更・削除されません（コピーのみ）
- SolidWORKS 2020 以降推奨

---

## Excel 資料の再生成

```bash
pip install openpyxl
python generate_guide.py
```
