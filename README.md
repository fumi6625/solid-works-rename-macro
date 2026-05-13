# SolidWORKS アセンブリ コピー＆リネーム マクロ

SolidWORKS の親アセンブリと子コンポーネントファイルをまとめて新しい名前でコピーする VBA マクロです。  
コピーされた親アセンブリを開くと、コピーされた子ファイルと正しくリンクされた状態になります。

---

## ファイル構成

| ファイル | 説明 |
|---|---|
| `RenameAndCopy.bas` | マクロのメインロジック（VBAモジュール） |
| `frmRenameDialog.frm` | ユーザー入力フォーム（VBA UserForm） |
| `マクロ使用方法ガイド.xlsx` | 操作説明資料（Excel） |
| `generate_guide.py` | Excel資料を生成するPythonスクリプト |

---

## 動作概要

```
① SolidWORKS で親アセンブリを開く
        ↓
② マクロを起動
        ↓
③ ダイアログで新しいアセンブリ名と保存先フォルダを指定
        ↓
④ 実行ボタンをクリック
        ↓
⑤ 親ファイルと子ファイルのコピーが作成される
   ・子ファイルの検出ルール：ファイル名が「親ファイル名」で始まるもの
   ・コピーされた子ファイル名：「新しい親名」＋「元の接尾辞（α）」
        ↓
⑥ 「ファイル名を変更しコピーが完了しました。」と表示されて完了
```

### ファイル名の変換例

```
親: WIDGET-A.sldasm → GADGET-B.sldasm
子: WIDGET-A-01.sldprt → GADGET-B-01.sldprt
子: WIDGET-A-FRAME.sldprt → GADGET-B-FRAME.sldprt
※ OTHER-PART.sldprt → コピーされない（名前が親で始まらない）
```

---

## インストール方法

### 1. SolidWORKS でマクロエディタを開く

`Tools（ツール）` → `Macro（マクロ）` → `Edit Macro（マクロを編集）`

### 2. 新しいマクロプロジェクトを作成

新規作成し、ファイル名を `RenameAndCopy.swp` として保存。

### 3. ファイルをインポート

VBAエディタのメニュー `File → Import File` で以下の順にインポート：
1. `RenameAndCopy.bas`
2. `frmRenameDialog.frm`

### 4. UserForm のコントロールを配置

インポートした `frmRenameDialog` を開き、以下のコントロールを配置してください：

| コントロール名 | 種類 | 用途 |
|---|---|---|
| `lblCurrentLabel` | Label | 「現在のアセンブリ名：」 |
| `lblCurrentName` | Label | 現在のアセンブリ名（自動表示） |
| `lblNewName` | Label | 「新しいアセンブリ名：」 |
| `txtNewName` | TextBox | 新しい名前入力欄 |
| `lblNewNameNote` | Label | 「（拡張子 .sldasm は不要です）」 |
| `lblFolder` | Label | 「保存先フォルダ：」 |
| `txtFolder` | TextBox | フォルダパス表示・入力欄 |
| `btnBrowse` | CommandButton | 「参照...」 |
| `lblPreview` | Label | 対象子ファイル数のプレビュー表示 |
| `btnExecute` | CommandButton | 「実行」 |
| `btnCancel` | CommandButton | 「キャンセル」 |

### 5. 保存

`Ctrl+S` でプロジェクトを保存。

---

## 実行方法

1. SolidWORKS で親アセンブリ（`.sldasm`）を開く
2. `Tools → Macro → Run Macro` で `RenameAndCopy.swp` を選択
3. `StartRenameAndCopy` を選択して「実行」
4. ダイアログに新しいアセンブリ名と保存先フォルダを入力
5. 「実行」ボタンをクリック

---

## 注意事項

- **実行前に必ずバックアップを取得してください**
- 元のファイルは変更・削除されません（コピーのみ）
- 子ファイルが他のプロセスで開かれている場合はコピーに失敗することがあります
- SolidWORKS 2020 以降推奨

---

## Excel 操作説明資料の再生成

```bash
pip install openpyxl
python generate_guide.py
```
