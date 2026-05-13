"""
generate_bas.py
RenameAndCopy.bas を Shift-JIS (CP932) エンコーディングで生成する。

VBA エディタ（SolidWORKS）は Windows の ANSI エンコーディング
（日本語環境では CP932 / Shift-JIS）でファイルを読み込む。
このスクリプトを実行することで、日本語が文字化けしない .bas ファイルを生成する。

実行方法:
    python generate_bas.py
"""

VBA_CODE = r"""Attribute VB_Name = "RenameAndCopy"
' ============================================================
' SolidWORKS アセンブリ コピー＆リネーム マクロ
' ------------------------------------------------------------
' 機能:
'   親アセンブリを開いた状態でマクロを起動し、新しい名前と
'   保存先フォルダを指定すると、親アセンブリおよびファイル名が
'   「親ファイル名＋α」となっている子コンポーネントを
'   「新しい親ファイル名＋α」の名前でコピーする。
'   Pack and Go API を使用するため、コピー後の親アセンブリは
'   コピーされた子ファイルと正しくリンクされる。
'
' 対応バージョン: SolidWORKS 2020 以降
' インポート方法: VBA エディタで [ファイル] → [ファイルのインポート]
' ============================================================

Option Explicit

' ============================================================
' エントリーポイント
' ============================================================
Sub StartRenameAndCopy()

    Dim swApp   As Object
    Dim swModel As Object

    On Error GoTo ErrNoApp
    Set swApp = CreateObject("SldWorks.Application")
    On Error GoTo 0

    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "SolidWORKS でアセンブリファイル（.sldasm）を開いてから" & vbCrLf & _
               "マクロを実行してください。", vbCritical, "エラー"
        Exit Sub
    End If

    If swModel.GetType() <> 2 Then
        MsgBox "アセンブリファイル（.sldasm）を開いてから" & vbCrLf & _
               "マクロを実行してください。" & vbCrLf & vbCrLf & _
               "現在開いているファイルはアセンブリではありません。", _
               vbCritical, "エラー"
        Exit Sub
    End If

    Dim parentPath As String
    Dim parentDir  As String
    Dim parentBase As String

    parentPath = swModel.GetPathName()
    parentDir  = GetFolderFromPath(parentPath)
    parentBase = GetBaseNameFromPath(parentPath)

    ' ---- 子ファイル数カウント（確認ダイアログ用）----
    Dim col As New Collection
    CollectUniqueComponentPaths swModel, col

    Dim childCount As Long
    childCount = 0
    Dim p As Variant
    For Each p In col
        If IsChildOfParent(CStr(p), parentBase) Then
            childCount = childCount + 1
        End If
    Next p

    ' ---- ステップ1：開始確認 ----
    Dim startMsg As String
    startMsg = "■ アセンブリ コピー＆リネーム" & vbCrLf & vbCrLf & _
               "現在のアセンブリ名：" & parentBase & ".sldasm" & vbCrLf & _
               "対象子ファイル数　：" & childCount & " 件" & vbCrLf & vbCrLf & _
               "続行しますか？"

    If MsgBox(startMsg, vbQuestion + vbYesNo, "開始確認") <> vbYes Then
        Exit Sub
    End If

    ' ---- ステップ2：新しいアセンブリ名の入力 ----
    Dim newName As String
    newName = InputBox("新しいアセンブリ名を入力してください。" & vbCrLf & _
                       "（拡張子 .sldasm は不要です）", _
                       "新しいアセンブリ名", "")

    If Trim(newName) = "" Then
        MsgBox "キャンセルしました。", vbInformation, "キャンセル"
        Exit Sub
    End If
    newName = Trim(newName)

    If ContainsInvalidChars(newName) Then
        MsgBox "ファイル名に使用できない文字が含まれています。" & vbCrLf & _
               "次の文字は使用できません：" & vbCrLf & _
               "  \  /  :  *  ?  ""  <  >  |", _
               vbExclamation, "入力エラー"
        Exit Sub
    End If

    ' ---- ステップ3：保存先フォルダの選択 ----
    Dim shell     As Object
    Dim folderObj As Object
    Set shell = CreateObject("Shell.Application")
    Set folderObj = shell.BrowseForFolder( _
        0, "保存先フォルダを選択してください", 0, parentDir)

    If folderObj Is Nothing Then
        MsgBox "キャンセルしました。", vbInformation, "キャンセル"
        Exit Sub
    End If

    Dim destFolder As String
    destFolder = folderObj.Self.Path
    If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"

    ' ---- ステップ4：実行確認 ----
    Dim confirmMsg As String
    confirmMsg = "以下の内容でコピーを実行します。" & vbCrLf & vbCrLf & _
                 "現在の親ファイル名：" & parentBase & ".sldasm" & vbCrLf & _
                 "新しい親ファイル名：" & newName & ".sldasm" & vbCrLf & _
                 "保存先フォルダ　　：" & destFolder & vbCrLf & _
                 "対象子ファイル数　：" & childCount & " 件" & vbCrLf & vbCrLf & _
                 "実行しますか？" & vbCrLf & vbCrLf & _
                 "※ 実行前に元ファイルのバックアップを取得することを推奨します。"

    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "実行確認") <> vbYes Then
        Exit Sub
    End If

    ' ---- 実行（Pack and Go を使用）----
    Call ExecutePackAndGo(swModel, parentBase, newName, destFolder)

    Exit Sub

ErrNoApp:
    MsgBox "SolidWORKS が起動していません。" & vbCrLf & _
           "SolidWORKS を起動してからマクロを実行してください。", _
           vbCritical, "エラー"
End Sub

' ============================================================
' Pack and Go を使用してコピー＆参照更新を一括実行
' ============================================================
' SolidWORKS の Pack and Go API はファイルのコピーと
' アセンブリ内の参照パス更新を一括で処理する。
' 手動で ReplaceReferencedDocument を呼ぶ必要がない。
' ============================================================
Sub ExecutePackAndGo(swModel As Object, parentBase As String, _
                     newName As String, destFolder As String)

    ' Pack and Go オブジェクトを取得
    ' SolidWORKS のバージョンにより引数の型・有無が異なるため
    ' 複数のパターンを順に試みる
    Dim packAndGo As Object
    Dim pgErr     As Long

    On Error Resume Next

    Err.Clear
    Set packAndGo = swModel.Extension.GetPackAndGo(0&)
    pgErr = Err.Number

    If pgErr <> 0 Or packAndGo Is Nothing Then
        Err.Clear
        Set packAndGo = swModel.Extension.GetPackAndGo(1&)
        pgErr = Err.Number
    End If

    If pgErr <> 0 Or packAndGo Is Nothing Then
        Err.Clear
        Set packAndGo = swModel.Extension.GetPackAndGo(True)
        pgErr = Err.Number
    End If

    If pgErr <> 0 Or packAndGo Is Nothing Then
        Err.Clear
        Set packAndGo = swModel.Extension.GetPackAndGo()
        pgErr = Err.Number
    End If

    On Error GoTo 0

    If packAndGo Is Nothing Then
        MsgBox "Pack and Go API が利用できません。" & vbCrLf & _
               "エラー番号：" & pgErr & vbCrLf & _
               "SolidWORKS のバージョンを確認してください。", vbCritical, "エラー"
        Exit Sub
    End If

    ' アセンブリに含まれる全ファイルリストを取得
    Dim nCount As Long
    nCount = packAndGo.GetDocumentCount()

    If nCount = 0 Then
        MsgBox "コンポーネントが見つかりません。" & vbCrLf & _
               "アセンブリが正しく読み込まれているか確認してください。", _
               vbExclamation, "警告"
        Exit Sub
    End If

    ' Pack and Go API は Variant 型配列を使用する
    Dim fileNames As Variant
    packAndGo.GetFileNames fileNames

    ' 親名プレフィックスに一致するファイルのパスを新しい名前に変更
    ' （一致しないファイルは元パスのまま → コピーされず元の場所を参照）
    Dim i        As Long
    Dim renamed  As Long
    renamed = 0

    For i = 0 To nCount - 1
        Dim bn As String
        bn = GetBaseNameFromPath(CStr(fileNames(i)))

        If StrComp(Left(bn, Len(parentBase)), parentBase, vbTextCompare) = 0 Then
            Dim alpha As String
            alpha = GetAlphaSuffix(CStr(fileNames(i)), parentBase)
            fileNames(i) = destFolder & newName & alpha
            renamed = renamed + 1
        End If
    Next i

    If renamed = 0 Then
        MsgBox "コピー対象のファイルが見つかりませんでした。" & vbCrLf & vbCrLf & _
               "子ファイル名が親ファイル名（" & parentBase & "）で" & vbCrLf & _
               "始まっているか確認してください。", vbExclamation, "対象なし"
        Exit Sub
    End If

    ' 変更後のファイルリストを Pack and Go にセット
    packAndGo.SetFileNames fileNames

    ' Pack and Go を実行
    ' （ファイルコピー＋アセンブリ内参照パス更新を自動処理）
    ' Save() のエラーコードも Variant 型配列で受け取る
    Dim errors  As Variant
    Dim nErrors As Long

    On Error GoTo PackAndGoFailed
    nErrors = packAndGo.Save(errors)
    On Error GoTo 0

    ' 結果表示
    If nErrors > 0 Then
        ' エラーが出たファイルを収集
        Dim errList As String
        errList = ""
        Dim j As Long
        For j = 0 To nCount - 1
            If IsArray(errors) Then
                If j <= UBound(errors) Then
                    If CLng(errors(j)) <> 0 Then
                        errList = errList & "・" & GetFileNameFromPath(CStr(fileNames(j))) & vbCrLf
                    End If
                End If
            End If
        Next j
        MsgBox "処理は完了しましたが、一部エラーがありました（" & nErrors & "件）：" & _
               vbCrLf & vbCrLf & errList & vbCrLf & _
               "保存先フォルダを確認してください。", vbExclamation, "警告"
    Else
        MsgBox "ファイル名を変更しコピーが完了しました。" & vbCrLf & vbCrLf & _
               "保存先　　　　：" & destFolder & vbCrLf & _
               "新しい親ファイル：" & newName & ".sldasm" & vbCrLf & _
               "コピーファイル数：" & renamed & " 件", _
               vbInformation, "完了"
    End If

    Exit Sub

PackAndGoFailed:
    MsgBox "Pack and Go の実行中にエラーが発生しました。" & vbCrLf & _
           "エラー番号：" & Err.Number & vbCrLf & _
           "内容：" & Err.Description, vbCritical, "エラー"

End Sub

' ============================================================
' 子コンポーネントのパスを重複なく収集する（確認ダイアログ用）
' ============================================================
Sub CollectUniqueComponentPaths(swModel As Object, ByRef col As Collection)

    Dim comps As Variant
    Dim c     As Variant
    Dim path  As String
    Dim dup   As Boolean
    Dim item  As Variant

    comps = swModel.GetComponents(False)
    If IsEmpty(comps) Then Exit Sub

    For Each c In comps
        path = c.GetPathName()
        If path <> "" Then
            dup = False
            For Each item In col
                If item = path Then
                    dup = True
                    Exit For
                End If
            Next item
            If Not dup Then col.Add path
        End If
    Next c

End Sub

' ============================================================
' ファイル名の先頭が親ファイル名と一致するか判定
' ============================================================
Function IsChildOfParent(filePath As String, parentBase As String) As Boolean

    Dim fileName As String
    fileName = GetBaseNameFromPath(filePath)

    If Len(fileName) >= Len(parentBase) Then
        If StrComp(Left(fileName, Len(parentBase)), parentBase, vbTextCompare) = 0 Then
            IsChildOfParent = True
            Exit Function
        End If
    End If

    IsChildOfParent = False

End Function

' ============================================================
' 子ファイルからα部分（接尾辞＋拡張子）を取得
' 例: 親="WIDGET-A", 子="WIDGET-A-01.sldprt" → "-01.sldprt"
' ============================================================
Function GetAlphaSuffix(filePath As String, parentBase As String) As String

    Dim baseName As String
    baseName = GetBaseNameFromPath(filePath)

    Dim alpha As String
    alpha = Mid(baseName, Len(parentBase) + 1)

    GetAlphaSuffix = alpha & GetExtension(filePath)

End Function

' ============================================================
' ファイル名に使用できない文字が含まれているか確認
' ============================================================
Function ContainsInvalidChars(name As String) As Boolean

    Dim invalidChars As String
    Dim i            As Integer
    invalidChars = "\/:*?""<>|"

    For i = 1 To Len(invalidChars)
        If InStr(name, Mid(invalidChars, i, 1)) > 0 Then
            ContainsInvalidChars = True
            Exit Function
        End If
    Next i

    ContainsInvalidChars = False

End Function

' ============================================================
' ユーティリティ：フォルダパスを取得
' ============================================================
Function GetFolderFromPath(filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        GetFolderFromPath = Left(filePath, pos)
    Else
        GetFolderFromPath = ""
    End If
End Function

' ============================================================
' ユーティリティ：拡張子なしファイル名を取得
' ============================================================
Function GetBaseNameFromPath(filePath As String) As String
    Dim fileName As String
    Dim dotPos   As Long
    fileName = GetFileNameFromPath(filePath)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetBaseNameFromPath = Left(fileName, dotPos - 1)
    Else
        GetBaseNameFromPath = fileName
    End If
End Function

' ============================================================
' ユーティリティ：ファイル名（拡張子あり）を取得
' ============================================================
Function GetFileNameFromPath(filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        GetFileNameFromPath = Mid(filePath, pos + 1)
    Else
        GetFileNameFromPath = filePath
    End If
End Function

' ============================================================
' ユーティリティ：拡張子を取得（例: ".sldprt"）
' ============================================================
Function GetExtension(filePath As String) As String
    Dim fileName As String
    Dim dotPos   As Long
    fileName = GetFileNameFromPath(filePath)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetExtension = Mid(fileName, dotPos)
    Else
        GetExtension = ""
    End If
End Function

' ============================================================
' ユーティリティ：ファイル存在確認
' ============================================================
Function FileExists(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function
"""

def main():
    output_path = "RenameAndCopy.bas"
    with open(output_path, "w", encoding="cp932") as f:
        f.write(VBA_CODE)
    print(f"生成完了（CP932エンコーディング）：{output_path}")

if __name__ == "__main__":
    main()
