Attribute VB_Name = "RenameAndCopy"
' ============================================================
' SolidWORKS アセンブリ コピー＆リネーム マクロ
' ------------------------------------------------------------
' 機能:
'   親アセンブリを開いた状態でマクロを起動し、新しい名前と
'   保存先フォルダを指定すると、親アセンブリおよびファイル名が
'   「親ファイル名＋α」となっている子コンポーネントを
'   「新しい親ファイル名＋α」の名前でコピーする。
'   コピーされた親アセンブリは子ファイルと正しくリンクされる。
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

    ' ---- 子ファイル数カウント ----
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

    ' ---- 実行 ----
    Call ExecuteRenameAndCopy(swApp, swModel, parentPath, parentDir, _
                              parentBase, newName, destFolder, col)
    Exit Sub

ErrNoApp:
    MsgBox "SolidWORKS が起動していません。" & vbCrLf & _
           "SolidWORKS を起動してからマクロを実行してください。", _
           vbCritical, "エラー"
End Sub

' ============================================================
' メイン実行処理
' ============================================================
Sub ExecuteRenameAndCopy(swApp As Object, swModel As Object, _
                         parentPath As String, parentDir As String, _
                         parentBase As String, newName As String, _
                         destFolder As String, col As Collection)

    Dim oldPath      As String
    Dim newPath      As String
    Dim alpha        As String
    Dim newAssyPath  As String
    Dim newModel     As Object
    Dim i            As Long
    Dim successCount As Long
    Dim failCount    As Long
    Dim overwriteAll As Boolean
    Dim skipAll      As Boolean

    successCount = 0
    failCount    = 0
    overwriteAll = False
    skipAll      = False

    ' 対象ファイルのマッピング作成
    Dim mapCount  As Long
    Dim mapOld()  As String
    Dim mapNew()  As String
    mapCount = 0

    ReDim mapOld(col.Count - 1)
    ReDim mapNew(col.Count - 1)

    Dim p As Variant
    For Each p In col
        oldPath = CStr(p)
        If IsChildOfParent(oldPath, parentBase) Then
            alpha = GetAlphaSuffix(oldPath, parentBase)
            mapOld(mapCount) = oldPath
            mapNew(mapCount) = destFolder & newName & alpha
            mapCount = mapCount + 1
        End If
    Next p

    ' ---- 子ファイルをコピー ----
    For i = 0 To mapCount - 1
        Dim src As String
        Dim dst As String
        src = mapOld(i)
        dst = mapNew(i)

        If Not FileExists(src) Then
            failCount = failCount + 1
            Debug.Print "スキップ（ファイルなし）: " & src

        ElseIf FileExists(dst) And Not overwriteAll And Not skipAll Then
            Dim ans As Integer
            ans = MsgBox("同名ファイルが既に存在します：" & vbCrLf & dst & vbCrLf & vbCrLf & _
                         "[はい] 上書き" & vbCrLf & _
                         "[いいえ] このファイルのみスキップ" & vbCrLf & _
                         "[キャンセル] 以降全てスキップ", _
                         vbQuestion + vbYesNoCancel, "上書き確認")
            Select Case ans
                Case vbYes
                    If Not SafeCopyFile(src, dst) Then
                        failCount = failCount + 1
                        MsgBox "コピーに失敗しました：" & vbCrLf & src, vbExclamation, "コピー失敗"
                    Else
                        successCount = successCount + 1
                    End If
                Case vbNo
                    ' このファイルだけスキップ
                Case vbCancel
                    skipAll = True
            End Select

        ElseIf Not skipAll Then
            If Not SafeCopyFile(src, dst) Then
                failCount = failCount + 1
                MsgBox "コピーに失敗しました：" & vbCrLf & src, vbExclamation, "コピー失敗"
            Else
                successCount = successCount + 1
                Debug.Print "コピー完了: " & src & " → " & dst
            End If
        End If
    Next i

    ' ---- 親アセンブリをコピー ----
    newAssyPath = destFolder & newName & ".sldasm"

    If FileExists(newAssyPath) Then
        Dim assyAns As Integer
        assyAns = MsgBox("親アセンブリの保存先に同名ファイルが存在します：" & vbCrLf & _
                         newAssyPath & vbCrLf & "上書きしますか？", _
                         vbQuestion + vbYesNo, "上書き確認")
        If assyAns <> vbYes Then
            MsgBox "処理をキャンセルしました。", vbInformation, "キャンセル"
            Exit Sub
        End If
    End If

    If Not SafeCopyFile(parentPath, newAssyPath) Then
        MsgBox "親アセンブリのコピーに失敗しました：" & vbCrLf & parentPath, vbCritical, "エラー"
        Exit Sub
    End If

    ' ---- コピーした親を開いて参照を更新 ----
    Dim openErrors   As Long
    Dim openWarnings As Long
    Set newModel = swApp.OpenDoc6(newAssyPath, 2, 1, "", openErrors, openWarnings)

    If newModel Is Nothing Then
        MsgBox "コピーした親アセンブリを開けませんでした。" & vbCrLf & _
               "手動で外部参照を更新してください。", vbExclamation, "警告"
        Exit Sub
    End If

    For i = 0 To mapCount - 1
        If FileExists(mapNew(i)) Then
            newModel.ReplaceReferencedDocument mapOld(i), mapNew(i)
        End If
    Next i

    ' 保存して閉じる
    Dim saveErrors   As Long
    Dim saveWarnings As Long
    newModel.Save3 1, saveErrors, saveWarnings
    swApp.CloseDoc newAssyPath

    ' ---- 完了メッセージ ----
    MsgBox "ファイル名を変更しコピーが完了しました。" & vbCrLf & vbCrLf & _
           "保存先　　　　：" & destFolder & vbCrLf & _
           "新しい親ファイル：" & newName & ".sldasm" & vbCrLf & _
           "子ファイルコピー：" & successCount & " 件", _
           vbInformation, "完了"

End Sub

' ============================================================
' 子コンポーネントのパスを重複なく収集する
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

' ============================================================
' ファイルコピー（SolidWORKS ロック中ファイル対応版）
' ============================================================
' SolidWORKS が開いているファイルは FileCopy でエラー70が発生するため、
' VBA バイナリI/O（Open...For Binary）でコピーする。
' バイナリI/O は共有読み取りアクセスを使用するため、
' 他のプロセスがファイルを開いていてもコピーできる。
' ============================================================
Function SafeCopyFile(src As String, dst As String) As Boolean

    Const CHUNK_SIZE As Long = 524288  ' 512KB ずつ読み書き

    Dim fn1      As Integer
    Dim fn2      As Integer
    Dim buf()    As Byte
    Dim remaining As Long
    Dim chunkLen As Long
    Dim fileLen  As Long

    SafeCopyFile = False

    On Error GoTo CopyFailed

    fn1 = FreeFile
    Open src For Binary Access Read As #fn1
    fileLen = LOF(fn1)

    fn2 = FreeFile
    Open dst For Binary Access Write As #fn2

    ' ファイルサイズが 0 の場合もそのままコピー（空ファイル）
    remaining = fileLen
    Do While remaining > 0
        chunkLen = remaining
        If chunkLen > CHUNK_SIZE Then chunkLen = CHUNK_SIZE
        ReDim buf(chunkLen - 1)
        Get #fn1, , buf
        Put #fn2, , buf
        remaining = remaining - chunkLen
    Loop

    Close #fn1
    Close #fn2
    SafeCopyFile = True
    Exit Function

CopyFailed:
    On Error Resume Next
    Close #fn1
    Close #fn2
    Debug.Print "SafeCopyFile 失敗: " & src & " → " & dst & " (Err=" & Err.Number & ": " & Err.Description & ")"
    SafeCopyFile = False

End Function
