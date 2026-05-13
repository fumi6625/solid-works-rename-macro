Attribute VB_Name = "RenameAndCopy"
' ============================================================
' SolidWORKS アセンブリ コピー＆リネーム マクロ
' ------------------------------------------------------------
' 機能:
'   親アセンブリを開いた状態でマクロを起動し、新しい名前と
'   保存先フォルダを指定すると、親アセンブリおよびファイル名が
'   「親ファイル名＋α」となっている子コンポーネントを
'   「新しい親ファイル名＋α」の名前でコピーする。
'   コピー後の親アセンブリはコピーされた子ファイルを参照する。
'
' 対応バージョン: SolidWORKS 2015 以降
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
    Dim shellObj  As Object
    Dim folderObj As Object
    Set shellObj  = CreateObject("Shell.Application")
    Set folderObj = shellObj.BrowseForFolder( _
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
    Call ExecuteCopyAndRename(swApp, swModel, parentBase, newName, destFolder)

    Exit Sub

ErrNoApp:
    MsgBox "SolidWORKS が起動していません。" & vbCrLf & _
           "SolidWORKS を起動してからマクロを実行してください。", _
           vbCritical, "エラー"
End Sub

' ============================================================
' ファイルコピー＋参照更新の実行
' ============================================================
Sub ExecuteCopyAndRename(swApp As Object, swModel As Object, parentBase As String, _
                          newName As String, destFolder As String)

    Dim parentPath As String
    parentPath = swModel.GetPathName()

    ' ---- 1. 子ファイルをコピー（バイナリI/Oでロック中も対応）----
    Dim col As New Collection
    CollectUniqueComponentPaths swModel, col

    Dim p         As Variant
    Dim copiedCnt As Long
    copiedCnt = 0

    For Each p In col
        If IsChildOfParent(CStr(p), parentBase) Then
            Dim destChild As String
            destChild = destFolder & newName & GetAlphaSuffix(CStr(p), parentBase)
            If Not SafeFileCopy(CStr(p), destChild) Then
                MsgBox "ファイルのコピーに失敗しました：" & vbCrLf & CStr(p), vbCritical, "エラー"
                Exit Sub
            End If
            copiedCnt = copiedCnt + 1
        End If
    Next p

    ' ---- 2. 親アセンブリをコピー ----
    Dim newParentPath As String
    newParentPath = destFolder & newName & GetExtension(parentPath)

    If Not SafeFileCopy(parentPath, newParentPath) Then
        MsgBox "親アセンブリのコピーに失敗しました。", vbCritical, "エラー"
        Exit Sub
    End If

    ' ---- 3. コピーした親アセンブリを開く ----
    ' （元ファイルが残っているため参照エラーなしで開ける）
    Dim openErr  As Long
    Dim openWarn As Long
    Dim newModel As Object
    Set newModel = swApp.OpenDoc6(newParentPath, 2, 0, "", openErr, openWarn)

    If newModel Is Nothing Then
        MsgBox "コピーした親アセンブリを開けませんでした。" & vbCrLf & _
               "エラー番号：" & openErr & vbCrLf & vbCrLf & _
               "ファイルのコピーは完了しています。" & vbCrLf & _
               "保存先：" & destFolder, vbCritical, "エラー"
        Exit Sub
    End If

    ' ---- 4. コンポーネントの参照を新しいパスに更新 ----
    Dim newComps As Variant
    newComps = newModel.GetComponents(True)

    Dim compOk As Boolean
    compOk = False
    On Error Resume Next
    Dim ub As Long
    ub = UBound(newComps)
    If Err.Number = 0 Then compOk = True
    On Error GoTo 0

    If compOk Then

        ' ---- 置換リストを構築 ----
        Dim replArr()  As Variant
        Dim pathArr()  As Variant
        Dim ucArr()    As Variant
        Dim cnArr()    As Variant
        Dim replCnt    As Long
        replCnt = 0

        ReDim replArr(ub)
        ReDim pathArr(ub)
        ReDim ucArr(ub)
        ReDim cnArr(ub)

        Dim ci As Long
        For ci = LBound(newComps) To ub
            On Error Resume Next
            Dim comp As Object
            Set comp = newComps(ci)
            On Error GoTo 0

            If Not comp Is Nothing Then
                Dim oldPath As String
                oldPath = ""
                On Error Resume Next
                oldPath = comp.GetPathName()
                On Error GoTo 0

                If oldPath <> "" And IsChildOfParent(oldPath, parentBase) Then
                    Set replArr(replCnt) = comp
                    pathArr(replCnt) = destFolder & newName & GetAlphaSuffix(oldPath, parentBase)
                    ucArr(replCnt)   = False
                    cnArr(replCnt)   = ""
                    replCnt = replCnt + 1
                End If
            End If
        Next ci

        If replCnt > 0 Then
            ReDim Preserve replArr(replCnt - 1)
            ReDim Preserve pathArr(replCnt - 1)
            ReDim Preserve ucArr(replCnt - 1)
            ReDim Preserve cnArr(replCnt - 1)

            ' ReplaceComponents2 で参照を一括更新
            On Error Resume Next
            newModel.ReplaceComponents2 replArr, pathArr, ucArr, cnArr, False
            Dim rc2Err As Long
            rc2Err = Err.Number
            On Error GoTo 0

            ' 失敗した場合はコンポーネントごとに ReplaceReferencedDocument を試みる
            If rc2Err <> 0 Then
                Dim ri As Long
                For ri = 0 To replCnt - 1
                    Dim rSrc As Object
                    Set rSrc = replArr(ri)
                    Dim rOld As String
                    Dim rNew As String
                    rOld = rSrc.GetPathName()
                    rNew = CStr(pathArr(ri))
                    On Error Resume Next
                    newModel.Extension.ReplaceReferencedDocument rOld, rNew
                    On Error GoTo 0
                Next ri
            End If
        End If

    End If

    ' ---- 5. 保存して閉じる ----
    Dim saveErr  As Long
    Dim saveWarn As Long
    On Error Resume Next
    newModel.Save3 1, saveErr, saveWarn
    On Error GoTo 0

    swApp.CloseDoc newParentPath

    ' ---- 完了メッセージ ----
    MsgBox "ファイル名を変更しコピーが完了しました。" & vbCrLf & vbCrLf & _
           "保存先　　　　：" & destFolder & vbCrLf & _
           "新しい親ファイル：" & newName & ".sldasm" & vbCrLf & _
           "コピーファイル数：" & copiedCnt & " 件", _
           vbInformation, "完了"

End Sub

' ============================================================
' バイナリI/Oによるファイルコピー（ロック中ファイルにも対応）
' ============================================================
Function SafeFileCopy(srcPath As String, destPath As String) As Boolean

    Const CHUNK_SIZE As Long = 32768

    Dim srcNo    As Integer
    Dim destNo   As Integer
    Dim buf()    As Byte
    Dim fileLen  As Long
    Dim pos      As Long
    Dim chunkLen As Long

    On Error GoTo CopyFailed

    srcNo = FreeFile
    Open srcPath For Binary Access Read Shared As #srcNo

    destNo = FreeFile
    Open destPath For Binary Access Write As #destNo

    fileLen = LOF(srcNo)
    pos = 0

    Do While pos < fileLen
        chunkLen = CHUNK_SIZE
        If pos + chunkLen > fileLen Then chunkLen = fileLen - pos
        ReDim buf(chunkLen - 1)
        Get #srcNo, pos + 1, buf
        Put #destNo, pos + 1, buf
        pos = pos + chunkLen
    Loop

    Close #srcNo
    Close #destNo
    SafeFileCopy = True
    Exit Function

CopyFailed:
    On Error Resume Next
    Close #srcNo
    Close #destNo
    On Error GoTo 0
    SafeFileCopy = False

End Function

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
' ユーティリティ：拡張子を取得（例: ".sldasm"）
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
