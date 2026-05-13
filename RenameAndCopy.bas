Attribute VB_Name = "RenameAndCopy"
' ============================================================
' SolidWORKS アセンブリ コピー＆リネーム マクロ
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

    Dim col As New Collection
    CollectUniqueComponentPaths swModel, col

    Dim childCount As Long
    childCount = 0
    Dim p As Variant
    For Each p In col
        If IsChildOfParent(CStr(p), parentBase) Then childCount = childCount + 1
    Next p

    Dim startMsg As String
    startMsg = "■ アセンブリ コピー＆リネーム" & vbCrLf & vbCrLf & _
               "現在のアセンブリ名：" & parentBase & ".sldasm" & vbCrLf & _
               "対象子ファイル数　：" & childCount & " 件" & vbCrLf & vbCrLf & _
               "続行しますか？"
    If MsgBox(startMsg, vbQuestion + vbYesNo, "開始確認") <> vbYes Then Exit Sub

    Dim newName As String
    newName = InputBox("新しいアセンブリ名を入力してください。" & vbCrLf & _
                       "（拡張子 .sldasm は不要です）", "新しいアセンブリ名", "")
    If Trim(newName) = "" Then
        MsgBox "キャンセルしました。", vbInformation, "キャンセル"
        Exit Sub
    End If
    newName = Trim(newName)

    If ContainsInvalidChars(newName) Then
        MsgBox "ファイル名に使用できない文字が含まれています。" & vbCrLf & _
               "次の文字は使用できません：" & vbCrLf & _
               "  \  /  :  *  ?  ""  <  >  |", vbExclamation, "入力エラー"
        Exit Sub
    End If

    Dim shellObj  As Object
    Dim folderObj As Object
    Set shellObj  = CreateObject("Shell.Application")
    Set folderObj = shellObj.BrowseForFolder(0, "保存先フォルダを選択してください", 0, parentDir)
    If folderObj Is Nothing Then
        MsgBox "キャンセルしました。", vbInformation, "キャンセル"
        Exit Sub
    End If

    Dim destFolder As String
    destFolder = folderObj.Self.Path
    If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"

    Dim confirmMsg As String
    confirmMsg = "以下の内容でコピーを実行します。" & vbCrLf & vbCrLf & _
                 "現在の親ファイル名：" & parentBase & ".sldasm" & vbCrLf & _
                 "新しい親ファイル名：" & newName & ".sldasm" & vbCrLf & _
                 "保存先フォルダ　　：" & destFolder & vbCrLf & _
                 "対象子ファイル数　：" & childCount & " 件" & vbCrLf & vbCrLf & _
                 "実行しますか？" & vbCrLf & vbCrLf & _
                 "※ 実行前に元ファイルのバックアップを取得することを推奨します。"
    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "実行確認") <> vbYes Then Exit Sub

    Call ExecuteCopyAndRename(swApp, swModel, parentBase, newName, destFolder)

    Exit Sub

ErrNoApp:
    MsgBox "SolidWORKS が起動していません。", vbCritical, "エラー"
End Sub

' ============================================================
' ファイルコピー＋参照更新
' ============================================================
Sub ExecuteCopyAndRename(swApp As Object, swModel As Object, parentBase As String, _
                          newName As String, destFolder As String)

    Dim parentPath As String
    parentPath = swModel.GetPathName()

    ' ---- 1. 全コンポーネントを収集してコピー ----
    Dim col As New Collection
    CollectUniqueComponentPaths swModel, col

    Dim p          As Variant
    Dim copiedCnt  As Long
    Dim copiedAsms As New Collection
    copiedCnt = 0

    For Each p In col
        If IsChildOfParent(CStr(p), parentBase) Then
            Dim destChild As String
            destChild = destFolder & newName & GetAlphaSuffix(CStr(p), parentBase)
            If Not SafeFileCopy(CStr(p), destChild) Then
                MsgBox "コピー失敗：" & vbCrLf & CStr(p), vbCritical, "エラー"
                Exit Sub
            End If
            copiedCnt = copiedCnt + 1
            If LCase(GetExtension(CStr(p))) = ".sldasm" Then
                copiedAsms.Add destChild
            End If
        End If
    Next p

    ' ---- 2. 親アセンブリをコピー ----
    Dim newParentPath As String
    newParentPath = destFolder & newName & GetExtension(parentPath)
    If Not SafeFileCopy(parentPath, newParentPath) Then
        MsgBox "親アセンブリのコピーに失敗しました。", vbCritical, "エラー"
        Exit Sub
    End If

    ' ---- 3. 参照更新：親アセンブリ ----
    Dim ok As Boolean
    ok = UpdateAssemblyReferences(swApp, newParentPath, parentBase, newName, destFolder)

    ' ---- 4. 参照更新：コピーされたサブアセンブリ ----
    Dim asm As Variant
    For Each asm In copiedAsms
        UpdateAssemblyReferences swApp, CStr(asm), parentBase, newName, destFolder
    Next asm

    ' ---- 完了 ----
    If ok Then
        MsgBox "ファイル名を変更しコピーが完了しました。" & vbCrLf & vbCrLf & _
               "保存先　　　　：" & destFolder & vbCrLf & _
               "新しい親ファイル：" & newName & ".sldasm" & vbCrLf & _
               "コピーファイル数：" & copiedCnt & " 件", vbInformation, "完了"
    Else
        MsgBox "ファイルのコピーは完了しました。" & vbCrLf & _
               "ただし、参照の更新に失敗しました。" & vbCrLf & vbCrLf & _
               "保存先：" & destFolder & vbCrLf & vbCrLf & _
               "新しい親アセンブリを開き、" & vbCrLf & _
               "「ファイル → 参照の置換」で手動設定してください。", _
               vbExclamation, "参照更新失敗"
    End If

End Sub

' ============================================================
' アセンブリを開いて子参照を更新して保存する
' ポイント: ReplaceComponents2 の前に新しい子ファイルを
'          OpenDoc6 でロードしておくことが必須
' ============================================================
Function UpdateAssemblyReferences(swApp As Object, assemblyPath As String, _
                                   parentBase As String, newName As String, _
                                   destFolder As String) As Boolean
    UpdateAssemblyReferences = False

    Dim openErr  As Long
    Dim openWarn As Long

    ' ---- アセンブリを開く ----
    Dim asmModel As Object
    Set asmModel = swApp.OpenDoc6(assemblyPath, 2, 0, "", openErr, openWarn)
    If asmModel Is Nothing Then Exit Function

    ' ---- トップレベルコンポーネントを取得 ----
    Dim comps As Variant
    comps = asmModel.GetComponents(True)

    Dim compUb As Long
    compUb = -1
    On Error Resume Next
    compUb = UBound(comps)
    On Error GoTo 0

    ' コンポーネントなし → そのまま保存して終了
    If compUb < 0 Then
        On Error Resume Next
        asmModel.Save3 0, openErr, openWarn
        swApp.CloseDoc assemblyPath
        On Error GoTo 0
        UpdateAssemblyReferences = True
        Exit Function
    End If

    ' ---- 置換リストを構築（ファイルパス重複除去）----
    Dim replArr()    As Variant
    Dim pathArr()    As Variant
    Dim ucArr()      As Variant
    Dim cnArr()      As Variant
    Dim newPaths()   As String     ' 事前ロード用
    ReDim replArr(compUb)
    ReDim pathArr(compUb)
    ReDim ucArr(compUb)
    ReDim cnArr(compUb)
    ReDim newPaths(compUb)

    Dim replCnt  As Long
    replCnt = 0

    Dim seenOld() As String
    ReDim seenOld(compUb)
    Dim seenCnt  As Long
    seenCnt = 0

    Dim ci As Long
    For ci = 0 To compUb
        Dim comp As Object
        Set comp = Nothing
        On Error Resume Next
        Set comp = comps(ci)
        On Error GoTo 0
        If comp Is Nothing Then GoTo NextComp

        Dim oldPath As String
        oldPath = ""
        On Error Resume Next
        oldPath = comp.GetPathName()
        On Error GoTo 0
        If oldPath = "" Then GoTo NextComp
        If Not IsChildOfParent(oldPath, parentBase) Then GoTo NextComp

        ' 重複スキップ（同一ファイルの複数インスタンスは1エントリで足りる）
        Dim seen As Boolean
        seen = False
        Dim si As Long
        For si = 0 To seenCnt - 1
            If LCase(seenOld(si)) = LCase(oldPath) Then seen = True: Exit For
        Next si
        If seen Then GoTo NextComp

        seenOld(seenCnt) = oldPath
        seenCnt = seenCnt + 1

        Dim newChildPath As String
        newChildPath = destFolder & newName & GetAlphaSuffix(oldPath, parentBase)

        Set replArr(replCnt) = comp
        pathArr(replCnt)     = newChildPath
        ucArr(replCnt)       = False
        cnArr(replCnt)       = ""
        newPaths(replCnt)    = newChildPath
        replCnt = replCnt + 1

NextComp:
    Next ci

    ' コピー対象なし
    If replCnt = 0 Then
        On Error Resume Next
        asmModel.Save3 0, openErr, openWarn
        swApp.CloseDoc assemblyPath
        On Error GoTo 0
        UpdateAssemblyReferences = True
        Exit Function
    End If

    ReDim Preserve replArr(replCnt - 1)
    ReDim Preserve pathArr(replCnt - 1)
    ReDim Preserve ucArr(replCnt - 1)
    ReDim Preserve cnArr(replCnt - 1)
    ReDim Preserve newPaths(replCnt - 1)

    ' ============================================================
    ' 【重要】ReplaceComponents2 の前に新しい子ファイルを
    '         OpenDoc6 で SolidWORKS にロードしておく
    '         ロードされていないと ReplaceComponents2 が False を返す
    ' ============================================================
    Dim loadedPaths() As String
    ReDim loadedPaths(replCnt - 1)
    Dim loadedCnt As Long
    loadedCnt = 0

    Dim li As Long
    For li = 0 To replCnt - 1
        Dim docType As Long
        docType = GetDocTypeFromPath(newPaths(li))
        If docType > 0 Then
            Dim loadErr  As Long
            Dim loadWarn As Long
            Dim childDoc As Object
            Set childDoc = Nothing
            On Error Resume Next
            Set childDoc = swApp.OpenDoc6(newPaths(li), docType, 0, "", loadErr, loadWarn)
            On Error GoTo 0
            If Not childDoc Is Nothing Then
                loadedPaths(loadedCnt) = newPaths(li)
                loadedCnt = loadedCnt + 1
            End If
        End If
    Next li

    ' 親アセンブリをアクティブにする
    On Error Resume Next
    swApp.ActivateDoc3 GetFileNameFromPath(assemblyPath), True, 0, openErr
    On Error GoTo 0

    ' ---- ReplaceComponents2 で参照を一括更新 ----
    ' matchByName=True で同一ファイルの全インスタンスを置換
    Dim bReplaced As Boolean
    Dim rc2Err    As Long
    bReplaced = False
    rc2Err = 0

    On Error Resume Next
    bReplaced = asmModel.ReplaceComponents2(replArr, pathArr, ucArr, cnArr, True)
    rc2Err = Err.Number
    On Error GoTo 0

    ' リビルドして変更を確定
    If bReplaced And rc2Err = 0 Then
        On Error Resume Next
        asmModel.ForceRebuild3 False
        On Error GoTo 0
    End If

    ' ---- 保存して閉じる ----
    Dim saveErr  As Long
    Dim saveWarn As Long
    On Error Resume Next
    asmModel.Save3 0, saveErr, saveWarn
    swApp.CloseDoc assemblyPath
    On Error GoTo 0

    ' ---- 事前ロードした子ファイルを閉じる ----
    Dim lj As Long
    For lj = 0 To loadedCnt - 1
        On Error Resume Next
        swApp.CloseDoc loadedPaths(lj)
        On Error GoTo 0
    Next lj

    UpdateAssemblyReferences = (bReplaced And rc2Err = 0)

End Function

' ============================================================
' 拡張子から SolidWORKS ドキュメント種別を返す
' ============================================================
Function GetDocTypeFromPath(filePath As String) As Long
    Select Case LCase(GetExtension(filePath))
        Case ".sldprt": GetDocTypeFromPath = 1
        Case ".sldasm": GetDocTypeFromPath = 2
        Case ".slddrw": GetDocTypeFromPath = 3
        Case Else:      GetDocTypeFromPath = 0
    End Select
End Function

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
' 全コンポーネントパスを重複なく収集
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
                If item = path Then dup = True: Exit For
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
' ============================================================
Function GetAlphaSuffix(filePath As String, parentBase As String) As String
    Dim baseName As String
    baseName = GetBaseNameFromPath(filePath)
    GetAlphaSuffix = Mid(baseName, Len(parentBase) + 1) & GetExtension(filePath)
End Function

' ============================================================
' ファイル名に使用できない文字が含まれているか確認
' ============================================================
Function ContainsInvalidChars(name As String) As Boolean
    Dim invalidChars As String
    Dim i As Integer
    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        If InStr(name, Mid(invalidChars, i, 1)) > 0 Then
            ContainsInvalidChars = True: Exit Function
        End If
    Next i
    ContainsInvalidChars = False
End Function

Function GetFolderFromPath(filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, "\")
    If pos > 0 Then GetFolderFromPath = Left(filePath, pos) Else GetFolderFromPath = ""
End Function

Function GetBaseNameFromPath(filePath As String) As String
    Dim fileName As String
    Dim dotPos   As Long
    fileName = GetFileNameFromPath(filePath)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then GetBaseNameFromPath = Left(fileName, dotPos - 1) Else GetBaseNameFromPath = fileName
End Function

Function GetFileNameFromPath(filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, "\")
    If pos > 0 Then GetFileNameFromPath = Mid(filePath, pos + 1) Else GetFileNameFromPath = filePath
End Function

Function GetExtension(filePath As String) As String
    Dim fileName As String
    Dim dotPos   As Long
    fileName = GetFileNameFromPath(filePath)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then GetExtension = Mid(fileName, dotPos) Else GetExtension = ""
End Function
