Attribute VB_Name = "RenameAndCopy"
' ============================================================
' SolidWORKS アセンブリ コピー＆リネーム マクロ
' ------------------------------------------------------------
' 機能:
'   親アセンブリを開いた状態でマクロを起動し、新しい名前と
'   保存先フォルダを指定すると、親アセンブリおよびファイル名
'   が「親ファイル名＋α」となっている子コンポーネントを
'   「新しい親ファイル名＋α」の名前でコピーする。
'   コピーされた親アセンブリを開くと、コピーされた子ファイルと
'   正しくリンクされた状態になっている。
'
' 対応バージョン: SolidWORKS 2020 以降
' ============================================================

Option Explicit

' フォームと共有するグローバル変数
Public g_swApp      As Object
Public g_swModel    As Object
Public g_parentBase As String   ' 親ファイルのベース名（拡張子・パスなし）
Public g_parentPath As String   ' 親ファイルのフルパス
Public g_parentDir  As String   ' 親ファイルのフォルダパス

' ============================================================
' エントリーポイント
' ============================================================
Sub StartRenameAndCopy()

    ' SolidWORKS アプリケーションを取得
    On Error GoTo ErrNoApp
    Set g_swApp = CreateObject("SldWorks.Application")
    On Error GoTo 0

    ' アクティブドキュメントを取得
    Set g_swModel = g_swApp.ActiveDoc

    If g_swModel Is Nothing Then
        MsgBox "SolidWORKS でアセンブリファイル（.sldasm）を開いてからマクロを実行してください。", _
               vbCritical, "エラー"
        Exit Sub
    End If

    ' アセンブリかどうか確認（type 2 = Assembly）
    If g_swModel.GetType() <> 2 Then
        MsgBox "アセンブリファイル（.sldasm）を開いてからマクロを実行してください。" & vbCrLf & _
               "現在開いているファイルはアセンブリではありません。", _
               vbCritical, "エラー"
        Exit Sub
    End If

    ' 親ファイルの情報を設定
    g_parentPath = g_swModel.GetPathName()
    g_parentDir  = GetFolderFromPath(g_parentPath)
    g_parentBase = GetBaseNameFromPath(g_parentPath)

    ' フォームを表示
    frmRenameDialog.Show

    Exit Sub

ErrNoApp:
    MsgBox "SolidWORKS が起動していません。SolidWORKS を起動してからマクロを実行してください。", _
           vbCritical, "エラー"
End Sub

' ============================================================
' メイン実行処理（フォームの「実行」ボタンから呼び出される）
' ============================================================
Sub ExecuteRenameAndCopy(newName As String, destFolder As String)

    Dim swAssembly   As Object
    Dim components() As Object
    Dim comp         As Object
    Dim oldPath      As String
    Dim newPath      As String
    Dim alpha        As String
    Dim newAssyPath  As String
    Dim newModel     As Object
    Dim errors       As Long
    Dim warnings     As Long
    Dim i            As Long
    Dim overwriteAll As Boolean
    Dim skipAll      As Boolean

    ' 保存先フォルダの末尾に区切り文字を付与
    If Right(destFolder, 1) <> "\" Then
        destFolder = destFolder & "\"
    End If

    Set swAssembly = g_swModel

    ' ========== 子コンポーネントのパス収集（重複排除）==========
    Dim uniquePaths As New Collection
    CollectUniqueComponentPaths swAssembly, uniquePaths

    ' ========== 対象ファイルをフィルタリング ==================
    ' 親ファイル名をプレフィックスとして持つ子コンポーネントのみ対象
    Dim mappingOld() As String
    Dim mappingNew() As String
    Dim mapCount     As Long
    mapCount = 0

    ReDim mappingOld(uniquePaths.Count - 1)
    ReDim mappingNew(uniquePaths.Count - 1)

    Dim p As Variant
    For Each p In uniquePaths
        oldPath = CStr(p)
        If IsChildOfParent(oldPath, g_parentBase) Then
            alpha = GetAlphaSuffix(oldPath, g_parentBase)
            newPath = destFolder & newName & alpha
            mappingOld(mapCount) = oldPath
            mappingNew(mapCount) = newPath
            mapCount = mapCount + 1
        End If
    Next p

    ' ========== ファイルコピー（子コンポーネント）==============
    overwriteAll = False
    skipAll = False

    Dim successCount As Long
    Dim failCount    As Long
    successCount = 0
    failCount = 0

    For i = 0 To mapCount - 1
        Dim src As String
        Dim dst As String
        src = mappingOld(i)
        dst = mappingNew(i)

        If Not FileExists(src) Then
            failCount = failCount + 1
            Debug.Print "スキップ（ファイルなし）: " & src
        ElseIf FileExists(dst) Then
            If skipAll Then
                ' 既にスキップが選択済み → 何もしない
            ElseIf overwriteAll Then
                FileCopy src, dst
                successCount = successCount + 1
            Else
                Dim ans As Integer
                ans = MsgBox("既に同名ファイルが存在します：" & vbCrLf & dst & vbCrLf & vbCrLf & _
                             "[はい] 上書き　[いいえ] スキップ　[キャンセル] 全てスキップ", _
                             vbQuestion + vbYesNoCancel, "上書き確認")
                If ans = vbYes Then
                    FileCopy src, dst
                    successCount = successCount + 1
                ElseIf ans = vbCancel Then
                    skipAll = True
                End If
            End If
        Else
            FileCopy src, dst
            successCount = successCount + 1
            Debug.Print "コピー完了: " & src & " → " & dst
        End If
    Next i

    ' ========== 親アセンブリをコピー ==========================
    newAssyPath = destFolder & newName & ".sldasm"

    If FileExists(newAssyPath) Then
        Dim assyAns As Integer
        assyAns = MsgBox("親アセンブリの保存先に既に同名ファイルが存在します：" & vbCrLf & newAssyPath & vbCrLf & _
                         "上書きしますか？", vbQuestion + vbYesNo, "上書き確認")
        If assyAns = vbYes Then
            FileCopy g_parentPath, newAssyPath
        Else
            MsgBox "処理をキャンセルしました。", vbInformation, "キャンセル"
            Exit Sub
        End If
    Else
        FileCopy g_parentPath, newAssyPath
    End If

    ' ========== コピーした親アセンブリを開いて参照を更新 =========
    Dim openErrors   As Long
    Dim openWarnings As Long
    Set newModel = g_swApp.OpenDoc6(newAssyPath, 2, 1, "", openErrors, openWarnings)

    If newModel Is Nothing Then
        MsgBox "コピーした親アセンブリを開けませんでした。" & vbCrLf & _
               "手動で参照を更新してください。", vbExclamation, "警告"
        Exit Sub
    End If

    ' 参照パスを更新
    For i = 0 To mapCount - 1
        If FileExists(mappingNew(i)) Then
            newModel.Extension.ReplaceReferencedDocument mappingOld(i), mappingNew(i)
            Debug.Print "参照更新: " & mappingOld(i) & " → " & mappingNew(i)
        End If
    Next i

    ' 保存
    Dim saveErrors   As Long
    Dim saveWarnings As Long
    newModel.Save3 1, saveErrors, saveWarnings

    ' 閉じる
    g_swApp.CloseDoc newAssyPath

    ' ========== 完了メッセージ ================================
    MsgBox "ファイル名を変更しコピーが完了しました。" & vbCrLf & vbCrLf & _
           "コピー先フォルダ：" & destFolder & vbCrLf & _
           "新しい親ファイル名：" & newName & ".sldasm" & vbCrLf & _
           "子ファイルコピー数：" & successCount & " 件", _
           vbInformation, "完了"

End Sub

' ============================================================
' 子コンポーネントのパスを重複なく収集する
' ============================================================
Sub CollectUniqueComponentPaths(swAssembly As Object, ByRef col As Collection)

    Dim comps As Variant
    Dim c     As Variant
    Dim path  As String
    Dim dup   As Boolean
    Dim item  As Variant

    comps = swAssembly.GetComponents(False) ' False = 全レベルを再帰的に取得

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
            If Not dup Then
                col.Add path
            End If
        End If
    Next c

End Sub

' ============================================================
' ファイルの先頭が親ファイルベース名と一致するか判定
' ============================================================
Function IsChildOfParent(filePath As String, parentBase As String) As Boolean

    Dim fileName As String
    fileName = GetBaseNameFromPath(filePath) ' 拡張子なしのファイル名

    ' 大文字小文字を区別しない比較
    If Len(fileName) >= Len(parentBase) Then
        If StrComp(Left(fileName, Len(parentBase)), parentBase, vbTextCompare) = 0 Then
            IsChildOfParent = True
            Exit Function
        End If
    End If

    IsChildOfParent = False

End Function

' ============================================================
' 子ファイルからα部分（新しい親名に付ける接尾辞）を取得
' 例: 親="WIDGET-A", 子="WIDGET-A-01.sldprt" → "-01.sldprt"
' ============================================================
Function GetAlphaSuffix(filePath As String, parentBase As String) As String

    Dim fileName As String
    fileName = GetFileNameFromPath(filePath) ' 拡張子ありのファイル名
    Dim baseName As String
    baseName = GetBaseNameFromPath(filePath) ' 拡張子なしのファイル名

    Dim alpha As String
    alpha = Mid(baseName, Len(parentBase) + 1) ' 親名の後の部分

    Dim ext As String
    ext = GetExtension(filePath)

    GetAlphaSuffix = alpha & ext

End Function

' ============================================================
' 対象子ファイルの数を返す（フォームのプレビュー用）
' ============================================================
Function CountTargetChildren() As Long

    Dim col   As New Collection
    Dim count As Long
    Dim p     As Variant

    count = 0
    CollectUniqueComponentPaths g_swModel, col

    For Each p In col
        If IsChildOfParent(CStr(p), g_parentBase) Then
            count = count + 1
        End If
    Next p

    CountTargetChildren = count

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
