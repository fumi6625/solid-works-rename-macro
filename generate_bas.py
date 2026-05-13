"""
generate_bas.py
RenameAndCopy.bas を Shift-JIS (CP932) エンコーディングで生成する。
"""

VBA_CODE = r"""Attribute VB_Name = "RenameAndCopy"
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
' ファイルコピー＋バイナリ参照書き換え
' ============================================================
Sub ExecuteCopyAndRename(swApp As Object, swModel As Object, parentBase As String, _
                          newName As String, destFolder As String)

    Dim parentPath As String
    parentPath = swModel.GetPathName()

    ' ---- 1. コピー前にコンポーネントパスを全収集 ----
    Dim col As New Collection
    CollectUniqueComponentPaths swModel, col

    ' 置換テーブルを構築
    Dim oldPaths() As String
    Dim newPaths() As String
    Dim pCount     As Long
    ReDim oldPaths(col.Count)
    ReDim newPaths(col.Count)
    pCount = 0

    Dim p As Variant
    For Each p In col
        If IsChildOfParent(CStr(p), parentBase) Then
            oldPaths(pCount) = CStr(p)
            newPaths(pCount) = destFolder & newName & GetAlphaSuffix(CStr(p), parentBase)
            pCount = pCount + 1
        End If
    Next p

    If pCount = 0 Then
        MsgBox "コピー対象の子ファイルが見つかりませんでした。", vbExclamation, "対象なし"
        Exit Sub
    End If

    ' ---- 2. 子ファイルをコピー ----
    Dim copiedAsms As New Collection
    Dim i As Long
    For i = 0 To pCount - 1
        If Not SafeFileCopy(oldPaths(i), newPaths(i)) Then
            MsgBox "コピー失敗：" & vbCrLf & oldPaths(i), vbCritical, "エラー"
            Exit Sub
        End If
        If LCase(GetExtension(oldPaths(i))) = ".sldasm" Then
            copiedAsms.Add newPaths(i)
        End If
    Next i

    ' ---- 3. 親アセンブリをコピー ----
    Dim newParentPath As String
    newParentPath = destFolder & newName & GetExtension(parentPath)
    If Not SafeFileCopy(parentPath, newParentPath) Then
        MsgBox "親アセンブリのコピーに失敗しました。", vbCritical, "エラー"
        Exit Sub
    End If

    ' ---- 4. コピーした .sldasm のバイナリを書き換えて参照を更新 ----
    ' 親アセンブリを更新
    Dim ok As Boolean
    ok = BinaryReplaceUnicode(newParentPath, oldPaths, newPaths, pCount)

    ' コピーされたサブアセンブリも更新
    Dim asm As Variant
    For Each asm In copiedAsms
        BinaryReplaceUnicode CStr(asm), oldPaths, newPaths, pCount
    Next asm

    ' ---- 完了 ----
    If ok Then
        MsgBox "ファイル名を変更しコピーが完了しました。" & vbCrLf & vbCrLf & _
               "保存先　　　　：" & destFolder & vbCrLf & _
               "新しい親ファイル：" & newName & ".sldasm" & vbCrLf & _
               "コピーファイル数：" & pCount & " 件", vbInformation, "完了"
    Else
        MsgBox "ファイルのコピーは完了しました。" & vbCrLf & _
               "ただし、参照パスの書き換えに失敗しました。" & vbCrLf & vbCrLf & _
               "保存先：" & destFolder & vbCrLf & vbCrLf & _
               "新しい親アセンブリを開き、" & vbCrLf & _
               "「ファイル → 参照の置換」で手動設定してください。", _
               vbExclamation, "参照更新失敗"
    End If

End Sub

' ============================================================
' .sldasm バイナリ内の Unicode パス文字列を一括書き換え
'
' SolidWORKS はコンポーネントパスを UTF-16LE (Unicode) で
' バイナリに格納する。該当バイト列を検索して新パスで上書きする。
' ============================================================
Function BinaryReplaceUnicode(filePath As String, _
                               ByRef oldPaths() As String, _
                               ByRef newPaths() As String, _
                               ByVal pCount As Long) As Boolean
    BinaryReplaceUnicode = False

    ' ---- ファイル全体を読み込む ----
    Dim fNo   As Integer
    Dim fLen  As Long
    Dim data() As Byte

    On Error GoTo ReadErr
    fNo = FreeFile
    Open filePath For Binary Access Read As #fNo
    fLen = LOF(fNo)
    If fLen = 0 Then Close #fNo: Exit Function
    ReDim data(fLen - 1)
    Get #fNo, 1, data
    Close #fNo
    On Error GoTo 0

    Dim totalReplaced As Long
    totalReplaced = 0

    ' ---- 各パスを検索して書き換え ----
    Dim pi     As Long
    Dim i      As Long
    Dim j      As Long
    Dim k      As Long
    Dim oB()   As Byte
    Dim nB()   As Byte
    Dim oLen   As Long
    Dim nLen   As Long
    Dim delta  As Long
    Dim isMatch As Boolean
    Dim hasSpace As Boolean

    For pi = 0 To pCount - 1
        If oldPaths(pi) = newPaths(pi) Then GoTo NextPair

        ' VBA の String は内部的に UTF-16LE → そのままバイト配列へ
        oB = oldPaths(pi)
        nB = newPaths(pi)
        oLen = Len(oldPaths(pi)) * 2   ' バイト数
        nLen = Len(newPaths(pi)) * 2   ' バイト数
        delta = nLen - oLen

        i = 0
        Do While i <= fLen - oLen - 1

            ' 先頭 2 バイトで絞り込み（高速化）
            If data(i) = oB(0) And data(i + 1) = oB(1) Then

                ' 全バイト比較
                isMatch = True
                For j = 2 To oLen - 1
                    If data(i + j) <> oB(j) Then
                        isMatch = False
                        Exit For
                    End If
                Next j

                If isMatch Then

                    If delta <= 0 Then
                        ' 新パスが短い or 同じ → 上書き＋末尾をゼロ埋め
                        For j = 0 To nLen - 1
                            data(i + j) = nB(j)
                        Next j
                        For j = nLen To oLen - 1
                            data(i + j) = 0
                        Next j
                        Call TryUpdateLenPrefix(data, i, oLen, nLen)
                        totalReplaced = totalReplaced + 1
                        i = i + oLen
                        GoTo ContinueOuter

                    Else
                        ' 新パスが長い → 後続がゼロ（パディング）か確認
                        hasSpace = True
                        For k = 0 To delta - 1
                            If i + oLen + k >= fLen Then
                                hasSpace = False: Exit For
                            End If
                            If data(i + oLen + k) <> 0 Then
                                hasSpace = False: Exit For
                            End If
                        Next k

                        If hasSpace Then
                            For j = 0 To nLen - 1
                                data(i + j) = nB(j)
                            Next j
                            Call TryUpdateLenPrefix(data, i, oLen, nLen)
                            totalReplaced = totalReplaced + 1
                            i = i + nLen
                            GoTo ContinueOuter
                        End If
                    End If

                End If
            End If

            i = i + 1
ContinueOuter:
        Loop

NextPair:
    Next pi

    If totalReplaced = 0 Then Exit Function

    ' ---- 書き換えたデータをファイルに書き戻す ----
    On Error GoTo WriteErr
    fNo = FreeFile
    Open filePath For Binary Access Write As #fNo
    Put #fNo, 1, data
    Close #fNo
    On Error GoTo 0

    BinaryReplaceUnicode = True
    Exit Function

ReadErr:
    On Error Resume Next: Close #fNo: On Error GoTo 0
    Exit Function
WriteErr:
    On Error Resume Next: Close #fNo: On Error GoTo 0
End Function

' ============================================================
' 文字列の直前 4 バイトが長さフィールドであれば更新する
' （BSTR 形式: 4バイト長さプレフィックス + Unicode 文字列）
' ============================================================
Sub TryUpdateLenPrefix(ByRef data() As Byte, ByVal pos As Long, _
                        ByVal oldBLen As Long, ByVal newBLen As Long)
    If pos < 4 Then Exit Sub

    Dim prefix As Long
    prefix = CLng(data(pos - 4)) + _
             CLng(data(pos - 3)) * 256& + _
             CLng(data(pos - 2)) * 65536& + _
             CLng(data(pos - 1)) * 16777216&

    ' バイト数 or 文字数のどちらかが一致すれば更新
    Dim newPrefix As Long
    If prefix = oldBLen Then
        newPrefix = newBLen
    ElseIf prefix = oldBLen \ 2 Then
        newPrefix = newBLen \ 2
    Else
        Exit Sub
    End If

    data(pos - 4) = CByte(newPrefix And 255&)
    data(pos - 3) = CByte((newPrefix \ 256&) And 255&)
    data(pos - 2) = CByte((newPrefix \ 65536&) And 255&)
    data(pos - 1) = CByte((newPrefix \ 16777216&) And 255&)
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
"""


def main():
    output_path = "RenameAndCopy.bas"
    with open(output_path, "w", encoding="cp932") as f:
        f.write(VBA_CODE)
    print(f"生成完了（CP932エンコーディング）：{output_path}")


if __name__ == "__main__":
    main()
