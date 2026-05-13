VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRenameDialog
   Caption         =   "アセンブリ コピー＆リネーム"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "frmRenameDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRenameDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' frmRenameDialog - ユーザー入力フォーム
' ============================================================
' このファイルはVBAエディタへのインポート用ソースコードです。
' SolidWORKSのVBAエディタでUserFormを作成し、以下のコントロールを
' 手動で配置してください。
'
' コントロール一覧:
'   lblTitle        Label       "アセンブリ コピー＆リネーム"
'   lblCurrentLabel Label       "現在のアセンブリ名："
'   lblCurrentName  Label       （現在のアセンブリ名を表示）
'   lblNewName      Label       "新しいアセンブリ名："
'   lblNewNameNote  Label       "（拡張子 .sldasm は不要です）"
'   txtNewName      TextBox     新しい名前を入力
'   lblFolder       Label       "保存先フォルダ："
'   txtFolder       TextBox     保存先フォルダパスを表示・入力
'   btnBrowse       CommandButton "参照..."
'   lblPreview      Label       対象子ファイル数プレビュー
'   btnExecute      CommandButton "実行"
'   btnCancel       CommandButton "キャンセル"
' ============================================================

Option Explicit

' ============================================================
' フォーム初期化
' ============================================================
Private Sub UserForm_Initialize()
    ' 現在のアセンブリ名をラベルに表示
    Me.lblCurrentName.Caption = g_parentBase & ".sldasm"

    ' 保存先フォルダのデフォルトは親ファイルと同じフォルダ
    Me.txtFolder.Text = g_parentDir

    ' 子ファイル数をプレビュー表示
    Dim childCount As Long
    childCount = CountTargetChildren()
    Me.lblPreview.Caption = "対象となる子ファイル数：" & childCount & " 件"

    ' 新しい名前欄は空にしてフォーカス
    Me.txtNewName.Text = ""
    Me.txtNewName.SetFocus
End Sub

' ============================================================
' 「参照...」ボタン：フォルダ選択ダイアログ
' ============================================================
Private Sub btnBrowse_Click()
    Dim shell  As Object
    Dim folder As Object

    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, "保存先フォルダを選択してください", 0, g_parentDir)

    If Not folder Is Nothing Then
        Me.txtFolder.Text = folder.Self.Path
    End If
End Sub

' ============================================================
' 「実行」ボタン
' ============================================================
Private Sub btnExecute_Click()
    Dim newName   As String
    Dim destFolder As String

    newName    = Trim(Me.txtNewName.Text)
    destFolder = Trim(Me.txtFolder.Text)

    ' ===== バリデーション =====
    If newName = "" Then
        MsgBox "新しいアセンブリ名を入力してください。", vbExclamation, "入力エラー"
        Me.txtNewName.SetFocus
        Exit Sub
    End If

    ' ファイル名として使用できない文字チェック
    If ContainsInvalidChars(newName) Then
        MsgBox "ファイル名に使用できない文字が含まれています。" & vbCrLf & _
               "次の文字は使用できません：\ / : * ? "" < > |", _
               vbExclamation, "入力エラー"
        Me.txtNewName.SetFocus
        Exit Sub
    End If

    If destFolder = "" Then
        MsgBox "保存先フォルダを指定してください。", vbExclamation, "入力エラー"
        Me.txtFolder.SetFocus
        Exit Sub
    End If

    ' フォルダ存在確認
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(destFolder) Then
        Dim createAns As Integer
        createAns = MsgBox("指定されたフォルダが存在しません：" & vbCrLf & destFolder & vbCrLf & vbCrLf & _
                           "フォルダを作成しますか？", vbQuestion + vbYesNo, "フォルダ確認")
        If createAns = vbYes Then
            On Error Resume Next
            fso.CreateFolder destFolder
            If Err.Number <> 0 Then
                MsgBox "フォルダを作成できませんでした：" & Err.Description, vbCritical, "エラー"
                Err.Clear
                Exit Sub
            End If
            On Error GoTo 0
        Else
            Me.txtFolder.SetFocus
            Exit Sub
        End If
    End If

    ' ===== 実行確認 =====
    Dim childCount As Long
    childCount = CountTargetChildren()

    Dim confirmMsg As String
    confirmMsg = "以下の内容でコピーを実行します。" & vbCrLf & vbCrLf & _
                 "現在の親ファイル名：" & g_parentBase & ".sldasm" & vbCrLf & _
                 "新しい親ファイル名：" & newName & ".sldasm" & vbCrLf & _
                 "保存先フォルダ　　：" & destFolder & vbCrLf & _
                 "対象子ファイル数　：" & childCount & " 件" & vbCrLf & vbCrLf & _
                 "実行しますか？" & vbCrLf & vbCrLf & _
                 "※ 実行前に元ファイルのバックアップを取得することを推奨します。"

    Dim ans As Integer
    ans = MsgBox(confirmMsg, vbQuestion + vbYesNo, "実行確認")

    If ans <> vbYes Then
        Exit Sub
    End If

    ' ===== フォームを隠して処理実行 =====
    Me.Hide

    On Error GoTo ExecError
    Call ExecuteRenameAndCopy(newName, destFolder)
    On Error GoTo 0

    ' 処理完了後にフォームを閉じる
    Unload Me
    Exit Sub

ExecError:
    MsgBox "処理中にエラーが発生しました：" & vbCrLf & Err.Description, vbCritical, "エラー"
    Me.Show
End Sub

' ============================================================
' 「キャンセル」ボタン
' ============================================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

' ============================================================
' フォームのXボタンで閉じた場合
' ============================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then ' ユーザーがXボタンで閉じた
        Cancel = 0
    End If
End Sub

' ============================================================
' ファイル名に使用できない文字チェック
' ============================================================
Private Function ContainsInvalidChars(name As String) As Boolean
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
