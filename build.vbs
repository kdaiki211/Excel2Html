Option Explicit

' オプション
Const dbg = False
Const ProjectName = "Excel2Html"
Const Extension = "xlam"
Const SrcDir = "src"
Const OutDir = "build"

' 定数
Const xlOpenXMLAddIn = 55

' 変数
Dim sh   ' シェル (カレントディレクトリ取得用)
Dim fso  ' FileSystemObject
Dim fo   ' folder
Dim xl   ' Excel アプリ
Dim wb   ' 本スクリプトが新規作成するワークブック
Dim vb   ' VBComponents
Dim pwd  ' カレントディレクトリ
Dim f
Dim logStr

' メイン処理 ------------------------------------------------

Sub PostProcess(ByRef wb, ByRef xl)
    wb.Close
    xl.Quit
End Sub

' シェルを準備
Set sh = CreateObject("WScript.Shell")
pwd = sh.CurrentDirectory

' FSO 準備
Set fso = CreateObject("Scripting.FileSystemObject")

' Execl を開く
Set xl = CreateObject("Excel.Application")
xl.Visible = dbg ' Excel のウィンドウを表示するかどうか

' アドインを生成
Set wb = xl.Workbooks.Add

On Error Resume Next
Set vb = wb.VBProject.VBComponents

If Err.Number = 1004 Then ' セキュリティ設定がデフォルトの場合
    MsgBox "Failed to get VBProject.VBComponents." & vbNewLine & _
           "To solve this problem, go to Excel's setting and " & _
           "select Security pane, then check ""Trust Access to Visual Basic Project.""." & vbNewLine & _
           "Retry executing the build script after the setting.", _
           vbExclamation, _
           "Error"
    PostProcess wb, xl
    WScript.Quit Err.Number
End If
On Error GoTo 0

' 各種ファイルインポート
Set fo = fso.GetFolder(pwd & "\" & SrcDir)
For Each f in fo.Files
    If Right(f.Name, 4) <> ".frx" Then ' frx file will be imported automatically when loading frm file
        logStr = logStr & vbNewLine & f.Name
        vb.Import pwd & "\" & SrcDir & "\" & f.Name
    End If
Next

' 確認メッセージ抑止
xl.DisplayAlerts = False

' 出力フォルダ作成 (存在しない場合のみ)
If Not fso.FolderExists(pwd & "\" & OutDir) Then
    fso.CreateFolder(pwd & "\" & OutDir)
End If

' アドイン保存
On Error Resume Next
wb.SaveAs pwd & "\" & OutDir & "\" & ProjectName & "." & Extension, xlOpenXMLAddIn

If Err.Number <> 0 Then
    MsgBox "Failed to output the addin file.", vbExclamation, "Error"
    PostProcess wb, xl
    WScript.Quit Err.Number
End If
On Error GoTo 0

PostProcess wb, xl

' 出力結果表示
MsgBox "Success building." & vbNewLine & _
       "Following components are included in the output file." & vbNewLine & _
       logStr, _
       vbInformation, _
       "Success Building"
