VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_Excel2Html 
   ClientHeight    =   4380
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6480
   OleObjectBlob   =   "UI_Excel2Html.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UI_Excel2Html"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btn_cancel_Click()
    Excel2Html.CancelConverting
End Sub

Private Sub tboxSelectAll()
    With tbox_output
        ' テキストボックスを全て選択状態にしてコピペしやすくする
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub convertToHtml()
    Dim res As String
    Dim c As Control
    
    Me.Caption = ProductName & " - Processing..."
    
    ' 全てのコントロールを無効にする
    For Each c In Me.Controls
        c.Enabled = False
    Next c
    btn_cancel.Enabled = True
    
    ' セルが選択されているかを確認する
    If TypeName(Selection) <> "Range" Then
        MsgBox "セルが選択されていません." & vbNewLine & "HTML に変換したい範囲を選択してから再試行してください.", vbCritical, "セル未選択エラー"
        End
    End If
    
    ' 変換処理
    res = ConvertSelectedRangeToHtml

last:
    ' 全てのコントロールを有効にする
    For Each c In Me.Controls
        c.Enabled = True
    Next c
    
    ' 出力 HTML を表示
    If Excel2Html.CancelReq = False Then
        tbox_output.Text = res
        tboxSelectAll
    End If
    
    Me.Caption = ProductName & " " & ProductVersion & " - Convert Result"
End Sub

Private Sub btn_close_Click()
    Unload Me
End Sub

Private Sub btn_config_Click()
    UI_Config.Show
    convertToHtml
End Sub

Private Sub btn_preview_Click()
    UI_Preview.HtmlToPreview = tbox_output.Text
    UI_Preview.Show
End Sub

Private Sub tbox_output_mouseup(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tboxSelectAll
End Sub

Private Sub UserForm_Activate()
    Dim i As Integer
    
    Me.Caption = ProductName & " " & ProductVersion
    
    ' フォーム表示時に自動的に Excel → HTML 変換を行う
    convertToHtml
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 処理を中断
    Excel2Html.CancelConverting
End Sub
