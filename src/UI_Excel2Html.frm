VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_Excel2Html 
   Caption         =   "Excel2Html - Convert Result"
   ClientHeight    =   4692
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6480
   OleObjectBlob   =   "UI_Excel2Html.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_Excel2Html"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Modified As Boolean
Dim IsReady As Boolean

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
    
    cmb_indentType.Enabled = False
    cmb_indentOffset.Enabled = False
    
    res = ConvertSelectedRangeToHtml

    cmb_indentType.Enabled = True
    cmb_indentOffset.Enabled = True
    If Excel2Html.CancelReq = False Then
        ' 出力 HTML を表示
        tbox_output.Text = res
        tboxSelectAll
    End If
End Sub

Private Sub chk_center_Click()
    If IsReady Then
        Dim val As Integer
        val = IIf(chk_center.value, 1, 0)
        SetConfValue "AddCenterTag", val, False
        
        convertToHtml
        Modified = True
    End If
End Sub

Private Sub cmb_indentType_Change()
    If IsReady Then
        SetConfValue "IndentType", cmb_indentType.ListIndex, False
        convertToHtml
        Modified = True
    End If
End Sub

Private Sub cmb_indentOffset_Change()
    If IsReady Then
        SetConfValue "IndentOffset", cmb_indentOffset.ListIndex, False
        convertToHtml
        Modified = True
    End If
End Sub

Private Sub UserForm_Activate()
    Dim indentTypeIdx As Integer
    Dim indentOffsetIdx As Integer
    Dim addCenterTagVal As Integer
    
    IsReady = False
    Modified = False
    
    ' コンボボックス (IndentType) のアイテム追加
    cmb_indentType.AddItem "None"
    cmb_indentType.AddItem "Tab"
    cmb_indentType.AddItem "1 Space"
    cmb_indentType.AddItem "2 Spaces"
    cmb_indentType.AddItem "4 Spaces"
    
    ' コンボボックス (IndentOffset) のアイテム追加
    cmb_indentOffset.AddItem "None"
    cmb_indentOffset.AddItem "1 Indent"
    cmb_indentOffset.AddItem "2 Indents"
    cmb_indentOffset.AddItem "3 Indents"
    cmb_indentOffset.AddItem "4 Indents"
    
    ' 過去の選択値をロード
    indentTypeIdx = GetConfValue("IndentType", 0)
    indentOffsetIdx = GetConfValue("IndentOffset", 0)
    addCenterTagVal = GetConfValue("AddCenterTag", 1)
    
    ' コンボボックスの選択値を設定
    cmb_indentType.ListIndex = indentTypeIdx
    cmb_indentOffset.ListIndex = indentOffsetIdx
    chk_center.value = IIf(addCenterTagVal = 1, True, False)

    ' フォーム表示時に自動的に Excel → HTML 変換を行う
    convertToHtml
    
    IsReady = True
End Sub


Private Sub closeUserFormIfEscapeKeyPressed(ByVal KeyCode As MSForms.ReturnInteger)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    closeUserFormIfEscapeKeyPressed KeyCode
End Sub

Private Sub tbox_output_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    closeUserFormIfEscapeKeyPressed KeyCode
End Sub

Private Sub cmb_indentOffset_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    closeUserFormIfEscapeKeyPressed KeyCode
End Sub

Private Sub cmb_indentType_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    closeUserFormIfEscapeKeyPressed KeyCode
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 処理を中断
    Excel2Html.CancelConverting
    
    ' 設定値保存
    If Modified Then
        CommitAllConf
    End If
End Sub
