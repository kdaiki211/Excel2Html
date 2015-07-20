VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_Config 
   Caption         =   "Config"
   ClientHeight    =   2640
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   3180
   OleObjectBlob   =   "UI_Config.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UI_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsReady As Boolean ' Boolean �̏����l�� False
Dim Modified As Boolean

Private Sub btn_ok_Click()
    Unload Me
End Sub

Private Sub chk_center_Click()
    If IsReady Then
        Dim val As Integer
        val = IIf(chk_center.value, 1, 0)
        SetConfValue "AddCenterTag", val, False
        Modified = True
    End If
End Sub

Private Sub chk_nobr_Click()
    If IsReady Then
        Dim val As Integer
        val = IIf(chk_nobr.value, 1, 0)
        SetConfValue "Nobr", val, False
        Modified = True
    End If
End Sub

Private Sub cmb_indentType_Change()
    If IsReady Then
        SetConfValue "IndentType", cmb_indentType.ListIndex, False
        Modified = True
    End If
End Sub

Private Sub cmb_indentOffset_Change()
    If IsReady Then
        SetConfValue "IndentOffset", cmb_indentOffset.ListIndex, False
        Modified = True
    End If
End Sub

Private Sub txt_tblClass_Change()
    If IsReady Then
        SetConfValue "TableClass", txt_tblClass.Text, False
        Modified = True
    End If
End Sub

Private Sub txt_tblId_Change()
    If IsReady Then
        SetConfValue "TableId", txt_tblId.Text, False
        Modified = True
    End If
End Sub

Private Sub UserForm_Activate()
    Dim indentTypeIdx As Integer
    Dim indentOffsetIdx As Integer
    Dim addCenterTagVal As Integer
    Dim nobr As Integer
    Dim tblClass As String
    Dim tblId As String
    
    ' �R���{�{�b�N�X (IndentType) �̃A�C�e���ǉ�
    cmb_indentType.AddItem "None"
    cmb_indentType.AddItem "Tab"
    cmb_indentType.AddItem "1 Space"
    cmb_indentType.AddItem "2 Spaces"
    cmb_indentType.AddItem "4 Spaces"
    
    ' �R���{�{�b�N�X (IndentOffset) �̃A�C�e���ǉ�
    cmb_indentOffset.AddItem "None"
    cmb_indentOffset.AddItem "1 Indent"
    cmb_indentOffset.AddItem "2 Indents"
    cmb_indentOffset.AddItem "3 Indents"
    cmb_indentOffset.AddItem "4 Indents"
    
    ' �ߋ��̑I��l�����[�h
    indentTypeIdx = GetConfValue("IndentType", 0)
    indentOffsetIdx = GetConfValue("IndentOffset", 0)
    addCenterTagVal = GetConfValue("AddCenterTag", 1)
    nobr = GetConfValue("Nobr", 0)
    tblClass = GetConfValue("TableClass", "")
    tblId = GetConfValue("TableId", "")
    
    ' �I��l��ݒ�
    cmb_indentType.ListIndex = indentTypeIdx
    cmb_indentOffset.ListIndex = indentOffsetIdx
    chk_center.value = IIf(addCenterTagVal = 1, True, False)
    chk_nobr.value = IIf(nobr = 1, True, False)
    txt_tblClass.Text = tblClass
    txt_tblId.Text = tblId
    
    IsReady = True
    Modified = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' �ݒ�l�ۑ�
    If Modified Then
        CommitAllConf
    End If
End Sub
