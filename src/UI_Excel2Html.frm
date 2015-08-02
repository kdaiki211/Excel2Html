VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_Excel2Html 
   ClientHeight    =   4380
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6480
   OleObjectBlob   =   "UI_Excel2Html.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
        ' �e�L�X�g�{�b�N�X��S�đI����Ԃɂ��ăR�s�y���₷������
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub convertToHtml()
    Dim res As String
    Dim c As Control
    
    Me.Caption = ProductName & " - Processing..."
    
    ' �S�ẴR���g���[���𖳌��ɂ���
    For Each c In Me.Controls
        c.Enabled = False
    Next c
    btn_cancel.Enabled = True
    
    ' �Z�����I������Ă��邩���m�F����
    If TypeName(Selection) <> "Range" Then
        MsgBox "�Z�����I������Ă��܂���." & vbNewLine & "HTML �ɕϊ��������͈͂�I�����Ă���Ď��s���Ă�������.", vbCritical, "�Z�����I���G���["
        End
    End If
    
    ' �ϊ�����
    res = ConvertSelectedRangeToHtml

last:
    ' �S�ẴR���g���[����L���ɂ���
    For Each c In Me.Controls
        c.Enabled = True
    Next c
    
    ' �o�� HTML ��\��
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
    
    ' �t�H�[���\�����Ɏ����I�� Excel �� HTML �ϊ����s��
    convertToHtml
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' �����𒆒f
    Excel2Html.CancelConverting
End Sub
