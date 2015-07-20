Attribute VB_Name = "Excel2Html"
Option Explicit

' ���C���f���g�̎�� (���D���Ȃ��̂ɕύX��)
Const Idt As String = vbTab

' �����s�̎�� (���D���Ȃ��̂ɕύX��)
Const Br As String = vbNewLine

' ��<table>�^�O�̕t���L��
Const AddTableTag As Boolean = True


' �� HTML �^�O�o�̓��\�b�h�Q
Private Sub htmlPostProcess(ByRef s As String)
    If AddTableTag Then
        s = "<table>" & Br & s & "</table>"
    End If
End Sub

Private Sub htmlStartNewRow(ByRef s As String)
    s = s & Idt & "<tr>" & Br
End Sub

Private Sub htmlFinishCurRow(ByRef s As String)
    s = s & Idt & "</tr>" & Br
End Sub

Private Function padLeft(ByVal s As String, ByVal c As String, ByVal l As Integer) As String
    If l - Len(s) <= 0 Then
        padLeft = s
        Exit Function
    End If
    padLeft = String(l - Len(s), c) & s
End Function

Private Function bgr2Rgb(ByVal color As Variant) As Variant
    bgr2Rgb = ((color And &HFF) * (2 ^ 16)) Or _
               (color And &HFF00) Or _
               ((color And &HFF0000) / (2 ^ 16))
End Function

Private Sub htmlAddNewCell(ByRef s As String, _
                           ByVal newCellArea As Range)
    Dim cellValue As String
    Dim colspan As Integer, rowspan As Integer
    Dim bgColor As Variant
    Dim color As Variant
    Dim textAlign As Variant
    Dim isBold As Boolean
    
    ' �����Z���̃v���p�e�B���擾
    cellValue = newCellArea.Cells(1, 1).Text ' �Z���̕�����͕K�� Range �̍���Z�����g�p
    colspan = newCellArea.Columns.Count
    rowspan = newCellArea.Rows.Count
    bgColor = bgr2Rgb(newCellArea.Interior.color)
    color = bgr2Rgb(newCellArea.Font.color)
    textAlign = newCellArea.Cells(1, 1).HorizontalAlignment
    isBold = newCellArea.Cells(1, 1).Font.Bold
    
    s = s & Idt & Idt & "<td"
    
    ' �w�i�F
    If bgColor <> &HFFFFFF Then
        s = s & " bgColor=""#" & padLeft(Hex(bgColor), "0", 6) & """"
    End If
    
    ' �s�����̘A��
    If colspan > 1 Then
        s = s & " colspan=" & CStr(colspan)
    End If
    
    ' ������̘A��
    If rowspan > 1 Then
        s = s & " rowspan=" & CStr(rowspan)
    End If
    
    ' �e�L�X�g�̐��������̃A���C��
    If textAlign = xlCenter Then
        s = s & " align=""center"""
    ElseIf textAlign = xlRight Then
        s = s & " align=""right"""
    End If

    s = s & ">"
    
    ' font �^�O�J�n
    If color <> &H0 Then
        s = s & "<font color=""#" & padLeft(Hex(color), "0", 6) & """>"
    End If
    
    ' b �^�O�J�n
    If isBold Then
        s = s & "<b>"
    End If
    
    ' �Z���̕�����
    s = s & cellValue
    
    ' b �^�O�I��
    If isBold Then
        s = s & "</b>"
    End If
     
    ' font �^�O�I��
    If color <> &H0 Then
        s = s & "</font>"
    End If
    
    s = s & "</td>" & Br
End Sub


' �� ���C���֐�
Function ConvertSelectedRangeToHtml() As String
    ' �I��͈͂� 1 �Z��������
    With Selection
        Dim r As Integer, c As Integer
        Dim selTopLeft As Range ' �I��͈͂̍���Z��
        Dim outHtml As String ' �o�� HTML ������ (VBA �̎d�l���A�����l�� "")
        
        ' �I��͈͂̍���Z���擾
        Set selTopLeft = .Cells(1, 1)
        
        ' �I��͈͓��̃Z���� 1 ������ (���C������)
        For r = 0 To .Rows.Count - 1
            htmlStartNewRow outHtml
            
            For c = 0 To .Columns.Count - 1
                Dim curCell As Range, curArea As Range, curAreaTopLeft As Range
                
                Set curCell = selTopLeft.Offset(r, c) ' ���݌��Ă���Z�� (1 �Z��)
                Set curArea = curCell.MergeArea ' ���݌��Ă���Z���������錋���Z���̑S��
                Set curAreaTopLeft = curArea.Cells(1, 1) ' ���݌��Ă���Z���������錋���Z���̍���Z�� (1 �Z��)
                
                ' r �s c �񂪌����Z���̍���Z���̂Ƃ��̂� HTML �o�͂���
                If curCell = curAreaTopLeft Then
                    htmlAddNewCell outHtml, curArea
                End If
            Next c
            
            htmlFinishCurRow outHtml
        Next r
    End With
    
    htmlPostProcess outHtml
    ConvertSelectedRangeToHtml = outHtml ' �߂�l��Ԃ�
End Function


' �� ���[�U�[�t�H�[���\�����\�b�h
Public Sub Excel2Html()
    UI_Excel2Html.Show vbModeless
End Sub
