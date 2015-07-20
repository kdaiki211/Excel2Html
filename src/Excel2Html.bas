Attribute VB_Name = "Excel2Html"
Option Explicit

' �����s�̎�� (���D���Ȃ��̂ɕύX��)
Const Br As String = vbNewLine

' ��<table>�^�O�̕t���L��
Const AddTableTag As Boolean = True

' �C���f���g�̎��
Dim Idt As String
Dim OfstIdt As String
Dim AddCenterTag As Boolean ' <center>�^�O�̕t���L��

Public CancelReq As Boolean

' �c�[��
Private Function padLeft(ByVal s As String, ByVal c As String, ByVal l As Integer) As String
    If l - Len(s) <= 0 Then
        padLeft = s
        Exit Function
    End If
    padLeft = String(l - Len(s), c) & s
End Function

Private Function bgr2Rgb(ByVal color As Variant) As Variant
    Const rMask As Variant = 255 ' 0xFF
    Const gMask As Variant = 65280 ' 0xFF00
    Const bMask As Variant = 16711680 '0xFF0000
    bgr2Rgb = ((color And rMask) * (2 ^ 16)) Or _
               (color And gMask) Or _
               ((color And bMask) / (2 ^ 16))
End Function

' �� CSS �o�͗p�֐�
Private Function cvtToCssColor(ByVal color As Variant) As String
    Dim cssColor As Variant
    cssColor = bgr2Rgb(color)
    cvtToCssColor = "#" & padLeft(Hex(cssColor), "0", 6)
End Function

Private Function cvtToCssThickness(ByVal thickness As XlBorderWeight) As String
    Dim cssThickness As String
    
    Select Case thickness
        Case xlHairline
            cssThickness = "1px"
        Case xlMedium
            cssThickness = "2px"
        Case xlThick
            cssThickness = "3px"
        Case xlThin
            cssThickness = "1px"
        Case Else
            cssThickness = "1px"
    End Select
    
    cvtToCssThickness = cssThickness
End Function

Private Function getLineCss(ByRef rng As Range) As String
    Dim cl As Variant, ct As Variant, cr As Variant, cb As Variant
    Dim wl As XlBorderWeight, wt As XlBorderWeight, wr As XlBorderWeight, wb As XlBorderWeight
    Dim resultCss As String
    Dim isSameColor As Boolean
    Dim isSameThickness As Boolean
    
    ' ���̑����擾
    With rng
        ' color
        cl = .Borders(xlEdgeLeft).color
        ct = .Borders(xlEdgeTop).color
        cr = .Borders(xlEdgeRight).color
        cb = .Borders(xlEdgeBottom).color
        isSameColor = (cl = ct And ct = cr And cr = cb)
        
        ' weight
        wl = .Borders(xlEdgeLeft).Weight
        wt = .Borders(xlEdgeTop).Weight
        wr = .Borders(xlEdgeRight).Weight
        wb = .Borders(xlEdgeBottom).Weight
        isSameThickness = (wl = wt And wt = wr And wr = wb)
    End With
    
    resultCss = "border:solid"
    
    ' �㉺���E�̐����S�ē����F�̏ꍇ
    If isSameColor Then
        resultCss = resultCss & " " & cvtToCssColor(cl)
    End If
    
    ' �㉺���E�̐����S�ē��������̏ꍇ
    If isSameThickness Then
        resultCss = resultCss & " " & cvtToCssThickness(wl)
    End If
    
    resultCss = resultCss & ";"
    
    
    ' �㉺���E�̐����قȂ�F/�����̏ꍇ�A�ł��邾���Z�� CSS �R�[�h���o�͂ł���悤�ɐS������
    If (Not isSameColor) And (Not isSameThickness) Then
        resultCss = resultCss & "border-left:solid " & cvtToCssThickness(wl) & IIf(cl <> &H0, " " & cvtToCssColor(cl), "") & ";"
        resultCss = resultCss & "border-top:solid " & cvtToCssThickness(wt) & IIf(ct <> &H0, " " & cvtToCssColor(ct), "") & ";"
        resultCss = resultCss & "border-right:solid " & cvtToCssThickness(wr) & IIf(cr <> &H0, " " & cvtToCssColor(cr), "") & ";"
        resultCss = resultCss & "border-bottom:solid " & cvtToCssThickness(wb) & IIf(cb <> &H0, " " & cvtToCssColor(cb), "") & ";"
    ElseIf Not isSameColor Then
        resultCss = resultCss & IIf(cl <> &H0, "border-left:solid " & cvtToCssColor(cl), "") & ";"
        resultCss = resultCss & IIf(ct <> &H0, "border-top:solid " & cvtToCssColor(ct), "") & ";"
        resultCss = resultCss & IIf(cr <> &H0, "border-right:solid " & cvtToCssColor(cr), "") & ";"
        resultCss = resultCss & IIf(cb <> &H0, "border-bottom:solid " & cvtToCssColor(cb), "") & ";"
    ElseIf Not isSameThickness Then
        resultCss = resultCss & "border-left:solid " & cvtToCssThickness(wl) & IIf(cl <> &H0, " " & cvtToCssColor(cl), "") & ";"
        resultCss = resultCss & "border-top:solid " & cvtToCssThickness(wt) & IIf(ct <> &H0, " " & cvtToCssColor(ct), "") & ";"
        resultCss = resultCss & "border-right:solid " & cvtToCssThickness(wr) & IIf(cr <> &H0, " " & cvtToCssColor(cr), "") & ";"
        resultCss = resultCss & "border-bottom:solid " & cvtToCssThickness(wb) & IIf(cb <> &H0, " " & cvtToCssColor(cb), "") & ";"
    End If
    
last:
    getLineCss = resultCss
End Function

' �� HTML �^�O�o�̓��\�b�h�Q
Private Sub htmlPostProcess(ByRef s As String)
    If AddTableTag Then
        s = OfstIdt & IIf(AddCenterTag, Idt, "") & "<table style=""border-collapse:collapse"">" & Br & _
            s & _
            OfstIdt & IIf(AddCenterTag, Idt, "") & "</table>" & Br
    End If
    If AddCenterTag Then
        s = OfstIdt & "<center>" & Br & _
            s & _
            OfstIdt & "</center>" & Br
    End If
End Sub

Private Sub htmlStartNewRow(ByRef s As String)
    s = s & OfstIdt & IIf(AddCenterTag, Idt, "") & IIf(AddTableTag, Idt, "") & "<tr>" & Br
End Sub

Private Sub htmlFinishCurRow(ByRef s As String)
    s = s & OfstIdt & IIf(AddCenterTag, Idt, "") & IIf(AddTableTag, Idt, "") & "</tr>" & Br
End Sub

Private Sub htmlAddNewCell(ByRef s As String, _
                           ByVal newCellArea As Range)
    Dim cellValue As String
    Dim colspan As Integer, rowspan As Integer
    Dim bgColor As Long
    Dim color As Long
    Dim textAlign As Variant
    Dim verticalAlign As Variant
    Dim isBold As Boolean
    
    ' �����Z���̃v���p�e�B���擾
    cellValue = newCellArea.Cells(1, 1).Text ' �Z���̕�����͕K�� Range �̍���Z�����g�p
    colspan = newCellArea.Columns.Count
    rowspan = newCellArea.Rows.Count
    bgColor = newCellArea.Interior.color
    color = newCellArea.Font.color
    textAlign = newCellArea.Cells(1, 1).HorizontalAlignment
    verticalAlign = newCellArea.Cells(1, 1).VerticalAlignment
    isBold = newCellArea.Cells(1, 1).Font.Bold
    
    s = s & OfstIdt & IIf(AddCenterTag, Idt, "") & IIf(AddTableTag, Idt, "") & Idt & "<td"
    
    ' �s�����̘A��
    If colspan > 1 Then
        s = s & " colspan=" & CStr(colspan)
    End If
    
    ' ������̘A��
    If rowspan > 1 Then
        s = s & " rowspan=" & CStr(rowspan)
    End If
    
    ' CSS
    s = s & " style="""
    
    ' CSS: �w�i�F
    If bgColor <> &HFFFFFF Then ' ���ȊO
        s = s & "background:" & cvtToCssColor(bgColor) & ";"
    End If
    
    ' CSS: �����F
    If color <> &H0 Then ' ���ȊO
        s = s & "color:" & cvtToCssColor(color) & ";"
    End If
    
    ' CSS: �e�L�X�g�̐��������A���C��
    If textAlign = xlCenter Then
        s = s & "text-align:center;"
    ElseIf textAlign = xlRight Then
        s = s & "text-align:right;"
    End If
    
    ' CSS: �e�L�X�g�̐��������A���C��
    If verticalAlign = xlVAlignCenter Then
        s = s & "vertical-align:middle;"
    ElseIf verticalAlign = xlVAlignBottom Then
        s = s & "vertical-align:bottom;"
    Else
        s = s & "vertical-align:top;"
    End If
    
    ' CSS: ����
    If isBold Then
        s = s & "font-weight:bold;"
    End If
    
    ' CSS: border
    s = s & getLineCss(newCellArea) ' ���̐F�A���̑�����\�� CSS ��������擾

    s = s & """>"
    
    ' �Z���̕�����
    s = s & cellValue
    s = s & "</td>" & Br
End Sub

' �� �ݒ�t�@�C���ǂݍ���
Private Sub loadConfig()
    Dim indentType As Integer
    Dim indentOffset As Integer
    
    indentType = GetConfValue("IndentType", 0)
    indentOffset = GetConfValue("IndentOffset", 0)
    AddCenterTag = IIf(GetConfValue("AddCenterTag", 1) = 1, True, False)
    
    ' �C���f���g�Ɏg���������擾
    Select Case indentType
        Case 0
            Idt = ""
        Case 1
            Idt = vbTab
        Case 2
            Idt = " "
        Case 3
            Idt = "  "
        Case 4
            Idt = "    "
        Case Else
            ' �ُ�l�̎��̓f�t�H���g�l�ɖ߂�
            Idt = ""
            SetConfValue "IndentType", 0
    End Select
    
    ' �I�t�Z�b�g�̃C���f���g�𐶐�
    If indentOffset >= 0 And indentOffset <= 4 Then
        Dim i As Integer
        OfstIdt = ""
        For i = 1 To indentOffset
            OfstIdt = OfstIdt & Idt
        Next i
    Else
        ' �ُ�l�̎��̓f�t�H���g�l�ɖ߂�
        OfstIdt = ""
        SetConfValue "IndentOffset", 0
    End If
End Sub

' �� �i���ʒm
Public Sub CancelConverting()
    CancelReq = True
End Sub

Private Sub updateProgressBar(ByRef numOfProcessedCells As Long, ByRef numOfEntireCells)
    Dim barWidth As Integer
    barWidth = CDbl(numOfProcessedCells) / CDbl(numOfEntireCells) * CDbl(UI_Excel2Html.lbl_progress_bg.Width)
    UI_Excel2Html.lbl_progress_fg.Width = barWidth
    DoEvents
End Sub

' �� ���C���֐�
Public Function ConvertSelectedRangeToHtml() As String
    ' �O���[�o���ϐ�������
    CancelReq = False
    
    ' �ݒ�l���[�h
    loadConfig
    
    ' �v���O���X�o�[�\��
    UI_Excel2Html.lbl_progress_bg.Visible = True
    UI_Excel2Html.lbl_progress_fg.Visible = True
    UI_Excel2Html.btn_cancel.Visible = True
    
    ' �I��͈͂� 1 �Z��������
    With Selection
        Dim r As Long, c As Long
        Dim outHtml As String ' �o�� HTML ������ (VBA �̎d�l���A�����l�� "")
        Dim progressUpdateInterval As Long
        Dim checkPointR As Long
        Dim checkPointC As Long
        Dim numOfEntireCells As Double
        
        numOfEntireCells = CDbl(.Rows.Count) * CDbl(.Columns.Count)
        progressUpdateInterval = 100
        checkPointR = 0
        checkPointC = 0
        
        ' �I��͈͓��̃Z���� 1 ������ (���C������)
        For r = 0 To .Rows.Count - 1
        
            htmlStartNewRow outHtml
            
            For c = 0 To .Columns.Count - 1
                Dim curCell As Range, curArea As Range, curAreaTopLeft As Range
                
                Set curCell = .Cells(1 + r, 1 + c) ' ���݌��Ă���Z�� (1 �Z��)
                Set curArea = curCell.MergeArea ' ���݌��Ă���Z���������錋���Z���̑S��
                Set curAreaTopLeft = curArea.Cells(1, 1) ' ���݌��Ă���Z���������錋���Z���̍���Z�� (1 �Z��)
                
                ' r �s c �񂪌����Z���̍���Z���̂Ƃ��̂� HTML �o�͂���
                If curCell = curAreaTopLeft Then
                    htmlAddNewCell outHtml, curArea
                End If
                
                
                ' �i���\��
                If c >= checkPointC Then
                    updateProgressBar r * .Columns.Count + c, numOfEntireCells
                    checkPointC = checkPointC + progressUpdateInterval
                    If CancelReq = True Then
                        Exit For
                    End If
                End If
            Next c
            
            htmlFinishCurRow outHtml
            
            If r >= checkPointR Then
                updateProgressBar r * .Columns.Count + c, numOfEntireCells
                checkPointR = checkPointR + progressUpdateInterval
                If CancelReq = True Then
                    Exit For
                End If
            End If
        Next r
    End With
    
    If CancelReq = True Then
        MsgBox "Canceled. No HTML will be output.", vbExclamation
    End If
    
    htmlPostProcess outHtml
    ConvertSelectedRangeToHtml = outHtml ' �߂�l��Ԃ�
    
    ' �v���O���X�o�[��\��
    UI_Excel2Html.lbl_progress_bg.Visible = False
    UI_Excel2Html.lbl_progress_fg.Visible = False
    UI_Excel2Html.btn_cancel.Visible = False
End Function


' �� ���[�U�[�t�H�[���\�����\�b�h
Public Sub Excel2Html()
    UI_Excel2Html.Show
End Sub


