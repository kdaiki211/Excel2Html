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
    ' �e���̑���
    Dim ci(0 To 3) As Variant        ' ColorIndex
    Dim c(0 To 3) As Variant         ' Color
    Dim w(0 To 3) As XlBorderWeight  ' Weight
    Dim dr(0 To 3) As XlBordersIndex ' Direction
    Dim drnm(0 To 3) As String       ' Direction Name (CSS)
    Dim bs(0 To 3) As String         ' Border Style (CSS)
    
    Dim resultCss As String
    Dim isSameColorIndex As Boolean
    Dim isSameColor As Boolean
    Dim isSameThickness As Boolean
    Dim isSameBorderStyle As Boolean
    
    Dim i As Integer
    
    ' �萔���
    dr(0) = xlEdgeLeft
    dr(1) = xlEdgeTop
    dr(2) = xlEdgeRight
    dr(3) = xlEdgeBottom
    drnm(0) = "left"
    drnm(1) = "top"
    drnm(2) = "right"
    drnm(3) = "bottom"
    
    ' ���̑����擾
    For i = 0 To 3
        With rng.Borders(dr(i))
            ci(i) = .ColorIndex
            c(i) = .color
            w(i) = .Weight
            Select Case .LineStyle ' ����
                Case xlLineStyleNone, xlContinuous
                    bs(i) = "solid" ' �����̂ق��A���Ȃ��� solid �ň��� (������ 0 �Ƃ��ď���)
                Case xlDouble
                    bs(i) = "double" ' ��d��
                Case xlDot
                    bs(i) = "dotted" ' �_��
                Case xlDash, xlDashDot, xlDashDotDot, xlSlantDashDot
                    bs(i) = "dashed" ' �j��
                Case Else
                    bs(i) = "solid"
            End Select
        End With
    Next i
    
    ' ���̐F���S�ē������E���̑������S�ē��������擾
    isSameColorIndex = (ci(0) = ci(1) And ci(1) = ci(2) And ci(2) = ci(3))
    isSameColor = (c(0) = c(1) And c(1) = c(2) And c(2) = c(3))
    isSameBorderStyle = (bs(0) = bs(1) And bs(1) = bs(2) And bs(2) = bs(3))
    isSameThickness = True ' �ȉ��̏����Œl�����߂�
    For i = 1 To 3
        ' �����Ȑ��͔�r�ΏۊO
        If ci(i - 1) = xlColorIndexNone Or ci(i) = xlColorIndexNone Then
            GoTo continue
        End If
        ' �������������m�F���A������甲����
        If w(i - 1) <> w(i) Then
            isSameThickness = False
            Exit For
        End If
continue:
    Next i
    
    ' �l�ӂ������̏ꍇ
    If isSameColorIndex And ci(0) = xlColorIndexNone Then
        resultCss = ""
        GoTo last
    End If
    
    ' �㉺���E�̐����S�ē�������̏ꍇ
    If isSameBorderStyle Then
        resultCss = bs(0)
    End If
    ' �㉺���E�̐����S�ē����F�̏ꍇ (�����������̏ꍇ������)
    If isSameColorIndex And ci(0) <> xlColorIndexNone Then
        resultCss = resultCss & " " & cvtToCssColor(c(0))
    End If
    ' �㉺���E�̐����S�ē��������̏ꍇ
    If isSameThickness Then
        resultCss = resultCss & " " & cvtToCssThickness(w(0))
    End If
    resultCss = "border:" & Trim(resultCss) & ";"
    
    ' �㉺���E�̐����قȂ�F/�����̏ꍇ
    For i = 0 To 3
        If ci(i) = xlColorIndexNone Then
            ' �����Ȑ�
            resultCss = resultCss & "border-" & drnm(i) & ":0;"
        Else
            If Not isSameBorderStyle Then
                ' ���킪���ꂳ��Ă��Ȃ�
                resultCss = resultCss & "border-" & drnm(i) & "-style:" & bs(i) & ";"
            End If
            If Not isSameColorIndex Then
                ' ���F�����ꂳ��Ă��Ȃ�
                resultCss = resultCss & "border-" & drnm(i) & "-color:" & cvtToCssColor(c(i)) & ";"
            End If
            If Not isSameThickness Then
                ' ���������ꂳ��Ă��Ȃ�
                resultCss = resultCss & "border-" & drnm(i) & "-width:" & cvtToCssThickness(w(i)) & ";"
            End If
        End If
    Next i
    
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

' �w�肵���Z���̃v���p�e�B��\�� HTML ��Ԃ��܂��B
' �Ȃ��A�Z���Ɋ܂܂��S�Ă̕�����ɋ��ʂ���X�^�C�����܂߂ĕԂ��܂��B
Private Function getCellProperties(ByRef newCellArea As Range) As String
    Dim colspan As Integer, rowspan As Integer
    Dim bgColorIndex As Variant
    Dim bgColor As Long
    Dim textAlign As Variant
    Dim verticalAlign As Variant
    
    Dim cssTextDecoration As String
    
    Dim propColspan As String
    Dim propRowspan As String
    Dim propStyle As String
    Dim ret As String
    
    ' ��Ɏ擾�\�ȃv���p�e�B (Null ���Ԃ�Ȃ�) ���擾
    colspan = newCellArea.Columns.Count
    rowspan = newCellArea.Rows.Count
    textAlign = newCellArea.Cells(1, 1).HorizontalAlignment
    verticalAlign = newCellArea.Cells(1, 1).VerticalAlignment
    bgColorIndex = newCellArea.Interior.ColorIndex
    bgColor = newCellArea.Interior.color
    
    ' colspan / rowspan �v���p�e�B���̕�����쐬
    propColspan = IIf(colspan > 1, CStr(colspan), "")
    propRowspan = IIf(rowspan > 1, CStr(rowspan), "")
    
    ' style �v���p�e�B���̕�����쐬 (�Z�����̕����񂪕����I�ɈقȂ�X�^�C���̏ꍇ�A�e��v���p�e�B�� Null �ƂȂ�̂Œ���)
    With newCellArea.Cells(1, 1).Font ' �t�H���g�֘A���܂Ƃ߂�
        ' font-weight
        If Not IsNull(.Bold) Then
            propStyle = propStyle & IIf(.Bold = True, "font-weight:bold;", "")
        End If
        ' font-style
        If Not IsNull(.Italic) Then
            propStyle = propStyle & IIf(.Italic = True, "font-style:italic;", "")
        End If
        ' text-decoration
        If Not IsNull(.Underline) Then
            If .Underline = xlUnderlineStyleNone Then
                ' NOP
            Else
                ' ���炩�̉����������Ă���ꍇ�A�����I�Ɉ�d���̉��������� (CSS �œ�d�������̂͒��߂�)
                cssTextDecoration = cssTextDecoration & "underline "
            End If
        End If
        If Not IsNull(.Strikethrough) Then
            cssTextDecoration = cssTextDecoration & IIf(.Strikethrough = True, "line-through ", "")
        End If
        propStyle = propStyle & IIf(Len(cssTextDecoration) > 0, "text-decoration:" & Trim(cssTextDecoration) & ";", "")
        ' color
        If Not IsNull(.color) Then
            If .color <> &H0 Then ' �����F = �� �ȊO�̏ꍇ�����A�����F���w�肷�� CSS ��f��
                propStyle = propStyle & IIf(.color <> &H0, "color:" & cvtToCssColor(.color) & ";", "")
            End If
        End If
    End With
    
    ' background
    If bgColorIndex <> xlColorIndexNone Then ' �����ȊO
        propStyle = propStyle & "background:" & cvtToCssColor(bgColor) & ";"
    End If
    ' text-align
    If textAlign = xlCenter Then
        propStyle = propStyle & "text-align:center;"
    ElseIf textAlign = xlRight Then
        propStyle = propStyle & "text-align:right;"
    End If
    ' vertical-align
    If verticalAlign = xlVAlignTop Then
        propStyle = propStyle & "vertical-align:top;"
    ElseIf verticalAlign = xlVAlignBottom Then
        propStyle = propStyle & "vertical-align:bottom;"
    End If
    ' border
    propStyle = propStyle & getLineCss(newCellArea)
    
    ' �t�@�C�i���C�Y
    If propColspan <> "" Then
        ret = ret & "colspan=" & propColspan & " "
    End If
    If propRowspan <> "" Then
        ret = ret & "rowspan=" & propRowspan & " "
    End If
    If propStyle <> "" Then
        ret = ret & "style=""" & propStyle & """"
    End If
    ret = Trim(ret) ' �擪/�����̔��p�X�y�[�X������
    
    getCellProperties = ret
End Function

' �w�肵���Z���̕�������X�^�C���t���ŕԂ��܂��B
' �{�֐��́A�Z���̕�����̈ꕔ�����ɃX�^�C�����K�p����Ă���ꍇ�� <font> �^�O���g�p���ČʂɃX�^�C���K�p���� HTML ��Ԃ��܂��B
' �Z���S�̂ɓK�p���ꂽ�����֖{�֐��͊֗^���܂���B
Private Function getCellValueWithStyle(ByRef newCellArea As Range) As String
    Dim cellValue As String
    Dim ret As String
    cellValue = newCellArea.Cells(1, 1).Text ' �Z���̕�����͕K�� Range �̍���Z�����g�p
    
    ret = cellValue
    getCellValueWithStyle = ret
End Function

Private Sub htmlAddNewCell(ByRef s As String, _
                           ByRef newCellArea As Range)
    Dim prop As String
    Dim content As String
    
    ' <td> �̃v���p�e�B�擾
    prop = getCellProperties(newCellArea)
    ' <td> �` </td> ���� HTML �擾
    content = getCellValueWithStyle(newCellArea)
    
    ' <td> �^�O�ǋL
    s = s & OfstIdt & IIf(AddCenterTag, Idt, "") & IIf(AddTableTag, Idt, "") & Idt & _
        "<td" & IIf(Len(prop) > 0, " ", "") & prop & ">" & _
        content & "</td>" & Br
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


