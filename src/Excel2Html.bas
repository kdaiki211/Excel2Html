Attribute VB_Name = "Excel2Html"
Option Explicit


' ��<table>�^�O�̕t���L��
Const AddTableTag As Boolean = True

' ����Ƃ���t�H���g�T�C�Y (����ȊO�̃t�H���g�T�C�Y�̏ꍇ�̂� <td> �� style �� font-size ���w�肷��
Const DefFontSize As Integer = 11

' �C���f���g�̎��
Dim Idt As String
Dim OfstIdt As String
Dim OfstIdtBackup As String ' ���s�Ȃ��ŏo�͎��ɍŌ�� 1 �����t���Ă�����I�t�Z�b�g�C���f���g
' <center>�^�O�̕t���L��
Dim AddCenterTag As Boolean
' ���s�̎��
Dim Br As String

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

' �� CSS �o�͗p�֐�

' �n���ꂽ Range �̐��F�A�����A�����\�� CSS �������Ԃ��܂�
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

' �n���ꂽ�t�H���g�� CSS �������Ԃ��܂�
Private Function getFontCss(ByRef f As Font) As String
    Dim propStyle As String
    Dim cssTextDecoration As String
    
    ' style �v���p�e�B���̕�����쐬 (�Z�����̕����񂪕����I�ɈقȂ�X�^�C���̏ꍇ�A�e��v���p�e�B�� Null �ƂȂ�̂Œ���)
    With f
        ' font-size
        If Not IsNull(.Size) Then
            propStyle = propStyle & IIf(.Size <> DefFontSize, "font-size:" & CStr(.Size) & "pt;", "")
        End If
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
        ' vetical-align
        If Not IsNull(.Subscript) Then
            propStyle = propStyle & IIf(.Subscript = True, "vertical-align:sub;", "")
        End If
        If Not IsNull(.Superscript) Then
            propStyle = propStyle & IIf(.Superscript = True, "vertical-align:super;", "")
        End If
        ' color
        If Not IsNull(.color) Then
            If .color <> &H0 Then ' �����F = �� �ȊO�̏ꍇ�����A�����F���w�肷�� CSS ��f��
                propStyle = propStyle & IIf(.color <> &H0, "color:" & cvtToCssColor(.color) & ";", "")
            End If
        End If
    End With
    
    getFontCss = propStyle
End Function

' �w�肵���Z���̃t�H���g�A������̃A���C���A�Z���̘A�����A���̐F�E���E��ށA�w�i�F��\�� HTML �̃v���p�e�B�������Ԃ��܂��B
' �������A�Z����������̈ꕔ�������ɓK�p����Ă���X�^�C���͕ϊ��ΏۂƂ��܂���B
Private Function getCellProperties(ByRef newCellArea As Range) As String
    Dim colspan As Integer, rowspan As Integer
    Dim bgColorIndex As Variant
    Dim bgColor As Long
    Dim textAlign As Variant
    Dim verticalAlign As Variant
    
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
    
    ' �t�H���g�X�^�C���ݒ�
    propStyle = propStyle & getFontCss(newCellArea.Font)
    
    ' colspan / rowspan �v���p�e�B���̕�����쐬
    propColspan = IIf(colspan > 1, CStr(colspan), "")
    propRowspan = IIf(rowspan > 1, CStr(rowspan), "")
    
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

' �w�肵���Z�����ɕ����I�ȕ�����X�^�C�����K�p����Ă��邩�ǂ�����Ԃ��܂�
Private Function existPartialStyle(ByRef newCellArea As Range) As Boolean
    Dim ret As Boolean
    
    With newCellArea.Font
        If IsNull(.color) Or IsNull(.Bold) Or IsNull(.Italic) Or _
           IsNull(.Underline) Or IsNull(.Strikethrough) Or _
           IsNull(.Size) Or IsNull(.Subscript) Or IsNull(.Superscript) Or _
           IsNull(.ThemeFont) Or IsNull(.FontStyle) Or IsNull(.TintAndShade) Then
            ret = True
        End If
    End With
    
    existPartialStyle = ret
End Function

Private Function isSameFont(ByRef l As Font, ByRef r As Font) As Boolean
    Dim ret As Boolean
    ret = (l.Bold = r.Bold) And _
           (l.color = r.color) And _
           (l.ColorIndex = r.ColorIndex) And _
           (l.FontStyle = r.FontStyle) And _
           (l.Italic = r.Italic) And _
           (l.Size = r.Size) And _
           (l.Strikethrough = r.Strikethrough) And _
           (l.Subscript = r.Subscript) And _
           (l.Superscript = r.Superscript) And _
           (l.ThemeFont = r.ThemeFont) And _
           (l.Underline = r.Underline)

    isSameFont = ret
End Function

' HTML �p�̃G�X�P�[�v����s�̕ϊ����s���܂��B
Private Function cvtTextToHtml(ByVal str As String) As String
    str = Replace(str, "<", "&lt;")
    str = Replace(str, ">", "&gt;")
    str = Replace(str, vbNewLine, "<br>")
    str = Replace(str, vbCrLf, "<br>")
    str = Replace(str, vbLf, "<br>")
    str = Replace(str, vbCr, "<br>")
    
    cvtTextToHtml = str
End Function

' �w�肵���Z�����̕�������X�^�C���t���ŕԂ��܂��B
' �{�֐��́A�Z���̕�����̈ꕔ�����ɃX�^�C�����K�p����Ă���ꍇ�A <font> �^�O���g�p���ČʂɃX�^�C���K�p���� HTML ��Ԃ��܂��B
' �Z���S�̂ɋ��ʂ��ēK�p���ꂽ�����񏑎��ւ͖{�֐��͊֗^���܂���B
Private Function getCellValueWithStyle(ByVal newCellArea As Range) As String
    Dim cellValue As String
    Dim ret As String
    Dim i As Long
    Dim css As String
    Dim txtBuf As String
    Dim prevFont As Font
    Dim txtLen As Long
    
    ' �Z���̕�����𐮌` (�G�X�P�[�v�Ȃ�)
    cellValue = Replace(cellValue, "<", "&lt;")
    cellValue = Replace(cellValue, ">", "&gt;")
    cellValue = Replace(cellValue, vbNewLine, "<br>")
    cellValue = Replace(cellValue, vbCrLf, "<br>")
    cellValue = Replace(cellValue, vbLf, "<br>")
    cellValue = Replace(cellValue, vbCr, "<br>")
    
    ' �Z���̒l���Z�b�g
    If existPartialStyle(newCellArea) = True Then
        ' �����I�ȃX�^�C���K�p (TODO: ���������x���̂łǂ��ɂ�������)
        With newCellArea.Cells(1, 1)
            txtLen = Len(.Text)
            If txtLen > 0 Then
                Set prevFont = .Characters(1, 1).Font
            End If
            For i = 1 To txtLen + 1
                If i = txtLen + 1 Or Not isSameFont(prevFont, .Characters(i, 1).Font) Then
                    ' �t���b�V��
                    css = getFontCss(prevFont)
                    txtBuf = cvtTextToHtml(txtBuf)
                    If css = "" Then
                        ret = ret & txtBuf
                    Else
                        ret = ret & "<font style=""" & css & """>" & txtBuf & "</font>"
                    End If
                    txtBuf = ""
                End If
                Set prevFont = .Characters(i, 1).Font
                txtBuf = txtBuf & .Characters(i, 1).Text
            Next i
        End With
    Else
        cellValue = newCellArea.Cells(1, 1).Text ' �Z���̕�����͕K�� Range �̍���Z�����g�p
        ' �����I�ȃX�^�C���K�p�Ȃ�
        
        ret = cvtTextToHtml(cellValue)
    End If
    
    getCellValueWithStyle = ret
End Function

' �� HTML �^�O�o�̓��\�b�h�Q
Private Sub htmlStartNewRow(ByRef s As String)
    s = s & OfstIdt & IIf(AddCenterTag, Idt, "") & IIf(AddTableTag, Idt, "") & "<tr>" & Br
End Sub

Private Sub htmlFinishCurRow(ByRef s As String)
    s = s & OfstIdt & IIf(AddCenterTag, Idt, "") & IIf(AddTableTag, Idt, "") & "</tr>" & Br
End Sub

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

Private Sub htmlPostProcess(ByRef s As String)
    Dim additionalTableProperties As String
    Dim tblClass As String
    Dim tblId As String
    
    tblClass = GetConfValue("TableClass", "")
    tblId = GetConfValue("TableId", "")
    
    If tblClass <> "" Then
        additionalTableProperties = additionalTableProperties & "class=""" & tblClass & """ "
    End If
    If tblId <> "" Then
        additionalTableProperties = additionalTableProperties & "id=""" & tblId & """ "
    End If
    
    If AddTableTag Then
        s = OfstIdt & IIf(AddCenterTag, Idt, "") & "<table " & additionalTableProperties & "style=""border-collapse:collapse;font-size:" & DefFontSize & "pt"">" & Br & _
            s & _
            OfstIdt & IIf(AddCenterTag, Idt, "") & "</table>" & Br
    End If
    If AddCenterTag Then
        s = OfstIdt & "<center>" & Br & _
            s & _
            OfstIdt & "</center>" & Br
    End If
    ' ���s�Ȃ��ŏo�͂���ꍇ�A�I�t�Z�b�g�̃C���f���g�͕t���Ă�����
    If Br = "" Then
        s = OfstIdtBackup & s
    End If
End Sub

' �� �ݒ�t�@�C���ǂݍ���
Private Sub loadConfig()
    Dim indentType As Integer
    Dim indentOffset As Integer
    
    indentType = GetConfValue("IndentType", 0)
    indentOffset = GetConfValue("IndentOffset", 0)
    AddCenterTag = IIf(GetConfValue("AddCenterTag", 1) = 1, True, False)
    Br = IIf(GetConfValue("Nobr", 0) = 0, vbNewLine, "")
    
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
    
    If Br = "" Then ' ���s�Ȃ��ŏo�͂���ꍇ�́A�C���f���g���Ȃ��ɂ��� (�������A�擪�� 1 �����C���f���g������)
        Idt = ""
        OfstIdtBackup = OfstIdt
        OfstIdt = ""
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
    UI_Excel2Html.lbl_prgBarBg.Visible = True
    
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
                If curCell.Address = curAreaTopLeft.Address Then
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
    UI_Excel2Html.lbl_prgBarBg.Visible = False
End Function


' �� ���[�U�[�t�H�[���\�����\�b�h
Public Sub Excel2Html()
    UI_Excel2Html.Show
End Sub


