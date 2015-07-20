Attribute VB_Name = "Excel2Html"
Option Explicit


' ★<table>タグの付加有無
Const AddTableTag As Boolean = True

' ★基準とするフォントサイズ (これ以外のフォントサイズの場合のみ <td> の style に font-size を指定する
Const DefFontSize As Integer = 11

' インデントの種類
Dim Idt As String
Dim OfstIdt As String
Dim OfstIdtBackup As String ' 改行なしで出力時に最後に 1 つだけ付けてあげるオフセットインデント
' <center>タグの付加有無
Dim AddCenterTag As Boolean
' 改行の種類
Dim Br As String

Public CancelReq As Boolean

' ツール
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

' ■ CSS 出力用関数

' 渡された Range の線色、線幅、線種を表す CSS 文字列を返します
Private Function getLineCss(ByRef rng As Range) As String
    ' 各線の属性
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
    
    ' 定数代入
    dr(0) = xlEdgeLeft
    dr(1) = xlEdgeTop
    dr(2) = xlEdgeRight
    dr(3) = xlEdgeBottom
    drnm(0) = "left"
    drnm(1) = "top"
    drnm(2) = "right"
    drnm(3) = "bottom"
    
    ' 線の属性取得
    For i = 0 To 3
        With rng.Borders(dr(i))
            ci(i) = .ColorIndex
            c(i) = .color
            w(i) = .Weight
            Select Case .LineStyle ' 線種
                Case xlLineStyleNone, xlContinuous
                    bs(i) = "solid" ' 実線のほか、線なしも solid で扱う (太さを 0 として処理)
                Case xlDouble
                    bs(i) = "double" ' 二重線
                Case xlDot
                    bs(i) = "dotted" ' 点線
                Case xlDash, xlDashDot, xlDashDotDot, xlSlantDashDot
                    bs(i) = "dashed" ' 破線
                Case Else
                    bs(i) = "solid"
            End Select
        End With
    Next i
    
    ' 線の色が全て同じか・線の太さが全て同じかを取得
    isSameColorIndex = (ci(0) = ci(1) And ci(1) = ci(2) And ci(2) = ci(3))
    isSameColor = (c(0) = c(1) And c(1) = c(2) And c(2) = c(3))
    isSameBorderStyle = (bs(0) = bs(1) And bs(1) = bs(2) And bs(2) = bs(3))
    isSameThickness = True ' 以下の処理で値を求める
    For i = 1 To 3
        ' 透明な線は比較対象外
        If ci(i - 1) = xlColorIndexNone Or ci(i) = xlColorIndexNone Then
            GoTo continue
        End If
        ' 太さが同じか確認し、違ったら抜ける
        If w(i - 1) <> w(i) Then
            isSameThickness = False
            Exit For
        End If
continue:
    Next i
    
    ' 四辺が透明の場合
    If isSameColorIndex And ci(0) = xlColorIndexNone Then
        resultCss = ""
        GoTo last
    End If
    
    ' 上下左右の線が全て同じ線種の場合
    If isSameBorderStyle Then
        resultCss = bs(0)
    End If
    ' 上下左右の線が全て同じ色の場合 (ただし透明の場合を除く)
    If isSameColorIndex And ci(0) <> xlColorIndexNone Then
        resultCss = resultCss & " " & cvtToCssColor(c(0))
    End If
    ' 上下左右の線が全て同じ太さの場合
    If isSameThickness Then
        resultCss = resultCss & " " & cvtToCssThickness(w(0))
    End If
    resultCss = "border:" & Trim(resultCss) & ";"
    
    ' 上下左右の線が異なる色/太さの場合
    For i = 0 To 3
        If ci(i) = xlColorIndexNone Then
            ' 透明な線
            resultCss = resultCss & "border-" & drnm(i) & ":0;"
        Else
            If Not isSameBorderStyle Then
                ' 線種が統一されていない
                resultCss = resultCss & "border-" & drnm(i) & "-style:" & bs(i) & ";"
            End If
            If Not isSameColorIndex Then
                ' 線色が統一されていない
                resultCss = resultCss & "border-" & drnm(i) & "-color:" & cvtToCssColor(c(i)) & ";"
            End If
            If Not isSameThickness Then
                ' 線幅が統一されていない
                resultCss = resultCss & "border-" & drnm(i) & "-width:" & cvtToCssThickness(w(i)) & ";"
            End If
        End If
    Next i
    
last:
    getLineCss = resultCss
End Function

' 渡されたフォントの CSS 文字列を返します
Private Function getFontCss(ByRef f As Font) As String
    Dim propStyle As String
    Dim cssTextDecoration As String
    
    ' style プロパティ内の文字列作成 (セル内の文字列が部分的に異なるスタイルの場合、各種プロパティは Null となるので注意)
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
                ' 何らかの下線が引いてある場合、強制的に一重線の下線をつける (CSS で二重線引くのは諦める)
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
            If .color <> &H0 Then ' 文字色 = 黒 以外の場合だけ、文字色を指定する CSS を吐く
                propStyle = propStyle & IIf(.color <> &H0, "color:" & cvtToCssColor(.color) & ";", "")
            End If
        End If
    End With
    
    getFontCss = propStyle
End Function

' 指定したセルのフォント、文字列のアライン、セルの連結数、線の色・幅・種類、背景色を表す HTML のプロパティ文字列を返します。
' ただし、セル内文字列の一部分だけに適用されているスタイルは変換対象としません。
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
    
    ' 常に取得可能なプロパティ (Null が返らない) を取得
    colspan = newCellArea.Columns.Count
    rowspan = newCellArea.Rows.Count
    textAlign = newCellArea.Cells(1, 1).HorizontalAlignment
    verticalAlign = newCellArea.Cells(1, 1).VerticalAlignment
    bgColorIndex = newCellArea.Interior.ColorIndex
    bgColor = newCellArea.Interior.color
    
    ' フォントスタイル設定
    propStyle = propStyle & getFontCss(newCellArea.Font)
    
    ' colspan / rowspan プロパティ内の文字列作成
    propColspan = IIf(colspan > 1, CStr(colspan), "")
    propRowspan = IIf(rowspan > 1, CStr(rowspan), "")
    
    ' background
    If bgColorIndex <> xlColorIndexNone Then ' 透明以外
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
    
    ' ファイナライズ
    If propColspan <> "" Then
        ret = ret & "colspan=" & propColspan & " "
    End If
    If propRowspan <> "" Then
        ret = ret & "rowspan=" & propRowspan & " "
    End If
    If propStyle <> "" Then
        ret = ret & "style=""" & propStyle & """"
    End If
    ret = Trim(ret) ' 先頭/末尾の半角スペースを除去
    
    getCellProperties = ret
End Function

' 指定したセル内に部分的な文字列スタイルが適用されているかどうかを返します
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

' HTML 用のエスケープや改行の変換を行います。
Private Function cvtTextToHtml(ByVal str As String) As String
    str = Replace(str, "<", "&lt;")
    str = Replace(str, ">", "&gt;")
    str = Replace(str, vbNewLine, "<br>")
    str = Replace(str, vbCrLf, "<br>")
    str = Replace(str, vbLf, "<br>")
    str = Replace(str, vbCr, "<br>")
    
    cvtTextToHtml = str
End Function

' 指定したセル内の文字列をスタイル付きで返します。
' 本関数は、セルの文字列の一部だけにスタイルが適用されている場合、 <font> タグを使用して個別にスタイル適用した HTML を返します。
' セル全体に共通して適用された文字列書式へは本関数は関与しません。
Private Function getCellValueWithStyle(ByVal newCellArea As Range) As String
    Dim cellValue As String
    Dim ret As String
    Dim i As Long
    Dim css As String
    Dim txtBuf As String
    Dim prevFont As Font
    Dim txtLen As Long
    
    ' セルの文字列を整形 (エスケープなど)
    cellValue = Replace(cellValue, "<", "&lt;")
    cellValue = Replace(cellValue, ">", "&gt;")
    cellValue = Replace(cellValue, vbNewLine, "<br>")
    cellValue = Replace(cellValue, vbCrLf, "<br>")
    cellValue = Replace(cellValue, vbLf, "<br>")
    cellValue = Replace(cellValue, vbCr, "<br>")
    
    ' セルの値をセット
    If existPartialStyle(newCellArea) = True Then
        ' 部分的なスタイル適用 (TODO: 処理が超遅いのでどうにかしたい)
        With newCellArea.Cells(1, 1)
            txtLen = Len(.Text)
            If txtLen > 0 Then
                Set prevFont = .Characters(1, 1).Font
            End If
            For i = 1 To txtLen + 1
                If i = txtLen + 1 Or Not isSameFont(prevFont, .Characters(i, 1).Font) Then
                    ' フラッシュ
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
        cellValue = newCellArea.Cells(1, 1).Text ' セルの文字列は必ず Range の左上セルを使用
        ' 部分的なスタイル適用なし
        
        ret = cvtTextToHtml(cellValue)
    End If
    
    getCellValueWithStyle = ret
End Function

' ■ HTML タグ出力メソッド群
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
    
    ' <td> のプロパティ取得
    prop = getCellProperties(newCellArea)
    ' <td> ～ </td> 内の HTML 取得
    content = getCellValueWithStyle(newCellArea)
    
    ' <td> タグ追記
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
    ' 改行なしで出力する場合、オフセットのインデントは付けてあげる
    If Br = "" Then
        s = OfstIdtBackup & s
    End If
End Sub

' ■ 設定ファイル読み込み
Private Sub loadConfig()
    Dim indentType As Integer
    Dim indentOffset As Integer
    
    indentType = GetConfValue("IndentType", 0)
    indentOffset = GetConfValue("IndentOffset", 0)
    AddCenterTag = IIf(GetConfValue("AddCenterTag", 1) = 1, True, False)
    Br = IIf(GetConfValue("Nobr", 0) = 0, vbNewLine, "")
    
    ' インデントに使う文字を取得
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
            ' 異常値の時はデフォルト値に戻す
            Idt = ""
            SetConfValue "IndentType", 0
    End Select
    
    ' オフセットのインデントを生成
    If indentOffset >= 0 And indentOffset <= 4 Then
        Dim i As Integer
        OfstIdt = ""
        For i = 1 To indentOffset
            OfstIdt = OfstIdt & Idt
        Next i
    Else
        ' 異常値の時はデフォルト値に戻す
        OfstIdt = ""
        SetConfValue "IndentOffset", 0
    End If
    
    If Br = "" Then ' 改行なしで出力する場合は、インデントをなしにする (ただし、先頭に 1 つだけインデントをつける)
        Idt = ""
        OfstIdtBackup = OfstIdt
        OfstIdt = ""
    End If
End Sub

' ■ 進捗通知
Public Sub CancelConverting()
    CancelReq = True
End Sub

Private Sub updateProgressBar(ByRef numOfProcessedCells As Long, ByRef numOfEntireCells)
    Dim barWidth As Integer
    barWidth = CDbl(numOfProcessedCells) / CDbl(numOfEntireCells) * CDbl(UI_Excel2Html.lbl_progress_bg.Width)
    UI_Excel2Html.lbl_progress_fg.Width = barWidth
    DoEvents
End Sub

' ■ メイン関数
Public Function ConvertSelectedRangeToHtml() As String
    ' グローバル変数初期化
    CancelReq = False
    
    ' 設定値ロード
    loadConfig
    
    ' プログレスバー表示
    UI_Excel2Html.lbl_progress_bg.Visible = True
    UI_Excel2Html.lbl_progress_fg.Visible = True
    UI_Excel2Html.btn_cancel.Visible = True
    UI_Excel2Html.lbl_prgBarBg.Visible = True
    
    ' 選択範囲を 1 セルずつ走査
    With Selection
        Dim r As Long, c As Long
        Dim outHtml As String ' 出力 HTML 文字列 (VBA の仕様より、初期値は "")
        Dim progressUpdateInterval As Long
        Dim checkPointR As Long
        Dim checkPointC As Long
        Dim numOfEntireCells As Double
        
        numOfEntireCells = CDbl(.Rows.Count) * CDbl(.Columns.Count)
        progressUpdateInterval = 100
        checkPointR = 0
        checkPointC = 0
        
        ' 選択範囲内のセルを 1 つずつ走査 (メイン処理)
        For r = 0 To .Rows.Count - 1
        
            htmlStartNewRow outHtml
            
            For c = 0 To .Columns.Count - 1
                Dim curCell As Range, curArea As Range, curAreaTopLeft As Range
                
                Set curCell = .Cells(1 + r, 1 + c) ' 現在見ているセル (1 セル)
                Set curArea = curCell.MergeArea ' 現在見ているセルが属する結合セルの全体
                Set curAreaTopLeft = curArea.Cells(1, 1) ' 現在見ているセルが属する結合セルの左上セル (1 セル)
                
                ' r 行 c 列が結合セルの左上セルのときのみ HTML 出力する
                If curCell.Address = curAreaTopLeft.Address Then
                    htmlAddNewCell outHtml, curArea
                End If
                
                ' 進捗表示
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
    ConvertSelectedRangeToHtml = outHtml ' 戻り値を返す
    
    ' プログレスバー非表示
    UI_Excel2Html.lbl_progress_bg.Visible = False
    UI_Excel2Html.lbl_progress_fg.Visible = False
    UI_Excel2Html.btn_cancel.Visible = False
    UI_Excel2Html.lbl_prgBarBg.Visible = False
End Function


' ■ ユーザーフォーム表示メソッド
Public Sub Excel2Html()
    UI_Excel2Html.Show
End Sub


