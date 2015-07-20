Attribute VB_Name = "Excel2Html"
Option Explicit

' ★改行の種類 (お好きなものに変更可)
Const Br As String = vbNewLine

' ★<table>タグの付加有無
Const AddTableTag As Boolean = True

' インデントの種類
Dim Idt As String
Dim OfstIdt As String
Dim AddCenterTag As Boolean ' <center>タグの付加有無

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

' ■ CSS 出力用関数
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
    
    ' 線の属性取得
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
    
    ' 上下左右の線が全て同じ色の場合
    If isSameColor Then
        resultCss = resultCss & " " & cvtToCssColor(cl)
    End If
    
    ' 上下左右の線が全て同じ太さの場合
    If isSameThickness Then
        resultCss = resultCss & " " & cvtToCssThickness(wl)
    End If
    
    resultCss = resultCss & ";"
    
    
    ' 上下左右の線が異なる色/太さの場合、できるだけ短い CSS コードを出力できるように心がける
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

' ■ HTML タグ出力メソッド群
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
    
    ' 結合セルのプロパティを取得
    cellValue = newCellArea.Cells(1, 1).Text ' セルの文字列は必ず Range の左上セルを使用
    colspan = newCellArea.Columns.Count
    rowspan = newCellArea.Rows.Count
    bgColor = newCellArea.Interior.color
    color = newCellArea.Font.color
    textAlign = newCellArea.Cells(1, 1).HorizontalAlignment
    verticalAlign = newCellArea.Cells(1, 1).VerticalAlignment
    isBold = newCellArea.Cells(1, 1).Font.Bold
    
    s = s & OfstIdt & IIf(AddCenterTag, Idt, "") & IIf(AddTableTag, Idt, "") & Idt & "<td"
    
    ' 行方向の連結
    If colspan > 1 Then
        s = s & " colspan=" & CStr(colspan)
    End If
    
    ' 列方向の連結
    If rowspan > 1 Then
        s = s & " rowspan=" & CStr(rowspan)
    End If
    
    ' CSS
    s = s & " style="""
    
    ' CSS: 背景色
    If bgColor <> &HFFFFFF Then ' 白以外
        s = s & "background:" & cvtToCssColor(bgColor) & ";"
    End If
    
    ' CSS: 文字色
    If color <> &H0 Then ' 黒以外
        s = s & "color:" & cvtToCssColor(color) & ";"
    End If
    
    ' CSS: テキストの水平方向アライン
    If textAlign = xlCenter Then
        s = s & "text-align:center;"
    ElseIf textAlign = xlRight Then
        s = s & "text-align:right;"
    End If
    
    ' CSS: テキストの垂直方向アライン
    If verticalAlign = xlVAlignCenter Then
        s = s & "vertical-align:middle;"
    ElseIf verticalAlign = xlVAlignBottom Then
        s = s & "vertical-align:bottom;"
    Else
        s = s & "vertical-align:top;"
    End If
    
    ' CSS: 太字
    If isBold Then
        s = s & "font-weight:bold;"
    End If
    
    ' CSS: border
    s = s & getLineCss(newCellArea) ' 線の色、線の太さを表す CSS 文字列を取得

    s = s & """>"
    
    ' セルの文字列
    s = s & cellValue
    s = s & "</td>" & Br
End Sub

' ■ 設定ファイル読み込み
Private Sub loadConfig()
    Dim indentType As Integer
    Dim indentOffset As Integer
    
    indentType = GetConfValue("IndentType", 0)
    indentOffset = GetConfValue("IndentOffset", 0)
    AddCenterTag = IIf(GetConfValue("AddCenterTag", 1) = 1, True, False)
    
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
                If curCell = curAreaTopLeft Then
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
End Function


' ■ ユーザーフォーム表示メソッド
Public Sub Excel2Html()
    UI_Excel2Html.Show
End Sub


