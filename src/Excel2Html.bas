Attribute VB_Name = "Excel2Html"
Option Explicit

' ★インデントの種類 (お好きなものに変更可)
Const Idt As String = vbTab

' ★改行の種類 (お好きなものに変更可)
Const Br As String = vbNewLine

' ★<table>タグの付加有無
Const AddTableTag As Boolean = True


' ■ HTML タグ出力メソッド群
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
    
    ' 結合セルのプロパティを取得
    cellValue = newCellArea.Cells(1, 1).Text ' セルの文字列は必ず Range の左上セルを使用
    colspan = newCellArea.Columns.Count
    rowspan = newCellArea.Rows.Count
    bgColor = bgr2Rgb(newCellArea.Interior.color)
    color = bgr2Rgb(newCellArea.Font.color)
    textAlign = newCellArea.Cells(1, 1).HorizontalAlignment
    isBold = newCellArea.Cells(1, 1).Font.Bold
    
    s = s & Idt & Idt & "<td"
    
    ' 背景色
    If bgColor <> &HFFFFFF Then
        s = s & " bgColor=""#" & padLeft(Hex(bgColor), "0", 6) & """"
    End If
    
    ' 行方向の連結
    If colspan > 1 Then
        s = s & " colspan=" & CStr(colspan)
    End If
    
    ' 列方向の連結
    If rowspan > 1 Then
        s = s & " rowspan=" & CStr(rowspan)
    End If
    
    ' テキストの水平方向のアライン
    If textAlign = xlCenter Then
        s = s & " align=""center"""
    ElseIf textAlign = xlRight Then
        s = s & " align=""right"""
    End If

    s = s & ">"
    
    ' font タグ開始
    If color <> &H0 Then
        s = s & "<font color=""#" & padLeft(Hex(color), "0", 6) & """>"
    End If
    
    ' b タグ開始
    If isBold Then
        s = s & "<b>"
    End If
    
    ' セルの文字列
    s = s & cellValue
    
    ' b タグ終了
    If isBold Then
        s = s & "</b>"
    End If
     
    ' font タグ終了
    If color <> &H0 Then
        s = s & "</font>"
    End If
    
    s = s & "</td>" & Br
End Sub


' ■ メイン関数
Function ConvertSelectedRangeToHtml() As String
    ' 選択範囲を 1 セルずつ走査
    With Selection
        Dim r As Integer, c As Integer
        Dim selTopLeft As Range ' 選択範囲の左上セル
        Dim outHtml As String ' 出力 HTML 文字列 (VBA の仕様より、初期値は "")
        
        ' 選択範囲の左上セル取得
        Set selTopLeft = .Cells(1, 1)
        
        ' 選択範囲内のセルを 1 つずつ走査 (メイン処理)
        For r = 0 To .Rows.Count - 1
            htmlStartNewRow outHtml
            
            For c = 0 To .Columns.Count - 1
                Dim curCell As Range, curArea As Range, curAreaTopLeft As Range
                
                Set curCell = selTopLeft.Offset(r, c) ' 現在見ているセル (1 セル)
                Set curArea = curCell.MergeArea ' 現在見ているセルが属する結合セルの全体
                Set curAreaTopLeft = curArea.Cells(1, 1) ' 現在見ているセルが属する結合セルの左上セル (1 セル)
                
                ' r 行 c 列が結合セルの左上セルのときのみ HTML 出力する
                If curCell = curAreaTopLeft Then
                    htmlAddNewCell outHtml, curArea
                End If
            Next c
            
            htmlFinishCurRow outHtml
        Next r
    End With
    
    htmlPostProcess outHtml
    ConvertSelectedRangeToHtml = outHtml ' 戻り値を返す
End Function


' ■ ユーザーフォーム表示メソッド
Public Sub Excel2Html()
    UI_Excel2Html.Show vbModeless
End Sub
