Attribute VB_Name = "PropertyManager"
Option Explicit

Const ConfSheetName As String = "conf"
Const SearchRange As String = "A:B"
Const ResultColNum As Integer = 2

Public Function GetConfValue(ByVal key As String, ByVal defaultValueWhenNotFound As Variant) As Variant
    On Error GoTo FailGetConfValue

    ' プロパティが見つかったら値を返す
    GetConfValue = Application.WorksheetFunction.VLookup(key, ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange), ResultColNum, False)
    
    Exit Function
    
FailGetConfValue:
    On Error GoTo FatalError
    
    ' プロパティが見つからなかったらデフォルト値を設定して、その値を返す
    SetConfValue key, defaultValueWhenNotFound, False
    GetConfValue = defaultValueWhenNotFound
    
    Exit Function
    
FatalError:
    MsgBox "デフォルト値の保存に失敗しました。"
End Function

Public Sub SetConfValue(ByVal key As String, ByVal value As Variant, Optional save As Boolean = True)
    Dim rowNum As Integer
    Dim s As Worksheet
    Dim existsConfWorksheet As Boolean

    On Error GoTo PropertyNotFound
    
    ' 既にプロパティが存在したら上書き
    rowNum = Application.WorksheetFunction.Match(key, ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).EntireColumn(1), 0)
    With ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange)
        .NumberFormatLocal = "@" ' 文字列で管理
        .Cells(rowNum, ResultColNum).value = value
    End With
    
    If save Then
        ThisWorkbook.save
    End If
    
    Exit Sub

PropertyNotFound:
    Resume CheckConfSheet
    
CheckConfSheet:
    On Error GoTo FailCreateNewSheet
    ' シートが存在しなかったら、新たにシートを作成
    existsConfWorksheet = False
    For Each s In ThisWorkbook.Worksheets
        If s.Name = ConfSheetName Then
            existsConfWorksheet = True
            Exit For
        End If
    Next

    If Not existsConfWorksheet Then
        Dim newSheet As Worksheet
        Set newSheet = ThisWorkbook.Worksheets.Add
        newSheet.Name = ConfSheetName
    End If

AddNewProperty:
    On Error GoTo FailSetConfValue

    ' プロパティが存在しなかったら、新規プロパティとして行を追加
    If Application.WorksheetFunction.CountBlank(ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).EntireColumn(1)) = 0 Then
        GoTo FailSetConfValue ' 空行なし (全ての行の A 列が埋まっている)
    End If
    
    ' 空行検索
    With ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).Cells(1, 1)
        If .Text = "" Then
            rowNum = 1
        ElseIf .Offset(1, 0).Text = "" Then
            rowNum = 2
        Else
            rowNum = .End(xlDown).Row + 1
        End If
    End With
    
    ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).Cells(rowNum, 1).value = key
    With ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).Cells(rowNum, ResultColNum)
        .NumberFormatLocal = "@"
        .value = value
    End With
    
    If save Then
        ThisWorkbook.save
    End If
    
    Exit Sub
    
FailCreateNewSheet:
    MsgBox "Conf シートの作成に失敗しました。"
    Exit Sub
FailSetConfValue:
    MsgBox "プロパティ " & key & " に値 " & CStr(value) & " を設定することが出来ませんでした。"
End Sub

Public Sub CommitAllConf()
    ThisWorkbook.save
End Sub
