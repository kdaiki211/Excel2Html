Attribute VB_Name = "PropertyManager"
Option Explicit

Const ConfSheetName As String = "conf"
Const SearchRange As String = "A:B"
Const ResultColNum As Integer = 2

Public Function GetConfValue(ByVal key As String, ByVal defaultValueWhenNotFound As Variant) As Variant
    On Error GoTo FailGetConfValue

    ' �v���p�e�B������������l��Ԃ�
    GetConfValue = Application.WorksheetFunction.VLookup(key, ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange), ResultColNum, False)
    
    Exit Function
    
FailGetConfValue:
    On Error GoTo FatalError
    
    ' �v���p�e�B��������Ȃ�������f�t�H���g�l��ݒ肵�āA���̒l��Ԃ�
    SetConfValue key, defaultValueWhenNotFound, False
    GetConfValue = defaultValueWhenNotFound
    
    Exit Function
    
FatalError:
    MsgBox "�f�t�H���g�l�̕ۑ��Ɏ��s���܂����B"
End Function

Public Sub SetConfValue(ByVal key As String, ByVal value As Variant, Optional save As Boolean = True)
    Dim rowNum As Integer
    Dim s As Worksheet
    Dim existsConfWorksheet As Boolean

    On Error GoTo PropertyNotFound
    
    ' ���Ƀv���p�e�B�����݂�����㏑��
    rowNum = Application.WorksheetFunction.Match(key, ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).EntireColumn(1), 0)
    With ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange)
        .NumberFormatLocal = "@" ' ������ŊǗ�
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
    ' �V�[�g�����݂��Ȃ�������A�V���ɃV�[�g���쐬
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

    ' �v���p�e�B�����݂��Ȃ�������A�V�K�v���p�e�B�Ƃ��čs��ǉ�
    If Application.WorksheetFunction.CountBlank(ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).EntireColumn(1)) = 0 Then
        GoTo FailSetConfValue ' ��s�Ȃ� (�S�Ă̍s�� A �񂪖��܂��Ă���)
    End If
    
    ' ��s����
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
    MsgBox "Conf �V�[�g�̍쐬�Ɏ��s���܂����B"
    Exit Sub
FailSetConfValue:
    MsgBox "�v���p�e�B " & key & " �ɒl " & CStr(value) & " ��ݒ肷�邱�Ƃ��o���܂���ł����B"
End Sub

Public Sub CommitAllConf()
    ThisWorkbook.save
End Sub
