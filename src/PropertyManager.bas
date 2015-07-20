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
    SetConfValue key, defaultValueWhenNotFound
    GetConfValue = defaultValueWhenNotFound
    
    Exit Function
    
FatalError:
    MsgBox "�f�t�H���g�l�̕ۑ��Ɏ��s���܂����B"
End Function

Public Sub SetConfValue(ByVal key As String, ByVal value As Variant, Optional save As Boolean = True)
    Dim rowNum As Integer

    On Error GoTo NotFound
    
    ' ���Ƀv���p�e�B�����݂�����㏑��
    rowNum = Application.WorksheetFunction.Match(key, ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).EntireColumn(1), 0)
    ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).Cells(rowNum, ResultColNum) = value
    
    If save Then
        ThisWorkbook.save
    End If
    
    Exit Sub
    
NotFound:
    Resume Enroll
    
Enroll:
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
    
    ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).Cells(rowNum, 1) = key
    ThisWorkbook.Sheets(ConfSheetName).Range(SearchRange).Cells(rowNum, ResultColNum) = value
    
    If save Then
        ThisWorkbook.save
    End If
    
    Exit Sub
    
FailSetConfValue:
    MsgBox "�v���p�e�B " & key & " �ɒl " & CStr(value) & " ��ݒ肷�邱�Ƃ��o���܂���ł����B"
End Sub

Public Sub CommitAllConf()
    ThisWorkbook.save
End Sub