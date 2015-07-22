Const dbg = False

Dim sh          ' �V�F�� (�J�����g�f�B���N�g���擾�p)
Dim xl          ' Excel �A�v��
Dim fso         ' File System Object
Dim outFile     ' �o�̓t�@�C��
Dim outFileName ' �o�̓t�@�C����
Dim inFileName  ' ���̓t�@�C����
Dim inSheetName ' ���̓V�[�g��
Dim inRange     ' ���͔͈�
Dim arg

Dim ad   ' �A�h�C��
Dim wb   ' �ϊ��Ώۂ̕\�����郏�[�N�u�b�N
Dim html ' �o�� HTML �i�[�p�o�b�t�@
Const example = "Example: cvt.vbs [..\tools\test_table.xlsx]ver.2!C4:N24"

' �����`�F�b�N
If WScript.Arguments.Count = 0 Then
    MsgBox "Invalid arguments." & vbNewLine & _
           "Usage: cvt.vbs convert-range ..." & vbNewLine & _
           example
    WScript.Quit
End If

' �V�F��������
set sh = CreateObject("WScript.Shell")

' Execl ���J��
Set xl = CreateObject("Excel.Application")
xl.Visible = dbg ' Excel �̃E�B���h�E��\�����邩�ǂ���

' �A�h�C�����J��
Set ad = xl.Workbooks.Open(sh.CurrentDirectory & "\..\build\Excel2Html.xlam")

For Each arg In WScript.Arguments
    Dim ok
    Dim pSt, pEn
    Dim exc

    WScript.echo "* ARG = " & arg
    ' �����p�[�X
    pSt = InStr(arg, "[")
    pEn = InStrRev(arg, "]")
    ok = False
    If pSt = 0 Or pEn = 0 Then
        WScript.echo "Error: Invalid convert-range format."
    ElseIf pSt + 1 = pEn Then
        WScript.echo "Error: No filename specified."
    Else
        inFileName = Mid(arg, pSt + 1, pEn - pSt - 1)
        ok = True
    End If

    If ok Then
        ' �t���p�X�w��ł͂Ȃ����̓J�����g�f�B���N�g����₤
        If Mid(inFileName, 2, 1) <> ":" And Left(inFileName, 2) <> "\\" Then
            inFileName = sh.CurrentDirectory & "\" & inFileName
        End If
        outFileName = inFileName & ".html"

        ' �V�[�g���擾
        exc = InStrRev(arg, "!")
        If exc = 0 Then
            WScript.echo "Error: Exclamation mark not found"
            ok = False
        ElseIf pEn + 1 = exc Then
            WScript.echo "Error: Sheet name not specified."
            ok = False
        End If
    End If

    If ok Then
        inSheetName = Mid(arg, pEn + 1, exc - pEn - 1)
        inRange = Mid(arg, exc + 1)
        If Len(inRange) = 0 Then
            WScript.echo "Error: Range not specified"
            ok = False
        End If
    End If

    If ok Then
        ' �����Ώۃ��[�N�u�b�N���J��
        WScript.echo "Opening " & inFileName & "..."
        Set wb = xl.Workbooks.Open(inFileName)

        ' �o�̓t�@�C���I�[�v��
        Set fso  = CreateObject("Scripting.FileSystemObject")
        Set outFile = fso.CreateTextFile(outFileName, True)

        ' �ϊ��ΏۃZ���I�� (TODO: �֐��Ɉ����œn�����ق����X�^�C���b�V��)
        wb.Sheets(inSheetName).Range(inRange).Select

        ' Excel �\ -> HTML �֕ϊ�
        WScript.echo "Converting..."
        html = ad.Application.Run("Excel2Html.ConvertSelectedRangeToHtml")

        ' ���ʂ��t�@�C���֏o�͂��ĕ���
        outFile.WriteLine html
        outFile.Close

        ' ���[�N�u�b�N�����
        wb.Close
        WScript.echo "Success." & vbNewLine
    End If
Next

' ��n��
ad.Close
xl.Quit
WScript.echo "Done."
