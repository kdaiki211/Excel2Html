Const dbg = False

Dim sh   ' �V�F�� (�J�����g�f�B���N�g���擾�p)
Dim xl   ' Excel �A�v��
Dim fso  ' File System Object
Dim file ' �o�̓t�@�C��

Dim ad   ' �A�h�C��
Dim wb   ' �ϊ��Ώۂ̕\�����郏�[�N�u�b�N
Dim html ' �o�� HTML �i�[�p�o�b�t�@

' �V�F��������
set sh = CreateObject("WScript.Shell")

' Execl ���J��
Set xl = CreateObject("Excel.Application")
xl.Visible = dbg ' Excel �̃E�B���h�E��\�����邩�ǂ���

' �o�̓t�@�C���I�[�v��
Set fso  = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile("test_table.xlsx.html", True)

' �A�h�C�����J��
Set ad = xl.Workbooks.Open(sh.CurrentDirectory & "\..\build\Excel2Html.xlam")

' �����Ώۃ��[�N�u�b�N���J��
Set wb = xl.Workbooks.Open(sh.CurrentDirectory & "\..\tools\test_table.xlsx")

' �ϊ��ΏۃZ���I�� (TODO: �֐��Ɉ����œn�����ق����X�^�C���b�V��)
wb.Sheets("ver.2").Range("C4:N24").Select

' Excel �\ -> HTML �֕ϊ�
html = ad.Application.Run("Excel2Html.ConvertSelectedRangeToHtml")

' ���ʂ��t�@�C���֏o��
file.WriteLine html

MsgBox "test_table.xlsx.html �ɏo�͂��܂���"

' ��n��
wb.Close
ad.Close
xl.Quit
