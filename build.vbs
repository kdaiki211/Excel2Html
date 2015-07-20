Option Explicit

' �I�v�V����
Const dbg = False
Const ProjectName = "Excel2Html"
Const Extension = "xlam"
Const SrcDir = "src"
Const OutDir = "build"

' �萔
Const xlOpenXMLAddIn = 55

' �ϐ�
Dim sh   ' �V�F�� (�J�����g�f�B���N�g���擾�p)
Dim fso  ' FileSystemObject
Dim fo   ' folder
Dim xl   ' Excel �A�v��
Dim wb   ' �{�X�N���v�g���V�K�쐬���郏�[�N�u�b�N
Dim vb   ' VBComponents
Dim pwd  ' �J�����g�f�B���N�g��
Dim f
Dim logStr

' ���C������ ------------------------------------------------

Sub PostProcess(ByRef wb, ByRef xl)
    wb.Close
    xl.Quit
End Sub

' �V�F��������
Set sh = CreateObject("WScript.Shell")
pwd = sh.CurrentDirectory

' FSO ����
Set fso = CreateObject("Scripting.FileSystemObject")

' Execl ���J��
Set xl = CreateObject("Excel.Application")
xl.Visible = dbg ' Excel �̃E�B���h�E��\�����邩�ǂ���

' �A�h�C���𐶐�
Set wb = xl.Workbooks.Add

On Error Resume Next
Set vb = wb.VBProject.VBComponents

If Err.Number = 1004 Then ' �Z�L�����e�B�ݒ肪�f�t�H���g�̏ꍇ
    MsgBox "Failed to get VBProject.VBComponents." & vbNewLine & _
           "To solve this problem, go to Excel's setting and " & _
           "select Security pane, then check ""Trust Access to Visual Basic Project.""." & vbNewLine & _
           "Retry executing the build script after the setting.", _
           vbExclamation, _
           "Error"
    PostProcess wb, xl
    WScript.Quit Err.Number
End If
On Error GoTo 0

' �e��t�@�C���C���|�[�g
Set fo = fso.GetFolder(pwd & "\" & SrcDir)
For Each f in fo.Files
    If Right(f.Name, 4) <> ".frx" Then ' frx file will be imported automatically when loading frm file
        logStr = logStr & vbNewLine & f.Name
        vb.Import pwd & "\" & SrcDir & "\" & f.Name
    End If
Next

' �m�F���b�Z�[�W�}�~
xl.DisplayAlerts = False

' �o�̓t�H���_�쐬 (���݂��Ȃ��ꍇ�̂�)
If Not fso.FolderExists(pwd & "\" & OutDir) Then
    fso.CreateFolder(pwd & "\" & OutDir)
End If

' �A�h�C���ۑ�
On Error Resume Next
wb.SaveAs pwd & "\" & OutDir & "\" & ProjectName & "." & Extension, xlOpenXMLAddIn

If Err.Number <> 0 Then
    MsgBox "Failed to output the addin file.", vbExclamation, "Error"
    PostProcess wb, xl
    WScript.Quit Err.Number
End If
On Error GoTo 0

PostProcess wb, xl

' �o�͌��ʕ\��
MsgBox "Success building." & vbNewLine & _
       "Following components are included in the output file." & vbNewLine & _
       logStr, _
       vbInformation, _
       "Success Building"
