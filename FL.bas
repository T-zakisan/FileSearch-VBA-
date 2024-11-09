Attribute VB_Name = "FL"
'������������������������������������������������������������������������������������������������������������
'���[�g�f�B���N�g��
Const ROOT = "C:\Users\zakis\Downloads"
Const PASSWD = "zaq1zaq1"
'������������������������������������������������������������������������������������������������������������

'������Ԃɖ߂����c
Public Sub init()

  Application.ScreenUpdating = False

  '�s�v�ȃ��c���\��
  Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)" '���{��
  Application.DisplayStatusBar = False      '�X�e�[�^�X�o�[
  Application.DisplayFormulaBar = False     '�����o�[
  ActiveWindow.DisplayHeadings = False      '���o��
  ActiveWindow.DisplayWorkbookTabs = False  '�V�[�g�^�u


  Dim shs As Variant: shs = Array("ReadMe", "�t�@�C�����X�g")
  Dim ii As Integer
  For ii = LBound(shs, 1) To UBound(shs, 1)
    
    '�v���e�N�g(�V�[�g�ی�)
    ThisWorkbook.Worksheets(shs(ii)).Unprotect Password:=PASSWD
    ThisWorkbook.Worksheets(shs(ii)).Protect UserInterfaceOnly:=True, Password:=PASSWD
    
    '�S�����{�Z����������
    With ThisWorkbook.Worksheets(shs(ii)).Cells
      .ClearContents '�S����
      .Font.Name = "BIZ UD�S�V�b�N"
      .Font.Size = 11
      .Font.Bold = False
      .Font.Color = RGB(221, 221, 221)
      .Interior.Color = RGB(30, 32, 34)
      .HorizontalAlignment = xlLeft '����
      .VerticalAlignment = xlBottom '������
      .RowHeight = 22
    End With
    
    ThisWorkbook.Worksheets(shs(ii)).Range("A1").Font.Bold = True '�����F��
    ThisWorkbook.Worksheets(shs(ii)).Range("A1").Font.Color = RGB(255, 255, 0) '�����F�F��
  
  Next ii
  
  
  '��������ʃV�[�g�̐ݒ�
  Dim sh As Worksheet
  
  ''[ReadMe]�̏����ݒ�
  Set sh = ThisWorkbook.Worksheets(shs(0))
  sh.Range("A1:A2").Value = WorksheetFunction.Transpose(Array("���t�@�C�����X�g�ɖ߂�", ""))
  Call initReadMe(sh)

  
  ''[�t�@�C�����X�g]�̏����ݒ�
  Set sh = ThisWorkbook.Worksheets(shs(1))
  sh.Range("A1:A2").Value = WorksheetFunction.Transpose(Array("���g����������", "����"))
  Call getFile(sh, sh.Range("A2"), True, "")


  Set sh = Nothing
  Application.ScreenUpdating = True
  
End Sub




'������ �t�@�C��(Dir�܂�)���̎擾 ������
Public Sub getFile(sh As Worksheet, target As Range, Cancel As Boolean, path As String)
  
  
  '��O�����̏���
  If target.Count <> 1 Then Exit Sub  '�����Z���̑I���Ȃ�I��
  If target.Row = 1 Then Exit Sub     '1�s�ڂȂ�I��
  If target.Value = "" Then Exit Sub  '�ΏۃZ���l���󗓂Ȃ�I��


  Application.ScreenUpdating = False
  Dim mode As Integer
  If path = "" Then
    mode = 0
    path = ""
  Else
    mode = 1
    path = ROOT
  End If

  '�E�ׂ̑|��
  If mode = 1 Then
    Dim ii As Long
    For ii = target.Offset(0, 1).Column To Cells.SpecialCells(xlCellTypeLastCell).Column
      sh.Columns(ii).ClearContents '�f�[�^�̂ݍ폜
      sh.Columns(ii).Font.Color = RGB(221, 221, 221)  '�t�H���g�F
      sh.Columns(ii).Interior.Color = RGB(30, 32, 34) '�w�i�F
    Next ii
  End If
  
  
  '�J�����g�f�B���N�g��������
  If InStr(target.Value, "\") <> 0 And _
     InStr(target.Value, ".pdf") = 0 Then '�擪��\(�f�B���N�g��),���t�@�C������.pdf���܂܂�Ȃ��ꍇ
    
    sh.Columns(target.Column).Font.Color = RGB(221, 221, 221) '�t�H���g�F�F�f�t�H���g(��)
    If path <> "" Then target.Font.Color = RGB(0, 255, 0)     '�t�H���g�F�F��
    If mode = 1 Then sh.Range("A1").Font.Color = RGB(255, 255, 0)  '�t�H���g�F�F�f�t�H���g(��)
  End If
  
  
  '�����p�X����
  path = ""
  For ii = 1 To target.Column '��A����W�N���b�N������܂ł̂���Ԃ�
    Dim edRow As Long: edRow = sh.Cells(Rows.Count, ii).End(xlUp).Row '�擪�s�̎擾(����Ԃ��͈�)
    Dim stRow As Long: stRow = sh.Cells(1, ii).End(xlDown).Row        '�ŏI�s�̎擾(����Ԃ��͈�)
    If ii = 1 Then stRow = 2 '���[�g�����f�B���N�g���̏ꍇ��2�s�ڂ��炭��Ԃ��J�n
    Dim jj As Long
    For jj = stRow To edRow
      If sh.Cells(jj, ii).Font.Color = RGB(0, 255, 0) Then '�����F�F��(�I���f�B���N�g��)�̏ꍇ
        path = path & sh.Cells(jj, ii).Value '�p�X��ǋL
        Exit For 'for���甲���o��
      End If
    Next jj
  Next ii
  
  Call actFile(sh, target, path, mode)
  
  Application.ScreenUpdating = True
  
  Cancel = True
End Sub



'������ �t�@�C����ɉ��������A�N�V���� ������
Private Sub actFile(sh As Worksheet, target As Range, path As String, mode As Integer)

  '
  Dim cnt As Long: cnt = target.Row '�L�q�s
  Dim ofs As Integer: ofs = 0 '�L�q��
  If mode = 0 Then ofs = -1     '�L�q��(�J�����g�̏ꍇ)
  path = ROOT & path
  
  If InStr(target.Value, ".pdf") <> 0 Then
    
    'PDF�t�@�C���̏ꍇ�F�֘A�t�����ꂽ�v���O�����ŊJ��(Adobe Reader��u���E�U��)
    CreateObject("Shell.Application").ShellExecute path & "\" & target.Value
  ElseIf target.Value = "[!] Nothing" Then
    
    '"[!] Nothing"�F�������Ȃ� �����R�ȓ���̂��߂̏����킯
  Else
  
    'W�N���b�N�����Z�����f�B���N�g���̏ꍇ
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ff As Object
    
    If fso.FolderExists(path) = True Then '�f�B���N�g��������ꍇ
      
      '�����̃f�B���N�g�����X�g��\��
      For Each ff In fso.GetFolder(path).SubFolders
        sh.Cells(cnt, target.Offset(0, 1 + ofs).Column) = "\" & ff.Name '�f�B���N�g�����\��
        cnt = cnt + 1
      Next ff
      
      '�����f�B���N�g�����Ȃ��ꍇ�F������PDF�t�@�C�����X�g��\��
      If cnt = target.Row Then
        For Each ff In fso.GetFolder(path).Files
          
          If Mid(ff.Name, InStrRev(ff.Name, ".")) = ".pdf" Then
            sh.Cells(cnt, target.Offset(0, 1 + ofs).Column).Value = "\" & ff.Name           '�t�@�C�����\��
            sh.Cells(cnt, target.Offset(0, 1 + ofs).Column).Font.Color = RGB(255, 128, 128) '����(Adobe���ӎ�)
            cnt = cnt + 1
          End If
          
        Next ff
      End If
    End If
    
    '�f�B���N�g����PDF�t�@�C�����Ȃ��ꍇ
    If cnt = target.Row Then
      If InStr(target.Value, ".pdf") = 0 Then
        sh.Cells(cnt, target.Offset(0, 1 + ofs).Column).Value = "[!] Nothing"
        sh.Cells(cnt, target.Offset(0, 1 + ofs).Column).Font.Color = RGB(255, 255, 0)
      End If
    End If
    
    '�\�������t�@�C��(�f�B���N�g���܂�)���X�g�ɉ������񕝂̒���
    sh.Columns(target.Column + 1).Columns.AutoFit
  End If
  
  
  '�t�߂܂ŃX�N���[��
  Application.ScreenUpdating = True
'  If target.Offset(0, 1).Font.Color = RGB(255, 128, 128) Then
  If 1 Then
    If target.Column <> 1 Then
      ActiveWindow.ScrollColumn = target.Column - 1
    Else
      ActiveWindow.ScrollColumn = 1
    End If
    ActiveWindow.ScrollRow = target.Row - 1
  End If
  
End Sub



'������ �V�[�g[ReadMe]�̏�����(������Ă����������邽��) ������
Sub initReadMe(sh As Worksheet)
  
  sh.Cells.ColumnWidth = 3 '

  With sh.Cells.Range("A:A").Font
    .Size = 16
    .Bold = True
  End With
  sh.Cells.Range("A1").Font.Size = 11
  
  Call addArray(sh, Array("����͂ȂɁH", "", ""))
  Call addArray(sh, Array("�@", "����܂ł̎���ƃp�[�c���X�g����������щ{�����ł���VBA"))
  Call addArray(sh, Array("�@", "W�N���b�N�݂̂�PDF�t�@�C���܂ł��ǂ邱�Ƃ��ł���"))
  Call addArray(sh, Array("�@", "", ""))
  
  Call addArray(sh, Array("�g����", "", ""))
  Call addArray(sh, Array("�@", "1.", ""))
  Call addArray(sh, Array("�@", "", ""))
  Call addArray(sh, Array("�@", "2.�@��J�e�S����W�N���b�N", ""))
  Call addArray(sh, Array("�@", "", "���E���ɃJ�e�S�����X�g���\��"))
  Call addArray(sh, Array("�@", "3.�J�e�S����W�N���b�N", ""))
  Call addArray(sh, Array("�@", "", "���J�e�S�����X�g �������� PDF�t�@�C�����X�g(�����F�F��)���\��"))
  Call addArray(sh, Array("�@", "", "��PDF�t�@�C�����X�g���\������܂ŌJ��Ԃ�"))
  Call addArray(sh, Array("�@", "4.PDF�t�@�C��(�����F�F��)��W�N���b�N", ""))
  Call addArray(sh, Array("�@", "", "��pdf�t�@�C���Ɋ֘A�t����ꂽ�v���O����(Adobe Reader,�e��u���E�U��)�ŊJ�����"))
  Call addArray(sh, Array("�@", "", ""))
  
  
  Call addArray(sh, Array("�悭���鎿��", "", ""))
  Call addArray(sh, Array("�@", "Q.��ʃJ�e�S���ɖ߂肽���Ƃ�", ""))
  Call addArray(sh, Array("�@", "", "A.�J�������@��A���邢�̓J�e�S����W�N���b�N"))
  Call addArray(sh, Array("�@", "Q.�@�햼�������Ă��܂��� �������� ReadMe�̕����������Ă��܂����Ƃ�", ""))
  Call addArray(sh, Array("�@", "", "A.�t�@�C�����ēx�J������"))
  Call addArray(sh, Array("�@", "Q.�����������Č��Â炢", ""))
  Call addArray(sh, Array("�@", "", "A.[Ctrl] + �}�E�X�X�N���[�� �ł��D�݂̕����T�C�Y�ɁI"))
  Call addArray(sh, Array("�@", "", ""))
  
  Call addArray(sh, Array("����", "", ""))
  Call addArray(sh, Array("�@", "2024.11.11:����", ""))

End Sub



'������ �V�[�g[ReadMe]�̏�����(������Ă����������邽��) ������
Private Function addArray(sh As Worksheet, list As Variant) As Integer
  
  Dim mRow As Long: mRow = sh.Cells(Rows.Count, "A").End(xlUp).Row + 1
  If sh.Cells(Rows.Count, "A").End(xlUp).Row < sh.Cells(Rows.Count, "B").End(xlUp).Row Then
    mRow = sh.Cells(Rows.Count, "B").End(xlUp).Row
  End If
  sh.Range(sh.Cells(mRow, "A"), sh.Cells(mRow, UBound(list, 1) + 1)) = list
  
End Function

