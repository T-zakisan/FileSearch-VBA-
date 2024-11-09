VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'������ �t�@�C�����J�����Ƃ��̃C�x���g ������
Private Sub Workbook_Open()
  Call FL.init  '������
  ThisWorkbook.Worksheets("�t�@�C�����X�g").Select
End Sub



'������ W�N���b�N�����Ƃ��̃C�x���g ������
Private Sub Workbook_SheetBeforeDoubleClick(ByVal sh As Object, ByVal target As Range, Cancel As Boolean)
  
  '�V�[�g�̈ړ��̂�
  
  '�Z��[A1]�ȊO�͏I��
  If target.Address <> Range("A1").Address Then Exit Sub
  
  '�V�[�g�̈ړ�
  If sh.Name = "ReadMe" Then
    ThisWorkbook.Worksheets("�t�@�C�����X�g").Select
  Else
    ThisWorkbook.Worksheets("ReadMe").Select
  End If
  
  
  Cancel = True
End Sub



'������ �����A�N�e�B�u�ɂȂ����Ƃ��̃C�x���g ������
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
  Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)" '���{��
  Application.DisplayStatusBar = False      '�X�e�[�^�X�o�[
  Application.DisplayFormulaBar = False     '�����o�[
End Sub



'������ ������A�N�e�B�u�ɂȂ����Ƃ��̃C�x���g ������
Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
  Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)" '���{��
  Application.DisplayStatusBar = True      '�X�e�[�^�X�o�[
  Application.DisplayFormulaBar = True     '�����o�[
End Sub

