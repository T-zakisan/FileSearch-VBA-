VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'■■■ ファイルを開いたときのイベント ■■■
Private Sub Workbook_Open()
  Call FL.init  '初期化
  ThisWorkbook.Worksheets("ファイルリスト").Select
End Sub



'■■■ Wクリックしたときのイベント ■■■
Private Sub Workbook_SheetBeforeDoubleClick(ByVal sh As Object, ByVal target As Range, Cancel As Boolean)
  
  'シートの移動のみ
  
  'セル[A1]以外は終了
  If target.Address <> Range("A1").Address Then Exit Sub
  
  'シートの移動
  If sh.Name = "ReadMe" Then
    ThisWorkbook.Worksheets("ファイルリスト").Select
  Else
    ThisWorkbook.Worksheets("ReadMe").Select
  End If
  
  
  Cancel = True
End Sub



'■■■ 窓がアクティブになったときのイベント ■■■
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
  Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)" 'リボン
  Application.DisplayStatusBar = False      'ステータスバー
  Application.DisplayFormulaBar = False     '数式バー
End Sub



'■■■ 窓が非アクティブになったときのイベント ■■■
Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
  Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)" 'リボン
  Application.DisplayStatusBar = True      'ステータスバー
  Application.DisplayFormulaBar = True     '数式バー
End Sub


