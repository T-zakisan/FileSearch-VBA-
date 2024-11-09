VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


'Wクリックしたときのイベント
Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
  Call FL.getFile(Me, target, Cancel, target.Value)
End Sub

