Attribute VB_Name = "FL"
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'ルートディレクトリ
Const ROOT = "C:\Users\zakis\Downloads"
Const PASSWD = "zaq1zaq1"
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

'初期状態に戻すヤツ
Public Sub init()

  Application.ScreenUpdating = False

  '不要なヤツを非表示
  Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)" 'リボン
  Application.DisplayStatusBar = False      'ステータスバー
  Application.DisplayFormulaBar = False     '数式バー
  ActiveWindow.DisplayHeadings = False      '見出し
  ActiveWindow.DisplayWorkbookTabs = False  'シートタブ


  Dim shs As Variant: shs = Array("ReadMe", "ファイルリスト")
  Dim ii As Integer
  For ii = LBound(shs, 1) To UBound(shs, 1)
    
    'プロテクト(シート保護)
    ThisWorkbook.Worksheets(shs(ii)).Unprotect Password:=PASSWD
    ThisWorkbook.Worksheets(shs(ii)).Protect UserInterfaceOnly:=True, Password:=PASSWD
    
    '全消し＋セル属性整え
    With ThisWorkbook.Worksheets(shs(ii)).Cells
      .ClearContents '全消し
      .Font.Name = "BIZ UDゴシック"
      .Font.Size = 11
      .Font.Bold = False
      .Font.Color = RGB(221, 221, 221)
      .Interior.Color = RGB(30, 32, 34)
      .HorizontalAlignment = xlLeft '左寄せ
      .VerticalAlignment = xlBottom '下揃え
      .RowHeight = 22
    End With
    
    ThisWorkbook.Worksheets(shs(ii)).Range("A1").Font.Bold = True '文字：太
    ThisWorkbook.Worksheets(shs(ii)).Range("A1").Font.Color = RGB(255, 255, 0) '文字色：黄
  
  Next ii
  
  
  'ここから個別シートの設定
  Dim sh As Worksheet
  
  ''[ReadMe]の初期設定
  Set sh = ThisWorkbook.Worksheets(shs(0))
  sh.Range("A1:A2").Value = WorksheetFunction.Transpose(Array("■ファイルリストに戻る", ""))
  Call initReadMe(sh)

  
  ''[ファイルリスト]の初期設定
  Set sh = ThisWorkbook.Worksheets(shs(1))
  sh.Range("A1:A2").Value = WorksheetFunction.Transpose(Array("■使い方を見る", "分類"))
  Call getFile(sh, sh.Range("A2"), True, "")


  Set sh = Nothing
  Application.ScreenUpdating = True
  
End Sub




'■■■ ファイル(Dir含む)名の取得 ■■■
Public Sub getFile(sh As Worksheet, target As Range, Cancel As Boolean, path As String)
  
  
  '門前払いの条件
  If target.Count <> 1 Then Exit Sub  '複数セルの選択なら終了
  If target.Row = 1 Then Exit Sub     '1行目なら終了
  If target.Value = "" Then Exit Sub  '対象セル値が空欄なら終了


  Application.ScreenUpdating = False
  Dim mode As Integer
  If path = "" Then
    mode = 0
    path = ""
  Else
    mode = 1
    path = ROOT
  End If

  '右隣の掃除
  If mode = 1 Then
    Dim ii As Long
    For ii = target.Offset(0, 1).Column To Cells.SpecialCells(xlCellTypeLastCell).Column
      sh.Columns(ii).ClearContents 'データのみ削除
      sh.Columns(ii).Font.Color = RGB(221, 221, 221)  'フォント色
      sh.Columns(ii).Interior.Color = RGB(30, 32, 34) '背景色
    Next ii
  End If
  
  
  'カレントディレクトリを強調
  If InStr(target.Value, "\") <> 0 And _
     InStr(target.Value, ".pdf") = 0 Then '先頭に\(ディレクトリ),かつファイル名に.pdfが含まれない場合
    
    sh.Columns(target.Column).Font.Color = RGB(221, 221, 221) 'フォント色：デフォルト(白)
    If path <> "" Then target.Font.Color = RGB(0, 255, 0)     'フォント色：緑
    If mode = 1 Then sh.Range("A1").Font.Color = RGB(255, 255, 0)  'フォント色：デフォルト(黄)
  End If
  
  
  '検索パス生成
  path = ""
  For ii = 1 To target.Column '列AからWクリックした列までのくり返し
    Dim edRow As Long: edRow = sh.Cells(Rows.Count, ii).End(xlUp).Row '先頭行の取得(くり返し範囲)
    Dim stRow As Long: stRow = sh.Cells(1, ii).End(xlDown).Row        '最終行の取得(くり返し範囲)
    If ii = 1 Then stRow = 2 'ルート直下ディレクトリの場合は2行目からくり返し開始
    Dim jj As Long
    For jj = stRow To edRow
      If sh.Cells(jj, ii).Font.Color = RGB(0, 255, 0) Then '文字色：緑(選択ディレクトリ)の場合
        path = path & sh.Cells(jj, ii).Value 'パスを追記
        Exit For 'forから抜け出し
      End If
    Next jj
  Next ii
  
  Call actFile(sh, target, path, mode)
  
  Application.ScreenUpdating = True
  
  Cancel = True
End Sub



'■■■ ファイル種に応じたリアクション ■■■
Private Sub actFile(sh As Worksheet, target As Range, path As String, mode As Integer)

  '
  Dim cnt As Long: cnt = target.Row '記述行
  Dim ofs As Integer: ofs = 0 '記述列
  If mode = 0 Then ofs = -1     '記述列(カレントの場合)
  path = ROOT & path
  
  If InStr(target.Value, ".pdf") <> 0 Then
    
    'PDFファイルの場合：関連付けされたプログラムで開く(Adobe Readerやブラウザ等)
    CreateObject("Shell.Application").ShellExecute path & "\" & target.Value
  ElseIf target.Value = "[!] Nothing" Then
    
    '"[!] Nothing"：何もしない ※自然な動作のための条件わけ
  Else
  
    'Wクリックしたセルがディレクトリの場合
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ff As Object
    
    If fso.FolderExists(path) = True Then 'ディレクトリがある場合
      
      '直下のディレクトリリストを表示
      For Each ff In fso.GetFolder(path).SubFolders
        sh.Cells(cnt, target.Offset(0, 1 + ofs).Column) = "\" & ff.Name 'ディレクトリ名表示
        cnt = cnt + 1
      Next ff
      
      '直下ディレクトリがない場合：直下のPDFファイルリストを表示
      If cnt = target.Row Then
        For Each ff In fso.GetFolder(path).Files
          
          If Mid(ff.Name, InStrRev(ff.Name, ".")) = ".pdf" Then
            sh.Cells(cnt, target.Offset(0, 1 + ofs).Column).Value = "\" & ff.Name           'ファイル名表示
            sh.Cells(cnt, target.Offset(0, 1 + ofs).Column).Font.Color = RGB(255, 128, 128) 'やや赤(Adobeを意識)
            cnt = cnt + 1
          End If
          
        Next ff
      End If
    End If
    
    'ディレクトリもPDFファイルもない場合
    If cnt = target.Row Then
      If InStr(target.Value, ".pdf") = 0 Then
        sh.Cells(cnt, target.Offset(0, 1 + ofs).Column).Value = "[!] Nothing"
        sh.Cells(cnt, target.Offset(0, 1 + ofs).Column).Font.Color = RGB(255, 255, 0)
      End If
    End If
    
    '表示したファイル(ディレクトリ含む)リストに応じた列幅の調整
    sh.Columns(target.Column + 1).Columns.AutoFit
  End If
  
  
  '付近までスクロール
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



'■■■ シート[ReadMe]の初期化(消されても復活させるため) ■■■
Sub initReadMe(sh As Worksheet)
  
  sh.Cells.ColumnWidth = 3 '

  With sh.Cells.Range("A:A").Font
    .Size = 16
    .Bold = True
  End With
  sh.Cells.Range("A1").Font.Size = 11
  
  Call addArray(sh, Array("これはなに？", "", ""))
  Call addArray(sh, Array("　", "これまでの取説とパーツリストを検索および閲覧ができるVBA"))
  Call addArray(sh, Array("　", "WクリックのみでPDFファイルまでたどることができる"))
  Call addArray(sh, Array("　", "", ""))
  
  Call addArray(sh, Array("使い方", "", ""))
  Call addArray(sh, Array("　", "1.", ""))
  Call addArray(sh, Array("　", "", ""))
  Call addArray(sh, Array("　", "2.機種カテゴリをWクリック", ""))
  Call addArray(sh, Array("　", "", "→右側にカテゴリリストが表示"))
  Call addArray(sh, Array("　", "3.カテゴリをWクリック", ""))
  Call addArray(sh, Array("　", "", "→カテゴリリスト もしくは PDFファイルリスト(文字色：赤)が表示"))
  Call addArray(sh, Array("　", "", "※PDFファイルリストが表示するまで繰り返す"))
  Call addArray(sh, Array("　", "4.PDFファイル(文字色：赤)をWクリック", ""))
  Call addArray(sh, Array("　", "", "→pdfファイルに関連付けられたプログラム(Adobe Reader,各種ブラウザ等)で開かれる"))
  Call addArray(sh, Array("　", "", ""))
  
  
  Call addArray(sh, Array("よくある質問", "", ""))
  Call addArray(sh, Array("　", "Q.上位カテゴリに戻りたいとき", ""))
  Call addArray(sh, Array("　", "", "A.開きたい機種、あるいはカテゴリをWクリック"))
  Call addArray(sh, Array("　", "Q.機種名を消してしまった もしくは ReadMeの文字を消してしまったとき", ""))
  Call addArray(sh, Array("　", "", "A.ファイルを再度開き直す"))
  Call addArray(sh, Array("　", "Q.字が小さくて見づらい", ""))
  Call addArray(sh, Array("　", "", "A.[Ctrl] + マウススクロール でお好みの文字サイズに！"))
  Call addArray(sh, Array("　", "", ""))
  
  Call addArray(sh, Array("履歴", "", ""))
  Call addArray(sh, Array("　", "2024.11.11:初版", ""))

End Sub



'■■■ シート[ReadMe]の初期化(消されても復活させるため) ■■■
Private Function addArray(sh As Worksheet, list As Variant) As Integer
  
  Dim mRow As Long: mRow = sh.Cells(Rows.Count, "A").End(xlUp).Row + 1
  If sh.Cells(Rows.Count, "A").End(xlUp).Row < sh.Cells(Rows.Count, "B").End(xlUp).Row Then
    mRow = sh.Cells(Rows.Count, "B").End(xlUp).Row
  End If
  sh.Range(sh.Cells(mRow, "A"), sh.Cells(mRow, UBound(list, 1) + 1)) = list
  
End Function

