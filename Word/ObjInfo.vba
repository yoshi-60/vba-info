'Wordのオブジェクト一覧をファイルに出力する
Sub F_ObjOut()
  Dim myFD As FileDialog
  Dim myFDPath As String
  Dim myDoc As Document
  Dim myTbl As Table
  Dim myInShp As Object
  Dim myShp As Object
  Dim myTblNum As Integer
  Dim myShapeNum As Integer
  Dim myIlShapeNum As Integer
  Dim myTblCnt As Integer
  Dim myShapeCnt As Integer
  Dim myIlShapeCnt As Integer
  Dim myOleType As Object
  Dim myOleTypeStr As String
  Dim myOleClass As String
  Dim myWrapTypeStr As String
  Dim myShapeTop As Integer
  Dim myShapeLeft As Integer
  Dim myIlShapeLine As Integer
  Dim myIlShapeStr As String
  Dim myShapeStr As String
  Dim myOutFile As String
  Dim myOutPath As String
  Dim myOutFnum As Integer
  Dim myPageNum As Integer
  Dim rowNum As Long
  Dim colNum As Long
  Dim irow As Long
  Dim icol As Long
  Dim cellStr As String
  Dim rowStr As String
  Dim myShell As Object

  'フォルダバスを設定
  Set myShell = CreateObject("Wscript.Shell")
  myFDPath = myShell.SpecialFolders("MyDocuments")
  '出力ファイル名を設定
  myOutFile = "WordObjList.txt"

  'ファイルの選択ダイアログ
  Set myFD = Application.FileDialog(msoFileDialogFilePicker)

  'ダイアログボックスのタイトルを設定
  myFD.Title = "Wordファイルを選択してください"
  '複数ファイルの選択をオフ
  myFD.AllowMultiSelect = False
  '表示するファイルの種類の設定
  myFD.Filters.Clear
  myFD.Filters.Add "すべてのWordファイル", "*.doc; *.docx"
  '最初に表示するフォルタを設定
  myFD.InitialFileName = myFDPath
  'ファイルを選択して「OK」ボタンをクリックした場合の処理
  If myFD.Show = -1 Then
    Set myDoc = Documents.Open(myFD.SelectedItems(1))
    Debug.Print myDoc.Name
    MsgBox "ファイル： " & myDoc.Name & "を開きました"
  Else
    MsgBox "終了します"
    Exit Sub
  End If

  '表示するファイルの種類の設定を解除
  myFD.Filters.Clear
  Set myFD = Nothing

  '出力ファイルの設定
  myOutPath = myDoc.Path & Application.PathSeparator & myOutFile
  Debug.Print myOutPath
  MsgBox "出力ファイル： " & myOutPath

  '表，InlineShape，Shape の数をカウント
  myTblNum = myDoc.Tables.Count
  myIlShapeNum = myDoc.InlineShapes.Count
  myShapeNum = myDoc.Shapes.Count
  MsgBox "表の数： " & myTblNum & vbCr & "InlineShape： " & myIlShapeNum & vbCr & "Shape： " & myShapeNum

  '出力ファイルオーブン
  myOutFnum = FreeFile
  Open myOutPath For Output As myOutFnum

  '全ての表をループ
  myTblCnt = 0
  For Each myTbl In myDoc.Tables
    myTblCnt = myTblCnt + 1
    rowNum = myTbl.Rows.Count
    colNum = myTbl.Columns.Count
    myPageNum = myTbl.Range.Information(wdActiveEndPageNumber)
    MsgBox "No " & myTblCnt & myTbl.ID & "：" & myTbl.Title & "," & rowNum & "," & colNum
    Print #myOutFnum, "Table_No" & myTblCnt & " Page_" & myPageNum & ", Row: " & rowNum & ", Col: " & colNum
  Next myTbl

  '全てのInlineShapeをループ
  myIlShapeCnt = 0
  For Each myInShp In myDoc.InlineShapes
    myIlShapeCnt = myIlShapeCnt + 1
    myPageNum = myInShp.Range.Information(wdActiveEndPageNumber) 'wdActiveEndAdjustedPageNumber
    myOleTypeStr = F_IlShapeTypeStr(myInShp.Type)
    myIlShapeLine = myInShp.Range.Information(wdFirstCharacterLineNumber)
    If myInShp.Type = wdInlineShapeEmbeddedOLEObject Then
      myOleClass = " (" & myInShp.OLEFormat.ClassType & ")"
    Else
      myOleClass = ""
    End If
    MsgBox "inlineShape_No_" & myIlShapeCnt & "：" & myOleTypeStr
    Print #myOutFnum, "inlineShape_No" & myIlShapeCnt & " Page_" & myPageNum & " Line_" & myIlShapeLine & " Type_" & myInShp.Type & " " & myOleTypeStr & myOleClass
  Next myInShp

  '全てのShapeをルーブ
  myShapeCnt = 0
  For Each myShp In myDoc.Shapes
    myShapeCnt = myShapeCnt + 1
    myPageNum = myShp.Anchor.Information(wdActiveEndPageNumber)
    myOleTypeStr = F_ShapeTypeStr(myShp.Type)
    myWrapTypeStr = F_WrapTypeStr(myShp.WrapFormat.Type)
    myShapeTop = myShp.Top
    myShapeLeft = myShp.Left
    If myShp.Type = msoEmbeddedOLEObject Then
      myOleClass = " (" & myShp.OLEFormat.ClassType & ")"
    Else
      myOleClass = ""
    End If
    MsgBox "Shape_No_" & myShapeCnt & "：" & myOleTypeStr & "," & myOleClass & "," & myWrapTypeStr
    Print #myOutFnum, "Shape_No" & myShapeCnt & " Page_" & myPageNum & " Type_" & myShp.Type & " (" & myShapeLeft & "," & myShapeTop; ")," & _
      myOleTypeStr & myOleClass & " " & myWrapTypeStr
  Next myShp

  'ファイルクローズ
  Close myOutFnum
  myDoc.Close SaveChanges:=False
  Set myDoc = Nothing

End Sub

Function F_IlShapeTypeStr(typeNum As Integer) As String
  Dim typeStr(21, 1) As String
  Dim selNum As Integer
  
  If typeNum < 1 Then
    selNum = 0
  ElseIf typeNum > 20 Then
    selNum = 21
  Else
    selNum = typeNum
  End If
  
  typeStr(0, 0) = "undefined"
  typeStr(0, 1) = "不明なオブジェクト"
  typeStr(1, 0) = "wdInlineShapeEmbeddedOLEObject"
  typeStr(1, 1) = "埋め込み OLEオブジェクト"
  typeStr(2, 0) = "wdInlineShapeLinkedOLEObject"
  typeStr(2, 1) = "リンクされた OLEオブジェクト"
  typeStr(3, 0) = "wdInlineShapePicture"
  typeStr(3, 1) = "図 (ピクチャ)"
  typeStr(4, 0) = "wdInlineShapeLinkedPicture"
  typeStr(4, 1) = "リンクされた図"
  typeStr(5, 0) = "wdInlineShapeOLEControlObject"
  typeStr(5, 1) = "OLE コントロール オブジェクト"
  typeStr(6, 0) = "wdInlineShapeHorizontalLine"
  typeStr(6, 1) = "水平線"
  typeStr(7, 0) = "wdInlineShapePictureHorizontalLine"
  typeStr(7, 1) = "図 (水平線表示)"
  typeStr(8, 0) = "wdInlineShapeLinkedPictureHorizontalLine"
  typeStr(8, 1) = "リンクされた図 (水平線表示)"
  typeStr(9, 0) = "wdInlineShapePictureBullet"
  typeStr(9, 1) = "行頭文字として使用される図"
  typeStr(10, 0) = "wdInlineShapeScriptAnchor"
  typeStr(10, 0) = "スクリプトアンカー"
  typeStr(11, 0) = "wdInlineShapeOWSAnchor"
  typeStr(11, 1) = "OWS アンカー"
  typeStr(12, 0) = "wdInlineShapeChart"
  typeStr(12, 1) = "インライングラフ"
  typeStr(13, 0) = "wdInlineShapeDiagram"
  typeStr(13, 1) = "インライン図表"
  typeStr(14, 0) = "wdInlineShapeLockedCanvas"
  typeStr(14, 1) = "ロックされたインライン図形のキャンバス"
  typeStr(15, 0) = "wdInlineShapeSmartArt"
  typeStr(15, 1) = "SmartArtグラフィック"
  typeStr(16, 0) = "wdInlineShapeWebVideo"
  typeStr(16, 1) = "Web ビデオのポスターフレーム画像"
  typeStr(17, 0) = "undefined"
  typeStr(17, 1) = "不明なオブジェクト"
  typeStr(18, 0) = "undefined"
  typeStr(18, 1) = "不明なオブジェクト"
  typeStr(19, 0) = "wdInlineShape3DModel"
  typeStr(19, 1) = "3Dモデル"
  typeStr(20, 0) = "wdInlineShapeLinked3DModel"
  typeStr(20, 1) = "リンクされた 3Dモデル"
  typeStr(21, 0) = "undefined"
  typeStr(21, 1) = "不明なオブジェクト"

  F_IlShapeTypeStr = typeStr(selNum, 1)
End Function

Function F_ShapeTypeStr(typeNum As Integer) As String
  Dim typeStr(33, 1) As String
  Dim selNum As Integer
  
  If typeNum = -2 Then
    selNum = 32
  ElseIf typeNum < 0 Then
    selNum = 0
  ElseIf typeNum > 31 Then
    selNum = 33
  Else
    selNum = typeNum
  End If
  
  typeStr(0, 0) = "undefined"
  typeStr(0, 1) = "不明なオブジェクト"
  typeStr(1, 0) = "msoAutoShape"
  typeStr(1, 1) = "オートシェイプ"
  typeStr(2, 0) = "msoCallout"
  typeStr(2, 1) = "吹き出し"
  typeStr(3, 0) = " msoChart"
  typeStr(3, 1) = "グラフ"
  typeStr(4, 0) = "msoComment"
  typeStr(4, 1) = "コメント"
  typeStr(5, 0) = " msoFreeform"
  typeStr(5, 1) = "フリーフォーム"
  typeStr(6, 0) = "msoGroup"
  typeStr(6, 1) = "Group"
  typeStr(7, 0) = "msoEmbeddedOLEObject"
  typeStr(7, 1) = " 埋め込み OLE オブジェクト"
  typeStr(8, 0) = "msoFormControl"
  typeStr(8, 1) = "フォーム コントロール"
  typeStr(9, 0) = "msoLine"
  typeStr(9, 1) = "線"
  typeStr(10, 0) = "msoLinkedOLEObject"
  typeStr(10, 1) = "リンク OLE オブジェクト"
  typeStr(11, 0) = "msoLinkedPicture"
  typeStr(11, 1) = "リンク画像"
  typeStr(12, 0) = "msoOLEControlObject"
  typeStr(12, 1) = "OLE コントロール オブジェクト"
  typeStr(13, 0) = "msoPicture"
  typeStr(13, 1) = "画像"
  typeStr(14, 0) = "msoPlaceholder"
  typeStr(14, 1) = "プレースホルダー"
  typeStr(15, 0) = "msoTextEffect"
  typeStr(15, 1) = "テキスト効果"
  typeStr(16, 0) = "msoMedia"
  typeStr(16, 1) = "メディア"
  typeStr(17, 0) = "msoTextBox"
  typeStr(17, 1) = "テキスト ボックス"
  typeStr(18, 0) = "msoScriptAnchor"
  typeStr(18, 1) = "スクリプト アンカー"
  typeStr(19, 0) = "msoTable"
  typeStr(19, 1) = "テーブル"
  typeStr(20, 0) = "msoCanvas"
  typeStr(20, 1) = "キャンバス"
  typeStr(21, 0) = "msoDiagram"
  typeStr(21, 1) = "図 (ダイアグラム)"
  typeStr(22, 0) = "msoInk"
  typeStr(22, 1) = "インク"
  typeStr(23, 0) = "msoInkComment"
  typeStr(23, 1) = "インク コメント"
  typeStr(24, 0) = "msoIgxGraphic"
  typeStr(24, 1) = "SmartArt グラフィック"
  typeStr(25, 0) = "msoSlicer"
  typeStr(25, 1) = "Slicer"
  typeStr(26, 0) = "msoWebVideo"
  typeStr(26, 1) = "Web ビデオ"
  typeStr(27, 0) = "msoContentApp"
  typeStr(27, 1) = "コンテンツ Officeアドイン"
  typeStr(28, 0) = "msoGraphic"
  typeStr(28, 1) = "グラフィック"
  typeStr(29, 0) = "msoLinkedGraphic"
  typeStr(29, 1) = "リンクされたグラフィック"
  typeStr(30, 0) = "mso3DModel"
  typeStr(30, 1) = "3D モデル"
  typeStr(31, 0) = "msoLinked3DModel"
  typeStr(31, 1) = "リンクされた 3Dモデル"
  typeStr(32, 0) = "msoShapeTypeMixed"
  typeStr(32, 1) = "図形の種類の組み合わせ"
  typeStr(33, 0) = "undefined"
  typeStr(33, 1) = "不明なオブジェクト"


  F_ShapeTypeStr = typeStr(selNum, 1)
End Function

Function F_WrapTypeStr(typeNum As Integer) As String
  Dim typeStr(8, 1) As String
  Dim selNum As Integer
  
  If typeNum < 0 Then
    selNum = 8
  ElseIf typeNum > 8 Then
    selNum = 8
  Else
    selNum = typeNum
  End If
  
  typeStr(0, 0) = "wdWrapSquare"
  typeStr(0, 1) = "図形の周囲で折り返し"
  typeStr(1, 0) = "wdWrapTight"
  typeStr(1, 1) = "図形に近接している文字列を折り返します。"
  typeStr(2, 0) = "wdWrapThrough"
  typeStr(2, 1) = "図形の周囲で折り返します。"
  typeStr(3, 0) = "wdWrapNone"
  typeStr(3, 1) = "図形を文字列の前面に配置"
  typeStr(4, 0) = "wdWrapTopBottom"
  typeStr(4, 1) = "文字列を図形の上下に配置"
  typeStr(5, 0) = "wdWrapBehind"
  typeStr(5, 1) = "図形を文字列の背面に配置"
  typeStr(6, 0) = "undefined"
  typeStr(6, 1) = "配置情報"
  typeStr(7, 0) = "wdWrapInline"
  typeStr(7, 1) = "図形を行内に配置"
  typeStr(8, 0) = "undefined"
  typeStr(8, 1) = "不明な配置情報"
  
  F_WrapTypeStr = typeStr(selNum, 1)
End Function
