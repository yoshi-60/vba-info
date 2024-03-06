Sub F_ObjOut()

  Dim myFD As FileDialog
  Dim myFDPath As String
  Dim myDoc As Document
  Dim myTbl As Table
  Dim myTblNum As Integer
  Dim myShapeNum As Integer
  Dim myIlShapeNum As Integer
  Dim myTblCnt As Integer
  Dim myShapeCnt As Integer
  Dim myIlShapeCnt As Integer
  Dim myOleClasst As String
  Dim myOutFile As String
  Dim myOutPath As String
  Dim myOutFnum As Integer
  Dim rowNum As Long
  Dim colNum As Long
  Dim irow As Long
  Dim icol As Long
  Dim cellStr As String
  Dim rowStr As String
  
  'フォルダパスを設定
  myFDPath = "C:\Users\UserName\Documents"
  '出力ファイル名を設定
  myOutFile = "WordObjOut.txt"

  'ファイルの選択ダイアログ
  Set myFD = Application.FileDialog(msoFileDialogFilePicker)

  'ダイアログボックスのタイトルを設定
  myFD.Title = "Wordファイルを選択してください"
  '複数ファイルの選択をオフ
  myFD.AllowMultiSelect = False
  '表示するファイルの種類の設定
  myFD.Filters.Clear
  myFD.Filters.Add "すべてのWordファイル", "*.doc; *.docx"
  '最初に表示するフォルダを設定
  myFD.InitialFileName = myFDPath
  'ファイルを選択して「OK」ボタンをクリックした場合の処理
  If myFD.Show = -1 Then
    Set myDoc = Documents.Open(myFD.SelectedItems(1))
    Debug.Print myDoc.Name
    MsgBox "ファイル： " & myDoc.Name & " を開きました。"
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
  MsgBox "出力ファイル：" & myOutPath
  
  '表, InlineShape, Shape の数をカウント
  myTblNum = myDoc.Tables.Count
  myIlShapeNum = myDoc.InlineShapes.Count
  myShapeNum = myDoc.Shapes.Count
  MsgBox "表の数： " & myTblNum & vbCr & _
    "InlineShape： " & myIlShapeNum & vbCr & "Shape： " & myShapeNum

  '出力ファイルオープン
  myOutFnum = FreeFile
  Open myOutPath For Output As myOutFnum

  '全ての表をループ
  myTblCnt = 0
  For Each myTbl In myDoc.Tables
    myTblCnt = myTblCnt + 1
    rowNum = myTbl.Rows.Count
    colNum = myTbl.Columns.Count
    'Debug.Print "No_" & myTblCnt, myTbl.ID, ":", myTbl.Title, ",", rowNum, ",", colNum
    'MsgBox "No_" & myTblCnt & myTbl.ID & ":" & myTbl.Title & "," & rowNum & "," & colNum
    Print #myOutFnum, "Table_No" & myTblCnt & "," & rowNum & "," & colNum
  Next myTbl

  '全てのInlineShapeをループ
  myIlShapeCnt = 0
  For Each myInShp In myDoc.inlineShapes
    myIlShapeCnt = myIlShapeCnt + 1
    If myInShp.Type = msoEmbeddedOLEObject Then
        myOleType =  "EmbeddedObject"
        myOleClass =  myInShp.OLEFormat.ClassType
        MsgBox "inlineShape_No_" & myIlShapeCnt & ":" & myOleType & "," & myOleClass
    End If
  Next myInShp

  
  'ファイルクローズ
  Close myOutFnum
  myDoc.Close SaveChanges:=False
  Set myDoc = Nothing

End Sub
