Option Explicit

Sub Edit_Book()
  Dim wb0 As Workbook
  Dim wb1 As Workbook
  Dim ws0 As Worksheet
  Dim ws1 As Worksheet
  Dim ws0Name As String
  Dim bookName As String
  Dim sheetName As String
  Dim cellStr As String
  Dim nameCol As Integer
  Dim bookRow As Integer
  Dim sheetRow As Integer
  Dim cellRow As Integer
  Dim selRow As Integer
  Dim resultRow As Integer
  Dim param1Row As Integer
  Dim param2Row As Integer
  Dim param3Row As Integer
  Dim rowStaNum As Long
  Dim rowEndNum As Long
  Dim rowCount As Long
  Dim colStaNum As Long
  Dim colEndNum As Long
  Dim colCount As Long
  Dim selectNum As Integer
  Dim resultNum As Long
  Dim param1Num As Integer
  Dim param2Str As String
  Dim param3Str As String

  nameCol = 3
  bookRow = 3
  sheetRow = 4
  cellRow = 5
  selRow = 7
  param1Row = 8
  param2Row = 9
  param3Row = 10
  resultRow = 14

  ws0Name = ActiveSheet.Name
  Set wb0 = ThisWorkbook
  Set ws0 = wb0.Worksheets(ws0Name)

  bookName = ws0.Cells(bookRow, nameCol).Value
  sheetName = ws0.Cells(sheetRow, nameCol).Value
  cellStr = ws0.Cells(cellRow, nameCol).Value
  selectNum = ws0.Cells(selRow, nameCol).Value
  param1Num = ws0.Cells(param1Row, nameCol).Value
  param2Str = ws0.Cells(param2Row, nameCol).Value
  param3Str = ws0.Cells(param3Row, nameCol).Value

  Set wb1 = Workbooks(bookName)
  Set ws1 = wb1.Worksheets(sheetName)

  rowStaNum = ws1.Range(cellStr).Row
  colStaNum = ws1.Range(cellStr).Column
  rowCount = ws1.Range(cellStr).Rows.Count
  colCount = ws1.Range(cellStr).Columns.Count
  rowEndNum = ws1.Range(cellStr).Rows(rowCount).Row
  colEndNum = ws1.Range(cellStr).Columns(colCount).Column

  resultNum = 0
  Select Case selectNum
    Case 1
      Call Edit_Func1(ws1, rowStaNum, colStaNum, rowEndNum, colEndNum, param1Num, param2Str, resultNum)
    Case 2
      Call Edit_Func2(ws1, rowStaNum, colStaNum, rowEndNum, colEndNum, param1Num, param2Str, param3Str, resultNum)
    Case Else
      Debug.Print "Not Defined Select: ", selectNum
  End Select
  ws0.Cells(resultRow, nameCol) = resultNum

End Sub

Sub Edit_Func1(ws As Worksheet, rowSta As Long, colSta As Long, rowEnd As Long, colEnd As Long, p1Num As Integer, p2Str As String, retVal As Long)
  Dim cellStr As String
  Dim searchStr As String
  Dim findNum As Integer
  Dim outlineLevel As Integer
  Dim rowI As Long
  Dim funcCount As Long

  searchStr = p2Str
  outlineLevel = p1Num

  funcCount = 0
  For rowI = rowSta To rowEnd
    cellStr = ws.Cells(rowI, colSta).Value
    findNum = InStr(cellStr, searchStr)
    If findNum = 1 Then
      GoTo loopContinue
    Else
      ws.Rows(rowI).outlineLevel = p1Num
     funcCount = funcCount + 1
    End If
loopContinue:
  Next rowI
  retVal = funcCount
End Sub

Sub Edit_Func2(ws As Worksheet, rowSta As Long, colSta As Long, rowEnd As Long, colEnd As Long, p1Num As Integer, p2Str As String, p3Str As String, retVal As Long)
  Dim cellStr As String
  Dim searchStr As String
  Dim findNum As Integer
  Dim chgNum As Integer
  Dim rowI As Long
  Dim colI As Long
  Dim funcCount As Long
  Dim rgbNum As Long

  searchStr = p2Str
  rgbNum = CLng("&H" & p3Str)
  
  chgNum = Len(searchStr)
  funcCount = 0
  For rowI = rowSta To rowEnd
    For colI = colSta To colEnd
      cellStr = ws.Cells(rowI, colI).Value
      findNum = InStr(p1Num, cellStr, searchStr)
      If findNum > 0 Then
        ws.Cells(rowI, colI).Characters(findNum, chgNum).Font.Color = rgbNum
        ws.Cells(rowI, colI).Characters(findNum, chgNum).Font.Bold = True
        funcCount = funcCount + 1
      End If
    Next colI
  Next rowI
  retVal = funcCount
End Sub
