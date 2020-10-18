VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private log As Logger
Private objConfOutput As Conf
Private resultList As Collection
Private wbIn As Workbook

Public Function init(aLog As Logger, aObjConfOutput As Conf, aResultList As Collection, aWbIn As Workbook)
    Set log = aLog
    Set objConfOutput = aObjConfOutput
    Set resultList = aResultList
    Set wbIn = aWbIn
End Function

Public Function exec()
    log.logInfo "ResultOutput.exec start"
    
    Dim wsOut As Worksheet
    Set wsOut = initOutputSheet()
    
    ' 検索結果
    Dim rowOut As Long: rowOut = 1
    Dim i As Long, item As HitItem
    For i = 1 To resultList.Count
        Set item = resultList(i)
        rowOut = rowOut + 1
        Dim colOut As Long: colOut = 0
        colOut = colOut + 1: wsOut.Cells(rowOut, colOut).value = rowOut - 1
        Dim wkPath As String
        wkPath = IIf(objConfOutput.isRelativePath, Replace(item.path, objConfOutput.rootFolder, ""), item.path)
        wkPath = IIf(Left(wkPath, 1) = "\", Mid(wkPath, 2), wkPath)
        colOut = colOut + 1: wsOut.Cells(rowOut, colOut).value = wkPath
        colOut = colOut + 1: wsOut.Cells(rowOut, colOut).value = item.bookName
        colOut = colOut + 1: wsOut.Cells(rowOut, colOut).value = item.sheetName
        Dim wkType As String
        Select Case item.itemType
            Case item.ITEMTYPE_CELL_VALUE
                wkType = "値"
            Case item.ITEMTYPE_CELL_FORMULA
                wkType = "数式"
            Case item.ITEMTYPE_CELL_COMMENT
                wkType = "コメント"
            Case item.ITEMTYPE_SHAPE_TEXT
                wkType = "図形"
            Case Else
                MsgBox "unexpected itemType:" & item.itemType
                End
        End Select
        colOut = colOut + 1: wsOut.Cells(rowOut, colOut).value = wkType
        Dim wkAddress As String: wkAddress = Replace(item.itemAddress, "$", "")
        colOut = colOut + 1: Call wsOut.Hyperlinks.Add(anchor:=wsOut.Cells(rowOut, colOut), _
                                                        address:=item.path & "\" & item.bookName, _
                                                        SubAddress:="'" & item.sheetName & "'" & "!" & wkAddress, _
                                                        TextToDisplay:=wkAddress)
        Dim wkValue As String: wkValue = item.itemValue
        If Left(wkValue, 1) = "=" Then wkValue = "'" & wkValue
        colOut = colOut + 1: wsOut.Cells(rowOut, colOut).value = wkValue
    Next i
    Call formatOutputSheet(wsOut)
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells.SpecialCells(xlCellTypeLastCell)).Borders.LineStyle = xlContinuous
    log.logInfo "ResultOutput.exec end"
End Function

Private Function initOutputSheet()
    Dim ws As Worksheet
    Select Case objConfOutput.outputDest
    Case objConfOutput.OUTPUT_DEST_TO_INPUT_RESULT
        Set ws = wbIn.Worksheets("RESULT")
        ws.Cells.Clear
    Case objConfOutput.OUTPUT_DEST_TO_INPUT_NEW
        wbIn.Activate
        wbIn.Worksheets.Add
        Set ws = ActiveSheet
    Case objConfOutput.OUTPUT_DEST_TO_NEW_BOOK
        Workbooks.Add
        Set ws = ActiveSheet
    Case Else
    End Select
    Set initOutputSheet = ws
End Function

Private Function formatOutputSheet(wsOut As Worksheet)
    wsOut.Cells.Font.Name = objConfOutput.fontName
    wsOut.Cells.Font.Size = objConfOutput.fontSize

    ' ヘッダ
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, 7)) = Array("No", "path", "filename", "sheetname", "type", "address", "string")

    Dim row As Long: row = 1
    Dim col As Long: col = 0
    col = col + 1: wsOut.Cells(row, col).EntireColumn.ColumnWidth = 3
    col = col + 1: wsOut.Cells(row, col).EntireColumn.ColumnWidth = 20
    col = col + 1: wsOut.Cells(row, col).EntireColumn.ColumnWidth = 15
    col = col + 1: wsOut.Cells(row, col).EntireColumn.ColumnWidth = 8
    col = col + 1: wsOut.Cells(row, col).EntireColumn.ColumnWidth = 6
    col = col + 1: wsOut.Cells(row, col).EntireColumn.ColumnWidth = 5
    col = col + 1: wsOut.Cells(row, col).EntireColumn.ColumnWidth = 20

    wsOut.Activate
    wsOut.Cells(2, 2).Select
    ActiveWindow.FreezePanes = True

End Function




