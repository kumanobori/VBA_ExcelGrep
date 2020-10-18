VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HitItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ITEMTYPE_CELL_VALUE As Long
Public ITEMTYPE_CELL_FORMULA As Long
Public ITEMTYPE_CELL_COMMENT As Long
Public ITEMTYPE_SHAPE_TEXT As Long

' 外部から設定・参照するのでpublic
Public path As String
Public bookName As String
Public sheetName As String
Public itemType As Long
Public itemAddress As String
Public shapeId As Long
Public itemValue As String


Public Sub class_initialize()
    ITEMTYPE_CELL_VALUE = 1
    ITEMTYPE_CELL_FORMULA = 2
    ITEMTYPE_CELL_COMMENT = 3
    ITEMTYPE_SHAPE_TEXT = 4
End Sub

Public Function initCell(aPath As String, aBookName As String, aSheetName As String, aItemType As Long, aItemAddress As String, aItemValue As String)
    path = aPath
    bookName = aBookName
    sheetName = aSheetName
    itemType = aItemType
    shapeId = 0
    itemAddress = aItemAddress
    itemValue = aItemValue
End Function

Public Function initShape(aPath As String, aBookName As String, aSheetName As String, aItemType As Long, aShapeId As Long, aItemAddress As String, aItemValue As String)
    path = aPath
    bookName = aBookName
    sheetName = aSheetName
    itemType = aItemType
    shapeId = aShapeId
    itemAddress = aItemAddress
    itemValue = aItemValue
End Function

' 同一シート内のアイテム同士である前提で、同じものかどうか
Public Function equalsAsResult(c As HitItem)
    
    ' itemAddressとshapeIdは一致しなければならない
    If itemAddress <> c.itemAddress _
        Or shapeId <> c.shapeId Then
        equalsAsResult = False
        Exit Function
    End If
    
    ' 上記をパスしたうえでitemTypeが一致すれば同じとみなす
    If itemType = c.itemType Then
        equalsAsResult = True
        Exit Function
    Else
        ' 一致しない場合でも、組み合わせが「値」と「数式」の場合は同じとみなす
        If (itemType = ITEMTYPE_CELL_VALUE And c.itemType = ITEMTYPE_CELL_FORMULA) _
            Or (itemType = ITEMTYPE_CELL_FORMULA And c.itemType = ITEMTYPE_CELL_VALUE) Then
            equalsAsResult = True
            Exit Function
        Else
            equalsAsResult = False
        End If
    End If
End Function

Public Function toArray()
    toArray = Array(path, bookName, sheetName, itemType, itemAddress, shapeId, itemValue)
End Function
