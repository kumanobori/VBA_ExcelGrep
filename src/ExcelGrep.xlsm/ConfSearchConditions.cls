VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfSearchConditions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private searchConditions As Collection

Public Function init(options As Variant)
    Set searchConditions = New Collection
    Dim i As Long
    For i = LBound(options) To UBound(options)
        Dim searchCondition As confSearchCondition
        Set searchCondition = New confSearchCondition
        Call searchCondition.init(options, i)
        Call searchConditions.Add(searchCondition)
    Next i
End Function

' 検索条件のうち、有効設定されているものだけをコレクションで返す
Public Function fetchValidConditions()
    Dim conditions As New Collection
    Dim condition As confSearchCondition
    Dim i As Long
    For i = 1 To searchConditions.Count
        If searchConditions(i).isValid Then
            conditions.Add (searchConditions(i))
        End If
    Next i
    Set fetchValidConditions = conditions
End Function
