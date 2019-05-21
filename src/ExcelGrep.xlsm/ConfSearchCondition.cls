VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfSearchCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 検索条件として外部から参照するのでpublic
Public no As Long
Public isValid As Boolean
Public isSearchValue As Boolean
Public isSearchFormula As Boolean
Public isSearchComment As Boolean
Public isSearchShape As Boolean
Public isMatchCase As Boolean
Public isLookAt As Boolean
Public isMatchByte As Boolean
Public isRegex As Boolean
Public searchString As String

Public Function init(options As Variant, idx1 As Long)
    Dim i As Long: i = 0
    i = i + 1: no = options(idx1, i)
    i = i + 1: isValid = (options(idx1, i) = "Y")
    i = i + 1: isSearchValue = (options(idx1, i) = "Y")
    i = i + 1: isSearchFormula = (options(idx1, i) = "Y")
    i = i + 1: isSearchComment = (options(idx1, i) = "Y")
    i = i + 1: isSearchShape = (options(idx1, i) = "Y")
    i = i + 1: isMatchCase = (options(idx1, i) = "Y")
    i = i + 1: isLookAt = (options(idx1, i) = "Y")
    i = i + 1: isMatchByte = (options(idx1, i) = "Y")
    i = i + 1: isRegex = (options(idx1, i) = "Y")
    i = i + 1: searchString = options(idx1, i)
    
    Call validate
End Function

Private Function validate()
    If isValid Then
        If searchString = "" Then
            MsgBox "検索文字列が空欄です。検索は行えません。" & toString
        End If
    End If
End Function

Private Function toString()
    toString = "no=[" & no _
             & "], 有効=[" & isValid _
             & "], 値=[" & isSearchFormula _
             & "], 数式=[" & isSearchFormula _
             & "], コメント=[" & isSearchFormula _
             & "], 図形=[" & isSearchFormula _
             & "], 大小区別=[" & isMatchCase _
             & "], 完全一致=[" & isLookAt _
             & "], 全半区別=[" & isMatchByte _
             & "], 正規表現=[" & isRegex _
             & "], 検索文字列=[" & searchString _
             & "]"
End Function

