VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelInspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private log As Logger
Private listCondition As Collection
Private wrDummySearch  As Range ' �}�`��������Z���ɓ���Č������Ă݂邽�߂̃_�~�[�Z��
Private regex As Object

Public Function init(aLog As Logger, aListCondition As Collection, aWrDummySearch As Range)
    Set log = aLog
    Set listCondition = aListCondition
    Set wrDummySearch = aWrDummySearch
    Set regex = CreateObject("VBScript.RegExp")
End Function


Public Function exec(ByRef listResult As Collection, filepath As String)
    Dim logPrefix As String: logPrefix = "ExcelInspect.exec: "
    log.logInfo logPrefix & filepath & " start."
    
    Dim wb As Workbook
    Workbooks.Open fileName:=filepath, ReadOnly:=True
    Set wb = ActiveWorkbook
    
    ' �V�[�g���ƂɌ������{
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim listResultTmp As Collection: Set listResultTmp = New Collection
        Call inspectSheet(listResultTmp, ws)
        ' �������ʂ�ǉ�
        Dim i As Long
        For i = 1 To listResultTmp.Count
            Call listResult.Add(listResultTmp(i))
        Next i
    Next ws
    
    
    wb.Close saveChanges:=False
    
    log.logInfo logPrefix & filepath & " end."
End Function


Private Function inspectSheet(ByRef listResult As Collection, ws As Worksheet)
    Dim logPrefix As String: logPrefix = "ExcelInspect.inspectSheet: "
    log.logDebug logPrefix & "start. sheetName=" & ws.Name
    
    Dim i As Long
    ' �����������ƂɃ��[�v
    For i = 1 To listCondition.Count
        Dim condition As confSearchCondition: Set condition = listCondition(i)
        log.logDebug logPrefix & "ExcelInspect.inspectSheet condition " & condition.no
        
        ' ���K�\���̃p�^�[�����Z�b�g����
        regex.Pattern = condition.searchString
        
        ' �l������
        If condition.isSearchValue Then
            Call inspectByFind(ws, listResult, condition, xlValues)
        End If
        ' ����������
        If condition.isSearchFormula Then
            Call inspectByFind(ws, listResult, condition, xlFormulas)
        End If
        ' �R�����g������
        If condition.isSearchComment Then
            Call inspectByFind(ws, listResult, condition, xlComments)
        End If
        ' �}�`������
        If condition.isSearchShape Then
            Call inspectShape(ws, listResult, condition)
        End If
        
    Next i
    
    log.logDebug logPrefix & "end."
End Function

Private Function inspectByFind(ws As Worksheet, ByRef listResult As Collection, condition As confSearchCondition, optLookIn As Variant)
    Dim logPrefix As String: logPrefix = "ExcelInspect.inspectByFind: "
    log.logDebug logPrefix & "start. sheetName=" & ws.Name & ", optLookIn=" & optLookIn
    
    Dim wr As Range, wrFirst As Range
    
    ' ���K�\���̏ꍇ�͌����������*�Ƃ���B�i�S�Z�����q�b�g�����������Ő��K�\���}�b�`�������j
    Dim optWhat As String: optWhat = IIf(condition.isRegex, "*", condition.searchString)
    Dim optLookAt As String: optLookAt = IIf(condition.isLookAt, xlWhole, xlPart)
    ' ���񌟍�
    Set wr = ws.Cells.Find(what:=optWhat, LookIn:=optLookIn, lookat:=optLookAt, MatchCase:=condition.isMatchCase, matchbyte:=condition.isMatchByte)
    If Not wr Is Nothing Then
        Set wrFirst = wr
        Do
            ' �q�b�g�����Z�����Ώۂł���Ό��ʂɒǉ�
            If isCellTarget(wr, condition, optLookIn) Then
                Call resultFromCell(listResult, wr, optLookIn)
            End If
            ' 2��ڈȍ~�̌���
            Set wr = ws.Cells.FindNext(wr)
        Loop While wr.address <> wrFirst.address
    End If
    
    log.logDebug logPrefix & "end."
End Function
Private Function isCellTarget(wr As Range, condition As confSearchCondition, optLookIn As Variant)
    
    ' ���K�\���̏ꍇ�̓q�b�g�����Z�������K�\���}�b�`���Ȃ��ꍇfalse
    If condition.isRegex Then
        If Not isCellRegexMatched(wr, optLookIn) Then
            isCellTarget = False
            Exit Function
        End If
    End If
    
    isCellTarget = True
End Function
Private Function isCellRegexMatched(wr As Range, optLookIn As Variant)
    Dim strTarget As String: strTarget = fetchCellString(wr, optLookIn)
    isCellRegexMatched = isStringRegexMatched(strTarget)
End Function

Private Function fetchCellString(wr As Range, optLookIn As Variant)
    Dim str As String
    Select Case optLookIn
        Case xlValues
            str = wr.value
        Case xlFormulas
            str = wr.Formula
        Case xlComments
            str = wr.Comment.Text
    End Select
    fetchCellString = str
End Function
Private Function isStringRegexMatched(str As String)
    Dim matches As Variant: Set matches = regex.Execute(str)
    isStringRegexMatched = (matches.Count > 0)
End Function

Private Function resultFromCell(listResult As Collection, wr As Range, optLookIn As Variant)
    Dim logPrefix As String: logPrefix = "ExcelInspect.resultFromCell: "
    log.logDebug logPrefix & "start. address=" & wr.address & ", optLookIn=" & optLookIn
    
    ' �q�b�g�A�C�e���I�u�W�F�N�g���쐬
    Dim item As New HitItem
    Dim itemValue As String: itemValue = fetchCellString(wr, optLookIn)
    Dim itemType As Long
    Select Case optLookIn
        Case xlValues
            itemType = item.ITEMTYPE_CELL_VALUE
        Case xlFormulas
            itemType = item.ITEMTYPE_CELL_FORMULA
        Case xlComments
            itemType = item.ITEMTYPE_CELL_COMMENT
    End Select
    Call item.initCell(wr.Parent.Parent.path, wr.Parent.Parent.Name, wr.Parent.Name, itemType, wr.address, itemValue)
    
    ' �����q�b�g�A�C�e�������X�g�ɂȂ���Βǉ�
    Call addIfNotExist(listResult, item)
    
    log.logDebug "ExcelInspect.resultFromCell end."
End Function
Private Function addIfNotExist(ByRef listResult As Collection, item As HitItem)
    Dim logPrefix As String: logPrefix = "ExcelInspect.addIfNotExist: "
    log.logDebug logPrefix & "start. item=" & Join(item.toArray, ",")
    Dim i As Long, found As Boolean: found = False
    For i = 1 To listResult.Count
        If item.equalsAsResult(listResult(i)) Then
            found = True
        End If
    Next i
    If Not found Then
        log.logDebug logPrefix & "added."
        Call listResult.Add(item)
    Else
        log.logDebug logPrefix & "not added."
    End If

    log.logDebug logPrefix & "end."
End Function

Private Function inspectShape(ws As Worksheet, listResult As Collection, searchCondition As confSearchCondition)
    log.logDebug "ExcelInspect.inspectShape start. sheetName=" & ws.Name
    
    Dim obj As Shape
    Dim i As Long
    For i = 1 To ws.Shapes.Count
        Set obj = ws.Shapes(i)
        If isShapeTarget(obj, searchCondition) Then
            Call resultFromShape(listResult, obj)
        End If
    Next i
    log.logDebug "ExcelInspect.inspectShape end."
End Function
Private Function isShapeTarget(obj As Shape, searchCondition As confSearchCondition)
    log.logDebug "ExcelInspect.inShapeTarget start. id=" & obj.ID & ", type=" & obj.Type
    
    If obj.Type = msoComment Then
        isShapeTarget = False
        log.logDebug "ExcelInspect.inShapeTarget end. msoComment false"
        Exit Function
    End If
    
    If obj.TextFrame2.HasText Then
        If isShapeTextMatch(obj.TextFrame2.TextRange.Text, searchCondition) Then
            isShapeTarget = True
            log.logDebug "ExcelInspect.inShapeTarget end. true"
            Exit Function
        End If
    End If
    isShapeTarget = False
    log.logDebug "ExcelInspect.inShapeTarget end. false"
End Function
Private Function isShapeTextMatch(shapeText As String, searchCondition As confSearchCondition)
    If searchCondition.isRegex Then
        isShapeTextMatch = isStringRegexMatched(shapeText)
    Else
        wrDummySearch.value = shapeText
        Dim optLookAt As String: optLookAt = IIf(searchCondition.isLookAt, xlWhole, xlPart)
        Dim result As Range
        Set result = wrDummySearch.Find(what:=searchCondition.searchString, LookIn:=xlValues, lookat:=optLookAt, MatchCase:=searchCondition.isMatchCase, matchbyte:=searchCondition.isMatchByte)
        wrDummySearch.value = ""
        isShapeTextMatch = (Not result Is Nothing)
    End If
End Function
Private Function resultFromShape(listResult As Collection, obj As Shape)
    log.logDebug "ExcelInspect.resultFromShape start."
    
    ' �q�b�g�A�C�e���I�u�W�F�N�g���쐬
    Dim item As New HitItem
    Call item.initShape(obj.Parent.Parent.path, obj.Parent.Parent.Name, obj.Parent.Name, item.ITEMTYPE_SHAPE_TEXT, _
                        obj.ID, obj.TopLeftCell.address, obj.TextFrame2.TextRange.Text)
    
    ' �����q�b�g�A�C�e�������X�g�ɂȂ���Βǉ�
    Call addIfNotExist(listResult, item)
    
    log.logDebug "ExcelInspect.resultFromShape end."
End Function

