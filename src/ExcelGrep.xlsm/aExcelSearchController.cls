VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aExcelSearchController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private log As Logger
Private objConf As Conf
Private objSearchConditions As ConfSearchConditions
Private wbIn As Workbook

' �R���g���[���[������
Public Function init(aLog As Logger, aConf As Variant, aSearchCondition As Variant, aWbIn As Workbook)
    
    Set log = aLog
    log.logDebug "init start"
    Set objConf = New Conf
    Set objSearchConditions = New ConfSearchConditions
    Set wbIn = aWbIn
    
    Call objConf.init(aConf)
    Call objSearchConditions.init(aSearchCondition)
    log.logInfo "ExcelSearchController.init end"
End Function

' �R���g���[���[���s
Public Function exec()
    Dim prefix As String: prefix = "ExcelSearchController.exec: "
    log.logInfo prefix & "start"
    
    ' �Ώۃt�@�C�����X�g���擾
    Dim fileList As New Collection
    Set fileList = generateFileList()
    If objConf.isOnlyFileSearch Then
        log.logInfo prefix & "isOnlyFileSearch is valid. term."
        Exit Function
    End If
    
    ' �������s
    Dim resultList As New Collection
    Set resultList = inspect(fileList)
    
    ' ���ʏo��
    Call outputResult(resultList)
    
    log.logInfo prefix & "end"
End Function

' �Ώۃt�@�C���̃��X�g���쐬
Private Function generateFileList()
    Dim prefix As String: prefix = "ExcelSearchController.generateFileList: "
    log.logInfo prefix & "start"
    
    Dim objFileSearch As New FileSearch
    Call objFileSearch.init(log, objConf)
    Dim listResult As Collection
    Set listResult = objFileSearch.search()
    
    Dim i As Long
    For i = 1 To listResult.Count
        log.logInfo prefix & "target:" & listResult(i)
    Next i
    Set generateFileList = listResult
    log.logInfo prefix & "end"
End Function

' ����
Private Function inspect(fileList As Collection)
    Dim prefix As String: prefix = "ExcelSearchController.search: "
    log.logInfo prefix & "start"
    
    ' �}�`�����p�_�~�[�u�b�N
    Workbooks.Add
    Dim wbDummy As Workbook: Set wbDummy = ActiveWorkbook
    Dim wrDummy As Range: Set wrDummy = ActiveSheet.Cells(1, 1)
    
    Dim objExcelInspect As New ExcelInspect
    Call objExcelInspect.init(log, objSearchConditions.fetchValidConditions, wrDummy)
    Dim eachFilePath As String, listResult As New Collection
    ' 1�t�@�C������
    Dim i As Long
    For i = 1 To fileList.Count
        Call objExcelInspect.exec(listResult, fileList(i))
    Next i
    Set inspect = listResult
    
    ' �����_�~�[�p�u�b�N����
    wbDummy.Close saveChanges:=False
    
    log.logInfo prefix & "end"
End Function



Private Function outputResult(resultList As Collection)
    log.logInfo "ExcelSearchController.outputResult start"
    
    Dim resultOutput As New resultOutput
    Call resultOutput.init(log, objConf, resultList, wbIn)
    Call resultOutput.exec
    
    log.logInfo "ExcelSearchController.outputResult end"
End Function
