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

' コントローラー初期化
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

' コントローラー実行
Public Function exec()
    Dim prefix As String: prefix = "ExcelSearchController.exec: "
    log.logInfo prefix & "start"
    
    ' 対象ファイルリストを取得
    Dim fileList As New Collection
    Set fileList = generateFileList()
    If objConf.isOnlyFileSearch Then
        log.logInfo prefix & "isOnlyFileSearch is valid. term."
        Exit Function
    End If
    
    ' 検索実行
    Dim resultList As New Collection
    Set resultList = inspect(fileList)
    
    ' 結果出力
    Call outputResult(resultList)
    
    log.logInfo prefix & "end"
End Function

' 対象ファイルのリストを作成
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

' 検索
Private Function inspect(fileList As Collection)
    Dim prefix As String: prefix = "ExcelSearchController.search: "
    log.logInfo prefix & "start"
    
    ' 図形検索用ダミーブック
    Workbooks.Add
    Dim wbDummy As Workbook: Set wbDummy = ActiveWorkbook
    Dim wrDummy As Range: Set wrDummy = ActiveSheet.Cells(1, 1)
    
    Dim objExcelInspect As New ExcelInspect
    Call objExcelInspect.init(log, objSearchConditions.fetchValidConditions, wrDummy)
    Dim eachFilePath As String, listResult As New Collection
    ' 1ファイルごと
    Dim i As Long
    For i = 1 To fileList.Count
        Call objExcelInspect.exec(listResult, fileList(i))
    Next i
    Set inspect = listResult
    
    ' 検索ダミー用ブック閉じる
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

