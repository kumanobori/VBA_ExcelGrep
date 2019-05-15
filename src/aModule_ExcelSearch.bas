Attribute VB_Name = "aModule_ExcelSearch"
Option Explicit

Const CONF_START = "D2"
Const CONF_SEARCH_CONDITION_START = "F14" ' 「No」のセル
Const CONF_LOG_START = "C28"

Sub ExcelSearch()
    Application.ScreenUpdating = False

    ' 設定ブックとシートを定義
    Dim wbConf As Workbook, wsConf As Worksheet
    Set wbConf = ActiveWorkbook
    Set wsConf = ActiveSheet
    
    ' 各設定の先頭セルを取得
    Dim wrConfStart As Range: Set wrConfStart = wsConf.Range(CONF_START)
    Dim wrSearchConditionStart As Range: Set wrSearchConditionStart = wsConf.Range(CONF_SEARCH_CONDITION_START)
    
    ' 各設定項目を取得
    Dim aryConf As Variant: aryConf = wsConf.Range(wrConfStart, wrConfStart.End(xlDown))
    
    ' 検索条件を取得
    Dim conditionFinalColumn As Long: conditionFinalColumn = wrSearchConditionStart.End(xlToRight).Column
    Dim conditionFinalRow As Long: conditionFinalRow = wrSearchConditionStart.End(xlDown).row
    Dim aryCondition As Variant: aryCondition = wsConf.Range(wrSearchConditionStart.Offset(1, 0), wsConf.Cells(conditionFinalRow, conditionFinalColumn))
    
    ' ロガー
    Dim log As Logger: Set log = New Logger
    Call log.init(log.DESTINATION_CELL, log.LEVEL_DEBUG)
    Call log.initCell(wsConf.Range(CONF_LOG_START))
    
    Dim obj As New aExcelSearchController
    Call obj.init(log, aryConf, aryCondition, wbConf)
    Call obj.exec
    MsgBox "done"
End Sub

