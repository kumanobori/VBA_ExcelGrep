VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DESTINATION_DEBUG_PRINT As Long
Public DESTINATION_RETURN As Long
Public DESTINATION_CELL As Long

Public LEVEL_DEBUG As Long
Public LEVEL_INFO As Long
Public LEVEL_ERROR As Long

Public destination As Long
Public level As Long
Public wrOut As Range

Public Sub class_initialize()
    DESTINATION_DEBUG_PRINT = 1
    DESTINATION_RETURN = 2
    LEVEL_DEBUG = 1
    LEVEL_INFO = 2
    LEVEL_ERROR = 3
End Sub

Public Function init(aDestination As Long, aLevel As Long)
    destination = aDestination
    level = aLevel
End Function
Public Function initCell(wr As Range)
    Set wrOut = wr
    Dim ws As Worksheet: Set ws = wrOut.Parent
    ws.Range(wrOut, ws.Cells(ws.Cells.SpecialCells(xlCellTypeLastCell).row, wrOut.Column)).ClearContents
End Function
Private Function log(aStr As String, aLevel As String)
    
    Dim s As String: s = Format(Now, "yyyy-mm-dd HH:MM:SS ") & aLevel & aStr
    Select Case destination
        Case DESTINATION_DEBUG_PRINT
            Debug.Print s
        Case DESTINATION_RETURN
            log = s
        Case DESTINATION_CELL
            wrOut.value = s
            Set wrOut = wrOut.Offset(1, 0)
    End Select
End Function


Public Function logDebug(s As String)
    If level <= LEVEL_DEBUG Then
        logDebug = log(s, "[DEBUG] ")
    End If
End Function


Public Function logInfo(s As String)
    If level <= LEVEL_INFO Then
        logInfo = log(s, "[INFO]  ")
    End If
End Function

Public Function logError(s As String)
    If level <= LEVEL_ERROR Then
        logError = log(s, "[ERROR] ")
    End If
End Function

