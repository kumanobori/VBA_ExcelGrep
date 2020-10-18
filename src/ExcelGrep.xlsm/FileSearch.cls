VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private confIsFileNameRegex As Boolean
Private confIsSearchSubfolder As Boolean
Private confRootFolder As String
Private confTargetFileName As String

Private log As Logger
Private fso As Object
Private regex As Object

Public Function init(aLog As Logger, Conf As Conf)
    
    ' 引数項目
    Set log = aLog
    confIsFileNameRegex = Conf.isFileNameRegex
    confIsSearchSubfolder = Conf.isSearchSubfolder
    confRootFolder = Conf.rootFolder
    confTargetFileName = Conf.targetFileName
    
    ' 内部で共通的に使う変数の初期化
    Set fso = CreateObject("Scripting.FileSystemObject")
    If confIsFileNameRegex Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = confTargetFileName
    End If
    
    log.logDebug "FileSearch.init  root=" & confRootFolder & ", target=" & confTargetFileName & ", regex=" & confIsFileNameRegex & ", searchSubFolder=" & confIsSearchSubfolder
End Function

Public Function search()
    Dim listResult As New Collection
    Call searchFolder(listResult, fso.getFolder(confRootFolder))
    Set search = listResult
End Function
Private Function searchFolder(ByRef listResult As Collection, folder As Object)
    log.logDebug "FileSearch.searchFolder " & folder.path & " start."
    
    ' フォルダ内のファイルを処理
    Dim listFile As Object: Set listFile = folder.Files
    Dim eachFile As Object
    For Each eachFile In listFile
        If isTargetFile(eachFile) Then
            listResult.Add eachFile.path
        End If
    Next eachFile
    
    ' サブディレクトリを処理
    If confIsSearchSubfolder Then
        Dim listFolder As Object: Set listFolder = folder.SubFolders
        Dim eachFolder As Object
        For Each eachFolder In listFolder
            Call searchFolder(listResult, eachFolder)
        Next eachFolder
    End If
    
    log.logDebug "FileSearch.searchFolder " & folder.path & " end."
End Function

Private Function isTargetFile(objFile As Object)
    log.logDebug "FileSearch.isTargetFile " & objFile.path & " start."
    Dim fileName As String: fileName = objFile.Name
    If confIsFileNameRegex Then
        Dim matches As Variant
        Set matches = regex.Execute(fileName)
        isTargetFile = (matches.Count > 0)
    Else
        isTargetFile = fileName Like confTargetFileName
    End If
    log.logDebug "FileSearch.isTargetFile " & objFile.path & " end with " & isTargetFile
End Function
