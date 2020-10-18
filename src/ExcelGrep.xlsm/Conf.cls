VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Conf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 外部から定数的に使いたいので、publicにしてコンストラクタ内で初期化してある
Public OUTPUT_DEST_TO_INPUT_RESULT As Long
Public OUTPUT_DEST_TO_INPUT_NEW As Long
Public OUTPUT_DEST_TO_NEW_BOOK As Long

' 外部から参照するのでpublic
Public outputDest As Long
Public isRelativePath As Boolean
Public fontName As String
Public fontSize As Double

Public rootFolder As String
Public isSearchSubfolder As Boolean
Public targetFileName As String
Public isFileNameRegex As Boolean
Public isOnlyFileSearch As Boolean

Public Sub class_initialize()
    OUTPUT_DEST_TO_INPUT_RESULT = 1
    OUTPUT_DEST_TO_INPUT_NEW = 2
    OUTPUT_DEST_TO_NEW_BOOK = 3
End Sub

Public Function init(options As Variant)
    Dim i As Long: i = 0
    
    i = i + 1: rootFolder = options(i, 1)
    i = i + 1: isSearchSubfolder = (options(i, 1) = "Y")
    i = i + 1: targetFileName = options(i, 1)
    i = i + 1: isFileNameRegex = (options(i, 1) = "Y")
    i = i + 1: isOnlyFileSearch = (options(i, 1) = "Y")
    
    i = i + 1: outputDest = Left(options(i, 1), 1) ' 選択肢の1文字目(数字)
    i = i + 1: isRelativePath = (options(i, 1) = "Y")
    i = i + 1: fontName = options(i, 1)
    i = i + 1: fontSize = options(i, 1)

End Function
