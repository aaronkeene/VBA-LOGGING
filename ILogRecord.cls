VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logging")

Option Explicit



Public Property Get LoggerName() As String
End Property

Public Property Get LogLevel() As enLogLevel
End Property

Public Property Get Message() As String
End Property

Public Property Get TimeStamp() As Date
End Property
