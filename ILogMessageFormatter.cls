VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logger.LogSink")
Option Explicit


Public Property Get LoggingLevel() As vbLoggingLevel
End Property

Public Sub Emit(ByVal FormattedMessage As String)
End Sub

Public Sub SetLoggingLevel(ByVal LoggingLevel As vbLoggingLevel)
End Sub
