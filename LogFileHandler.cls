VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogFileHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Logger.LogSink")
'@PredeclaredId

Option Explicit
Implements ILogSink



Private Type ThisClass
    LoggingLevel As vbLoggingLevel
    FileWriter   As IFileWriter
End Type

Private Self As ThisClass



Public Function Create(ByVal LogLevel As vbLoggingLevel, FileWriter As IFileWriter) As ILogSink
    Dim Result As New FileLogSink
    
    With Result
        .LoggingLevel = LogLevel
        Set .FileWriter = FileWriter
    End With
    
    Set Create = Result
End Function


'''' Class Set Methods ''''

Friend Property Let LoggingLevel(ByVal value As vbLoggingLevel)
    Self.LoggingLevel = value
End Property

Friend Property Set FileWriter(value As IFileWriter)
    Set Self.FileWriter = value
End Property


'''' Interface Methods ''''

Private Property Get ILogSink_LoggingLevel() As vbLoggingLevel
    ILogSink_LoggingLevel = Self.LoggingLevel
End Property

Private Sub ILogSink_Emit(ByVal FormattedMessage As String)
    Self.FileWriter.Append FormattedMessage
End Sub

Private Sub ILogSink_SetLoggingLevel(ByVal LoggingLevel As vbLoggingLevel)
    Self.LoggingLevel = LoggingLevel
End Sub
