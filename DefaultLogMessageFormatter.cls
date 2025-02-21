VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultLogFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logger.LogFormatter")

Option Explicit
Implements ILogMessageFormatter


'''' Class Interface Methods ''''

Private Function ILogMessageFormatter_FormatMessage(ByVal LogLevel As vbLoggingLevel, ByVal LogMessage As String, Optional ByVal SourceMethodName As String = vbNullString) As String
    ILogMessageFormatter_FormatMessage = PrintString(LogLevel, LogMessage, SourceMethodName)
End Function


'''' Class Private Methods ''''

Private Function LogLevelString(ByVal Level As vbLoggingLevel) As String
    Select Case Level
        Case vbLoggingLevel.logTRACE:       LogLevelString = "TRACE    "
        Case vbLoggingLevel.logINFO:        LogLevelString = "INFO     "
        Case vbLoggingLevel.logSUCCESS:     LogLevelString = "SUCCESS  "
        Case vbLoggingLevel.logWARNING:     LogLevelString = "WARNING  "
        Case vbLoggingLevel.logERROR:       LogLevelString = "ERROR    "
        Case vbLoggingLevel.logCRITICAL:    LogLevelString = "CRITICAL "
    End Select
End Function


Private Function PrintString(ByVal LogLevel As vbLoggingLevel, ByVal LogMessage As String, Optional ByVal SourceMethodName As String = vbNullString) As String
    If SourceMethodName = vbNullString Then
        PrintString = Format(Now(), "yyyy-mm-dd HH:mm:ss") & " | " & LogLevelString(LogLevel) & "| Method:=Unspecified | " & LogMessage
        Exit Function
    End If
    
    PrintString = Format(Now(), "yyyy-mm-dd HH:mm:ss") & " | " & LogLevelString(LogLevel) & "| Method:=" & SourceMethodName & " | " & LogMessage
End Function
