VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Logger")
'@PredeclaredId

Option Explicit

Public Enum vbLoggingLevel
    logNOTSET
    logTRACE
    logINFO
    logSUCCESS
    logWARNING
    logERROR
    logCRITICAL
End Enum

Private Type ThisClass
    LogSinks     As Dictionary
    Formatter    As ILogMessageFormatter
End Type

Private Self As ThisClass


'''' Class Intialize / Terminate Methods ''''

Private Sub Class_Initialize()
    Set Self.LogSinks = New Dictionary
    Set Self.Formatter = New DefaultLogMessageFormatter
    
    AddLogSink "Console", ConsoleLogSink.Create(vbLoggingLevel.logNOTSET)
End Sub


Private Sub Class_Terminate()
    Set Self.LogSinks = Nothing
End Sub


''''' Public Methods '''''

Public Sub AddLogSink(ByVal LogName As String, ByVal LogSink As ILogSink)
    If LogSinkExists(LogName) Then
        Logger.Warning "Duplicate LogSink [" & LogName & "], LogSink not added."
        Exit Sub
    End If
    
    Self.LogSinks.Add LogName, LogSink
    
    If LogName = "Console" Then
        Exit Sub
    End If
    
    Logger.Trace "LogSink [" & LogName & "] added to Logger."
End Sub


Public Sub RemoveLogSink(ByVal LogName As String)
    If Not Self.LogSinks.Exists(LogName) Then
        Logger.Warning "Attempting to remove LogSink [" & LogName & "] which does not exist."
        Exit Sub
    End If
    
    Self.LogSinks.Remove LogName
    Logger.Trace "LogSink [" & LogName & "] removed from Logger."
End Sub


Public Sub SetGlobalLogLevel(ByVal LogLevel As vbLoggingLevel)
    Dim key     As Variant
    Dim Sink    As ILogSink
    
    For Each key In Self.LogSinks.Keys
        Set Sink = Self.LogSinks(key)
        Sink.SetLoggingLevel LogLevel
    Next
End Sub


Public Sub SetSinkLogLevel(ByVal LogSinkName As String, ByVal LogLevel As vbLoggingLevel)
    Dim Sink As ILogSink
    
    If Not Self.LogSinks.Exists(LogSinkName) Then
        Exit Sub
    End If
    
    Set Sink = Self.LogSinks(LogSinkName)
    Sink.SetLoggingLevel LogLevel
End Sub


''''' Public Log Methods '''''

Public Sub Trace(ByVal LogMessage As String, Optional ByVal SourceMethodName As String = vbNullString)
    ProcessLog vbLoggingLevel.logTRACE, LogMessage, SourceMethodName
End Sub


Public Sub Info(ByVal LogMessage As String, Optional ByVal SourceMethodName As String = vbNullString)
    ProcessLog vbLoggingLevel.logINFO, LogMessage, SourceMethodName
End Sub


Public Sub Sucess(ByVal LogMessage As String, Optional ByVal SourceMethodName As String = vbNullString)
    ProcessLog vbLoggingLevel.logSUCCESS, LogMessage, SourceMethodName
End Sub


Public Sub Warning(ByVal LogMessage As String, Optional ByVal SourceMethodName As String = vbNullString)
    ProcessLog vbLoggingLevel.logWARNING, LogMessage, SourceMethodName
End Sub


Public Sub Error(ByVal LogMessage As String, Optional ByVal SourceMethodName As String = vbNullString)
    ProcessLog vbLoggingLevel.logERROR, LogMessage, SourceMethodName
End Sub


Public Sub Critical(ByVal LogMessage As String, Optional ByVal SourceMethodName As String = vbNullString)
    ProcessLog vbLoggingLevel.logCRITICAL, LogMessage, SourceMethodName
End Sub


''''' Private Methods '''''

Private Function LogSinkExists(ByVal LogName As String) As Boolean
    Dim key As Variant

    For Each key In Self.LogSinks.Keys
        If key = LogName Then
            LogSinkExists = True
            Exit Function
        End If
    Next
End Function


Private Sub ProcessLog(ByVal LogLevel As vbLoggingLevel, ByVal LogMessage As String, Optional ByVal SourceMethodName As String = vbNullString)
    Dim Log As ILogSink
    Dim key As Variant

    For Each key In Self.LogSinks.Keys
        Set Log = Self.LogSinks(key)
        If Log.LoggingLevel <= LogLevel Then
            Log.Emit FormattedMessage(LogLevel, LogMessage, SourceMethodName)
        End If
    Next
End Sub


Private Function FormattedMessage(ByVal LogLevel As vbLoggingLevel, ByVal LogMessage As String, Optional ByVal SourceMethodName As String = vbNullString)
      FormattedMessage = Self.Formatter.FormatMessage(LogLevel, LogMessage, SourceMethodName)
End Function
