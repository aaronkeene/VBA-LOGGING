VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Logging")
'@PredeclaredId

Option Explicit
Implements ILogger



Private Type TLogger
    Name        As String
    LogLevel    As enLogLevel
    Handlers    As New Dictionary
End Type



Private self As TLogger



Public Function Create(ByVal Name As String, Optional ByVal LogLevel As enLogLevel) As Logger
    Dim result As New Logger
        
    With result
        .Name = Name
        .LogLevel = LogLevel
    End With
    
    Set Create = result
End Function



'==============================
'Properties
'==============================


Friend Property Let Name(ByVal value As String)
    self.Name = value
End Property

Friend Property Let LogLevel(ByVal value As enLogLevel)
    self.LogLevel = value
End Property

Public Property Let ILogger_SetLogLevel(ByVal value As enLogLevel)
    LogLevel = value
End Property



'==============================
'Public Methods
'==============================


Public Sub ILogger_AddHandler(handler As ILogHandler)
    If self.Handlers.Exists(handler.Name) Then
        Exit Sub
    End If
    
    self.Handlers.Add handler.Name, handler
End Sub


Public Sub ILogger_LogCritical(ByVal LogMessage As String)
    If Not ILogger_LogEnabledForLevel(LOG_LEVEL_CRITICAL) Then
        Exit Sub
    End If
    
    Dim Record As New LogRecord
    Set Record = Record.Create(self.Name, LOG_LEVEL_CRITICAL, LogMessage)
    
    ILogger_HandleRecord Record
End Sub


Public Sub ILogger_LogError(ByVal LogMessage As String)
    If Not ILogger_LogEnabledForLevel(LOG_LEVEL_ERROR) Then
        Exit Sub
    End If
    
    Dim Record As New LogRecord
    Set Record = Record.Create(self.Name, LOG_LEVEL_ERROR, LogMessage)
    
    ILogger_HandleRecord Record
End Sub


Public Sub ILogger_LogWarning(ByVal LogMessage As String)
    If Not ILogger_LogEnabledForLevel(LOG_LEVEL_WARNING) Then
        Exit Sub
    End If
    
    Dim Record As New LogRecord
    Set Record = Record.Create(self.Name, LOG_LEVEL_WARNING, LogMessage)
    
    ILogger_HandleRecord Record
End Sub


Public Sub ILogger_LogInfo(ByVal LogMessage As String)
    If Not ILogger_LogEnabledForLevel(LOG_LEVEL_INFO) Then
        Exit Sub
    End If
    
    Dim Record As New LogRecord
    Set Record = Record.Create(self.Name, LOG_LEVEL_INFO, LogMessage)
    
    ILogger_HandleRecord Record
End Sub


Public Sub ILogger_LogDebug(ByVal LogMessage As String)
    If Not ILogger_LogEnabledForLevel(LOG_LEVEL_DEBUG) Then
        Exit Sub
    End If
    
    Dim Record As New LogRecord
    Set Record = Record.Create(self.Name, LOG_LEVEL_DEBUG, LogMessage)
    
    ILogger_HandleRecord Record
End Sub



'==============================
'Private Methods
'==============================

Private Function ILogger_LogEnabledForLevel(ByVal LogLevel As enLogLevel) As Boolean
    If LogLevel >= self.LogLevel Then
        ILogger_LogEnabledForLevel = True
        Exit Function
    End If
End Function


Private Function ILogger_HasHandlers() As Boolean
    If self.Handlers.Count > 0 Then
        ILogger_HasHandlers = True
        Exit Function
    End If
End Function


Private Sub ILogger_HandleRecord(Record As LogRecord)
    If Not ILogger_HasHandlers Then: Exit Sub
    
    Dim key As Variant
    Dim handler As ILogHandler
    
    For Each key In self.Handlers.Keys
        Set handler = self.Handlers(key)
        handler.HandleRecord Record
    Next
End Sub
