VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logging")

Option Explicit



Public Property Let SetLogLevel(ByVal value As enLogLevel):                     End Property



Public Sub AddHandler(handler As ILogHandler):                                  End Sub
'Public Sub RemoveHandler(handler As ILogHandler):                 End Sub
'Public Sub AddFilter(Filter as ILogFilter):                       End Sub
'Public Sub RemoveFilter(Filter as ILogFilter:                     End Sub



Public Sub LogCritical(ByVal LogMessage As String):                             End Sub
Public Sub LogError(ByVal LogMessage As String):                                End Sub
Public Sub LogWarning(ByVal LogMessage As String):                              End Sub
Public Sub LogInfo(ByVal LogMessage As String):                                 End Sub
Public Sub LogDebug(ByVal LogMessage As String):                                End Sub


'@Ignore ParameterNotUsed, ProcedureNotUsed, EmptyMethod, NonReturningFunction
Private Function LogEnabledForLevel(ByVal LogLevel As enLogLevel) As Boolean:   End Function

'@Ignore ParameterNotUsed, ProcedureNotUsed, EmptyMethod, NonReturningFunction
Private Function HasHandlers() As Boolean:                                      End Function

'@Ignore ProcedureNotUsed, EmptyMethod, ParameterNotUsed
Private Sub HandleRecord(Record As LogRecord):                                  End Sub
