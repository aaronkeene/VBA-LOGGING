VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logging")

Option Explicit



Public Property Get Name() As String:                                           End Property
Public Property Let SetLogLevel(ByVal value As enLogLevel):                     End Property
Public Property Set SetFormatter(ByVal value As ILogFormatter):                 End Property



Public Sub HandleRecord(Record As LogRecord):                                   End Sub



'@Ignore ProcedureNotUsed, EmptyMethod, ParameterNotUsed
Private Sub EmitRecord(ByVal Record As LogRecord):                              End Sub

'@Ignore ProcedureNotUsed, EmptyMethod, ParameterNotUsed
Private Function FormatRecord(ByVal Record As LogRecord) As String:             End Function

'@Ignore ParameterNotUsed, ProcedureNotUsed, EmptyMethod, NonReturningFunction
Private Function HasFormatter() As Boolean:                                     End Function

