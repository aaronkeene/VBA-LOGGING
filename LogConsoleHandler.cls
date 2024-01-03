VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogConsoleHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Logging")
'@PredeclaredId

Option Explicit
Implements ILogHandler



Private Type TLogHandlerConsole
    Name As String
    LogLevel As enLogLevel
    Formatter As ILogFormatter
End Type



Private self As TLogHandlerConsole



Public Function Create(Optional Name As String, Optional ByVal LogLevel As enLogLevel) As ILogHandler
    Const LOGHANDLER_NAME As String = "console_handler"
        
    'check if name pasesd in, otherside use default name;
    If Name = vbNullString Then
        Name = LOGHANDLER_NAME
    End If
    
    Dim result As New LogConsoleHandler
    
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


Friend Property Set Formatter(ByVal Formatter As ILogFormatter)
    Set self.Formatter = Formatter
End Property


Public Property Get ILogHandler_Name() As String
    ILogHandler_Name = self.Name
End Property


Public Property Let ILogHandler_SetLogLevel(ByVal value As enLogLevel)
    LogLevel = value
End Property


Public Property Set ILogHandler_SetFormatter(ByVal value As ILogFormatter)
    Set self.Formatter = value
End Property



'==============================
'Public Methods
'==============================


Public Sub ILogHandler_HandleRecord(Record As LogRecord)
    'TODO add filter
    Loghandler_EmitRecord Record
End Sub



'==============================
'Private Methods
'==============================


Private Sub Loghandler_EmitRecord(ByVal Record As LogRecord)
    Debug.Print ILogHandler_FormatRecord(Record)
End Sub


Private Function ILogHandler_FormatRecord(ByVal Record As LogRecord) As String
    If Not ILogHandler_HasFormtatter Then
        Set ILogHandler_SetFormatter = New DefaultLogFormatter
    End If
        
    ILogHandler_FormatRecord = self.Formatter.Format(Record)
End Function


Private Function ILogHandler_HasFormtatter() As Boolean
    If Not self.Formatter Is Nothing Then
        ILogHandler_HasFormtatter = True
        Exit Function
    End If
End Function
