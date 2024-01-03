VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogFileHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Logging")
'@PredeclaredId

Option Explicit
Implements ILogHandler


Private Type TLogFileHandler
    FilePath    As String
    Name        As String
    LogLevel    As enLogLevel
    Formatter   As New ILogFormatter
End Type


Private self As TLogFileHandler


Public Function Create(ByVal FilePath As String, Optional ByVal Name As String, Optional ByVal LogLevel As enLogLevel, Optional Formatter As ILogFormatter) As ILogHandler
    
    'Filepath
'@TODO need to implement VBA list to avoid using names

    'check if name pasesd in, otherside use default name;
    If Name = vbNullString Then
        Name = "file_handler"
    End If
    
    'check if formatter is passed in, otherwise use defualt log formatter
    If Formatter Is Nothing Then
        Set Formatter = New DefaultLogFormatter
    End If
    
    Dim result As New LogFileHandler
    
    With result
        .FilePath = FilePath
        .Name = Name
        .LogLevel = LogLevel
        Set .Formatter = Formatter
    End With
    
    Set Create = result

End Function





'Friend properties

Friend Property Let Name(ByVal value As String)
    self.Name = value
End Property


Friend Property Let LogLevel(ByVal value As enLogLevel)
    self.LogLevel = value
End Property


Friend Property Set Formatter(ByVal Formatter As ILogFormatter)
    Set self.Formatter = Formatter
End Property


Friend Property Let FilePath(ByVal value As String)
    self.FilePath = value
End Property




'==============================
'Properties
'==============================

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
    'filter
    Loghandler_EmitRecord Record
End Sub

'==============================
'Private Methods
'==============================

Private Sub Loghandler_EmitRecord(ByVal Record As LogRecord)
    Debug.Print self.Formatter.Format(Record)
End Sub















Private Property Get ILogHandler_Name() As String
    ILogHandler_Name = self.Name
End Property


Private Property Get LogHandler_LogLevel() As enLogLevel
    LogHandler_LogLevel = self.LogLevel
End Property
