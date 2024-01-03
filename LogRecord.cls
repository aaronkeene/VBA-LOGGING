VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logging")

Option Explicit
Implements ILogRecord



Private Type TLogRecord
    LoggerName  As String
    LogLevel    As String
    Message     As String
    TimeStamp   As Date
End Type



Private self As TLogRecord

    
    
Public Function Create(ByVal LoggerName As String, ByVal LogLevel As enLogLevel, ByVal Message As String) As LogRecord
    Dim result As New LogRecord
        
    With result
        .LoggerName = LoggerName
        .LogLevel = LogLevel
        .Message = Message
        .TimeStamp = Now()
    End With
    
    Set Create = result
End Function



'==============================
'Properties
'==============================


Friend Property Let LoggerName(ByVal value As String)
    self.LoggerName = value
End Property


Friend Property Let LogLevel(ByVal value As enLogLevel)
    self.LogLevel = value
End Property


Friend Property Let Message(ByVal value As String)
    self.Message = value
End Property


Friend Property Let TimeStamp(ByVal value As Date)
    self.TimeStamp = value
End Property


Public Property Get ILogRecord_LoggerName() As String
    ILogRecord_LoggerName = self.LoggerName
End Property


Public Property Get ILogRecord_LogLevel() As enLogLevel
    ILogRecord_LogLevel = self.LogLevel
End Property


Public Property Get ILogRecord_Message() As String
    ILogRecord_Message = self.Message
End Property


Public Property Get ILogRecord_TimeStamp() As Date
    ILogRecord_TimeStamp = self.TimeStamp
End Property

