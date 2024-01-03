VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultLogFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logging")

Option Explicit
Implements ILogFormatter



Public Function ILogFormatter_Format(ByVal Record As ILogRecord) As String
    Dim Output As String
    Dim LogName As String
    Dim TimeStamp As String
    Dim LogLevel As String
    
   
    LogLevel = Logging.GetLevelName(Record.LogLevel)
    LogName = Record.LoggerName
    TimeStamp = Strings.Format$(Record.TimeStamp, "yyyy-mm-dd hh:nn:ss")
    
    Output = LogLevel & " - " & LogName & " - " & TimeStamp & " - " & Record.Message
    
    ILogFormatter_Format = Output
    
End Function
