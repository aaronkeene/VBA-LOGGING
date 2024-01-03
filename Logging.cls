VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Logging")
'@PredeclaredId

Option Explicit

Private Const ROOT_LOGGER_NAME As String = "root_logger"

Public Enum enLogLevel
    LOG_LEVEL_CRITICAL = 50
    LOG_LEVEL_ERROR = 40
    LOG_LEVEL_WARNING = 30
    LOG_LEVEL_INFO = 20
    LOG_LEVEL_DEBUG = 10
    LOG_LEVEL_NOTSET = 0
End Enum


Private Type TLogging
    Loggers     As New Dictionary
End Type


Private self As TLogging



'==============================
'Public Methods
'==============================


Public Function GetLogger(Optional Name As String) As ILogger
    If Name = vbNullString Then
        Name = ROOT_LOGGER_NAME
    End If
   
    If self.Loggers.Exists(Name) Then
        Set GetLogger = self.Loggers.Item(Name)
        Exit Function
    End If
    
    If Name = ROOT_LOGGER_NAME Then
        CreateRootLogger
        Set GetLogger = self.Loggers.Item(ROOT_LOGGER_NAME)
        Exit Function
    End If
    
    CreateNewLogger (Name)
    
    Set GetLogger = self.Loggers(Name)
End Function


Private Sub CreateRootLogger()
    CreateNewLogger ROOT_LOGGER_NAME, LOG_LEVEL_WARNING
    
    Dim Logger As ILogger
    Set Logger = self.Loggers(ROOT_LOGGER_NAME)
    
    Logger.AddHandler LogConsoleHandler.Create()
End Sub


Public Sub LogCritical(ByVal Message As String)
    Dim Logger As ILogger
    Set Logger = GetLogger()
    
    Logger.LogCritical (Message)
End Sub


Public Sub LogError(ByVal Message As String)
    Dim Logger As ILogger
    Set Logger = GetLogger()
    
    Logger.LogError (Message)
End Sub


Public Sub LogWarning(ByVal Message As String)
    Dim Logger As ILogger
    Set Logger = GetLogger()
    
    Logger.LogWarning (Message)
End Sub


Public Sub LogInfo(ByVal Message As String)
    Dim Logger As ILogger
    Set Logger = GetLogger()
    
    Logger.LogInfo (Message)
End Sub


Public Sub LogDebug(ByVal Message As String)
    Dim Logger As ILogger
    Set Logger = GetLogger()
    
    Logger.LogDebug (Message)
End Sub



'==============================
'Private Methods
'==============================


Private Sub CreateNewLogger(Name As String, Optional ByVal LogLevel As enLogLevel = LOG_LEVEL_NOTSET)
    self.Loggers.Add Name, Logger.Create(Name, LogLevel)
End Sub


Friend Function GetLevelName(ByVal LogLevel As enLogLevel) As String
    Select Case LogLevel
        Case Is = 50: GetLevelName = "CRITCAL"
        Case Is = 40: GetLevelName = "ERROR"
        Case Is = 30: GetLevelName = "WARNING"
        Case Is = 20: GetLevelName = "INFO"
        Case Is = 10: GetLevelName = "DEBUG"
        Case Is = 0: GetLevelName = "NOTSET"
    End Select
End Function
