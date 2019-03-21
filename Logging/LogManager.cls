VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Logger.Implementation")
'@PredeclaredId
'@Exposed
Option Explicit

Public Enum LogLevel
    TraceLevel = 0
    DebugLevel
    InfoLevel
    WarnLevel
    ErrorLevel
    FatalLevel
End Enum

Private Type TLogManager
    Formatter As ILogMessageFormatter
    Loggers As New Dictionary
End Type

Private this As TLogManager

Public Property Get Formatter() As ILogMessageFormatter
    Set Formatter = this.Formatter
End Property

Public Property Set Formatter(ByVal value As ILogMessageFormatter)
    Set this.Formatter = value
End Property

Public Sub Register(ByVal logger As ILogger)
    If Not this.Loggers.Exists(logger.Name) Then
        this.Loggers.Add logger.Name, logger
    Else
        Err.Raise vbObjectError + 1098, "LogManager.Register", "There is already a logger registered with name '" & logger.Name & "'."
    End If
End Sub

Public Function IsEnabled(ByVal level As LogLevel) As Boolean

    Dim logger As ILogger
    Dim item As Variant
    For Each item In this.Loggers.Items
        Set logger = item
        If level >= logger.MinLevel Then
            IsEnabled = True
            Exit Function
        End If
    Next

End Function

Public Sub Log(ByVal level As LogLevel, ByVal message As String, Optional ByVal loggerName As String)

    Dim logger As ILogger
    If loggerName = vbNullString Then

        Dim item As Variant
        For Each item In this.Loggers.Items
            Set logger = item
            LogWith logger, level, message
        Next

    ElseIf this.Loggers.Exists(loggerName) Then

        LogWith this.Loggers(loggerName), level, message

    Else
        Err.Raise vbObjectError + 1099, "LogManager.Log", "There is no registered logger named '" & loggerName & "'."
    End If

End Sub

Private Sub LogWith(ByVal logger As ILogger, ByVal level As LogLevel, ByVal message As String)
    If level >= logger.MinLevel Then
        logger.Log FormatMessage(level, logger.Name, message)
    End If
End Sub

Friend Function FormatMessage(ByVal level As LogLevel, ByVal loggerName As String, ByVal message As String) As String
    FormatMessage = this.Formatter.FormatMessage(level, loggerName, message)
End Function

Private Sub Class_Initialize()
    Set this.Formatter = New DefaultLogMessageFormatter
End Sub

