VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logger.Loggers")
Option Explicit

Private Type TFileLogger
    Name As String
    MinLevel As LogLevel
    LogFile As String
    Writer As TextWriter
End Type
Private this As TFileLogger

Implements ILogger

Public Function Create(ByVal loggerName As String, ByVal loggerMinLevel As LogLevel, ByVal path As String) As ILogger

    Dim result As New FileLogger
    result.Name = loggerName
    result.MinLevel = loggerMinLevel
    Set result.Writer = TextWriter.Create
    result.LogFile = path
    Set Create = result

End Function

Friend Property Get Name() As String
    Name = this.Name
End Property

Friend Property Let Name(ByVal value As String)
    this.Name = value
End Property

Friend Property Get MinLevel() As LogLevel
    MinLevel = this.MinLevel
End Property

Friend Property Let MinLevel(ByVal value As LogLevel)
    this.MinLevel = value
End Property

Friend Property Get LogFile() As String
    LogFile = this.LogFile
End Property

Friend Property Let LogFile(ByVal value As String)
    this.LogFile = value
End Property

Friend Property Get Writer() As TextWriter
    Set Writer = this.Writer
End Property

Friend Property Set Writer(ByVal value As TextWriter)
    Set this.Writer = value
End Property

Private Sub ILogger_Log(ByVal output As String)
    If this.Writer.OpenFile(this.LogFile) Then

        this.Writer.WriteLine output
        this.Writer.CloseFile

    Else
        Err.Raise vbObjectError + 1092, "ILogger.Log", "FileLogger.LogFile could not be opened."
    End If
End Sub

Private Property Get ILogger_MinLevel() As LogLevel
    ILogger_MinLevel = this.MinLevel
End Property

Private Property Get ILogger_Name() As String
    ILogger_Name = this.Name
End Property
