VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logger.Loggers")
Option Explicit

Private Type TDebugLogger
    Name As String
    MinLevel As LogLevel
End Type
Private this As TDebugLogger

Implements ILogger

Public Function Create(ByVal loggerName As String, ByVal loggerMinLevel As LogLevel) As ILogger

    Dim result As New DebugLogger
    result.Name = loggerName
    result.MinLevel = loggerMinLevel
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

Private Sub ILogger_Log(ByVal output As String)
    Debug.Print output
End Sub

Private Property Get ILogger_MinLevel() As LogLevel
    ILogger_MinLevel = this.MinLevel
End Property

Private Property Get ILogger_Name() As String
    ILogger_Name = this.Name
End Property
