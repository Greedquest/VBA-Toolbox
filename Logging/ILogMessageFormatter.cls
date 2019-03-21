VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogMessageFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logger.Interfaces")
Option Explicit

Public Function FormatMessage(ByVal level As LogLevel, ByVal loggerName As String, ByVal message As String) As String
End Function

