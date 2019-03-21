VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Logger.Interfaces")
Option Explicit

Public Sub Log(ByVal output As String)
End Sub

Public Property Get Name() As String
End Property

Public Property Get MinLevel() As LogLevel
End Property
