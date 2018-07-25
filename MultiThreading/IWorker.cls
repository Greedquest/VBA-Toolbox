VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Main.Interfaces")
Option Explicit

'class acts as interface for any thread task
'Execute runs the task
'Events are raised by the task if it interfaces properly

Public Property Set Events(ByRef value As IWorkerEvents)
End Property

Public Sub Execute(Optional argument As Variant)
End Sub

Public Function CreateNew() As IWorker
End Function
