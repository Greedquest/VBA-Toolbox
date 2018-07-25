VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkerEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Main.Interfaces")
Option Explicit

'class holds all the events that a thread task can raise

Public Event Complete(returnVal As Variant)
Public Event statusChange(statusVal As Variant)
Public Event Started()

Public Sub Complete(Optional returnVal As Variant)
    RaiseEvent Complete(returnVal)
End Sub

Public Sub statusChange(statusVal As Variant)
    RaiseEvent statusChange(statusVal)
End Sub

Public Sub Started()
    RaiseEvent Started
End Sub
