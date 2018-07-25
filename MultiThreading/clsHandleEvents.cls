VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHandleEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Main.Code")
Option Explicit
'class to convert calls from the thread handle into events which the multi thread group can tap into

Public Event Complete(obj As clsThreadHandle, returnVal As Variant)
Public Event Opened(threadID As String)          'when thread is actually opened
Public Event Closed(threadID As String)          'when thread is closed

Public Sub NotifyComplete(obj As clsThreadHandle, Optional returnVal As Variant)
    RaiseEvent Complete(obj, returnVal)
End Sub

Public Sub NotifyThreadOpened(threadID As String)
    RaiseEvent Opened(threadID)
End Sub

Public Sub NotifyThreadClosed(threadID As String)
    RaiseEvent Closed(threadID)
End Sub

Private Sub Class_Terminate()
Debug.Print "Events Terminated"
End Sub
