VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultMethodCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Callbacks.DefaultMethod")

Option Explicit

'@DefaultMember
Public Sub TimerProc(ByVal windowHandle As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
Attribute TimerProc.VB_UserMemId = 0
    Debug.Print printf("Called from timer {0}!", timerID)
End Sub
