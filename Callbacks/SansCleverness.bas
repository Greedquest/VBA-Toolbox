Attribute VB_Name = "SansCleverness"
Option Explicit
'@Folder("Callbacks")

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                         ByVal lpPrevWndFunc As Long, _
                         ByVal HWnd As Long, _
                         ByVal msg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long

Private Declare Function SetTimer Lib "user32" ( _
                         ByVal HWnd As Long, _
                         ByVal nIDEvent As Long, _
                         ByVal uElapse As Long, _
                         ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" ( _
                         ByVal HWnd As Long, _
                         ByVal nIDEvent As Long) As Long

'@Description("Windows Timer Message https://docs.microsoft.com/windows/desktop/winmsg/wm-timer")
Public Enum WindowsMessage
    WM_TIMER = &H113
    WM_NOTIFY = &H4E 'arbitrary, sounds nice though
End Enum

Private paramDict As Dictionary

Public Sub DelayedCallClassMethod(ByVal object As Object, ByVal methodName As String, Optional ByVal delayMillis As Long = 0)
    
    'construct param object from arguments
    Dim params As CallbackParams
    Set params = CallbackParams.Create(object, methodName, VbMethod)
    Dim timerID As Long
    
    If delayMillis = 0 Then
        timerID = uniqueId
        paramList.Add timerID, params
        CallWindowProc AddressOf ClassMethodCallbackProc, 0, WM_NOTIFY, timerID, 0
    ElseIf delayMillis > 0 Then
        timerID = SetTimer(0, 0, delayMillis, AddressOf ClassMethodCallbackProc)
        paramList.Add timerID, params 'BUG: could callback run before params are added?
    Else
        Err.Raise 5                              'bad argument
    End If
    
End Sub

Private Property Get paramList() As Dictionary
    If paramDict Is Nothing Then Set paramDict = CreateObject("Scripting.Dictionary")
    Set paramList = paramDict
End Property

Private Property Get uniqueId() As Long
    Static index As Long
    Do While paramList.Exists(index)
        index = index + 1
    Loop
    uniqueId = index
End Property

Private Sub ClassMethodCallbackProc(ByVal windowHandle As Long, ByVal message As WindowsMessage, ByVal timerID As Long, ByVal tickCount As Long)
        

    Select Case message
    Case WM_TIMER
        KillTimer 0, timerID
    Case WM_NOTIFY
        'No timer to end, so do nothing
    Case Else
        Err.Raise 5
        Exit Sub
    End Select
    
    'BUG if we get 2 messages (not killed fast enough perhaps), may try to access item what's already removed?
    Dim params As CallbackParams
    Set params = paramList.Item(timerID)
    paramList.Remove timerID
    CallClassMethod params
    
End Sub

Private Sub CallClassMethod(ByVal params As CallbackParams)
    CallByName params.object, params.procName, params.callType
End Sub
