Attribute VB_Name = "Callback"
'@Folder("Toolbox.Startup")
Option Explicit

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                         ByVal lpPrevWndFunc As Long, _
                         ByVal HWnd As Long, _
                         ByVal msg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
 
Public Declare Function SetTimer Lib "user32" ( _
                        ByVal HWnd As Long, _
                        ByVal nIDEvent As Long, _
                        ByVal uElapse As Long, _
                        ByVal lpTimerFunc As Long) As Long

Public Declare Function KillTimer Lib "user32" ( _
                        ByVal HWnd As Long, _
                        ByVal nIDEvent As Long) As Long
                        
Public Type objectCallback
    Object As Object
    ProcName As String
    CallType As VbCallType
    Args As Variant
    TimerID As Long
End Type

Private Sub SendMessage(ByVal callbackPointer As Long, ByRef callbackObject As objectCallback)
    callbackObject.TimerID = 0
    CallWindowProc callbackPointer, callbackObject, 0&, 0&, 0&
    
End Sub

Private Sub StartTimer(ByVal pauseMillis As Long, ByVal callbackPointer As Long, ByRef callbackObject As objectCallback)
    'return timer id
    callbackObject.TimerID = SetTimer(0&, VarPtr(callbackObject), pauseMillis, callbackPointer)
End Sub

Private Sub EndTimer(ByVal TimerID As Long)
    On Error Resume Next
    KillTimer 0&, TimerID
End Sub


Private Sub ClassMethodCallback(ByRef callbackObject As objectCallback, ByVal unused1 As Long, ByVal unused2 As Long, ByVal unused3 As Long)
    CallByName callbackObject.Object, callbackObject.ProcName, callbackObject.CallType, callbackObject.Args
    EndTimer callbackObject.TimerID
End Sub

Public Sub CallClassMethod(ByVal Object As Object, ByVal methodName As String, Optional ByVal delayMillis As Long)
    Dim params As objectCallback
    params.CallType = VbMethod
    Set params.Object = Object
    params.ProcName = methodName
    
    If delayMillis = 0 Then
        SendMessage AddressOf ClassMethodCallback, params
    If delayMillis > 0 Then
        Err.Raise 5
        'StartTimer delayMillis, AddressOf ClassMethodCallback, params
    Else
        Err.Raise 5 'bad argument
    End If
End Sub
'
'Private Function ProcPtr(ByVal nAddress As Long) As Long
'    'Just return the address we just got
'    ProcPtr = nAddress
'End Function
'
'Public Sub CallFunction(ByVal address As LongPtr)
'    Dim sMessage As String
'    Dim nSubAddress As Long
'
'
'    'Get the address to the sub we are going to call
'    nSubAddress = ProcPtr(AddressOf ShowMessage)
'    'Do the magic!
'    CallWindowProc nSubAddress, VarPtr(sMessage), 0&, 0&, 0&
'End Sub
'
'
Sub t()
    Dim x As New Class1
    CallByName x, "chirp", VbMethod
End Sub
