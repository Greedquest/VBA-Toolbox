Attribute VB_Name = "Callback"
'@Folder("Callbacks")
Option Explicit

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



Private windowHandle As Long
'Private storedParams As objectCallback

Private Sub SendMessage(ByVal callbackPointer As Long, ByRef CallbackParams As objectCallback)

    Debug.Print "Sending Message..."
    CallbackParams.timerID = 0
    CallWindowProc callbackPointer, VarPtr(CallbackParams), 0, 0, 0

End Sub

Private Sub StartTimer(ByVal pauseMillis As Long, ByVal callbackPointer As Long, ByRef CallbackParams As objectCallback, Optional ByVal HWnd As Long = 0)

    If HWnd = 0 Then
        Debug.Print "Calibrating..."
        Debug.Print 0, 0, pauseMillis, callbackPointer
        Debug.Print printf("Setting {1}ms timer with ID {0}", SetTimer(0, 0, pauseMillis, callbackPointer), pauseMillis)
    Else
        Debug.Print "Excecuting Custom Method..."
        Debug.Print , VarPtr(CallbackParams)
        Debug.Print HWnd, VarPtr(CallbackParams), pauseMillis, callbackPointer
        CallbackParams.timerID = SetTimer(HWnd, VarPtr(CallbackParams), pauseMillis, callbackPointer)
        Debug.Print printf("Setting {1}ms timer with ID {0}", CallbackParams.timerID, pauseMillis)
    End If

    '    Debug.Print 0&, VarPtr(callbackObject), pauseMillis, callbackPointer 'these are the args
    '    callbackObject.timerID = SetTimer(0&, VarPtr(callbackObject), pauseMillis, callbackPointer)

End Sub

Private Sub EndTimer(ByVal timerID As Long, Optional ByVal HWnd As Long = 0)
    On Error Resume Next                         'TODO, what error is this?
    KillTimer HWnd, timerID
    Debug.Print printf("Killing {0} on handle {1}", timerID, HWnd)
End Sub

Private Sub CalibrateHandleProc(ByVal HWnd As Long, ByVal uMsg As Long, ByVal timerID As Long, ByVal tickCount As Long)
    EndTimer timerID, 0
    windowHandle = HWnd
    Debug.Print printf("Timer calibrated to handle {0}", windowHandle)
End Sub

Private Sub ClassMethodCallbackProc(ByRef callbackObject As objectCallback, ByVal unused1 As Long, ByVal timerID As Long, ByVal unused3 As Long)
    EndTimer timerID
    CallByName callbackObject.object, callbackObject.procName, callbackObject.callType
End Sub

Private Sub TimerIDTestProc(ByVal HWnd As Long, ByVal uMsg As Long, ByVal callbackObjectPointer As Long, ByVal tickCount As Long)

    EndTimer callbackObjectPointer, HWnd

    Dim param As objectCallback
    objectCallbackDeReference callbackObjectPointer, param
    Debug.Print param.procName

End Sub

Private Sub MessageTestProc(ByVal HWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Debug.Print "Message Read:"
    Debug.Print HWnd, uMsg, wParam, lParam
End Sub

Public Sub CallClassMethod(ByVal object As Object, ByVal methodName As String, Optional ByVal delayMillis As Long = 0, Optional ByVal calibrationDelayMillis As Long = 200)

    '    Dim params As objectCallback
    '    params.CallType = VbMethod
    '    Set params.Object = Object
    '    params.ProcName = methodName
    '
    '    storedParams = params
    Dim storedParams As objectCallback
    storedParams.procName = methodName


    If delayMillis = 0 Then
        SendMessage AddressOf MessageTestProc, storedParams
    ElseIf delayMillis > 0 Then
        'Err.Raise 5
        'StartTimer delayMillis, AddressOf ClassMethodCallback, params
        If windowHandle = 0 Then
            StartTimer calibrationDelayMillis, AddressOf CalibrateHandleProc, storedParams
        Else
            StartTimer delayMillis, AddressOf TimerIDTestProc, storedParams, windowHandle
        End If
    Else
        Err.Raise 5                              'bad argument
    End If
End Sub

Sub testCaller()
    CallClassMethod New TestObj, "chirp", 500
End Sub

Private Property Get paramsMemoryOffset() As Long
    Dim a As objectCallback
    paramsMemoryOffset = LenB(a)
End Property

Sub testObjReconstruction()

    Dim a As objectCallback
    a.procName = "barry"
    Debug.Print a.procName

    Dim b As objectCallback
    objectCallbackDeReference VarPtr(a), b
    Debug.Print b.procName

End Sub

Sub testMemLocations()
    'we know safeArrays occupy continuous areas in memory
    'so to check actual memory footprint you can just take the difference in location between elements
    Dim obArray(1 To 2) As objectCallback
    Debug.Assert VarPtr(obArray(2)) - VarPtr(obArray(1)) = paramsMemoryOffset
End Sub

Sub testDefaultMember()
    Dim a As DefaultMethodCallback
    Set a = New DefaultMethodCallback
    CallWindowProc VarPtr(a), 10, WM_NOTIFY, 11, 17
End Sub




