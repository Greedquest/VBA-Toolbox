VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "asyncTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Tick()
Public Event Timeout()
Public Event Complete()

Private Const waitIndefinitely As Double = -1

Private Type tTimer
    tickFrequency As Double                      'in seconds
    timeoutTime As Date
    neverTimeout As Boolean
    conditionAddress As LongPtr
End Type

Private this As tTimer

Public Sub await(ByRef waitUntil As Boolean, Optional ByVal tickFrequency As Double = 1, Optional maxWait As Double = 10)
    this.conditionAddress = VarPtr(waitUntil)
    If maxWait = waitIndefinitely Then
        this.neverTimeout = True
    ElseIf maxWait >= 0 Then
        this.timeoutTime = Now + TimeSerial(0, 0, maxWait)
    Else
        Err.Raise 5
    End If
    
    startTicking tickFrequency, Me
End Sub

Public Sub Tick()
    If Peek(this.conditionAddress, vbBoolean) Then 'check if val has changed
        stopTicking
        RaiseEvent Complete
    ElseIf Now > this.timeoutTime Then
        stopTicking
        RaiseEvent Timeout
    Else
        RaiseEvent Tick
    End If
End Sub

Private Sub Class_Terminate()
    stopTicking
End Sub

