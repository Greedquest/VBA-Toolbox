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

Private Const waitIndefinitely As Double = -1

Private Type tTimer
    tickFrequency As Double 'in seconds
    timeoutTime As Date
    neverTimeout As Boolean
End Type

Private this As tTimer

Public Sub await(ByRef waitUntil As Boolean, Optional ByVal tickFrequency As Double = 1, Optional maxWait As Double = 10)

    If maxWait = waitIndefinitely Then
        this.neverTimeout = True
    ElseIf maxWait >= 0 Then
        this.timeoutTime = Now + maxWait
    Else
        Err.Raise 5
    End If

    startTicking tickFrequency, Me
End Sub

Public Sub Tick()
    If Now > this.timeoutTime Then
        stopTicking
        RaiseEvent Timeout
    Else
        RaiseEvent Tick
    End If
End Sub
