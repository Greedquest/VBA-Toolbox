Attribute VB_Name = "tickGenerator"
Option Explicit

Public Declare Function SetTimer Lib "user32" ( _
                        ByVal HWnd As Long, ByVal nIDEvent As Long, _
                        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" ( _
                        ByVal HWnd As Long, ByVal nIDEvent As Long) As Long

Private Type tGenerator
    caller As asyncTimer
    timerID As Long
End Type

Private this As tGenerator

Public Sub startTicking(ByVal tickFrequency As Double, ByVal caller As asyncTimer)
    Set this.caller = caller
    this.timerID = SetTimer(0, 0, tickFrequency * 1000, AddressOf Tick)
End Sub

Public Sub Tick(ByVal HWnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTimer As Long)
    this.caller.Tick
End Sub

Public Sub stopTicking()
    On Error Resume Next
    KillTimer 0, this.timerID
End Sub
