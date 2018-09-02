Attribute VB_Name = "tickGenerator"
Option Explicit

'Module generates tick callbacks with a given frequency until told to stop

Private Type tGenerator
    keepTicking As Boolean
    tickFrequency As Double
    caller As asyncTimer
End Type

Private this As tGenerator

Public Sub startTicking(ByVal frequency As Double, ByVal caller As asyncTimer)
    this.tickFrequency = frequency
    this.keepTicking = True
    Set this.caller = caller
    Tick
End Sub
Public Sub Tick()
    this.caller.Tick
    If this.keepTicking Then
        Application.OnTime Now + this.tickFrequency, "Tick"
    End If
End Sub

Public Sub stopTicking()
    this.keepTicking = False
End Sub
