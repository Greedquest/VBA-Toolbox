Attribute VB_Name = "codeReviewDemo"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub StatusBarProgress()
    Const runningTime As Single = 5000 'in milliseconds
    Const numberOfSteps As Long = 100
    With New AsciiProgressBar
        .Init base:="Loading: ", formatMask:="{0}{2}%{1}|"
        Dim i As Long
        For i = 1 To numberOfSteps
            .Update i / numberOfSteps
            Application.StatusBar = .repr
            'Or equivalently:
            'Application.StatusBar = .Update(i / numberOfSteps)
            Sleep runningTime / numberOfSteps
            DoEvents
        Next i
    End With
End Sub
