Attribute VB_Name = "codeReviewDemo"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub StatusBarProgressInit()               'uses New syntax - could equally use existing instance
    Const runningTime As Single = 100            'in milliseconds
    Const numberOfSteps As Long = 100
    With New AsciiProgressBar
        .Init base:="Running: ", formatMask:="{0}{2}%{1}|"
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
    Application.StatusBar = "Complete!"
    Sleep 1000
    Application.StatusBar = False
End Sub

Public Sub StatusBarProgress()
    Const runningTime As Single = 100            'in milliseconds
    Const numberOfSteps As Long = 100
    With AsciiProgressBar.Create(base:="Running: ", formatMask:="{0}{2}%{1}|")
        '.Init base:="Running: ", formatMask:="{0}{2}%{1}|"
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
    Application.StatusBar = "Complete!"
    Sleep 1000
    Application.StatusBar = False
End Sub

Public Sub errorTest()
    With Toolbox.AsciiProgressBar.Create(size:=20, whitespace:="#")
Debug.Print .Update(0.713)
    End With
End Sub

