Attribute VB_Name = "StopwatchProvider"
Option Explicit
'@Folder Stopwatch

Public Property Get newStopwatch() As Stopwatch
    Set newStopwatch = New Stopwatch
End Property
