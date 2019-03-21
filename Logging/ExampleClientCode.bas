Attribute VB_Name = "ExampleClientCode"
'@Folder("Logger.Examples")
Option Explicit

Public Sub TestLogger()

    On Error GoTo CleanFail

    LogManager.Register DebugLogger.Create("MyLogger", DebugLevel)
    LogManager.Register FileLogger.Create("TestLogger", ErrorLevel, "C:\Dev\VBA\log.txt")

    LogManager.Log TraceLevel, "logger has been created."
    LogManager.Log InfoLevel, "it works!"

    Debug.Print LogManager.IsEnabled(TraceLevel)

    Dim boom As Integer
    boom = 1 / 0

CleanExit:
    LogManager.Log DebugLevel, "we're done here.", "TestLogger"
    Exit Sub

CleanFail:
    LogManager.Log ErrorLevel, Err.Description
    Resume CleanExit

End Sub
