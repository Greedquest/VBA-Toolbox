Attribute VB_Name = "ExampleClientCode"
'@Folder("Logger.Examples")
Option Explicit

Public Sub TestLogger()

    On Error GoTo CleanFail

    LogManager.Register DebugLogger.Create("MyLogger", DebugLevel)
    LogManager.Register FileLogger.Create("TestLogger", ErrorLevel, "C:\Users\guy\OneDrive - University Of Cambridge\Uni\2nd Year\Labs\IDP\log.txt")

    LogManager.Log TraceLevel, "logger has been created."
    LogManager.Log InfoLevel, "it works!"

    Debug.Print LogManager.IsEnabled(TraceLevel)

    '@Ignore VariableNotUsed
    Dim boom As Long
    boom = 1 / 0

CleanExit:
    LogManager.Log DebugLevel, "we're done here.", "TestLogger"
    Exit Sub

CleanFail:
    LogManager.Log ErrorLevel, Err.Description
    Resume CleanExit

End Sub

