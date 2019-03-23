Attribute VB_Name = "StartupRoutines"
'@Folder("Toolbox.Startup")
Option Explicit

Public Sub InitialiseLoggers()
    LogManager.Register DebugLogger.Create("the EYES", TraceLevel)
End Sub
