Attribute VB_Name = "testRunner"
Option Explicit
Private runner As testClass

Sub runTest()
    Set runner = New testClass
    runner.test
End Sub

