VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "testClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ticker As asyncTimer
Attribute ticker.VB_VarHelpID = -1

Private Sub ticker_Tick()
Debug.Print "Tick"
End Sub

Sub test()
    Set ticker = New asyncTimer
    ticker.await False
End Sub

