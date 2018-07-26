Attribute VB_Name = "TextFormatting"
Option Explicit

Public Sub EngineerFormat()
    If TypeOf Selection Is Range Then
        Selection.NumberFormat = "##0.0E+0"
    End If
End Sub

