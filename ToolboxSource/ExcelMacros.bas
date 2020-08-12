Attribute VB_Name = "ExcelMacros"
'@Folder("_Excel")
Option Explicit

Public Sub EngineerFormat()
    If TypeOf Selection Is Range Then
        Selection.NumberFormat = "##0.0E+0"
    End If
End Sub
