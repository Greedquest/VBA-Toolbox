Attribute VB_Name = "TextFormatting"
Option Explicit

Public Sub EngineerFormat()
    If TypeOf Selection Is Range Then
        Selection.NumberFormat = "##0.0E+0"
    End If
End Sub

Public Function printf(ByVal mask As String, ParamArray tokens()) As String
'Format string with by substituting into mask - stackoverflow.com/a/17233834/6609896
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    printf = mask
End Function
