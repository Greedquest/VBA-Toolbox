Attribute VB_Name = "TextFormatting"
'@Folder("Toolbox.Common")
Option Explicit

'@Ignore AssignedByValParameter
Public Function printf(ByVal mask As String, ParamArray tokens()) As String
    'Format string with by substituting into mask - stackoverflow.com/a/17233834/6609896
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    printf = mask
End Function

Public Property Let Assign(ByRef variable As Variant, ByVal value As Variant)
    If IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If
End Property
