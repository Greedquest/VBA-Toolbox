Attribute VB_Name = "Patches"
'@Folder("Logger.Utils.TextWriter")
Option Explicit

Public Function TryGetValue(k As Variant, ByRef outValue As Variant, ByVal dict As Dictionary) As Boolean

    If dict.Exists(k) Then
        letset outValue, dict(key)
        TryGetValue = True

    Else

        TryGetValue = False

    End If

End Function

Public Sub RemoveByValue(ByVal lookupVal As Variant, ByVal dict As Dictionary)
    Dim key As Variant
    For Each key In dict.Keys
        If dict(key) = lookupVal Then
            dict.Remove key
            Exit For
        End If
    Next
End Sub


Private Sub letset(ByRef variable As Variant, ByVal value As Variant)
    If IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If
End Sub
