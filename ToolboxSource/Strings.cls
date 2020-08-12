VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Static module for doing string manipulation such as printing and formatting"
'@Folder("Strings")
'@ModuleDescription("Static module for doing string manipulation such as printing and formatting")
'@PredeclaredId
'@Exposed
Option Explicit

Public Function Format(ByVal mask As String, ParamArray tokens()) As String
    Format = StringFormat(mask, Array(tokens)(0))
End Function

Public Function Contains(ByVal haystack As Variant, ByVal needle As Variant, Optional ByVal caseSensitive As Boolean = False) As Boolean

    Dim compareMethod As Long

    If caseSensitive Then
        compareMethod = vbBinaryCompare
    Else
        compareMethod = vbTextCompare
    End If
    'Have you thought about Null?
    Contains = (InStr(1, CStr(haystack), CStr(needle), compareMethod) <> 0)

End Function

Public Function ContainsAny(ByVal haystack As Variant, ByVal caseSensitive As Boolean, ParamArray needles() As Variant) As Boolean

    Dim i As Long

    For i = LBound(needles) To UBound(needles)
        If Contains(CStr(haystack), CStr(needles(i)), caseSensitive) Then
            ContainsAny = True
            Exit Function
        End If
    Next

    ContainsAny = False                          'Not really necessary, default is False..

End Function

Public Function MatchesAny(ByVal source As String, ParamArray potentialMatches()) As Boolean

    'String-typed local copies of passed parameter values:
    Dim find As String
    Dim i As Long
    Dim found As Boolean

    For i = LBound(potentialMatches) To UBound(potentialMatches)
        find = CStr(potentialMatches(i))
        found = (source = find)
        If found Then Exit For
    Next

    MatchesAny = found

End Function

Public Function StartsWith(ByVal start As String, ByVal string_source As Variant) As Boolean

    StartsWith = Left$(CStr(string_source), Len(CStr(start))) = CStr(start)

End Function

