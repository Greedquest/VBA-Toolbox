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

Public Function Format(ByVal mask As String, ParamArray tokens()) As String
    Format = StringFormat(mask, Array(tokens)(0))
End Function

Public Function Contains(haystack, needle, Optional ByVal caseSensitive As Boolean = False) As Boolean

    Dim compareMethod As Integer

    If caseSensitive Then
        compareMethod = vbBinaryCompare
    Else
        compareMethod = vbTextCompare
    End If
    'Have you thought about Null?
    Contains = (InStr(1, CStr(haystack), CStr(needle), compareMethod) <> 0)

End Function

Public Function ContainsAny(haystack, ByVal caseSensitive As Boolean, ParamArray needles()) As Boolean

    Dim i As Integer

    For i = LBound(needles) To UBound(needles)
        If Contains(CStr(haystack), CStr(needles(i)), caseSensitive) Then
            ContainsAny = True
            Exit Function
        End If
    Next

    ContainsAny = False                          'Not really necessary, default is False..

End Function

Public Function MatchesAny(ByVal string_source, ParamArray potential_matches()) As Boolean

    'String-typed local copies of passed parameter values:
    Dim find As String, src As String, i As Integer, found As Boolean
    src = CStr(string_source)

    For i = LBound(potential_matches) To UBound(potential_matches)
        find = CStr(potential_matches(i))
        found = (src = find)
        If found Then Exit For
    Next

    MatchesAny = found

End Function

Public Function StartsWith(ByVal start As String, ByVal string_source) As Boolean

    StartsWith = Left$(CStr(string_source), Len(CStr(start))) = CStr(start)

End Function

