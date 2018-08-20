Attribute VB_Name = "FilterListUtils"
'@Folder("SynchronisedList.Utils.FilterList")
Option Explicit

Public Function isIterable(ByVal obj As Variant) As Boolean
    On Error Resume Next
    Dim iterator As Variant
    For Each iterator In obj
        Exit For
    Next
    isIterable = Err.Number = 0
End Function

