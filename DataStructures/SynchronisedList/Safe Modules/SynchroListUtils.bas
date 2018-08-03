Attribute VB_Name = "SynchroListUtils"
'@Folder(SynchronisedList.Utils)
Option Explicit

Public Function FlattenArray(iterableToFlatten As Variant, Optional level As Long = 0) As Collection
    If Not isIterable(iterableToFlatten) Then Err.raise 5
    Dim item As Variant
    Dim flattenedResult As New Collection        'store added items in a temp collection which preserves order
    For Each item In iterableToFlatten
        If isIterable(item) Then
            Dim contentToAdd As Variant
Debug.Print level
            For Each contentToAdd In FlattenArray(item, level + 1)
                flattenedResult.Add contentToAdd
            Next contentToAdd
        Else
            flattenedResult.Add item
        End If
    Next item
    Set FlattenArray = flattenedResult
End Function

Public Sub LetSet(ByRef variable As Variant, ByVal Value As Variant)
    If IsObject(Value) Then
        Set variable = Value
    Else
        variable = Value
    End If
End Sub

Public Function IsNothing(valueToTest As Variant) As Boolean
    If Not IsObject(valueToTest) Then
        IsNothing = False
    ElseIf valueToTest Is Nothing Then
        IsNothing = True
    Else
        IsNothing = False
    End If
End Function

Public Function removeDuplicates(ByVal inArray As Variant, ByVal dataSet As FilterList) As Variant
    If IsArray(inArray) Then
        Dim lowerBound As Long
        lowerBound = LBound(inArray)
        Dim upperBound As Long
        upperBound = UBound(inArray)
        If upperBound >= lowerBound Then
            Dim result()
            ReDim result(lowerBound To upperBound)
            Dim i As Long, matchCount As Long
            For i = lowerBound To upperBound
                If Not dataSet.Contains(inArray(i)) Then
                    LetSet result(lowerBound + matchCount), inArray(i)
                    matchCount = matchCount + 1
                End If
            Next i
            If matchCount = 0 Then
                Set removeDuplicates = Nothing
            Else
                ReDim Preserve result(lowerBound To lowerBound + matchCount - 1)
                removeDuplicates = result
            End If
        Else
            Set removeDuplicates = Nothing
        End If
    Else
        Err.Description = "You must pass an array, not a " & TypeName(inArray)
        Err.raise 5
    End If
End Function

