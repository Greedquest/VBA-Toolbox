Attribute VB_Name = "SynchroListUtils"
'@Folder(SynchronisedList.Utils)
Option Explicit

'Public Function FlattenArray(iterableToFlatten As Variant, Optional level As Long = 0) As Collection
'    If Not isIterable(iterableToFlatten) Then Err.Raise 5
'    Dim item As Variant
'    Dim flattenedResult As New Collection        'store added items in a temp collection which preserves order
'    For Each item In iterableToFlatten
'        If isIterable(item) Then
'            Dim contentToAdd As Variant
'Debug.Print level
'            For Each contentToAdd In FlattenArray(item, level + 1)
'                flattenedResult.Add contentToAdd
'            Next contentToAdd
'        Else
'            flattenedResult.Add item
'        End If
'    Next item
'    Set FlattenArray = flattenedResult
'End Function


Public Function flattenParamArray(ParamArray passedParams() As Variant) As Variant
    
    Dim argSet As Variant
    argSet = passedParams(0)
    If NumElements(Array(passedParams)(0)) <> 1 Then
        Err.Description = "Only pass one paramarray to the function"
        Err.Raise 5
    Else
        argSet = passedParams(0)
    End If
    Dim result() As Variant
    Dim i As Long
    Dim noErrors As Boolean
    noErrors = True
    For i = LBound(argSet) To UBound(argSet)
        If IsArray(argSet(i)) Then
            noErrors = noErrors And ConcatenateArrays(result, argSet(i))
        Else
            noErrors = noErrors And ConcatenateArrays(result, Array(argSet(i)))
        End If
        If Not noErrors Then
            Err.Description = "Unable to merge all the items in paramarray, possible type conflict"
            Err.Raise 5
        End If
    Next
    
    flattenParamArray = result
End Function

Public Sub LetSet(ByRef Variable As Variant, ByVal Value As Variant)
    If IsObject(Value) Then
        Set Variable = Value
    Else
        Variable = Value
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
        Err.Raise 5
    End If
End Function

