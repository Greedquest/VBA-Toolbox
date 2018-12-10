Attribute VB_Name = "SynchroListUtils"
'@Folder(SynchronisedList.Utils)
Option Explicit

Public Enum BufferMode
    slAdding
    slRemoving
    slAmmending
End Enum

Public Function ArrayToFilterList(ByVal itemArray As Variant) As FilterList
    If IsArrayEmpty(itemArray) Then
        Exit Function
    Else
        Dim result As New FilterList
        Dim i As Long
        For i = LBound(itemArray) To UBound(itemArray)
            result.Add itemArray(i)
        Next i
        Set ArrayToFilterList = result
    End If
End Function

Public Function flattenParamArray(ParamArray passedParams() As Variant) As Variant
    
    Dim argSet As Variant
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
        If IsArrayNotObj(argSet(i)) Then
            noErrors = noErrors And ConcatenateArrays(result, argSet(i))
        ElseIf isIterable(argSet(i)) Then
            noErrors = noErrors And ConcatenateArrays(result, IterableToArray(argSet(i)))
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

Public Function IsArrayNotObj(valueToTest As Variant) As Boolean
    IsArrayNotObj = IsArray(valueToTest) And Not IsObject(valueToTest)
End Function

Public Function removeDuplicates(ByVal inArray As Variant, ByVal dataSet As FilterList) As Variant
    'Function checks if inArray items are in dataset, removes them from the array if so
    'Returns Nothing if all items dropped or inArray is empty
    'Errors if inArray is not an array
    If IsArray(inArray) Then
        Dim lowerBound As Long
        lowerBound = LBound(inArray)
        Dim upperBound As Long
        upperBound = UBound(inArray)
        'return either the filtered array or nothing
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

Public Function IterableToArray(ByVal iterableObject As Variant, Optional ByVal itemCount As Long = 8, Optional ByVal base As Long = 0) As Variant
    If Not isIterable(iterableObject) Then Err.Description = "You can only convert iterable objects to arrays, not " & TypeName(iterableObject): Err.Raise 5
    Dim item
    Dim result()
    Dim arraySize As Long
    arraySize = IIf(itemCount = 0, 1, itemCount)
    ReDim result(base To arraySize + base - 1)
    Dim Count As Long
    Count = base
    For Each item In iterableObject
        If Count > UBound(result) Then
            arraySize = arraySize * 2
            ReDim Preserve result(base To arraySize + base - 1)
        End If
        LetSet result(Count), item
        Count = Count + 1
    Next
    arraySize = Count - 1
    If UBound(result) > arraySize Then
        ReDim Preserve result(base To arraySize)
    End If
    IterableToArray = result
End Function

