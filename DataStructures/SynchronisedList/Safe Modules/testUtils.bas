Attribute VB_Name = "testUtils"
'@Folder("Tests.Utils")
Option Explicit

Public Function getEmptyDummyClasses(Optional N As Long = 5) As DummyGridItem()
    Dim results() As DummyGridItem
    ReDim results(1 To N)
    Dim dummyItem As DummyGridItem
    Dim i As Long
    For i = 1 To N
        Set dummyItem = New DummyGridItem
        dummyItem.Name = "DummyItem " & i
        Set results(i) = dummyItem
    Next i
    getEmptyDummyClasses = results
End Function

Public Function getDummyClassesWithProperties(Optional N As Variant, Optional propertyPairs As Variant) As DummyGridItem()
    'propertypairs is zero indexed array of 0 indexed (name,val) arrays
   
    '!!! Need Array(ItemDefinition1(property1(name,val),property2(name,val),...),Item2(...),...)
    'Ubound(PropertyPairs) must equal n
    If IsMissing(N) Then
        N = 5
        If Not IsMissing(propertyPairs) Then
            N = UBound(propertyPairs) + 1
        End If
    Else
        If Not IsMissing(propertyPairs) Then
            If UBound(propertyPairs) <> N - 1 Then
                Err.Description = "property pairs aren't of same size as n"
                Err.Raise 5
            End If
        End If
    End If

    Dim results() As DummyGridItem
    ReDim results(1 To N)
    Dim dummyItem As DummyGridItem
    Dim i As Long
    For i = 1 To N
        Set dummyItem = New DummyGridItem
        dummyItem.Name = "DummyItem " & i
        Dim sortingInterface As ISortable
        Set sortingInterface = dummyItem
        With sortingInterface
            If IsMissing(propertyPairs) Then
                sortingInterface.Properties.InitRandom
            Else
                Dim propertyPair
                Set propertyPair = propertyPairs(i - 1)
                '.Properties.addProperty propertyPair(0), propertyPair(1)
            End If
        End With
            
        Set results(i) = dummyItem
    Next i
    getDummyClassesWithProperties = results
End Function

Public Function IterableToArray(ByVal iterableObject As Variant, Optional ByVal itemCount As Long = 8, Optional ByVal base As Long = 0) As Variant
    If Not isIterable(iterableObject) Then Err.Description = "You can only convert iterable objects to arrays": Err.Raise 5
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

