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

Public Function IterableToArray(ByVal iterableObject As Variant) As Variant
    If Not isIterable(iterableObject) Then Err.Description = "You can only convert iterable objects to arrays": Err.Raise 5
    Dim item
    Dim result()
    Dim Count As Long
    Count = 0
    For Each item In iterableObject
        If isIterable(item) Then LetSet item, IterableToArray(item)
        ReDim Preserve result(0 To Count)
        LetSet result(Count), item
        Count = Count + 1
    Next
    IterableToArray = result
End Function

