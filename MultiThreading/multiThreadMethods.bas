Attribute VB_Name = "multiThreadMethods"
'@Folder("Main.Code")
Option Explicit

Public Function addIterableToQueue(iterator As Variant, ByRef resultQueue As Queue) As Long
    'function to take iterable group and add it to the queue
    'returns the number of items added
    Dim item As Variant
    Dim itemsAdded As Long
    itemsAdded = 0
    For Each item In iterator
        resultQueue.enqueue item
        itemsAdded = itemsAdded + 1
    Next item
    addIterableToQueue = itemsAdded
End Function

Function isIterable(obj As Variant) As Boolean
    On Error Resume Next
    Dim iterator As Variant
    For Each iterator In obj
        Exit For
    Next
    isIterable = Err.Number = 0
End Function
