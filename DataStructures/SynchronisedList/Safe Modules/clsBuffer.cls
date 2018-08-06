VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SynchronisedList")
Option Explicit
'''
'Class breaks up data stream into chunks
'Events raised every n items
'Events for adding, removing and ammending items
'1-indexed
'''

Public Event ItemsAdded()
Public Event ItemsAmmended()
Public Event ItemsRemoved()

Private addingThreshold As Long
Private ammendingThreshold As Long
Private removingThreshold As Long

Private Type TBuffer
    BufferQueue As Object
    AmmendmentsQueue As Object
    RemovalQueue As Object
End Type

Private this As TBuffer

Private Sub Enqueue(item As Variant, queue As Object)
    queue.Enqueue item
End Sub

Private Sub Dequeue(queue As Object, ByRef writeTo As Variant)
    'On Error Resume Next
    If IsObject(queue.peek) Then
        Set writeTo = queue.Dequeue
    Else
        writeTo = queue.Dequeue
    End If
End Sub

Public Sub AddItems(iterableItems As Variant)
    'add iterable load of items to buffer
    ' - raise event every n items (according to adding trigger
    '   to push through buffer
    ' - Also raise event when nothing more to add TODO buffer with async timer
    Dim item As Variant
    Dim justRaised As Boolean
    If Not isIterable(iterableItems) Then Err.Description = "Items added to buffer must be iterable, not " & TypeName(iterableItems): Err.Raise 5
    For Each item In iterableItems
        justRaised = False
        Enqueue item, this.BufferQueue
        If this.BufferQueue.Count >= Me.AddingTrigger Then
            RaiseEvent ItemsAdded
            justRaised = True
        End If
    Next item
    If Not justRaised Then RaiseEvent ItemsAdded 'clear out tail end
End Sub

Public Sub AmmendItems(iterableItems As Variant)
    'Mark items as ammended (by key)
    Dim item As Variant
    Dim justRaised As Boolean
    For Each item In iterableItems
        justRaised = False
        Enqueue item, this.AmmendmentsQueue
        If this.AmmendmentsQueue.Count >= Me.AmmendingTrigger Then
            RaiseEvent ItemsAmmended
            justRaised = True
        End If
    Next item
    If Not justRaised Then RaiseEvent ItemsAmmended
End Sub

Public Sub RemoveItems(iterableItems As Variant)
    'Mark items for removal (byRef)
    Dim item As Variant
    Dim justRaised As Boolean
    For Each item In iterableItems
        justRaised = False
        Enqueue item, this.RemovalQueue
        If this.RemovalQueue.Count >= Me.RemovingTrigger Then
            RaiseEvent ItemsRemoved
            justRaised = True
        End If
    Next item
    If Not justRaised Then RaiseEvent ItemsRemoved 'clear out tail end

End Sub

Public Function AddedItems()
    'get array of items out
    Dim results()
    Dim noItems As Long: noItems = this.BufferQueue.Count
    If noItems > 0 Then
        ReDim results(1 To noItems)
        Dim i As Long
        For i = 1 To noItems
            Dequeue this.BufferQueue, results(i)
        Next i
    
    Else
        ReDim results(0 To 0)
    End If
    AddedItems = results
End Function

Public Function RemovalItems()
    Dim results()
    Dim noItems As Long: noItems = this.RemovalQueue.Count
    ReDim results(1 To noItems)
    Dim i As Long
    For i = 1 To noItems
        Dequeue this.RemovalQueue, results(i)
    Next i
    RemovalItems = results
End Function

Public Function AmmendedItems()
    Dim results()
    Dim noItems As Long: noItems = this.AmmendmentsQueue.Count
    ReDim results(1 To noItems)
    Dim i As Long
    For i = 1 To noItems
        Dequeue this.AmmendmentsQueue, results(i)
    Next i
    AmmendedItems = results
End Function

Public Property Get RemovingTrigger() As Long
    RemovingTrigger = removingThreshold
End Property

Public Property Let RemovingTrigger(ByVal Value As Long)
    removingThreshold = Value
End Property

Public Property Get AmmendingTrigger() As Long
    AmmendingTrigger = ammendingThreshold
End Property

Public Property Let AmmendingTrigger(ByVal Value As Long)
    ammendingThreshold = Value
End Property

Public Property Get AddingTrigger() As Long
    AddingTrigger = addingThreshold
End Property

Public Property Let AddingTrigger(ByVal Value As Long)
    addingThreshold = Value
End Property

Private Sub Class_Initialize()
    'make queues
    Set this.BufferQueue = CreateObject("System.Collections.Queue")
    Set this.RemovalQueue = CreateObject("System.Collections.Queue")
    Set this.AmmendmentsQueue = CreateObject("System.Collections.Queue")
    'set defaults
    Me.AddingTrigger = 5
    Me.RemovingTrigger = 5
    Me.AmmendingTrigger = 5
End Sub

