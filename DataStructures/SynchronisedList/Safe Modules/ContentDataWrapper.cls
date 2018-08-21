VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentDataWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("SynchronisedList")

'''
'Class to hold data currently after sorting/filtering
'Trigger re-ordering event
'When Items added:
' - Duplicates are removed so item can't be added to grid twice
' - Entire contentData is optionally re-sorted
'When Items are removed:
' - These need references deleted and to be filtered out
' - Won't change sorting order
'Therefore: Filter, Sort and Arrange all need to be implemented separately


Private WithEvents contentDataBuffer As ListBuffer
Attribute contentDataBuffer.VB_VarHelpID = -1

Public Event NeedsFiltering()
Public Event NeedsAligning(ByVal firstChangeIndex As Long) 'Range to re-align
Public Event NeedsSorting(ByVal itemToTrack As Variant)

'Interface events so that custom onAddition/onRemoval handlers can be used
Public Event ItemsAdded(ByVal itemSet As Variant)
Public Event ItemsRemoved(ByVal itemSet As Variant)

Private Type TContentData
    data As FilterList                           'Filtered Data in Grid
    Buffer As ListBuffer
    AlignOnRemoval As Boolean
    SortOnAddition As Boolean
End Type

Private this As TContentData
Private Const alignAll As Long = 0

Private Sub Class_Initialize()
    Set this.data = New FilterList
    Set contentDataBuffer = New ListBuffer
    Set this.Buffer = contentDataBuffer
    this.AlignOnRemoval = True
    this.SortOnAddition = True
End Sub

'''
'Buffer Interface
'''

Public Sub AddItems(ByVal IterableToAdd As Variant)
    this.Buffer.AddItems IterableToAdd
End Sub

Public Sub RemoveItems(ByVal IterableToRemove As Variant)
    this.Buffer.RemoveItems IterableToRemove
End Sub

Public Sub ClearContent()
    'this.Data.Clear won't change visibility necessarily unless good garbage collection
    this.Buffer.RemoveItems this.data.Clone.data 'this.Data.Clone.Data if removing data breaks for loop
End Sub

'''
'Buffer Events
'''
Private Sub contentDataBuffer_ItemsAdded()

'add items which match filter requirements
    Dim newItems
    newItems = this.Buffer.AddedItems
    LetSet newItems, removeDuplicates(newItems, this.data)
    If IsNothing(newItems) Then
        'do nothing
    Else
        Dim i As Long, firstChangeIndex As Long
        firstChangeIndex = this.data.Count       'without sorting, first change is just the end of the array
        For i = 1 To UBound(newItems)
            this.data.Add newItems(i)
        Next i
        RaiseEvent ItemsAdded(newItems)
        If this.SortOnAddition Then
            RaiseEvent NeedsSorting(newItems(1)) 'sort items afresh, tracking where first item goes
        Else
            RaiseEvent NeedsAligning(firstChangeIndex)
        End If
    End If
End Sub

Private Sub contentDataBuffer_ItemsRemoved()
    Dim markedItems                              'remove items from grid and mark for removal from source
    markedItems = contentDataBuffer.RemovalItems
    Dim firstChangeIndex As Long
    firstChangeIndex = this.data.IndexOf(markedItems(1))
    Dim item
    For Each item In markedItems
        this.data.Remove item
    Next item
    RaiseEvent ItemsRemoved(markedItems)
    If this.AlignOnRemoval Then RaiseEvent NeedsAligning(firstChangeIndex)
End Sub

'''
'Grid Methods
'''

Public Sub Filter(ByVal FilterObject As Variant, ByVal FilterMode As lstFilterMode, ByVal Comparer As IComparer)
    'Filter data references
    Dim unFilteredData As FilterList
    Set unFilteredData = this.data.Clone
    this.data.Filter FilterObject, FilterMode, Comparer
    
    Dim firstChangeIndex As Long
    firstChangeIndex = -1                        '0 index
    Dim gridItem As Variant
    For Each gridItem In unFilteredData.data
        firstChangeIndex = firstChangeIndex + 1
        If this.data.Contains(gridItem) = False Then Exit For
    Next gridItem
    RaiseEvent NeedsAligning(firstChangeIndex)   'off by 1
End Sub

Public Sub Sort(ByVal Comparer As IComparer, ByVal SortMode As lstSortMode, Optional ByVal itemToTrack As Variant)
    this.data.Sort Comparer, SortMode
    If IsMissing(itemToTrack) Or IsNothing(itemToTrack) Then
        RaiseEvent NeedsAligning(alignAll)       'to everything
    Else
        RaiseEvent NeedsAligning(this.data.IndexOf(itemToTrack))
    End If
End Sub

Public Property Get RawData() As FilterList
    Set RawData = this.data.Clone
End Property

