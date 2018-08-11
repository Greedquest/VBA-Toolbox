VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectedDataWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SynchronisedList")

'''
'Class to hold data currently displayed/ in the grid
'Should be able to trigger display changes to images (make visible or not)
'Trigger re-arrangements of images
'When Items added:
' - These are filtered into the grid, so only need sorting
'When Items ammended:
' - The source undergoes a complete re-filter and outputs a new set for the grid
' - The items are ammended byRef so may have updated their own UI
' - However in case their properties have changed, a re-sorting is required here
'When Items are removed:
' - These need references deleted and to be filtered out
' - Won't change sorting order
'Therefore: Filter, Sort and Arrange all need to be implemented separately


Private WithEvents gridDataBuffer As ListBuffer
Attribute gridDataBuffer.VB_VarHelpID = -1

Public Event NeedsFiltering()
Public Event NeedsAligning(ByVal firstChangeIndex As Long) 'Range to re-align
Public Event NeedsSorting(ByRef itemToTrack As Variant)

'Interface events so that custom onAddition/onRemoval handlers can be used
Public Event ItemsAdded(ByRef itemSet As Variant)
Public Event ItemsRemoved(ByRef itemSet As Variant)

Private Type TSelectedData
    data As FilterList                           'Filtered Data in Grid
    Buffer As ListBuffer
    AlignOnRemoval As Boolean
    SortOnAddition As Boolean
End Type

Private this As TSelectedData
Private Const alignAll = 0

Private Sub Class_Initialize()
    Set this.data = New FilterList
    Set gridDataBuffer = New ListBuffer
    Set this.Buffer = gridDataBuffer
    this.AlignOnRemoval = True
    this.SortOnAddition = True
End Sub

'''
'Buffer Interface
'''

Public Sub AddItems(IterableToAdd As Variant)
    this.Buffer.AddItems IterableToAdd
End Sub

Public Sub RemoveItems(IterableToRemove As Variant)
    this.Buffer.RemoveItems IterableToRemove
End Sub

Public Sub ClearGrid()
    'this.Data.Clear won't change visibility necessarily unless good garbage collection
    this.Buffer.RemoveItems this.data.Clone.data 'this.Data.Clone.Data if removing data breaks for loop
End Sub

'''
'Buffer Events
'''
Private Sub gridDataBuffer_ItemsAdded()
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

Private Sub gridDataBuffer_ItemsRemoved()
    Dim markedItems                              'remove items from grid and mark for removal from source
    markedItems = gridDataBuffer.RemovalItems
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

Public Sub filter(FilterObject As Variant, FilterMode As lstFilterMode, Comparer As IComparer)
    'Filter data references
    Dim unFilteredData As FilterList
    Set unFilteredData = this.data.Clone
    this.data.filter FilterObject, FilterMode, Comparer
    
    Dim firstChangeIndex As Long
    firstChangeIndex = -1                        '0 index
    Dim gridItem As Variant
    For Each gridItem In unFilteredData.data
        firstChangeIndex = firstChangeIndex + 1
        If this.data.Contains(gridItem) = False Then Exit For
    Next gridItem
    RaiseEvent NeedsAligning(firstChangeIndex)   'off by 1
End Sub

Public Sub Sort(ByVal Comparer As IComparer, ByVal SortMode As lstSortMode, Optional ByRef itemToTrack As Variant)
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

