VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SynchronizedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SynchronisedList")
Option Explicit

Public Event OrderChanged(ByVal firstChangeIndex As Long) '0 indexed location of first different item
Public Event OnAddition(ByRef itemArray As Variant)
Public Event AfterRemoval(ByRef itemArray As Variant)

Private WithEvents sourceDataWrapper As clsSourceWrapper
Attribute sourceDataWrapper.VB_VarHelpID = -1
Private WithEvents gridDataWrapper As clsGridDataWrapper
Attribute gridDataWrapper.VB_VarHelpID = -1

Private Type TList
    SourceData As clsSourceWrapper
    GridData As clsGridDataWrapper
    FilterSortMethods As GridFilterRunner
    SortOnAddition As Boolean
End Type

Private Const initialFiltering = lstFilterMode.lstKeepAll 'ensures we can filter without FilterComparer set
Private Const initialSorting = lstSortMode.lstNoSorting 'ensures we can sort without SortComparer set

Private this As TList

Private Sub Class_Initialize()
    'set up source and its events
    Set sourceDataWrapper = New clsSourceWrapper
    Set this.SourceData = sourceDataWrapper
    'set up Grid Data and its events
    Set gridDataWrapper = New clsGridDataWrapper
    Set this.GridData = gridDataWrapper
    'set up FilterRunner
    Set this.FilterSortMethods = New GridFilterRunner
    this.FilterSortMethods.SetFilterMode , , initialFiltering
    this.FilterSortMethods.SetSortMode , initialSorting
End Sub

Public Sub Add(ParamArray gridItems())
    'adds any iterables to the source
    If UBound(gridItems) < 0 Then Err.Raise 5    'check for non empty array
    this.SourceData.AddItems flattenParamArray(gridItems)
End Sub

Public Sub Remove(ParamArray gridItems())
    'removes any iterables from the source
    If UBound(gridItems) < 0 Then Err.Raise 5    'check for non empty array
    this.SourceData.RemoveItems flattenParamArray(gridItems) 'should auto trigger removal in grid
End Sub

Public Sub MarkAsAmmended(ParamArray gridItems())
    'Ammends items by removing them from the grid and re-adding from the source
    If UBound(gridItems) < 0 Then Err.Raise 5    'check for non empty array
    this.SourceData.AmmendItems flattenParamArray(gridItems) 'should auto trigger removal in grid
End Sub

Public Sub Sort(Optional newComparer As IComparer, Optional overrideSortMode As lstSortMode = 0)
    this.FilterSortMethods.SetSortMode newComparer, overrideSortMode
    this.FilterSortMethods.sortGrid this.GridData, Nothing
    'this.SourceData.ReSortAndFilter 'doubles grid data
End Sub

Public Sub filter(Optional compareTo As Variant = vbNullString, Optional newComparer As IComparer, Optional overrideFilterMode As lstFilterMode = 0)
    this.FilterSortMethods.SetFilterMode compareTo, newComparer, overrideFilterMode
    this.SourceData.ReSortAndFilter
End Sub

Public Property Get SourceData() As FilterList
    'If this.SourceData Is Nothing Then Set this.SourceData = New clsSourceWrapper
    Set SourceData = this.SourceData.RawData
End Property

Public Property Get GridData() As FilterList
    'If this.GridData Is Nothing Then Set this.GridData = New clsGridDataWrapper
    Set GridData = this.GridData.RawData
End Property

Private Sub gridDataWrapper_ItemsAdded(itemArray As Variant)
    If this.SortOnAddition Then this.FilterSortMethods.sortGrid this.GridData, itemArray(1)
    RaiseEvent OnAddition(itemArray)
End Sub

Private Sub gridDataWrapper_ItemsRemoved(itemSet As Variant)
    RaiseEvent AfterRemoval(itemSet)
End Sub

Private Sub gridDataWrapper_NeedsAligning(ByVal firstChangeIndex As Long)
    RaiseEvent OrderChanged(firstChangeIndex)
End Sub

Private Sub gridDataWrapper_NeedsFiltering()
    this.FilterSortMethods.filterGrid this.GridData
End Sub

Private Sub gridDataWrapper_NeedsSorting(itemToTrack As Variant)
    this.FilterSortMethods.sortGrid this.GridData, itemToTrack
End Sub

Private Sub sourceDataWrapper_DataAdded()
    'When something flagged as added to source (and has made it through the source filter
    'put it on the grid
    this.GridData.AddItems this.SourceData.AddedData.Data
    'sort and align triggered
End Sub

Private Sub sourceDataWrapper_DataRemoved(RemovedItems As Variant)
    this.GridData.RemoveItems RemovedItems
End Sub

Private Sub sourceDataWrapper_NeedsSortingAndFiltering(PartialFilterSet As Variant)
    'Empty partial set means complete resort/filter
    If IsNothing(PartialFilterSet) Then this.FilterSortMethods.filterGrid this.GridData
    'the next will trigger a resort when new data is added
    this.FilterSortMethods.SortAndFilterSourceToOutput this.SourceData, PartialFilterSet
End Sub

