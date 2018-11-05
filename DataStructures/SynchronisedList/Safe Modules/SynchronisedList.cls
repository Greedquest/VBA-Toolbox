VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SynchronisedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SynchronisedList")
Option Explicit

Public Event OrderChanged(ByVal firstChangeIndex As Long) '0 indexed location of first different item
Public Event OnAddition(ByRef itemArray As Variant)
Public Event AfterRemoval(ByRef itemArray As Variant)

Private WithEvents sourceWrapper As SourceDataWrapper
Attribute sourceWrapper.VB_VarHelpID = -1
Private WithEvents contentWrapper As ContentDataWrapper
Attribute contentWrapper.VB_VarHelpID = -1

Private Type TList
    SourceData As SourceDataWrapper
    ContentData As ContentDataWrapper
    FilterSortMethods As FilterRunner
    SortOnAddition As Boolean
End Type

Private Const initialFiltering = lstFilterMode.lstKeepAll 'ensures we can filter without FilterComparer set
Private Const initialSorting = lstSortMode.lstNoSorting 'ensures we can sort without SortComparer set

Private this As TList

Private Sub Class_Initialize()
    'set up source and its events
    Set sourceWrapper = New SourceDataWrapper
    Set this.SourceData = sourceWrapper
    'set up Grid Data and its events
    Set contentWrapper = New ContentDataWrapper
    Set this.ContentData = contentWrapper
    'set up FilterRunner
    Set this.FilterSortMethods = New FilterRunner
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
    this.FilterSortMethods.sortContent this.ContentData, Nothing
    'this.SourceData.ReSortAndFilter 'doubles grid data
End Sub

Public Sub Filter(Optional compareTo As Variant = vbNullString, Optional newComparer As IComparer, Optional overrideFilterMode As lstFilterMode = 0)
    this.FilterSortMethods.SetFilterMode compareTo, newComparer, overrideFilterMode
    this.SourceData.ReSortAndFilter
End Sub

Public Property Get SourceData() As FilterList
    'If this.SourceData Is Nothing Then Set this.SourceData = New clsSourceWrapper
    Set SourceData = this.SourceData.RawData
End Property

Public Property Get ContentData() As FilterList
    'If this.GridData Is Nothing Then Set this.GridData = New clsGridDataWrapper
    Set ContentData = this.ContentData.RawData
End Property

Private Sub contentWrapper_ItemsAdded(itemArray As Variant)
    If this.SortOnAddition Then this.FilterSortMethods.sortContent this.ContentData, itemArray(1)
    RaiseEvent OnAddition(itemArray)
End Sub

Private Sub contentWrapper_ItemsRemoved(itemSet As Variant)
    RaiseEvent AfterRemoval(itemSet)
End Sub

Private Sub contentWrapper_NeedsAligning(ByVal firstChangeIndex As Long)
'what is this blasphemy!!!!
    this.ContentData
    RaiseEvent OrderChanged(firstChangeIndex)
End Sub

Private Sub contentWrapper_NeedsFiltering()
    this.FilterSortMethods.filterContent this.ContentData
End Sub

Private Sub contentWrapper_NeedsSorting(itemToTrack As Variant)
    this.FilterSortMethods.sortContent this.ContentData, itemToTrack
End Sub

Private Sub sourceWrapper_DataAdded()
    'When something flagged as added to source (and has made it through the source filter
    'put it on the grid
    this.ContentData.AddItems this.SourceData.AddedData.data
    'sort and align triggered
End Sub

Private Sub sourceWrapper_DataRemoved(RemovedItems As Variant)
    this.ContentData.RemoveItems RemovedItems
End Sub

Private Sub sourceWrapper_NeedsSortingAndFiltering(PartialFilterSet As Variant)
    'Empty partial set means complete resort/filter
    If IsNothing(PartialFilterSet) Then this.FilterSortMethods.filterContent this.ContentData
    'the next will trigger a resort when new data is added
    this.FilterSortMethods.SortAndFilterSourceToOutput this.SourceData, PartialFilterSet
End Sub

