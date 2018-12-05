VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder(SynchronisedList.Utils)
Option Explicit

Private Type tFilter
    SortComparer As IComparer
    FilterComparer As IComparer
    FilterMode As lstFilterMode
    SortMode As lstSortMode
    filterAgainst As Variant
    filterModeSet As Boolean                     'check to ensure we don't filter with unset IComparer
    sortModeSet As Boolean
End Type

Private Const defaultFilterMode = lstFilterMode.lstRemoveMatching
Private Const defaultSortMode = lstSortMode.lstSortAscending

Private this As tFilter

Private Sub Class_Initialize()
    'Set this.SortComparer = New propertyComparer
    this.FilterMode = defaultFilterMode
    this.SortMode = defaultSortMode
    this.filterModeSet = False
    this.sortModeSet = False
End Sub

Public Property Get FilterComparer() As IComparer
    Set FilterComparer = this.FilterComparer
End Property

Public Property Set FilterComparer(ByVal Value As IComparer)
    Set this.FilterComparer = Value
    this.filterModeSet = Not Value Is Nothing    'filter mode set as long as valid IComparer set
End Property

Public Property Get SortComparer() As IComparer
    Set SortComparer = this.SortComparer
End Property

Public Property Set SortComparer(ByVal Value As IComparer)
    Set this.SortComparer = Value
    this.sortModeSet = Not Value Is Nothing      'sort mode set as long as valid IComparer set
End Property

Public Sub SetFilterMode(Optional ByVal compareTo As Variant = vbNullString, Optional ByVal newComparer As IComparer, Optional ByVal overrideFilterMode As lstFilterMode = 0)
    'NOTE: why vbnullstring? shouldn't this be its own default
    'If overrides not set, revert to default mode
    If Not newComparer Is Nothing Then Set FilterComparer = newComparer
    If overrideFilterMode = 0 Then
        this.FilterMode = defaultFilterMode
    Else
        this.FilterMode = overrideFilterMode
    End If
    LetSet this.filterAgainst, compareTo
End Sub

Public Sub SortInPlace(ByVal list As FilterList)
    If Not isSafeToSort Then
        Err.Description = "No IComparer has been set for Sorting, please setSortMode"
        Err.Raise 5
    Else
        list.Sort this.SortComparer, this.SortMode
    End If
End Sub

Public Sub FilterInPlace(ByVal list As FilterList)
    If Not isSafeToFilter Then
        Err.Description = "No IComparer has been set for Filtering, please setFilterMode"
        Err.Raise 5
    Else
        list.Filter this.filterAgainst, this.FilterMode, this.FilterComparer
    End If
End Sub

Public Sub SetSortMode(Optional ByVal newComparer As IComparer, Optional ByVal overrideSortMode As lstSortMode = 0)
    'If overrides not set, revert to default mode
    If Not newComparer Is Nothing Then Set SortComparer = newComparer
    If overrideSortMode = 0 Then
        this.SortMode = defaultSortMode
    Else
        this.SortMode = overrideSortMode
    End If
End Sub

Public Sub FilterSourceToOutput(source As SourceDataWrapper, Optional PartialSet As Variant)
    If Not isSafeToFilter Then
        Err.Description = "No IComparer has been set for Filtering, please setFilterMode"
        Err.Raise 5
    ElseIf Not isSafeToSort Then
        Err.Description = "No IComparer has been set for Sorting, please setSortMode"
        Err.Raise 5
    Else
        source.FilterToOutput this.filterAgainst, this.FilterMode, FilterComparer, SortComparer, this.SortMode, PartialSet
    End If
End Sub

Public Sub filterContent(content As ContentDataWrapper)
    If isSafeToFilter Then
        content.Filter this.filterAgainst, this.FilterMode, FilterComparer
    Else
        Err.Description = "No IComparer has been set for filtering, please setFilterMode"
        Err.Raise 5
    End If
End Sub

Public Sub sortContent(content As ContentDataWrapper, itemToTrack As Variant)
    content.Sort SortComparer, this.SortMode, itemToTrack
End Sub

Private Function isSafeToFilter() As Boolean
    isSafeToFilter = this.filterModeSet Or this.FilterMode = lstRemoveAll Or this.FilterMode = lstKeepAll
End Function

Private Function isSafeToSort() As Boolean
    isSafeToSort = this.sortModeSet Or this.SortMode = lstNoSorting
End Function

