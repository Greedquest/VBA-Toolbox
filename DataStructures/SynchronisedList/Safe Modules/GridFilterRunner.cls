VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GridFilterRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder(SynchronisedList.Utils)
Option Explicit

Private Type tFilter
    sortComparer As IComparer
    filterComparer As IComparer
    FilterMode As lstFilterMode
    SortMode As lstSortMode
    filterAgainst As Variant
    FilterModeSet As Boolean                     'check to ensure we don't filter with unset IComparer
    SortModeSet As Boolean
End Type

Private Const defaultFilterMode = lstFilterMode.lstRemoveMatching
Private Const defaultSortMode = lstSortMode.lstSortAscending

Private this As tFilter

Private Sub Class_Initialize()
    Set this.sortComparer = New propertyComparer
    this.FilterMode = defaultFilterMode
    this.SortMode = defaultSortMode
    this.FilterModeSet = False
    this.SortModeSet = False
End Sub

Public Property Get filterComparer() As IComparer
    Set filterComparer = this.filterComparer
End Property

Public Property Set filterComparer(Value As IComparer)
    Set this.filterComparer = Value
    this.FilterModeSet = Not Value Is Nothing    'filter mode set as long as valid IComparer set
End Property

Public Property Get sortComparer() As IComparer
    Set sortComparer = this.sortComparer
End Property

Public Property Set sortComparer(Value As IComparer)
    Set this.sortComparer = Value
    this.SortModeSet = Not Value Is Nothing      'sort mode set as long as valid IComparer set
End Property

Public Sub SetFilterMode(Optional compareTo As Variant = vbNullString, Optional newComparer As IComparer, Optional overrideFilterMode As lstFilterMode = 0)
    'If overrides not set, revert to default mode
    If Not newComparer Is Nothing Then Set filterComparer = newComparer
    If overrideFilterMode = 0 Then
        this.FilterMode = defaultFilterMode
    Else
        this.FilterMode = overrideFilterMode
    End If
    If IsObject(compareTo) Then
        Set this.filterAgainst = compareTo
    Else
        this.filterAgainst = compareTo
    End If
End Sub

Public Sub SetSortMode(Optional newComparer As IComparer, Optional overrideSortMode As lstSortMode = 0)
    
    If Not newComparer Is Nothing Then Set sortComparer = newComparer
    If overrideSortMode = 0 Then
        this.SortMode = defaultSortMode
    Else
        this.SortMode = overrideSortMode
    End If
End Sub

Public Sub SortAndFilterSourceToOutput(source As clsSourceWrapper, Optional PartialSet As Variant)
    If Not isSafeToFilter Then
        Err.Description = "No IComparer has been set for Filtering, please setFilterMode"
        Err.raise 5
    ElseIf Not isSafeToSort Then
        Err.Description = "No IComparer has been set for Sorting, please setSortMode"
        Err.raise 5
    Else
        source.SortAndFilterToOutput this.filterAgainst, this.FilterMode, filterComparer, sortComparer, this.SortMode, PartialSet
    End If
End Sub

Public Sub filterGrid(grid As clsGridDataWrapper)
    If isSafeToFilter Then
        grid.Filter this.filterAgainst, this.FilterMode, filterComparer
    Else
        Err.Description = "No IComparer has been set for filtering, please setFilterMode"
        Err.raise 5
    End If
End Sub

Public Sub sortGrid(grid As clsGridDataWrapper, itemToTrack As Variant)
    grid.Sort sortComparer, this.SortMode, itemToTrack
End Sub

Private Function isSafeToFilter() As Boolean
    isSafeToFilter = this.FilterModeSet Or this.FilterMode = lstRemoveAll Or this.FilterMode = lstKeepAll
End Function

Private Function isSafeToSort() As Boolean
    isSafeToSort = this.SortModeSet Or this.SortMode = lstNoSorting
End Function

