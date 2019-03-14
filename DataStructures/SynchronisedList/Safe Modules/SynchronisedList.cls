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

'TODO rename these to smt like on show on hide, since this is not removing/adding overall, just to the content
Public Event OrderChanged(ByVal firstChangeIndex As Long) '0 indexed location of first different item
Public Event OnAddition(ByRef itemArray As Variant) 'when stuff added to content wrapper
Public Event AfterRemoval(ByRef itemArray As Variant) 'when stuff removed from content

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents sourceWrapper As SourceDataWrapper
Attribute sourceWrapper.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents contentWrapper As ContentDataWrapper
Attribute contentWrapper.VB_VarHelpID = -1

Private Type TList
    SourceData As SourceDataWrapper
    ContentData As ContentDataWrapper
    FilterSortMethods As FilterRunner
    SortOnAddition As Boolean
End Type

Private Const initialFiltering As Long = lstFilterMode.lstKeepAll 'ensures we can filter without FilterComparer set
Private Const initialSorting As Long = lstSortMode.lstNoSorting 'ensures we can sort without SortComparer set

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
    this.FilterSortMethods.SetFilterMode overrideFilterMode:=initialFiltering
    this.FilterSortMethods.SetSortMode overrideSortMode:=initialSorting
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

Public Sub Sort(Optional newComparer As IComparer, Optional overrideSortMode As lstSortMode = lstUndefinedSorting)
    this.FilterSortMethods.SetSortMode newComparer, overrideSortMode
    this.FilterSortMethods.sortContent this.ContentData, Nothing
    'this.SourceData.ReSortAndFilter 'TODO what does this do, it breaks no unit tests whether there or not!'doubles grid data
End Sub

Public Sub Filter(Optional compareTo As Variant = vbNullString, Optional newComparer As IComparer, Optional overrideFilterMode As lstFilterMode = lstUndefinedFiltering)
    this.FilterSortMethods.SetFilterMode compareTo, newComparer, overrideFilterMode
    this.SourceData.PushToOutput
End Sub

Public Property Get SourceData() As FilterList
    'If this.SourceData Is Nothing Then Set this.SourceData = New clsSourceWrapper
    Set SourceData = this.SourceData.RawData
End Property

Public Property Get ContentData() As FilterList
    'If this.GridData Is Nothing Then Set this.GridData = New clsGridDataWrapper
    Set ContentData = this.ContentData.RawData
End Property

Private Sub contentWrapper_BufferEvent(ByVal mode As BufferMode, ByVal bufferItems As FilterList)
    'fresh from buffer, sort and filter in place
    Select Case mode
    Case slAdding
        this.FilterSortMethods.FilterInPlace bufferItems
        If bufferItems.Count > 0 Then
            this.FilterSortMethods.SortInPlace bufferItems
        End If
    Case slRemoving
        If bufferItems.Count > 0 Then
            this.FilterSortMethods.SortInPlace bufferItems
        End If
    Case Else
        Err.Raise 5
    End Select
End Sub

Private Sub contentWrapper_ItemsAdded(ByVal itemSet As Variant)
    
    'If this.SortOnAddition Then this.FilterSortMethods.sortContent this.ContentData, itemArray(1)
    'For now at least, this is controlled with events from the data
    'Which means we need to pass info on behaviour down
    'But that's what we're doing for the buffers anyway, so no problem
    RaiseEvent OnAddition(itemSet) 'BUG this is not necessarily a 1 indexed array
End Sub

Private Sub contentWrapper_ItemsRemoved(ByVal itemSet As Variant)
    RaiseEvent AfterRemoval(itemSet)
End Sub

Private Sub contentWrapper_NeedsAligning(ByVal firstChangeIndex As Long)

    RaiseEvent OrderChanged(firstChangeIndex)

    '    'wehenever something happens that triggers a re-align, we clear source from that point
    '    'and push through the content. This way the buffer can do its job of managing the rate of
    '    'stuff going to the UI
    '    this.ContentData.ClearFromIndex firstChangeIndex
    '    RaiseEvent OrderChanged(firstChangeIndex)
    '    this.SourceData.ReSortAndFilter
    
End Sub

Private Sub contentWrapper_NeedsFiltering()
    this.FilterSortMethods.filterContent this.ContentData
End Sub

Private Sub contentWrapper_NeedsSorting(ByVal itemToTrack As Variant)
    this.FilterSortMethods.sortContent this.ContentData, itemToTrack
End Sub

Private Sub sourceWrapper_DataAdded()
    'When something flagged as added to source (and has made it through the source filter
    'put it on the grid
    'The onAddition event is raised only by content, not source
    this.ContentData.AddItems this.SourceData.AddedData.data
    'sort and align triggered
End Sub

Private Sub sourceWrapper_DataRemoved(ByVal RemovedItems As Variant)
    this.ContentData.RemoveItems RemovedItems
End Sub

Private Sub sourceWrapper_NeedsPushingToOutput(PartialFilterSet As Variant)
    'Empty partial set means complete resort/filter - filter mode is changed
    'in which case we also want to filter the content first, then
    If IsNothing(PartialFilterSet) Then this.FilterSortMethods.filterContent this.ContentData
    'the next will trigger a resort when new data is added
    this.FilterSortMethods.FilterSourceToOutput this.SourceData, PartialFilterSet
End Sub


'''
'Settings
'''
Public Property Get Settings() As IListSettings
    Set Settings = Me
End Property

Private Property Get IListSettings_SourceBufferRemovalTrigger() As Long
    IListSettings_SourceBufferRemovalTrigger = this.SourceData.Buffer.RemovingTrigger
End Property

Private Property Let IListSettings_SourceBufferRemovalTrigger(ByVal size As Long)
    this.SourceData.Buffer.RemovingTrigger = size
End Property

Private Property Get IListSettings_SourceBufferAddingTrigger() As Long
    IListSettings_SourceBufferAddingTrigger = this.SourceData.Buffer.AddingTrigger
End Property

Private Property Let IListSettings_SourceBufferAddingTrigger(ByVal size As Long)
    this.SourceData.Buffer.AddingTrigger = size
End Property

Private Property Get IListSettings_SourceBufferAmmendingTrigger() As Long
    IListSettings_SourceBufferAmmendingTrigger = this.SourceData.Buffer.AmmendingTrigger
End Property

Private Property Let IListSettings_SourceBufferAmmendingTrigger(ByVal size As Long)
    this.SourceData.Buffer.AmmendingTrigger = size
End Property

Private Property Get IListSettings_ContentBufferRemovalTrigger() As Long
    IListSettings_ContentBufferRemovalTrigger = this.ContentData.Buffer.RemovingTrigger
End Property

Private Property Let IListSettings_ContentBufferRemovalTrigger(ByVal size As Long)
    this.ContentData.Buffer.RemovingTrigger = size
End Property

Private Property Get IListSettings_ContentBufferAddingTrigger() As Long
    IListSettings_ContentBufferAddingTrigger = this.ContentData.Buffer.AddingTrigger
End Property

Private Property Let IListSettings_ContentBufferAddingTrigger(ByVal size As Long)
    this.ContentData.Buffer.AddingTrigger = size
End Property

Private Property Get IListSettings_ContentBufferAmmendingTrigger() As Long
    IListSettings_ContentBufferAmmendingTrigger = this.ContentData.Buffer.AmmendingTrigger
End Property

Private Property Let IListSettings_ContentBufferAmmendingTrigger(ByVal size As Long)
    this.ContentData.Buffer.AmmendingTrigger = size
End Property
