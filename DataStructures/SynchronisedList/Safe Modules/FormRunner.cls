VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CodeReview")
Option Explicit

Private WithEvents userInterface As ExampleForm
Attribute userInterface.VB_VarHelpID = -1
Private WithEvents synchro As SynchronisedList
Attribute synchro.VB_VarHelpID = -1

Private Type tRunner
    headerRange As Range
    UIForm As ExampleForm
    data As SynchronisedList
    SortComparer As CallByNameComparer
    FilterComparer As CallByNameComparer
    filterObj As dummyRange
End Type

Private this As tRunner

Private Sub Class_Initialize()

    Set synchro = New SynchronisedList
    Set this.data = synchro
    Set this.SortComparer = New CallByNameComparer
    Set this.FilterComparer = New CallByNameComparer
    Set this.filterObj = New dummyRange
End Sub

Public Sub init(ByVal dataTable As ListObject)
    Set this.headerRange = dataTable.HeaderRowRange
    
    Set userInterface = New ExampleForm
    Set this.UIForm = userInterface
    this.UIForm.dataDisplayBox.ColumnCount = this.headerRange.Cells.Count 'set number of columns
    
    this.UIForm.populateFilterBox this.headerRange
    this.UIForm.populateSortBox this.headerRange
    
    'show form and start adding data to it
    this.UIForm.Show False
    synchro.Add dataTable.DataBodyRange.Rows
    

End Sub

Private Sub synchro_AfterRemoval(itemArray As Variant)
    Dim i As Long
    'would be better to use order changed to clear then re-add
    For i = LBound(itemArray) To UBound(itemArray)
        'already removed from content data
        this.UIForm.RemoveItem this.data.SourceData.IndexOf(itemArray(i))
    Next i
End Sub

'Private Sub synchro_OrderChanged(ByVal firstChangeIndex As Long)
''this is no good! The list list should only raise incremental changes from the buffer,
''Instead we build up the list nicely but re-order all at once
'    this.UIForm.ClearFromIndex firstChangeIndex 'listbox is 0 indexed too
'    Dim i As Long
'    For i = firstChangeIndex To this.data.ContentData.Count - 1 '0 indexed
'        this.UIForm.AddItem this.data.ContentData(i)
'    Next
'    this.UIForm.Repaint
'    DoEvents
'End Sub

Private Sub synchro_OnAddition(itemArray As Variant)
    Dim i As Long
    For i = LBound(itemArray) To UBound(itemArray)
        this.UIForm.AddItem itemArray(i)
    Next i
End Sub

Private Sub userInterface_filterModeSet(ByVal FilterBy As String, ByVal FilterValue As String)
    this.FilterComparer.init "Cells", VbGet, 1, columnIndexFromName(FilterBy)
    this.filterObj.Cells(1, columnIndexFromName(FilterBy)) = FilterValue
    this.data.Filter this.filterObj, this.FilterComparer
End Sub

Private Sub userInterface_sortModeSet(ByVal SortBy As String)
    this.SortComparer.init "Cells", VbGet, 1, columnIndexFromName(SortBy)
    this.data.Sort this.SortComparer
End Sub

Private Function columnIndexFromName(ByVal colName As String) As Long
    Const EXACT_MATCH As Long = 0
    Dim result
    On Error Resume Next
    result = WorksheetFunction.Match(colName, this.headerRange, EXACT_MATCH)
    columnIndexFromName = IIf(Err.Number = 0, result, 1)
End Function

