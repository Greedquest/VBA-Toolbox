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
Private WithEvents synchro As SynchronizedList
Attribute synchro.VB_VarHelpID = -1

Private Type tRunner
    headerRange As Range
    UIForm As ExampleForm
    data As SynchronizedList
End Type
Private this As tRunner


Private Sub Class_Initialize()
    Set synchro = New SynchronizedList
    Set this.data = synchro
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


Private Sub synchro_OrderChanged(ByVal firstChangeIndex As Long)
    this.UIForm.ClearFromIndex firstChangeIndex + 1 'listbox is 1 indexed
    Dim i As Long
    For i = firstChangeIndex To this.data.GridData.Count - 1 '0 indexed
        this.UIForm.AddItem this.data.GridData(i)
    Next
    this.UIForm.Repaint
    DoEvents
End Sub
