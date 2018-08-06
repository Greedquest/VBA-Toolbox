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
    Dim tableItem As ListRow
    For Each tableItem In dataTable.ListRows
        synchro.Add tableItem.Range
    Next
    
    Set userInterface = New ExampleForm
    Set this.UIForm = userInterface
    
    this.UIForm.DisplayData this.data.SourceData.ToArray
End Sub

