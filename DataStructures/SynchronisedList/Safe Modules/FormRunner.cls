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

Public Sub init(ByVal dataTable As ListObject)
    Dim tableItem As ListRow
    For Each tableItem In dataTable.ListRows
        synchro.Add toArray(tableItem)
    Next
End Sub

Private Function toArray(tableItem As ListRow) As Variant
    toArray = WorksheetFunction.Transpose(tableItem.Range)
End Function
