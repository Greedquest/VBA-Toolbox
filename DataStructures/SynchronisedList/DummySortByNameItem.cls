VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DummySortByNameItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("GridProject.Tests.TestClasses")
Public Name As String
Public Number As Long
Public Time As Double
Public CreationDate As Date


Private Sub Class_Initialize()
    Name = RandomString
    Number = Application.WorksheetFunction.RandBetween(10, 100)
    Time = Now                                   'ordered
    CreationDate = GetRndDate(Now - 500, Now + 500)
End Sub

