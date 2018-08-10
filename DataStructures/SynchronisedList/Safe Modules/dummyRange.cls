VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dummyRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CodeReview")
Option Explicit

Private cellVals As Object 'contains dummy cell data

Public Property Get Cells(ByVal rowNum As Long, ByVal colNum As Long) As String
    Cells = cellVals.item(getKey(rowNum, colNum))
End Property

Public Property Let Cells(ByVal rowNum As Long, ByVal colNum As Long, ByVal newVal As String)
    cellVals.item(getKey(rowNum, colNum)) = newVal
End Property

Private Sub Class_Initialize()
    Set cellVals = CreateObject("Scripting.Dictionary")
End Sub
Private Function getKey(ByVal r As Long, ByVal c As Long) As String
    getKey = "item" & r & "_" & c
End Function
