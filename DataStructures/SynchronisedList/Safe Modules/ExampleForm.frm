VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleForm 
   Caption         =   "Example App"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7965
   OleObjectBlob   =   "ExampleForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExampleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CodeReview")
Option Explicit

Public Event sortModeSet(ByVal SortBy As String)
Public Event filterModeSet(ByVal FilterBy As String, ByVal FilterValue As String)

'GUI
Private boolEnter As Boolean


'
'Public OrdersList As mscorlib.ArrayList
'Private pc As propertyComparer
'
'Private Sub UserForm_Initialize()
'    Dim cell As Range
'    Set OrdersList = New ArrayList
'    Set pc = New propertyComparer
'
'    With Worksheets("Orders")
'        For Each cell In .Range("A2", .Range("A" & .Rows.Count).End(xlUp))
'            OrdersList.Add cell.Resize(1, 8)
'        Next
'
'        For Each cell In .Range("A1").Resize(1, 8)
'            cboSortBy.AddItem cell.Value
'        Next
'
'    End With
'
'    cboSortBy.AddItem "Row"
'
'    FillOrdersListBox
'End Sub
'

'
'Private Sub btnReverse_Click()
'    OrdersList.Reverse
'    FillOrdersListBox
'End Sub
'
'Private Sub cboSortBy_Change()
'    If cboSortBy.ListIndex = -1 Then Exit Sub
'
'    Select Case cboSortBy.ListIndex
'        Case Is < 8
'            pc.Init "Cells", VbGet, 1, cboSortBy.ListIndex + 1
'        Case 8
'            pc.Init "Row", VbGet
'    End Select
'
'    OrdersList.Sort_2 pc
'    FillOrdersListBox
'End Sub

'Form


'Form Control Methods

Public Sub DisplayData(ByRef dataArray As Variant)
    dataDisplayBox.List = WorksheetFunction.Transpose(WorksheetFunction.Transpose(dataArray))
End Sub

Public Sub RemoveItem(ByVal itemIndex As Long)
    With dataDisplayBox
        .List
        .RemoveItem itemIndex
    End With
End Sub

Public Sub AddItem(ByVal itemIndex As Long)
    With dataDisplayBox
        .AddItem itemIndex
    End With
End Sub

'Form GUI

Private Sub FilterValue_Enter()
    boolEnter = True
End Sub

Private Sub FilterValue_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)
    If boolEnter = True Then
        With FilterValue
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        boolEnter = False
    End If
End Sub
