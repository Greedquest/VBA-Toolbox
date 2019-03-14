VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleForm 
   Caption         =   "Example App"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7968
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

Sub populateSortBox(ByVal options As Variant)
    Me.SortBy.list = doubleTranspose(options)
End Sub

Sub populateFilterBox(ByVal options As Variant)
    Me.FilterBy.list = doubleTranspose(options)
End Sub

Public Sub DisplayData(ByRef dataArray As Variant)
    If IsArray(dataArray) And ArraySupport.NumberOfArrayDimensions(dataArray) = 1 Then
        dataDisplayBox.list = doubleTranspose(dataArray)
    Else
        Err.Raise 5
    End If
End Sub

Public Sub RemoveItem(ByVal itemIndex As Long)
    dataDisplayBox.RemoveItem itemIndex
End Sub

Public Sub AddItem(itemArray As Variant)

    If IsArray(itemArray) Then                   'assume 1 indexed
        Dim transposedArray
        transposedArray = doubleTranspose(itemArray)
        With dataDisplayBox
            .AddItem
            Dim i As Long
            For i = 0 To .ColumnCount - 1
                .list(.listCount - 1, i) = transposedArray(i + 1)
            Next
        End With
    Else
        Err.Raise 5
    End If
End Sub

Public Sub ClearFromIndex(startingIndex As Long)
    Dim i As Long
    Dim listCount As Long
    listCount = dataDisplayBox.listCount
    'nothing to clear if first change > end of list 0 indexed
    If listCount = startingIndex Then Exit Sub
    For i = listCount - 1 To startingIndex Step -1 'count backwards
        RemoveItem i
    Next
End Sub

Private Function doubleTranspose(ByVal arrayToTranspose As Variant) As Variant
    doubleTranspose = WorksheetFunction.Transpose(WorksheetFunction.Transpose(arrayToTranspose))
End Function

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

Private Sub SortButton_Click()
    RaiseEvent sortModeSet(Me.SortBy.Value)
End Sub

Private Sub FilterButton_Click()
    RaiseEvent filterModeSet(Me.FilterBy.Value, Me.FilterValue.Value)
End Sub


