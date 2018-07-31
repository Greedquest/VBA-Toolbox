VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A custom colllection class to facilitate sorting and filtering"
Option Explicit

'@Folder(SynchronisedList.Utils.FilterList)

Private dataSet As Object
Public Enum lstFilterMode
    lstKeepMatching = 1
    lstRemoveMatching
    lstKeepAll
    lstRemoveAll
End Enum

Public Enum lstSortMode
    lstSortAscending = 1
    lstSortDescending
    lstNoSorting
End Enum

Private Sub Class_Initialize()
    Set dataSet = CreateObject("System.Collections.ArrayList")
End Sub

Public Property Get Data() As Object
    Set Data = dataSet
End Property

Public Property Set Data(Value As Object)
    Set dataSet = Value
End Property

Public Function Contains(ByVal Value As Variant) As Boolean
    Contains = dataSet.Contains(Value)
End Function

Public Function Clone() As FilterList
    Dim result As New FilterList
    Set result.Data = Me.Data.Clone
    Set Clone = result
End Function

Public Sub Add(ByVal Value As Variant)
    dataSet.Add Value
End Sub

Public Sub Remove(ByVal Value As Variant)
    dataSet.Remove Value
End Sub

Public Sub Sort(ByVal Comparer As IComparer, ByVal SortMode As lstSortMode)
    Select Case SortMode
    Case lstSortAscending
        dataSet.Sort_2 Comparer
    Case lstSortDescending
        dataSet.Sort_2 Comparer
        Reverse
    Case lstNoSorting
        'do nothing
    Case Else
        Err.Description = "Invalid filter mode"
        Err.raise 5
    End Select
End Sub

Public Sub Clear()
    dataSet.Clear
End Sub

Public Sub Reverse()
    dataSet.Reverse
End Sub

Public Property Get IndexOf(ByVal Value As Variant) As Long
    IndexOf = dataSet.IndexOf_3(Value)           '0 indexed
End Property

Public Property Get Count() As Long
    Count = dataSet.Count
End Property

Public Sub Filter(ByVal testValue As Variant, ByVal FilterMode As lstFilterMode, ByVal Comparer As IComparer)
    Dim i As Long
    Dim item As Variant
    Select Case FilterMode
    Case lstRemoveMatching
        For i = dataSet.Count - 1 To 0 Step -1
            LetSet item, dataSet(i)
            If Comparer.Compare(testValue, item) = 0 Then Me.Remove item
        Next i
    Case lstKeepMatching
        For i = dataSet.Count - 1 To 0 Step -1
            LetSet item, dataSet(i)
            If Comparer.Compare(testValue, item) <> 0 Then Me.Remove item
        Next i
    Case lstRemoveAll
        Me.Clear
    Case lstKeepAll
        'do nothing
    Case Else
        Err.Description = "Invalid filter mode"
        Err.raise 5
    End Select
End Sub

Private Sub Class_Terminate()
    Set dataSet = Nothing
End Sub

