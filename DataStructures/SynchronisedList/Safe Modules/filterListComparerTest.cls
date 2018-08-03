VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "filterListComparerTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Utils.DummyItems)

Implements ISortable
Private propertySet As New Properties



'Private Function IComparable_CompareTo(ByVal obj As Variant) As Long
'    If obj.value < Me.value Then
'        IComparable_CompareTo = 1
'    ElseIf obj.value > Me.value Then
'        IComparable_CompareTo = -1
'    Else
'        IComparable_CompareTo = 0
'    End If
'End Function

Public Sub printSelf()
Debug.Print Value
End Sub

Public Property Let Value(ByVal Number As Long)
    propertySet.addProperty "value", Number
End Property

Public Property Get Value() As Long
    Value = propertySet.itemFromName("value")
End Property

Private Property Get ISortable_Properties() As Properties
    Set ISortable_Properties = propertySet
End Property

Private Property Get ISortable_SortingProperty() As Variant
    ISortable_SortingProperty = Value
End Property

Private Property Let ISortable_SortingProperty(ByVal SortMode As Variant)
    'do nothing, always sort by value
    If Not UCase(SortMode) = "VALUE" Then
        Err.Description = "filter list test class can only be sorted by value"
        Err.raise 4
    End If
End Property

