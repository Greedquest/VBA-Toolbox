VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bufferTestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("GridProject.Tests.TestClasses")
Private WithEvents testBuffer As clsBuffer
Attribute testBuffer.VB_VarHelpID = -1
Private itemsRemaining As Long
Implements IComparable
Public comparableVal As Double

Public Sub TestEvents()
    Set testBuffer = New clsBuffer
    With testBuffer
        .AddingTrigger = 2
        itemsRemaining = 5
        .AddItems [a1:a5]
    End With
End Sub

Private Sub Class_Initialize()
    comparableVal = Rnd
End Sub

Private Function IComparable_CompareTo(ByVal obj As Variant) As Long
    If Me.comparableVal < obj.comparableVal Then
        IComparable_CompareTo = -1
    ElseIf Me.comparableVal > obj.comparableVal Then
        IComparable_CompareTo = 1
    Else
        IComparable_CompareTo = 0
    End If
End Function

Private Sub testBuffer_ItemsAdded()
Debug.Assert UBound(testBuffer.AddedItems) = Application.WorksheetFunction.min(testBuffer.AddingTrigger, itemsRemaining)
    itemsRemaining = itemsRemaining - 2
End Sub

