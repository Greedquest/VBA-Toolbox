VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "propertyComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("GridProject.Main.IComparerFactory")
Option Explicit
Implements mscorlib.IComparer
Private propertyName As String

Public Property Let ComparisonProperty(Value As String)
    propertyName = Value
End Property

Private Function IComparer_Compare(ByVal firstItem As Variant, ByVal secondItem As Variant) As Long
    Dim X As ISortable, Y As ISortable
    Set X = firstItem
    Set Y = secondItem
    Dim x1 As Variant
    Dim y1 As Variant
    x1 = X.Properties.itemFromName(propertyName)
    y1 = Y.Properties.itemFromName(propertyName)
    If TypeName(x1) = "String" Then
        IComparer_Compare = StrComp(x1, y1, vbTextCompare)
    Else
        If x1 > y1 Then
            IComparer_Compare = 1
        ElseIf x1 < y1 Then
            IComparer_Compare = -1
        Else
            IComparer_Compare = 0
        End If
    End If
End Function

