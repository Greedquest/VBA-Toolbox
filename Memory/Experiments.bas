Attribute VB_Name = "Experiments"
'@Folder("VBAProject")
'@IgnoreModule ProcedureNotUsed:Macros
Option Explicit

Private Sub getPointer()
       
'    Dim pointers(1 To 4) As Pointer
'
'    Dim obj As Collection
'    Set obj = New Collection
'    Set pointers(1) = Pointer.FromReference(obj)
'
'    Dim fnPointer As LongPtr
'    fnPointer = VBA.CLngPtr(AddressOf getPointer)
'    Set pointers(2) = Pointer.FromAddress(fnPointer)
'
'    Dim varPointer As LongPtr
'    varPointer = VarPtr(fnPointer)
'    Set pointers(3) = Pointer.FromAddress(varPointer)
'
'    Dim objPointer As LongPtr
'    objPointer = ObjPtr(obj)
'    Set pointers(4) = Pointer.FromAddress(objPointer)
'
''    Dim numericData As Double
''    Set pointers(5) = Pointer.FromValue(VarPtr(numericData), LenB(numericData))
'
'    Dim i As Long
'    For i = LBound(pointers) To UBound(pointers)
'        pointers(i).debugPrint
'    Next i
    
End Sub

Public Sub test()
    Dim a As Long, b As Long
    a = 5
    b = 6

    Dim aPointer As Pointer
    Set aPointer = Pointer.Create(VarPtr(a), VarType(a))

    Dim bPointer As Pointer
    Set bPointer = Pointer.Create(VarPtr(b), VarType(b))

    aPointer.Value = bPointer.Value

    Debug.Assert a = b 'succeeds

End Sub
