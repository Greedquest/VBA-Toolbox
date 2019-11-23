Attribute VB_Name = "Experiments"
'@Folder("VBAProject")
'@IgnoreModule ProcedureNotUsed:Macros
Option Explicit

Private Sub getPointer()
    
    Dim testObj As New Pointer
    Dim inspectedObject As Pointer
    
    Set inspectedObject = Pointer.Create(ObjPtr(testObj), vblongptr, 5)
    
    
    
End Sub

Public Sub test()
    Dim a As Double, b As Double
    a = 5.11
    b = 6.22

    Dim apPointer As Pointer, aPointer As Pointer
    Set apPointer = Pointer.Create(VarPtr(VarPtr(a)), varType(a), 2)
    
    Set aPointer = apPointer.DeRef

    Dim bPointer As Pointer
    Set bPointer = Pointer.Create(VarPtr(b), varType(b))
    
    Debug.Print "BEFORE"
    Debug.Print "&&a:";: apPointer.DebugPrint
    Debug.Print "&a:";: aPointer.DebugPrint
    Debug.Print "&b:";: bPointer.DebugPrint
    
    aPointer.Value = bPointer.Value
    'apPointer.DeRef.Value = bPointer.Value

    Debug.Print "AFTER"
    Debug.Print "&&a:";: apPointer.DebugPrint
    Debug.Print "&a:";: aPointer.DebugPrint
    Debug.Print "&b:";: bPointer.DebugPrint
    
    Debug.Print "a: "; a, "b: "; b

End Sub


Sub testValueLet()

    Debug.Print String(30, "_")
    Debug.Print String(30, "-")
    
    Dim a As Double
    a = 3.14159
    
    Dim pA As Pointer
    Set pA = Pointer.Create(VarPtr(a), varType(a))
    
    Debug.Print "a: "; a
    pA.DebugPrint
    
    pA.Value = 2.73
    
    Debug.Print "a: "; a
    pA.DebugPrint
    
    Debug.Print String(30, "-")
    
    Dim ppA As Pointer
    Set ppA = Pointer.Create(VarPtr(VarPtr(a)), varType(a), 2)
    
    Debug.Print "a: "; a
    ppA.DebugPrint
    
    ppA.DeRef.Value = 99.999
    
    Debug.Print "a: "; a
    ppA.DeRef.DebugPrint
    
End Sub

Sub testObjectExploration()
    Dim someObj As Collection
    Set someObj = New Collection
    
    Dim ppVtable As Pointer
    Set ppVtable = Pointer.Create(ObjPtr(someObj), vblongptr, 3)
    
    Dim vtableFirst As Pointer
    Set vtableFirst = Pointer.Create(ppVtable.DeRef.Value, vbLong)
    
    ppVtable.DebugPrint
    ppVtable.DeRef.DebugPrint
    ppVtable.DeRef.DeRef.DebugPrint
    vtableFirst.DebugPrint
End Sub



