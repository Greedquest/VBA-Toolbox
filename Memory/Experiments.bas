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
    Dim a As Double, b As Double
    a = 5.11
    b = 6.22

    Dim apPointer As Pointer, aPointer As Pointer
    Set apPointer = Pointer.Create(VarPtr(VarPtr(a)), VarType(a), 2)
    
    Set aPointer = apPointer.DeRef

    Dim bPointer As Pointer
    Set bPointer = Pointer.Create(VarPtr(b), VarType(b))
    
    Debug.Print "&&a:", ;: apPointer.DebugPrint
    Debug.Print "&a:", ;: aPointer.DebugPrint
    Debug.Print "&b:", ;: bPointer.DebugPrint
    
    aPointer.Value = bPointer.Value
    'apPointer.DeRef.Value = bPointer.Value

    Debug.Print "&&a:", ;: apPointer.DebugPrint
    Debug.Print "&a:", ;: aPointer.DebugPrint
    Debug.Print "&b:", ;: bPointer.DebugPrint
    
    Debug.Print "a: "; a, "b: "; b

End Sub


Sub testValueLet()

    Debug.Print String(30, "_")
    Debug.Print String(30, "-")
    
    Dim a As Double
    a = &HAABBCCDD
    
    Dim pA As Pointer
    Set pA = Pointer.Create(VarPtr(a), VarType(a))
    
    Debug.Print "a: "; a
    pA.DebugPrint
    
    pA.Value = 2.73
    
    Debug.Print "a: "; a
    pA.DebugPrint
    
    Debug.Print String(30, "-")
    
    Dim ppA As Pointer
    Set ppA = Pointer.Create(VarPtr(VarPtr(a)), VarType(a), 2)
    
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
    Set ppVtable = Pointer.Create(ObjPtr(someObj), vbLongPtr, 3)
    
    Dim vtableFirst As Pointer
    Set vtableFirst = Pointer.Create(ppVtable.DeRef.Value, vbLong)
    
    ppVtable.DebugPrint
    ppVtable.DeRef.DebugPrint
    ppVtable.DeRef.DeRef.DebugPrint
    vtableFirst.DebugPrint
End Sub

Sub inspectVariant()
    Dim a As Double
    a = 1.7
    
    Dim aVar As Variant
    aVar = a
    
    Debug.Print TypeName(a), VarType(a), VariantType(a)
    
    Dim variantData() As Byte
    ReDim variantData(1 To 16)
    
    Dim sourceData() As Byte
    ReDim sourceData(1 To LenB(a))
    
    Debug.Print ArrPtr(variantData), VarPtr(variantData(1))
    
    
    MoveMemory ByVal VarPtr(sourceData(1)), a, UBound(sourceData)
    MoveMemory ByVal VarPtr(variantData(1)), aVar, UBound(variantData)
    
    Dim b As Currency
    Dim bVar As Variant
    
    VariantChangeTypeEx bVar, aVar, LOCALE_INVARIANT, 0, vbCurrency
    
    b = bVar
    
    
End Sub

