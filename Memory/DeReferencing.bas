Attribute VB_Name = "DeReferencing"
'@Folder("VBAProject")
Option Explicit

Private Enum BOOL
    API_FALSE = 0
    'Use NOT (result = API_FALSE) for API_TRUE, as TRUE is just non-zero
End Enum

Public Enum HRESULT
    S_OK = &H0                                   'Success.
    DISP_E_BADVARTYPE = &H8                      'The variant type is not a valid type of variant.
    DISP_E_OVERFLOW = &HA                        'The data pointed to by pvarSrc does not fit in the destination type.
    DISP_E_TYPEMISMATCH = &H5                    'The argument could not be coerced to the specified type.
    E_INVALIDARG = &H57                          'One of the arguments is not valid.
    E_OUTOFMEMORY = &HE                          'Insufficient memory to complete the operation.
End Enum

Private Enum VirtualProtectFlags                 'See Memory Protection constants: https://docs.microsoft.com/en-gb/windows/win32/memory/memory-protection-constants
    PAGE_EXECUTE_READWRITE = &H40
    PAGE_READONLY = &H2
    RESET_TO_PREVIOUS = -1
End Enum

Public Enum LCIDflags
    LOCALE_INVARIANT = &H7F
End Enum

Public Type VARIANT_STRUCT
    varType As Integer
    reserved(1 To 3) As DByte
    data As OByte
End Type

#If Win64 Then                                   'To decide whether to use 8 or 4 bytes per chunk of memory
    Private Declare PtrSafe Function GetMem Lib "msvbvm60" Alias "GetMem8" (ByRef source As Any, ByRef destination As Any) As Long
#Else
    Private Declare PtrSafe Function GetMem Lib "msvbvm60" Alias "GetMem4" (ByRef source As Any, ByRef destination As Any) As Long
#End If

Private Declare Function VariantChangeTypeEx Lib "oleaut32" (ByRef destination As Any, ByRef source As Any, ByVal lcid As LCIDflags, ByVal wFlags As Integer, ByVal varintType As Integer) As HRESULT

Private Declare PtrSafe Sub GetMem1 Lib "msvbvm60" (source As Any, destination As Any)
Private Declare PtrSafe Sub GetMem2 Lib "msvbvm60" (source As Any, destination As Any)
Private Declare PtrSafe Sub GetMem4 Lib "msvbvm60" (source As Any, destination As Any)
Private Declare PtrSafe Sub GetMem8 Lib "msvbvm60" (source As Any, destination As Any)
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)

Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (ByRef location As Any, ByVal numberOfBytes As Long, ByVal newProtectionFlags As VirtualProtectFlags, ByVal lpOldProtectionFlags As LongPtr) As BOOL

'@Description("Pointer dereferencing; reads/ writes a single 4 byte (32-bit) or 8 byte (64-bit) block of memory at the address specified. Performs any necessary unprotecting")
Public Property Let DeReference(ByVal address As LongPtr, ByVal Value As LongPtr)
    If ToggleMemoryProtection(address, LenB(Value), PAGE_EXECUTE_READWRITE) Then
        GetMem Value, ByVal address
        ToggleMemoryProtection address, LenB(Value)
    Else
        Err.Raise 5, Description:="That address is protected memory which cannot be accessed"
    End If
End Property

Public Property Get DeReference(ByVal address As LongPtr) As LongPtr
    If ToggleMemoryProtection(address, LenB(DeReference), PAGE_EXECUTE_READWRITE) Then
        GetMem ByVal address, DeReference
        ToggleMemoryProtection address, LenB(DeReference)
    Else
        Err.Raise 5, Description:="That address is protected memory which cannot be accessed"
    End If
End Property

'@Description("Read/write data of a certain type to and from corresponding variants")
Public Property Get ValueAt(ByVal address As LongPtr, ByVal dataType As VbVarType) As Variant
    Dim result As VARIANT_STRUCT
    If ToggleMemoryProtection(address, LenB(result.data), PAGE_EXECUTE_READWRITE) Then
        GetMem8 ByVal address, result.data       'read all the data - vartype will control what actually read
        ToggleMemoryProtection address, LenB(result.data)
    Else
        Err.Raise 5, Description:="That address is protected memory which cannot be accessed"
    End If
    
    result.varType = dataType
    MoveMemory ValueAt, result, LenB(result)
    
End Property

Public Property Let ValueAt(ByVal address As LongPtr, ByVal dataType As VbVarType, ByVal newValue As Variant)
    Dim typedValue As VARIANT_STRUCT
    If VariantChangeTypeEx(typedValue, newValue, LOCALE_INVARIANT, 0, dataType) <> S_OK Then
        Err.Raise 5, "Variant issues"
    End If
    
    'move the appropriate number of bytes from the data portion of the variant to the output variable
    If ToggleMemoryProtection(address, lengthFromType(dataType), PAGE_EXECUTE_READWRITE) Then
        MoveMemory ByVal address, typedValue.data, lengthFromType(dataType)
        ToggleMemoryProtection address, lengthFromType(dataType)
    Else
        Err.Raise 5, Description:="That address is protected memory which cannot be accessed"
    End If
    
End Property

Sub t()
    Dim someData As Long                         'could be a Double, Single, Byte etc. This works for any type except reference types (arrays, strings, objects)
    someData = &H1234ABCD                        'some number that fits in a Long
    
    Dim dataPointer As LongPtr
    dataPointer = VarPtr(someData)
    
    Dim dereferencedData As Variant
    dereferencedData = ValueAt(dataPointer, vbLong) 'interpret the data as a Long; vbLong = 3
    
    'dereferencedData now looks like 03 00 | 00 00 00 00 00 00 | 12 34 AB CD 00 00 00 00
    Debug.Assert dereferencedData = someData     'Passes
End Sub

Sub testReadWrite()
    Debug.Print String(80, "-")
    
    Const data As Double = 31.4159
    Dim testValue As Long
    Dim testVariant As Variant
    testValue = data
    
    testVariant = ValueAt(VarPtr(testValue), varType(testValue))
    
    Debug.Print "Raw: ";
    printVarInfo CVar(testValue)
    Debug.Print "Get: ";
    printVarInfo testVariant
    
    ValueAt(VarPtr(testValue), varType(testValue)) = data - 1
    Debug.Print "Let: ";
    printVarInfo CVar(testValue)
    
    
    
End Sub

Private Sub printVarInfo(ByRef var As Variant)
    Debug.Print Left$(TypeName(var), 5); varType(var), IIf(IsObject(var), "Objet", var), Hex$(VarPtr(var)), variantStructHex(VarPtr(var))
End Sub

Sub testVariousTypes()
    Debug.Print String(80, "-")
    Dim a As Single
    Dim b As Double
    Dim c As Long
    Dim d As Byte
    Dim e As String
    Dim f As New Collection
    
    Const data As Double = 31.4159
    
    a = data
    b = data
    c = data
    d = data
    e = data
    f.Add data
    
    Debug.Print TypeName(a); varType(a), a, Hex$(VarPtr(a)), HexCode(VarPtr(a), LenB(a))
    Debug.Print TypeName(b); varType(b), b, Hex$(VarPtr(b)), HexCode(VarPtr(b), LenB(b))
    Debug.Print TypeName(c); varType(c), c, Hex$(VarPtr(c)), HexCode(VarPtr(c), LenB(c))
    Debug.Print TypeName(d); varType(d), d, Hex$(VarPtr(d)), HexCode(VarPtr(d), LenB(d))
    Debug.Print TypeName(e); varType(e), e, Hex$(VarPtr(e)), HexCode(VarPtr(e), LenB(e)), Hex$(StrPtr(e))
    Debug.Print TypeName(f); varType(f), f(1), Hex$(VarPtr(f)), HexCode(VarPtr(f), 4), Hex$(ObjPtr(f))
    
    Dim var As Variant
    VariantChangeTypeEx var, CVar(a), LOCALE_INVARIANT, 0, varType(a)
    printVarInfo var
    
    VariantChangeTypeEx var, CVar(b), LOCALE_INVARIANT, 0, varType(b)
    printVarInfo var
    
    VariantChangeTypeEx var, CVar(c), LOCALE_INVARIANT, 0, varType(c)
    printVarInfo var
    
    VariantChangeTypeEx var, CVar(d), LOCALE_INVARIANT, 0, varType(d)
    printVarInfo var
    
    VariantChangeTypeEx var, CVar(e), LOCALE_INVARIANT, 0, varType(e)
    printVarInfo var
    
    VariantChangeTypeEx var, CVar(f), LOCALE_INVARIANT, 0, varType(f)
    printVarInfo var
    
End Sub

Public Property Get variantStructHex(ByVal address As LongPtr) As String
    Dim result As String
    Dim b
    For Each b In AsOByteArr(ByVal address).bytes
        result = result & WorksheetFunction.Dec2Hex(b, 2)
    Next b
    result = result & " | "
    For Each b In AsOByteArr(ByVal address + 8).bytes
        result = result & WorksheetFunction.Dec2Hex(b, 2)
    Next b
    variantStructHex = result
End Property

Public Property Get HexCode(ByVal address As LongPtr, ByVal length As Long) As String
    Dim bytes() As Byte
    ReDim bytes(1 To length)
    MoveMemory bytes(1), ByVal address, length
    Dim i As Long
    Dim result As String
    For i = LBound(bytes) To UBound(bytes)
        result = result & WorksheetFunction.Dec2Hex(bytes(i), 2)
    Next i
    HexCode = result
End Property

Private Static Function ToggleMemoryProtection(ByVal address As LongPtr, ByVal numberOfBytes As Long, Optional ByVal newMemoryFlag As VirtualProtectFlags = RESET_TO_PREVIOUS) As Boolean
    Dim previousMemoryState As VirtualProtectFlags
    Dim unprotectWasNOOP As Boolean
    If newMemoryFlag = RESET_TO_PREVIOUS Then
        If unprotectWasNOOP Then
            ToggleMemoryProtection = True
        Else
            ToggleMemoryProtection = VirtualProtect(ByVal address, numberOfBytes, previousMemoryState, VarPtr(newMemoryFlag)) <> API_FALSE
        End If
    Else
        ToggleMemoryProtection = VirtualProtect(ByVal address, numberOfBytes, newMemoryFlag, VarPtr(previousMemoryState)) <> API_FALSE
        'check whether unprotecting even had an effect - if not then no need to toggle back
        unprotectWasNOOP = newMemoryFlag = previousMemoryState
    End If
End Function

'@Description("Converts native types to their length in bytes - also accepts vbLongPtr")
Public Function lengthFromType(ByVal dataType As VbVarType) As Long
    Select Case dataType
        Case vbCurrency, vbLongLong, vbDouble
            lengthFromType = 8
        Case vbLong
            lengthFromType = 4
        Case vbSingle, vbInteger
            lengthFromType = 2
        Case vbBoolean, vbByte
            lengthFromType = 1
        Case Else
            Err.Raise 5, Description:="Unexpected dataType with unknown length"
    End Select
End Function

