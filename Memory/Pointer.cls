VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProject")

Option Explicit
'@PredeclaredId

Private Type pointerData
    address As LongPtr
    baseType As VbVarType
    depth As Long
End Type

Public Enum customVBVarTypes
    #If Win64 Then
        vblongptr = vbLongLong
    #Else
        vbLongLong = 20
        vblongptr = vbLong
    #End If
End Enum

Private this As pointerData

Public Function Create(ByVal address As LongPtr, ByVal dataType As VbVarType, Optional ByVal referenceDepth As Long = 1) As Pointer
    With New Pointer
        Set Create = .Self
        .Init address, dataType, referenceDepth
    End With
End Function

Friend Property Get Self() As Pointer
    Set Self = Me
End Property

Friend Sub Init(ByVal address As LongPtr, ByVal dataType As VbVarType, ByVal referenceDepth As Long)
    this.address = address
    this.depth = referenceDepth
    this.baseType = dataType
End Sub

Public Property Get DeRef() As Pointer
    If this.depth = 1 Then
        'this is a pointer to a variable, not another pointer - so don't dereference or it's riskkkkkyy
        Err.Raise 5, Description:="Pointer can't be dereferenced any further; either set a larger depth in the constructor, or use the Pointer.Value property to read the contents of the variable"
    Else
        Set DeRef = Pointer.Create(DeReference(this.address), this.baseType, this.depth - 1)
    End If
End Property

Public Property Get Value() As Variant
    Value = ValueAt(this.address, valueType)
End Property

Public Property Let Value(ByVal newValue As Variant)
    ValueAt(this.address, valueType) = newValue
End Property

Public Property Set Value(ByVal reference As IUnknown) 'REVIEW: is this weird?
    ValueAt(this.address, valueType) = ObjPtr(reference)
End Property

Public Property Get address() As LongPtr
    address = this.address
End Property

Public Property Get NextSibling() As Pointer
    Set NextSibling = Pointer.Create(this.address + lengthFromType(valueType), this.baseType, this.depth)
End Property

Public Property Get PreviousSibling() As Pointer
    Set PreviousSibling = Pointer.Create(this.address - lengthFromType(valueType), this.baseType, this.depth)
End Property

'@Description("The VarType of this pointer when dereferenced; if this is &a then datatype will be vartype(a), if &&a then datatype will be a pointer)
Private Property Get valueType() As VbVarType
    If this.depth = 1 Then                       'just a pointer to a value, so that value is this class' base type
        valueType = this.baseType
    Else
        valueType = vblongptr
    End If
End Property

Friend Sub DebugPrint()
    Debug.Print printf("Pointer @{3}{4}baseType: {0}{4}depth: {1}{4}value: {2}", this.baseType, this.depth, Value, this.address, vbCrLf & vbTab)
End Sub

