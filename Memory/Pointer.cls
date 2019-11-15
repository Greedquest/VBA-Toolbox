VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("VBAProject")
'@PredeclaredId
Option Explicit

Private Type pointerData
    address As LongPtr
    length As Long
End Type

Private this As pointerData

'Constructor methods
'Public Function FromValue(ByVal varPtrToData As LongPtr, ByVal length As Long) As Pointer
'    With New Pointer
'        Set FromValue = .Self
'        .Init varPtrToData, length
'    End With
'End Function

Public Function FromReference(ByVal obj As IUnknown) As Pointer
    With New Pointer
        Set FromReference = .Self
        .Init ObjPtr(obj), LenB(this.address) 'just LongPtr length
    End With
End Function

Public Function FromAddress(ByVal address As LongPtr) As Pointer
    With New Pointer
        Set FromAddress = .Self
        .Init address, LenB(address) 'just LongPtr length
    End With
End Function

Friend Sub Init(ByVal address As LongPtr, ByVal length As Long)
    this.address = address
    this.length = length
End Sub

Friend Property Get Self() As Pointer
    Set Self = Me
End Property

Public Property Get DeRef() As Pointer
    Set DeRef = Pointer.FromAddress(DeReference(this.address))
End Property

'@DefaultMember
Public Property Let Value(ByVal valueBytes As LongPtr)
    Init valueBytes, LenB(valueBytes)
End Property

Public Property Get Value() As LongPtr
    Value = this.address
End Property

'Debugging
Friend Sub debugPrint()
    Debug.Print printf("Pointer @{0} len {1}", this.address, this.length)
End Sub

