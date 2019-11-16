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

Option Explicit
'@PredeclaredId

Private Type pointerData
    address As LongPtr
    length As Long
    type As VbVarType
    depth As Long
End Type
    
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
    this.type = dataType
    this.length = lengthFromType(dataType)
End Sub

Private Function lengthFromType(ByVal dataType As VbVarType) As Long
    Select Case dataType
        Case vbCurrency, vbLongLong
            lengthFromType = 8
        Case vbLong, vbDouble
            lengthFromType = 4
        Case vbSingle, vbInteger
            lengthFromType = 2
        Case vbBoolean, vbByte
            lengthFromType = 1
        Case Else
            Err.Raise 5
    End Select
End Function
