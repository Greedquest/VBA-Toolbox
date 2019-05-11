VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Callbacks")

Option Explicit
'@PredeclaredId

Private Type objectCallback
    Object As Object
    ProcName As String
    CallType As VbCallType
    Args As Variant
End Type


Private this As objectCallback

Public Function Create(ByVal Object As Object, ByVal ProcName As String, byvalCallType As VbCallType)
    Set this.Object = Object
    this.ProcName = ProcName
    
End Function
