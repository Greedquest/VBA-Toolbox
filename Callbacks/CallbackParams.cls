VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallbackParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Callbacks")

Option Explicit
'@PredeclaredId
'@Exposed

Public object As Object
Public procName As String
Public callType As VbCallType
Public args As Variant

Public Function Create(ByVal object As Object, ByVal procName As String, ByVal callType As VbCallType) As CallbackParams
    With New CallbackParams
        Set .object = object
        .procName = procName
        .callType = callType
        .args = args
        Set Create = .Self
    End With
End Function

Public Property Get Self() As CallbackParams
    Set Self = Me
End Property
