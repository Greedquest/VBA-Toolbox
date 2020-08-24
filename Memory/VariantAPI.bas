Attribute VB_Name = "VariantAPI"
'@Folder("VBAProject")
Option Explicit

Public Declare Function InitVariantFromBuffer Lib "propsys.dll" (ByRef data As Any, ByVal length As Long, ByRef outVar As Variant) As HRESULT


Sub test()
Dim data As Long
data = &HFFFFFFFF

Dim b As Variant
Dim hr As HRESULT
hr = InitVariantFromBuffer(data, LenB(data), b)
If hr <> S_OK Then Stop
End Sub
