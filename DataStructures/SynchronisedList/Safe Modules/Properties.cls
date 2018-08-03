VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Utils")

'Dummy class to allow property comparisons
Private propertyData As Object                   'dictionary to hold values

Private Sub Class_Initialize()
    Set propertyData = CreateObject("Scripting.Dictionary")
End Sub

Public Sub InitRandom()

    addProperty "AspectRatio", 1366 / 768
    addProperty "FileCreationDate", PropertiesUtils.GetRndDate(Date - 1000, Date)
    addProperty "FileName", PropertiesUtils.RandomString
    'ImageData
    'FilePath
End Sub

Public Property Get PropertyList()               ' As String()
    PropertyList = propertyData.Keys
End Property

Public Function itemFromName(propertyName As String) As Variant
    itemFromName = propertyData(propertyName)
End Function

Public Sub addProperty(propertyName As String, Value As Variant)
    On Error Resume Next
    propertyData.Add propertyName, Value
    If Err.Number = 457 Then                     'key already exists
        On Error GoTo 0
        If IsObject(Value) Then
            Set propertyData(propertyName) = Value
        Else
            propertyData(propertyName) = Value
        End If
    End If
End Sub

