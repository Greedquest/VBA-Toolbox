VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DummyGridItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("GridProject.Tests.TestClasses")
'Implements IGridable
Implements ISortable
Public Name As String
Private Type TDummyItem
    PropertiesObj As Properties
End Type

Private this As TDummyItem


'Private Sub IGridable_ReAlign()
'    Err.raise 5                                  'TODO implement interface member
'End Sub
'
'Private Property Set IGridable_parentUF(ByVal Value As MSForms.UserForm)
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Let IGridable_Visible(ByVal Value As Boolean)
'Debug.Print "Warning, IGridable property used"
'    'Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Get IGridable_Visible() As Boolean
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Set IGridable_EventsObject(ByRef Value As clsGridEvents)
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Get IGridable_ItemWidth() As Long
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Let IGridable_ItemWidth(ByVal Value As Long)
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Get IGridable_ItemHeight() As Long
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Let IGridable_ItemHeight(ByVal Value As Long)
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Get IGridable_LeftCoordinate() As Long
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Let IGridable_LeftCoordinate(ByVal Value As Long)
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Let IGridable_LeftOffset(ByVal Value As Long)
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Get IGridable_TopCoordinate() As Long
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Let IGridable_TopCoordinate(ByVal Value As Long)
'    Err.raise 5                                  'TODO implement interface member
'End Property
'
'Private Property Let IGridable_TopOffset(ByVal Value As Long)
'    Err.raise 5                                  'TODO implement interface member
'End Property

Private Property Let ISortable_SortingProperty(ByVal Value As Variant)
    Err.raise 5                                  'TODO implement interface member
End Property

Private Property Get ISortable_SortingProperty() As Variant
    Err.raise 5                                  'TODO implement interface member
End Property

Private Property Get ISortable_Properties() As Properties
    If this.PropertiesObj Is Nothing Then Set this.PropertiesObj = New Properties
    Set ISortable_Properties = this.PropertiesObj
End Property

