VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabelTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'@Folder Stopwatch

Public Enum stopwatchLableType
    stp_LapTime = 1
    stp_Label
    stp_Start
    stp_Finish
End Enum

Private Type TLabelTree
    parentNode As LabelTree
    ChildNodes As Collection
    NodeName As String
    TimeSpent As Double
    TimeWasted As Double                         'time used by stopwatch runs
    Location As String
    LabelType As stopwatchLableType
End Type

Private this As TLabelTree
Public Property Get LabelType() As stopwatchLableType
    LabelType = this.LabelType
End Property

Public Property Let LabelType(ByVal value As stopwatchLableType)
    this.LabelType = value
End Property

Public Property Get Location() As String
    Location = this.Location
End Property

Public Property Let Location(ByVal value As String)
    this.Location = value
End Property

Public Property Get TimeSpent() As Double
    TimeSpent = this.TimeSpent
End Property

Public Property Let TimeSpent(ByVal value As Double)
    this.TimeSpent = value
End Property

Public Property Get TimeWasted() As Double
    TimeWasted = this.TimeWasted
End Property

Public Property Let TimeWasted(ByVal value As Double)
    this.TimeWasted = value
End Property

Public Property Get ChildNodes() As Collection
    Set ChildNodes = this.ChildNodes
End Property

Public Property Set ChildNodes(ByVal value As Collection)
    Set this.ChildNodes = value
End Property

Public Property Get NodeName() As String
    NodeName = this.NodeName
End Property

Public Property Let NodeName(ByVal value As String)
    this.NodeName = value
End Property

Public Property Get parentNode() As LabelTree
    Set parentNode = this.parentNode
End Property

Public Property Set parentNode(ByVal value As LabelTree)
    Set this.parentNode = value
End Property

Private Sub Class_Initialize()
    Set this.ChildNodes = New Collection
End Sub
