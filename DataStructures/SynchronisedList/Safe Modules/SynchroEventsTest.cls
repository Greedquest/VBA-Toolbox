VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SynchroEventsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.TestEventHandlers")
Public WithEvents synchro As SynchronisedList
Attribute synchro.VB_VarHelpID = -1
Private Type tSynchroTest
    OrderEventRaised As Long
    LastChangeIndex As Long
    PropertiesEventRaised As Long
    ChangeIndecies As Collection
End Type

Private this As tSynchroTest


Public Sub ClearCounts()
    this.LastChangeIndex = -1
    this.OrderEventRaised = 0
    this.PropertiesEventRaised = 0
    Set this.ChangeIndecies = New Collection
End Sub

Private Sub synchro_OrderChanged(ByVal firstChangeIndex As Long)
    this.OrderEventRaised = this.OrderEventRaised + 1
    this.LastChangeIndex = firstChangeIndex
    this.ChangeIndecies.Add firstChangeIndex
End Sub

Private Sub Synchro_PropertiesChanged()
    this.PropertiesEventRaised = this.PropertiesEventRaised + 1
End Sub

Public Property Get OrderEventRaised() As Long
    OrderEventRaised = this.OrderEventRaised
End Property

Public Property Get LastChangeIndex() As Long
    LastChangeIndex = this.LastChangeIndex
End Property

Public Property Get PropertiesEventRaised() As Long
    PropertiesEventRaised = this.PropertiesEventRaised
End Property

Public Property Get ChangeIndecies() As Collection
    Set ChangeIndecies = this.ChangeIndecies
End Property

Private Sub Class_Initialize()
    ClearCounts
End Sub

