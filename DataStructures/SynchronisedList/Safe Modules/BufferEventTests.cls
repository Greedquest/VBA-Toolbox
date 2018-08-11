VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BufferEventTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.TestEventHandlers")
Public WithEvents Buffer As ListBuffer
Attribute Buffer.VB_VarHelpID = -1
Private Type tBufferTest
    AddedEventRaised As Long
    AmmendedEventRaised As Long
    RemovedEventRaised As Long
End Type

Private this As tBufferTest

Private Sub Buffer_ItemsAdded()
    this.AddedEventRaised = this.AddedEventRaised + 1
End Sub

Private Sub Buffer_ItemsAmmended()
    this.AmmendedEventRaised = this.AmmendedEventRaised + 1
End Sub

Private Sub Buffer_ItemsRemoved()
    this.RemovedEventRaised = this.RemovedEventRaised + 1
End Sub

Public Sub ClearCounts()
    this.AddedEventRaised = 0
    this.AmmendedEventRaised = 0
    this.RemovedEventRaised = 0
End Sub

Public Property Get AddedEventRaised() As Long
    AddedEventRaised = this.AddedEventRaised
End Property

Public Property Get AmmendedEventRaised() As Long
    AmmendedEventRaised = this.AmmendedEventRaised
End Property

Public Property Get RemovedEventRaised() As Long
    RemovedEventRaised = this.RemovedEventRaised
End Property

Private Sub Class_Initialize()
    ClearCounts
End Sub

