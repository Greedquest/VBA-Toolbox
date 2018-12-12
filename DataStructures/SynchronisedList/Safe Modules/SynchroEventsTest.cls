VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SynchroEventsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Tests.TestEventHandlers")
Public WithEvents synchro As SynchronisedList
Attribute synchro.VB_VarHelpID = -1

Private Type tSynchroTest
    OrderEvents As FilterList
    AdditionEvents As FilterList
    RemovalEvents As FilterList
End Type

Private this As tSynchroTest

Public Sub ClearCounts()
    Set this.OrderEvents = New FilterList
    Set this.RemovalEvents = New FilterList
    Set this.AdditionEvents = New FilterList
End Sub

Private Sub Class_Initialize()
    ClearCounts
End Sub

Private Sub synchro_AfterRemoval(itemArray As Variant)
    this.RemovalEvents.Add itemArray
End Sub

Private Sub synchro_OnAddition(itemArray As Variant)
    this.AdditionEvents.Add itemArray
End Sub

Private Sub synchro_OrderChanged(ByVal firstChangeIndex As Long)
    this.OrderEvents.Add firstChangeIndex
End Sub

Public Property Get AdditionEvents() As FilterList
    Set AdditionEvents = this.AdditionEvents
End Property

Public Property Get RemovalEvents() As FilterList
    Set RemovalEvents = this.RemovalEvents
End Property

Public Property Get OrderEvents() As FilterList
    Set OrderEvents = this.OrderEvents
End Property
