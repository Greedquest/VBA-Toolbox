VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SynchronisedList.Interfaces")
Option Explicit

Public Property Get SourceBufferRemovalTrigger() As Long
End Property

Public Property Let SourceBufferRemovalTrigger(ByVal size As Long)
End Property

Public Property Get SourceBufferAddingTrigger() As Long
End Property

Public Property Let SourceBufferAddingTrigger(ByVal size As Long)
End Property

Public Property Get SourceBufferAmmendingTrigger() As Long
End Property

Public Property Let SourceBufferAmmendingTrigger(ByVal size As Long)
End Property

Public Property Get ContentBufferRemovalTrigger() As Long
End Property

Public Property Let ContentBufferRemovalTrigger(ByVal size As Long)
End Property

Public Property Get ContentBufferAddingTrigger() As Long
End Property

Public Property Let ContentBufferAddingTrigger(ByVal size As Long)
End Property

Public Property Get ContentBufferAmmendingTrigger() As Long
End Property

Public Property Let ContentBufferAmmendingTrigger(ByVal size As Long)
End Property

