VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SynchronisedList.Interfaces")
Option Explicit

Private ContentBufferAmmendment As Long
Private ContentBufferAdding As Long
Private ContentBufferRemoval As Long
Private SourceBufferAmmendment As Long
Private SourceBufferAdding As Long
Private SourceBufferRemoval As Long

Public Property Get SourceBufferRemovalTrigger() As Long
    SourceBufferRemovalTrigger = SourceBufferRemoval
End Property

Public Property Let SourceBufferRemovalTrigger(ByVal size As Long)
    SourceBufferRemoval = size
End Property



Public Property Get SourceBufferAddingTrigger() As Long
    SourceBufferAddingTrigger = SourceBufferAdding
End Property

Public Property Let SourceBufferAddingTrigger(ByVal size As Long)
    SourceBufferAdding = size
End Property



Public Property Get SourceBufferAmmendingTrigger() As Long
    SourceBufferAmmendingTrigger = SourceBufferAmmendment
End Property

Public Property Let SourceBufferAmmendingTrigger(ByVal size As Long)
    SourceBufferAmmendment = size
End Property



Public Property Get ContentBufferRemovalTrigger() As Long
    ContentBufferRemovalTrigger = ContentBufferRemoval
End Property

Public Property Let ContentBufferRemovalTrigger(ByVal size As Long)
    ContentBufferRemoval = size
End Property



Public Property Get ContentBufferAddingTrigger() As Long
    ContentBufferAddingTrigger = ContentBufferAdding
End Property

Public Property Let ContentBufferAddingTrigger(ByVal size As Long)
    ContentBufferAdding = size
End Property



Public Property Get ContentBufferAmmendmendingTrigger() As Long
    ContentBufferAmmendmendingTrigger = ContentBufferAmmendment
End Property

Public Property Let ContentBufferAmmendmendingTrigger(ByVal size As Long)
    ContentBufferAmmendment = size
End Property

