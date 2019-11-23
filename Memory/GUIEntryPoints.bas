Attribute VB_Name = "GUIEntryPoints"
'@Folder("GUI")
Option Explicit

Public Sub displayRandomPointer()
    Dim testObject As New Collection
    
    With MemoryGridGUI.Create(Pointer.Create(ObjPtr(testObject), vblongptr, 3))
        .Show False
    End With
End Sub
