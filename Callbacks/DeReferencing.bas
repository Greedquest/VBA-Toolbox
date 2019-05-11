Attribute VB_Name = "DeReferencing"
'@Folder("Callbacks")
Option Explicit

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                         ByVal lpPrevWndFunc As Long, _
                         ByVal HWnd As Long, _
                         ByVal msg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long

Public Sub objectCallbackDeReference(ByVal ptr As Long, ByRef result As objectCallback, Optional ByVal offset As Long = 0)
    CallWindowProc AddressOf objectCallbackProc, ptr - offset, VarPtr(result), 0, 0
End Sub

Private Sub objectCallbackProc(ByRef deReferencedType As objectCallback, ByRef result As objectCallback, ByVal unused1 As Long, ByVal unused2 As Long)
    result = deReferencedType
End Sub


