VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EscapeSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Strings.StringFormat")
Option Explicit

Private Type tEscapeSequence
    EscapeString As String
    ReplacementString As String
End Type

Private this As tEscapeSequence

Public Property Get EscapeString() As String
    EscapeString = this.EscapeString
End Property

Friend Property Let EscapeString(ByVal value As String)
    this.EscapeString = value
End Property

Public Property Get ReplacementString() As String
    ReplacementString = this.ReplacementString
End Property

Friend Property Let ReplacementString(ByVal value As String)
    this.ReplacementString = value
End Property

'Lord I wish VB6 had constructors!
Public Function Create(ByVal escape As String, ByVal replacement As String) As EscapeSequence
    Dim result As New EscapeSequence
    result.EscapeString = escape
    result.ReplacementString = replacement
    Set Create = result
End Function

