VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MemoryDisplayBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("GUI")
'@PredeclaredID
Option Explicit

Private Type displayBoxData
    guiElement As msforms.ListBox
End Type

Private this As displayBoxData

Public Function Create(ByVal parentForm As MemoryGridGUI, ByRef settings As displayBoxSettings, Optional ByVal parentLeft As Long = 0) As MemoryDisplayBox
    With New MemoryDisplayBox
        .Init parentForm, settings, parentLeft
        Set Create = .Self
    End With
End Function

Friend Property Get Self() As MemoryDisplayBox
    Set Self = Me
End Property

Public Sub Init(ByVal parentForm As MemoryGridGUI, ByRef settings As displayBoxSettings, ByVal leftCoord As Long)
    Set this.guiElement = parentForm.Controls.Add("Forms.ListBox.1")
    With this.guiElement
        .top = settings.topCoord
        .Left = leftCoord + settings.gap + IIf(leftCoord = 0, 0, settings.width)
        .height = settings.height
        .width = settings.width
        .ColumnCount = 2
    End With
End Sub

Public Property Get guiElement() As msforms.ListBox
    Set guiElement = this.guiElement
End Property

Private Sub AddEntry(ByVal newEntry As Pointer, ByVal addBelow As Boolean)
    Dim itemRow As Long
    If addBelow Then
        this.guiElement.AddItem newEntry.address
        itemRow = this.guiElement.ListCount - 1
    Else
        this.guiElement.AddItem newEntry.address, 0
        itemRow = 0
    End If
    this.guiElement.List(itemRow, 1) = newEntry.Value
    
End Sub

Public Sub display(ByVal content As Pointer, ByVal numberOfNeighbours As Long)
    Dim iteratorPointer As Pointer
    Set iteratorPointer = content
    
    Dim i As Long
    For i = -numberOfNeighbours To numberOfNeighbours
        If i < 0 Then
            Set iteratorPointer = iteratorPointer.PreviousSibling
            AddEntry iteratorPointer, False
        ElseIf i = 0 Then
            Set iteratorPointer = content
            AddEntry iteratorPointer, True
        Else
            Set iteratorPointer = iteratorPointer.NextSibling
            AddEntry iteratorPointer, True
        End If
    Next i
    
    this.guiElement.Selected(numberOfNeighbours) = True
End Sub

