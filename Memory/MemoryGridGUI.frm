VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MemoryGridGUI 
   Caption         =   "UserForm1"
   ClientHeight    =   5664
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   11112
   OleObjectBlob   =   "MemoryGridGUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MemoryGridGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("GUI")
Option Explicit


Private Type gridGUIData
    displayBoxSettings As displayBoxSettings
    displayBoxes As Collection
End Type

Private this As gridGUIData

Private Sub UserForm_Initialize()
    With this.displayBoxSettings
        .gap = TemplateListBox.Left
        .height = TemplateListBox.height
        .topCoord = TemplateListBox.top
        .width = TemplateListBox.width
    End With
    
'    TemplateListBox.Visible = True
'    TemplateListBox.AddItem
'    TemplateListBox.List(0, 0) = "Item1"
'    TemplateListBox.List(0, 1) = "Item2"
End Sub

Public Function Create(ByVal basePointer As Pointer) As MemoryGridGUI
    With New MemoryGridGUI
        .Init basePointer
        Set Create = .Self
    End With
End Function

Friend Property Get Self() As MemoryGridGUI
    Set Self = Me
End Property

Friend Sub Init(ByVal basePointer As Pointer)
    Set this.displayBoxes = New Collection 'clear all
    this.displayBoxes.Add MemoryDisplayBox.Create(Me, this.displayBoxSettings), "base"
    this.displayBoxes("base").display basePointer, 1
End Sub
