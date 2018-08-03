VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISortable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("GridProject.Interfaces")
Option Explicit
'''
'OPTIONAL Interface used to set property which is used for comparing/filtering
' - Also return a list of possible properties to sort by
' - Necessary for grid filtering/sorting but not otherwise essential for grid display
'''

Public Property Let SortingProperty(ByVal Value As Variant)
End Property

Public Property Get SortingProperty() As Variant
End Property

Public Property Get Properties() As Properties   'group of properties to compare against
End Property

