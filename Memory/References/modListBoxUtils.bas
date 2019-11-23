Attribute VB_Name = "modListBoxUtils"
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modListBoxUtils
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
'                  www.cpearson.com/Excel/ListBoxUtils.aspx
'
' This module provides several utility functions for MSFORMS
' ListBox controls.
' These procedures do not work on mutliple column list boxes.
' The ColumnCount property must be 1.
'
' Functions In This Module:
' -------------------------
' LBXInvertSelection            Inverts the selected items. Selected
'                               items are unselected, and unselected
'                               items are selected.
'
' LBXIsSelectionContiguous      Returns True if all selected items
'                               are contiguous, False otherwise.
'
' LBXMoveToTop                  Moves the selected items to the
'                               top of the list.
'
' LBXMoveUp                     Moves the selected items upwards
'                               in the list.
'
' LBXMoveDown                   Moves the selected items downwards
'                               in the list.
'
' LBXMoveToEnd                  Moves the selected items to the
'                               end of the list.
'
' LBXSort                       Sorts a list box. Requires the
'                               modQSortInPlace module. Sorts
'                               either the entire list box
'                               or a subset of the items.
'
' LBXSelectAllItems             Selects all items in the list box.
'
' LBXSelectionInfo              Populates ByRef variables with
'                               the number of selected items, the
'                               index number of the first selected
'                               item, and the index number of the
'                               last selected item.
'
' LBXSelectedItems              returns an array of Strings, each
'                               of which is a selected item in the
'                               list box. If no items are selected
'                               returns an unallocated array. Calling
'                               procedures should first call
'                               LBXSelectionInfo to determine if any
'                               items are selected.
'
' LBXUnSelectAllItems           Unselects all items in the list.
'
' LBXUnSelectIfNotContiguous    Unselects an item in the list box
'                               if it is not contiguous with the
'                               other selected items.
'
' LBXSelectCount                Returns the number of selected items.
'
' LBXSwapItems                  Swaps the positions of two items in
'                               the list box.
'
' LBXSelectedIndexes            Returns an array of Longs each of
'                               which is the index of a selected item.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function LBXSelectCount(LBX As MSForms.ListBox) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXSelectCount
' Returns the number of selected items in LBX.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim N As Long
Dim C As Long
With LBX
    For N = 0 To .ListCount - 1
        If .Selected(N) = True Then
            C = C + 1
        End If
    Next N
End With
LBXSelectCount = C
End Function

Public Function LBXSelectedIndexes(LBX As MSForms.ListBox) As Long()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXSelectedIndexes
' This returns an array of Longs that are the index numbers of
' the selected items.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim L() As Long
Dim C As Long
Dim N As Long
If LBXSelectCount(LBX) = 0 Then
    LBXSelectedIndexes = L
    Exit Function
End If
With LBX
    ReDim L(0 To .ListCount - 1)
    C = -1
    For N = 0 To .ListCount - 1
        If .Selected(N) Then
            C = C + 1
            L(C) = N
        End If
    Next N
End With
ReDim Preserve L(0 To C)
LBXSelectedIndexes = L

End Function



Public Sub LBXSwapItems(LBX As MSForms.ListBox, N1 As Long, N2 As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXSwapItems
' Swaps the items at positions N1 and N2.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim S1 As String
Dim S2 As String

If N1 < 0 Or N2 < 0 Then
    Exit Sub
End If
If N1 = N2 Then
    Exit Sub
End If
With LBX
    If N1 >= .ListCount Or N2 >= .ListCount Then
        Exit Sub
    End If
    S1 = .List(N1)
    S2 = .List(N2)
    .RemoveItem N1
    .AddItem S2, N1
    .RemoveItem N2
    .AddItem S1, N2
    .Selected(N1) = True
    .Selected(N2) = True
    .ListIndex = N2
End With

End Sub

Public Sub LBXInvertSelection(LBX As MSForms.ListBox)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXInvertSelection
' Inverts selected items. Selected items are unselected and selected
' items are unselected.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
With LBX
    ''''''''''''''''''''''''
    ' If list is empty, get
    ' out now.
    ''''''''''''''''''''''''
    If .ListCount = 0 Then
        Exit Sub
    End If
    For Ndx = 0 To .ListCount - 1
        .Selected(Ndx) = Not .Selected(Ndx)
    Next Ndx
End With

End Sub

Public Sub LBXMoveToTop(LBX As MSForms.ListBox)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXMoveToTop
' This moves the selected items to the top of the list box.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As String          ' temporary variable to hold value for swap
Dim Ndx As Long             ' index counter for LBX.List
Dim SelNdx As Long          ' index of selected items
Dim SelCount As Long        ' number of selected items
Dim FirstSelItem As Long    ' first selected item index
Dim LastSelItem As Long     ' last selected item index


''''''''''''''''''''''''
' If list is empty, get
' out now.
''''''''''''''''''''''''
If LBX.ListCount = 0 Then
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''
' Ensure the selected items
' are contiguous (no unselected rows
' within selected rows).
''''''''''''''''''''''''''''''''''''
If LBXIsSelectionContiguous(LBX:=LBX) = False Then
    Exit Sub
End If

With LBX
    If .ColumnCount > 1 Then
        ''''''''''''''''''''''''''
        ' No support for mutliple
        ' column listboxes.
        ''''''''''''''''''''''''''
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''
    ' If the list is empty, there
    ' is nothing to do. Get out.
    ''''''''''''''''''''''''''''''
    If .ListCount = 0 Then
        Exit Sub
    End If
    '''''''''''''''''''''''''''''
    ' If there is no selected
    ' item, there is nothing to
    ' do. Get Out.
    '''''''''''''''''''''''''''''
    If .ListIndex < 0 Then
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get information about the selected items
    ' in the list box LBX.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    LBXSelectionInfo LBX:=LBX, SelectedCount:=SelCount, _
        FirstSelectedItemIndex:=FirstSelItem, LastSelectedItemIndex:=LastSelItem
    
    ''''''''''''''''''''''''''''''''''
    ' If nothing is selected, get out.
    ''''''''''''''''''''''''''''''''''
    If SelCount = 0 Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''
    ' If no items are selected, get out.
    ' This should be picked up in
    ' SelCount = 0, but we test here for
    ' completeness.
    ''''''''''''''''''''''''''''''''''''
    If (FirstSelItem < 0) Or (LastSelItem < 0) Then
        Exit Sub
    End If
    
    SelNdx = 0
    Ndx = 0
    '''''''''''''''''''''''''''''''''''''''''
    ' Move the items up.
    '''''''''''''''''''''''''''''''''''''''''
    For SelNdx = FirstSelItem To LastSelItem
        Temp = .List(SelNdx)
        .RemoveItem SelNdx
        .AddItem Temp, Ndx
        Ndx = Ndx + 1
    Next SelNdx

    ''''''''''''''''''''''''''''''''
    ' Now, reselect the moved items.
    ''''''''''''''''''''''''''''''''
    LBXUnSelectAllItems LBX:=LBX
    For Ndx = 0 To (LastSelItem - FirstSelItem)
        .Selected(Ndx) = True
    Next Ndx

End With

End Sub


Public Sub LBXMoveUp(LBX As MSForms.ListBox)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXMoveUp
' This moves the selected items up one position in the list.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As String          ' temporary variable to hold value for swap
Dim Ndx As Long             ' index counter for LBX.List
Dim SelNdx As Long          ' index of selected items
Dim SelCount As Long        ' number of selected items
Dim FirstSelItem As Long    ' first selected item index
Dim LastSelItem As Long     ' last selected item index
Dim SaveNdx As Long         ' saved index to reselect items

''''''''''''''''''''''''
' If list is empty, get
' out now.
''''''''''''''''''''''''
If LBX.ListCount = 0 Then
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''
' Ensure the selected items
' are contiguous (no unselected rows
' within selected rows).
''''''''''''''''''''''''''''''''''''
If LBXIsSelectionContiguous(LBX:=LBX) = False Then
    Exit Sub
End If

SaveNdx = -1
With LBX
    
    If .ColumnCount > 1 Then
        ''''''''''''''''''''''''''
        ' No support for mutliple
        ' column listboxes.
        ''''''''''''''''''''''''''
        Exit Sub
    End If
    
    
    ''''''''''''''''''''''''''''''
    ' If the list is empty, there
    ' is nothing to do. Get out.
    ''''''''''''''''''''''''''''''
    If .ListCount = 0 Then
        Exit Sub
    End If
    '''''''''''''''''''''''''''''
    ' If there is no selected
    ' item, there is nothing to
    ' do. Get Out.
    '''''''''''''''''''''''''''''
    If .ListIndex < 0 Then
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get information about the selected items
    ' in the list box LBX.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    LBXSelectionInfo LBX:=LBX, SelectedCount:=SelCount, _
        FirstSelectedItemIndex:=FirstSelItem, LastSelectedItemIndex:=LastSelItem
    
    ''''''''''''''''''''''''''''''''''
    ' If nothing is selected, get out.
    ''''''''''''''''''''''''''''''''''
    If SelCount = 0 Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''
    ' If no items are selected, get out.
    ' This should be picked up in
    ' SelCount = 0, but we test here for
    ' completeness.
    ''''''''''''''''''''''''''''''''''''
    If (FirstSelItem < 0) Or (LastSelItem < 0) Then
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''
    ' If the first selected item is the
    ' first item in the list, get out.
    ''''''''''''''''''''''''''''''''''''
    If FirstSelItem = 0 Then
        Exit Sub
    End If
    
    SelNdx = 0
    Ndx = 0
    For SelNdx = FirstSelItem To LastSelItem
        Temp = .List(SelNdx)
        .RemoveItem SelNdx
        .AddItem Temp, SelNdx - 1
        If SaveNdx < 0 Then
            SaveNdx = SelNdx
        End If

    Next SelNdx
    
    LBXUnSelectAllItems LBX:=LBX
    For Ndx = SaveNdx - 1 To SaveNdx + (LastSelItem - FirstSelItem - 1)
        .Selected(Ndx) = True
    Next Ndx

End With


End Sub


Public Sub LBXMoveDown(LBX As MSForms.ListBox)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXMoveDown
' This move the selected items down one position in the list.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As String          ' temporary variable to hold value for swap
Dim Ndx As Long             ' index counter for LBX.List
Dim SelNdx As Long          ' index of selected items
Dim SelCount As Long        ' number of selected items
Dim FirstSelItem As Long    ' first selected item index
Dim LastSelItem As Long     ' last selected item index
Dim SaveNdx As Long         ' saved index to reselect items

''''''''''''''''''''''''
' If list is empty, get
' out now.
''''''''''''''''''''''''
If LBX.ListCount = 0 Then
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''
' Ensure the selected items
' are contiguous (no unselected rows
' within selected rows).
''''''''''''''''''''''''''''''''''''
If LBXIsSelectionContiguous(LBX:=LBX) = False Then
    Exit Sub
End If

With LBX
    If .ColumnCount > 1 Then
        ''''''''''''''''''''''''''
        ' No support for mutliple
        ' column listboxes.
        ''''''''''''''''''''''''''
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''
    ' If the list is empty, there
    ' is nothing to do. Get out.
    ''''''''''''''''''''''''''''''
    If .ListCount = 0 Then
        Exit Sub
    End If
    '''''''''''''''''''''''''''''
    ' If there is no selected
    ' item, there is nothing to
    ' do. Get Out.
    '''''''''''''''''''''''''''''
    If .ListIndex < 0 Then
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get information about the selected items
    ' in the list box LBX.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    LBXSelectionInfo LBX:=LBX, SelectedCount:=SelCount, _
        FirstSelectedItemIndex:=FirstSelItem, LastSelectedItemIndex:=LastSelItem
    
    ''''''''''''''''''''''''''''''''''
    ' If nothing is selected, get out.
    ''''''''''''''''''''''''''''''''''
    If SelCount = 0 Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''
    ' If no items are selected, get out.
    ' This should be picked up in
    ' SelCount = 0, but we test here for
    ' completeness.
    ''''''''''''''''''''''''''''''''''''
    If (FirstSelItem < 0) Or (LastSelItem < 0) Then
        Exit Sub
    End If
    
    SelNdx = 0
    Ndx = 0
    For SelNdx = LastSelItem To FirstSelItem Step -1
        If LastSelItem = .ListCount - 1 Then
            Exit Sub
        End If
        If SelNdx = .ListCount - 1 Then
            SaveNdx = SelNdx
            Exit For
        End If
        If LastSelItem = .ListCount - 1 Then
            Exit For
        End If
        
        Temp = .List(SelNdx)
        .RemoveItem SelNdx
        .AddItem Temp, SelNdx + 1
        SaveNdx = SelNdx + 1
        
    Next SelNdx

    LBXUnSelectAllItems LBX:=LBX
    For Ndx = SaveNdx To SaveNdx + (LastSelItem - FirstSelItem)
        .Selected(Ndx) = True
    Next Ndx

End With


End Sub


Public Sub LBXMoveToEnd(LBX As MSForms.ListBox)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXMoveToEnd
' This move the selected items to the end of the list.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As String          ' temporary variable to hold value for swap
Dim Ndx As Long             ' index counter for LBX.List
Dim SelNdx As Long          ' index of selected items
Dim SelCount As Long        ' number of selected items
Dim FirstSelItem As Long    ' first selected item index
Dim LastSelItem As Long     ' last selected item index
Dim SaveNdx As Long         ' saved index to reselect items at end

''''''''''''''''''''''''
' If list is empty, get
' out now.
''''''''''''''''''''''''
If LBX.ListCount = 0 Then
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''
' Ensure the selected items
' are contiguous (no unselected rows
' within selected rows).
''''''''''''''''''''''''''''''''''''
If LBXIsSelectionContiguous(LBX:=LBX) = False Then
    Exit Sub
End If

With LBX
    If .ColumnCount > 1 Then
        ''''''''''''''''''''''''''
        ' No support for mutliple
        ' column listboxes.
        ''''''''''''''''''''''''''
        Exit Sub
    End If
    

    ''''''''''''''''''''''''''''''
    ' If the list is empty, there
    ' is nothing to do. Get out.
    ''''''''''''''''''''''''''''''
    If .ListCount = 0 Then
        Exit Sub
    End If
    '''''''''''''''''''''''''''''
    ' If there is no selected
    ' item, there is nothing to
    ' do. Get Out.
    '''''''''''''''''''''''''''''
    If .ListIndex < 0 Then
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get information about the selected items
    ' in the list box LBX.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    LBXSelectionInfo LBX:=LBX, SelectedCount:=SelCount, _
        FirstSelectedItemIndex:=FirstSelItem, LastSelectedItemIndex:=LastSelItem
    
    ''''''''''''''''''''''''''''''''''
    ' If nothing is selected, get out.
    ''''''''''''''''''''''''''''''''''
    If SelCount = 0 Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''
    ' If no items are selected, get out.
    ' This should be picked up in
    ' SelCount = 0, but we test here for
    ' completeness.
    ''''''''''''''''''''''''''''''''''''
    If (FirstSelItem < 0) Or (LastSelItem < 0) Then
        Exit Sub
    End If
    
    SelNdx = 0
    Ndx = 0
    For SelNdx = LastSelItem To FirstSelItem Step -1
        Temp = .List(SelNdx)
        .RemoveItem SelNdx
        .AddItem Temp, .ListCount - Ndx
        SaveNdx = .ListCount - 1 - Ndx
        Ndx = Ndx + 1
    Next SelNdx

    LBXUnSelectAllItems LBX:=LBX
    For Ndx = SaveNdx To .ListCount - 1
        .Selected(Ndx) = True
    Next Ndx
    
    
End With

End Sub

Public Sub LBXSort(LBX As MSForms.ListBox, Optional FirstIndex As Long = -1, _
    Optional LastIndex As Long = -1, Optional Descending As Boolean = False, _
    Optional SelectedItemsOnly As Boolean = False)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXSort
' This calls QSortInPlace to sort the entries in the list box. If FirstIndex is
' supplied and is greater than or equal to 0, only entries at and after
' FirstIndex are sorted. If FirstIndex is omitted or less than 0, the sort starts
' with the first entry in the list box. This parameter is ignored if
' SelectedItemsOnly is True. If LastIndex is supplied and is greater than or equal
' to 0, only items at and before LastIndex are sorted. This parameter is ignored
' if SelectedItemsOnly is True. Descending is True or False indicating whether the
' list should be sorted in desending order. The default is False, indicating
' ascending order. SelectedItemsOnly is True or False indicating whether only the
' selected items should be sorted. If omitted, the items between (inclusive)
' FirstIndex and LastIndex are sorted. If FirstIndex > LastIndex, the entire
' list is sorted.
' If SelectedItemsOnly is True, then the procedure uses LBXIsSelectionContiguous to
' ensure that the selected items are contiguous. If they are contiguous, they are
' sorted. If they are not contiguous, the item are not sort.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Arr() As String
Dim First As Long
Dim Last As Long
Dim Ndx As Long

''''''''''''''''''''''''
' If list is empty, get
' out now.
''''''''''''''''''''''''
If LBX.ListCount = 0 Then
    Exit Sub
End If

''''''''''''''''''''''''''''''''''''
' Ensure the selected items
' are contiguous (no unselected rows
' within selected rows).
''''''''''''''''''''''''''''''''''''
If SelectedItemsOnly = True Then
    If LBXIsSelectionContiguous(LBX:=LBX) = False Then
        Exit Sub
    End If
End If


With LBX
    If .ColumnCount > 1 Then
        ''''''''''''''''''''''''''
        ' No support for mutliple
        ' column listboxes.
        ''''''''''''''''''''''''''
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' If FirstIndex is less than 0, set
    ' it to 0, the first item in the list.
    ''''''''''''''''''''''''''''''''''''''
    If FirstIndex < 0 Then
        First = 0
    Else
        First = FirstIndex
    End If
    ''''''''''''''''''''''''''''''''''''''
    ' If LastIndex is less than 0, set
    ' it to ListCount -1 , the last item
    ' in the list.
    ''''''''''''''''''''''''''''''''''''''
    If LastIndex < 0 Then
        Last = .ListCount - 1
    Else
        Last = LastIndex
    End If
    
    
    If First = Last Then
        ''''''''''''''''''''''''''''''''''
        ' There is nothing to do. Get out.
        ''''''''''''''''''''''''''''''''''
        Exit Sub
    End If
    
    If .ListCount <= 1 Then
        ''''''''''''''''''''''''''''''''''
        ' There is nothing to do. Get out.
        ''''''''''''''''''''''''''''''''''
        Exit Sub
    End If
        
    '''''''''''''''''''''''''''''''''''''
    ' Load the list contents into an array.
    '''''''''''''''''''''''''''''''''''''
    ReDim Arr(0 To .ListCount - 1)
    For Ndx = 0 To .ListCount - 1
        Arr(Ndx) = .List(Ndx)
    Next Ndx
    
    '''''''''''''''''''''''''''''''''
    ' Sort the array.
    '''''''''''''''''''''''''''''''''
    QSortInPlace InputArray:=Arr, _
                LB:=FirstIndex, _
                UB:=LastIndex, _
                Descending:=Descending, _
                CompareMode:=vbTextCompare, _
                NoAlerts:=True
    ''''''''''''''''''''''''''''''''
    ' Clear the list box and reload
    ' with the sorted array Arr.
    '''''''''''''''''''''''''''''''
    .Clear
    For Ndx = 0 To (UBound(Arr) - LBound(Arr))
        .AddItem Arr(Ndx)
    Next Ndx
End With


End Sub


Public Sub LBXSelectionInfo(LBX As MSForms.ListBox, ByRef SelectedCount As Long, _
    ByRef FirstSelectedItemIndex As Long, ByRef LastSelectedItemIndex As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SelectionInfo
' This procedure provides information about the selected
' items in the listbox referenced by LBX. The variable
' SelectedCount will be populated with the number of selected
' items, the variable FirstSelectedItem will be popuplated
' with the index number of the first (from the top down)
' selected item, and the variable LastSelectedItem will return
' the index number of the last (from the top down) selected
' item. If no item(s) are selected or ListIndex < 0,
' SelectedCount is set to 0, and FirstSelectedItem and
' LastSelectedItem are set to -1.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim FirstItem As Long: FirstItem = -1
Dim LastItem As Long:   LastItem = -1
Dim SelCount As Long:   SelCount = 0
Dim Ndx As Long

''''''''''''''''''''''''
' If list is empty, get
' out now.
''''''''''''''''''''''''
If LBX.ListCount = 0 Then
    Exit Sub
End If

With LBX
    If .ListCount = 0 Then
        SelectedCount = 0
        FirstSelectedItemIndex = -1
        LastSelectedItemIndex = -1
        Exit Sub
    End If
    If .ListIndex < 0 Then
        SelectedCount = 0
        FirstSelectedItemIndex = -1
        LastSelectedItemIndex = -1
        Exit Sub
    End If
    For Ndx = 0 To .ListCount - 1
        If .Selected(Ndx) = True Then
            If FirstItem < 0 Then
                FirstItem = Ndx
            End If
            SelCount = SelCount + 1
            LastItem = Ndx
        End If
    Next Ndx
End With
    
SelectedCount = SelCount
FirstSelectedItemIndex = FirstItem
LastSelectedItemIndex = LastItem

End Sub

Public Function LBXSelectedItems(LBX As MSForms.ListBox) As String()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SelectedItems
' This returns a 0-based array of strings, each of which is a selected
' item in the list box. If LBX is empty, the result is an unallocated
' array. The caller should first call SelectionInfo to determine whether
' there are any selected items prior to calling this procedure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim SelCount As Long
Dim FirstIndex As Long
Dim LastIndex As Long
Dim SelItems() As String
Dim Ndx As Long
Dim ArrNdx As Long

''''''''''''''''''''''''
' If list is empty, get
' out now.
''''''''''''''''''''''''
If LBX.ListCount = 0 Then
    Exit Function
End If

If LBX.ColumnCount > 1 Then
    ''''''''''''''''''''''''''
    ' No support for mutliple
    ' column listboxes.
    ''''''''''''''''''''''''''
    Exit Function
End If

LBXSelectionInfo LBX:=LBX, SelectedCount:=SelCount, _
    FirstSelectedItemIndex:=FirstIndex, LastSelectedItemIndex:=LastIndex

''''''''''''''''''''''''''''''''''''
' If nothing was selected, get out.
''''''''''''''''''''''''''''''''''''
If SelCount = 0 Then
    Exit Function
End If


ArrNdx = 0
'''''''''''''''''''''''''''''''''''
' Redim the result array to the
' number of selected items. This
' array is 0-based.
'''''''''''''''''''''''''''''''''''
ReDim SelItems(0 To SelCount - 1)

With LBX
    For Ndx = 0 To .ListCount - 1
        If .Selected(Ndx) = True Then
            SelItems(ArrNdx) = .List(Ndx)
            ArrNdx = ArrNdx + 1
        End If
    Next Ndx
End With

LBXSelectedItems = SelItems

End Function

Public Sub LBXUnSelectAllItems(LBX As MSForms.ListBox)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' UnSelectAllItems
' This procedure unselects all items in the listbox LBX.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
With LBX
    If .ListCount = 0 Then
        Exit Sub
    End If
    For Ndx = 0 To .ListCount - 1
        .Selected(Ndx) = False
    Next Ndx
End With

End Sub

Public Sub LBXSelectAllItems(LBX As MSForms.ListBox)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SelectAllItems
' This procedure selects all items in the listbox LBX.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
With LBX
    ''''''''''''''''''''''''
    ' If list is empty, get
    ' out now.
    ''''''''''''''''''''''''
    If .ListCount = 0 Then
        Exit Sub
    End If
    For Ndx = 0 To .ListCount - 1
        .Selected(Ndx) = True
    Next Ndx
End With

End Sub

Public Function LBXIsSelectionContiguous(LBX As MSForms.ListBox) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXIsSelectionContiguous
' This returns True if all selected items are contiguous, or False if
' they are not.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim SelCount As Long
Dim FirstItem As Long
Dim LastItem As Long
Dim Ndx As Long

''''''''''''''''''''''''
' If list is empty, get
' out now.
''''''''''''''''''''''''
If LBX.ListCount = 0 Then
    Exit Function
End If

LBXSelectionInfo LBX:=LBX, SelectedCount:=SelCount, FirstSelectedItemIndex:=FirstItem, _
     LastSelectedItemIndex:=LastItem

If SelCount > 0 Then
    For Ndx = FirstItem To LastItem
        If LBX.Selected(Ndx) = False Then
            LBXIsSelectionContiguous = False
            Exit Function
        End If
    Next Ndx
End If

LBXIsSelectionContiguous = True

End Function

Public Sub LBXUnSelectIfNotContiguous(LBX As MSForms.ListBox, ListIndex As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXUnSelectIfNotContiguous
' This procedure prevents selection of non-contiguous items in the specified
' list box.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim FirstItem As Long
Dim LastItem As Long
Dim SelCount As Long
Dim Ndx As Long
Dim UnSelectedFound As Boolean

''''''''''''''''''''''''
' If list is empty, get
' out now.
''''''''''''''''''''''''
If LBX.ListCount = 0 Then
    Exit Sub
End If

LBXSelectionInfo LBX:=LBX, SelectedCount:=SelCount, _
    FirstSelectedItemIndex:=FirstItem, LastSelectedItemIndex:=LastItem

With LBX
    If .ListCount = 0 Then
        Exit Sub
    End If

    For Ndx = FirstItem + 1 To LastItem Step 1
        If .Selected(Ndx) = False Then
            UnSelectedFound = True
        End If
        If UnSelectedFound = True Then
            .Selected(Ndx) = False
        End If
    Next Ndx
    
End With

End Sub

Public Function LBXIsListSorted(LBX As MSForms.ListBox, _
    Optional Descending As Boolean = False, _
    Optional CompareMode As VbCompareMethod = vbTextCompare) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXIsListSorted
' This function returns True if the List is sorted in either
' ascending order or descending order, depending on the value of
' the Descending parameter. If the list is not sorted, it returns
' False. If the List is empty, the result is True. Adjacent
' duplicate items are allowed an by themselves do not indicate that
' the List is not sorted.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
With LBX
    ''''''''''''''''''''''''''''
    ' See if the list is empty.
    ''''''''''''''''''''''''''''
    If .ListCount = 0 Then
        LBXIsListSorted = True
        Exit Function
    End If
    
    For Ndx = 0 To .ListCount - 2
        '''''''''''''''''''''''''''''''''''''''''''''
        ' Loop through all but the last two entries.
        ' The code will compare List(Ndx) with
        ' List(Ndx+1) to determine whether the List
        ' is sorted.
        '''''''''''''''''''''''''''''''''''''''''''''
        If Descending = False Then
            If Ndx < .ListCount - 2 Then
                '''''''''''''''''''''''''''''''''''''''''''
                ' Test to see if .List(Ndx) is greater than
                ' .List(Ndx+1). If .List(Ndx) is greater
                ' than List(Ndx+1), the these two elements
                ' are not in ascending sorted order, so
                ' return False and get out.
                ''''''''''''''''''''''''''''''''''''''''''''
                If StrComp(.List(Ndx), .List(Ndx + 1), CompareMode) > 0 Then
                    LBXIsListSorted = False
                    Exit Function
                End If
            End If
        Else
            If Ndx < .ListCount - 2 Then
                '''''''''''''''''''''''''''''''''''''''''
                ' Test to see if List(Ndx) is less than
                ' List(Ndx+1). If List(Ndx) is less than
                ' List(Ndx+1), then these two elements
                ' are not in descending sorted order, so
                ' returns False and get out.
                '''''''''''''''''''''''''''''''''''''''''
                If StrComp(.List(Ndx), .List(Ndx + 1), CompareMode) < 0 Then
                    LBXIsListSorted = False
                    Exit Function
                End If
            End If
        End If
    Next Ndx
    
    '''''''''''''''''''''''''''''''''
    ' If we make it out of the loop,
    ' all items are in sorted order,
    ' so return True.
    '''''''''''''''''''''''''''''''''
    LBXIsListSorted = True
End With

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modQSortInPlace
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This module contains the QSortInPlace procedure and private supporting procedures.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function QSortInPlace( _
    ByRef InputArray As Variant, _
    Optional ByVal LB As Long = -1&, _
    Optional ByVal UB As Long = -1&, _
    Optional ByVal Descending As Boolean = False, _
    Optional ByVal CompareMode As VbCompareMethod = vbTextCompare, _
    Optional ByVal NoAlerts As Boolean = False) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QSortInPlace
'
' This function sorts the array InputArray in place -- this is, the original array in the
' calling procedure is sorted. It will work with either string data or numeric data.
' It need not sort the entire array. You can sort only part of the array by setting the LB and
' UB parameters to the first (LB) and last (UB) element indexes that you want to sort.
' LB and UB are optional parameters. If omitted LB is set to the LBound of InputArray, and if
' omitted UB is set to the UBound of the InputArray. If you want to sort the entire array,
' omit the LB and UB parameters, or set both to -1, or set LB = LBound(InputArray) and set
' UB to UBound(InputArray).
'
' By default, the sort method is case INSENSTIVE (case doens't matter: "A", "b", "C", "d").
' To make it case SENSITIVE (case matters: "A" "C" "b" "d"), set the CompareMode argument
' to vbBinaryCompare (=0). If Compare mode is omitted or is any value other than vbBinaryCompare,
' it is assumed to be vbTextCompare and the sorting is done case INSENSITIVE.
'
' The function returns TRUE if the array was successfully sorted or FALSE if an error
' occurred. If an error occurs (e.g., LB > UB), a message box indicating the error is
' displayed. To suppress message boxes, set the NoAlerts parameter to TRUE.
'
''''''''''''''''''''''''''''''''''''''
' MODIFYING THIS CODE:
''''''''''''''''''''''''''''''''''''''
' If you modify this code and you call "Exit Procedure", you MUST decrment the RecursionLevel
' variable. E.g.,
'       If SomethingThatCausesAnExit Then
'           RecursionLevel = RecursionLevel - 1
'           Exit Function
'       End If
'''''''''''''''''''''''''''''''''''''''
'
' Note: If you coerce InputArray to a ByVal argument, QSortInPlace will not be
' able to reference the InputArray in the calling procedure and the array will
' not be sorted.
'
' This function uses the following procedures. These are declared as Private procedures
' at the end of this module:
'       IsArrayAllocated
'       IsSimpleDataType
'       IsSimpleNumericType
'       QSortCompare
'       NumberOfArrayDimensions
'       ReverseArrayInPlace
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Buffer As Variant
Dim CurLow As Long
Dim CurHigh As Long
Dim CurMidpoint As Long
Dim Ndx As Long
Dim pCompareMode As VbCompareMethod

'''''''''''''''''''''''''
' Set the default result.
'''''''''''''''''''''''''
QSortInPlace = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This variable is used to determine the level
' of recursion  (the function calling itself).
' RecursionLevel is incremented when this procedure
' is called, either initially by a calling procedure
' or recursively by itself. The variable is decremented
' when the procedure exits. We do the input parameter
' validation only when RecursionLevel is 1 (when
' the function is called by another function, not
' when it is called recursively).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Static RecursionLevel As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Keep track of the recursion level -- that is, how many
' times the procedure has called itself.
' Carry out the validation routines only when this
' procedure is first called. Don't run the
' validations on a recursive call to the
' procedure.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
RecursionLevel = RecursionLevel + 1

If RecursionLevel = 1 Then
    ''''''''''''''''''''''''''''''''''
    ' Ensure InputArray is an array.
    ''''''''''''''''''''''''''''''''''
    If IsArray(InputArray) = False Then
        If NoAlerts = False Then
            MsgBox "The InputArray parameter is not an array."
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' InputArray is not an array. Exit with a False result.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        RecursionLevel = RecursionLevel - 1
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test LB and UB. If < 0 then set to LBound and UBound
    ' of the InputArray.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If LB < 0 Then
        LB = LBound(InputArray)
    End If
    If UB < 0 Then
        UB = UBound(InputArray)
    End If
    
    Select Case NumberOfArrayDimensions(InputArray)
        Case 0
            ''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions indicates an unallocated
            ' dynamic array.
            ''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The InputArray is an empty, unallocated array."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case 1
            ''''''''''''''''''''''''''''''''''''''''''
            ' We sort ONLY single dimensional arrays.
            ''''''''''''''''''''''''''''''''''''''''''
        Case Else
            ''''''''''''''''''''''''''''''''''''''''''
            ' We sort ONLY single dimensional arrays.
            ''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The InputArray is multi-dimensional." & _
                      "QSortInPlace works only on single-dimensional arrays."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
    End Select
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that InputArray is an array of simple data
    ' types, not other arrays or objects. This tests
    ' the data type of only the first element of
    ' InputArray. If InputArray is an array of Variants,
    ' subsequent data types may not be simple data types
    ' (e.g., they may be objects or other arrays), and
    ' this may cause QSortInPlace to fail on the StrComp
    ' operation.
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
        If NoAlerts = False Then
            MsgBox "InputArray is not an array of simple data types."
            RecursionLevel = RecursionLevel - 1
            Exit Function
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ensure that the LB parameter is valid.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case LB
        Case Is < LBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is less than the LBound of the InputArray"
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is > UBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is greater than the UBound of the InputArray"
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is > UB
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is greater than the UB upper bound parameter."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
    End Select

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ensure the UB parameter is valid.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case UB
        Case Is > UBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The UB upper bound parameter is greater than the upper bound of the InputArray."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is < LBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The UB upper bound parameter is less than the lower bound of the InputArray."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is < LB
            If NoAlerts = False Then
                MsgBox "the UB upper bound parameter is less than the LB lower bound parameter."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
    End Select

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' if UB = LB, we have nothing to sort, so get out.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If UB = LB Then
        QSortInPlace = True
        RecursionLevel = RecursionLevel - 1
        Exit Function
    End If

End If ' RecursionLevel = 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure that CompareMode is either vbBinaryCompare  or
' vbTextCompare. If it is neither, default to vbTextCompare.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (CompareMode = vbBinaryCompare) Or (CompareMode = vbTextCompare) Then
    pCompareMode = CompareMode
Else
    pCompareMode = vbTextCompare
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Begin the actual sorting process.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CurLow = LB
CurHigh = UB

CurMidpoint = (LB + UB) \ 2 ' note integer division (\) here

Temp = InputArray(CurMidpoint)

Do While (CurLow <= CurHigh)
    
    Do While QSortCompare(V1:=InputArray(CurLow), V2:=Temp, CompareMode:=pCompareMode) < 0
        CurLow = CurLow + 1
        If CurLow = UB Then
            Exit Do
        End If
    Loop
    
    Do While QSortCompare(V1:=Temp, V2:=InputArray(CurHigh), CompareMode:=pCompareMode) < 0
        CurHigh = CurHigh - 1
        If CurHigh = LB Then
           Exit Do
        End If
    Loop

    If (CurLow <= CurHigh) Then
        Buffer = InputArray(CurLow)
        InputArray(CurLow) = InputArray(CurHigh)
        InputArray(CurHigh) = Buffer
        CurLow = CurLow + 1
        CurHigh = CurHigh - 1
    End If
Loop

If LB < CurHigh Then
    QSortInPlace InputArray:=InputArray, LB:=LB, UB:=CurHigh, _
        Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
End If

If CurLow < UB Then
    QSortInPlace InputArray:=InputArray, LB:=CurLow, UB:=UB, _
        Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
End If

'''''''''''''''''''''''''''''''''''''
' If Descending is True, reverse the
' order of the array, but only if the
' recursion level is 1.
'''''''''''''''''''''''''''''''''''''
If Descending = True Then
    If RecursionLevel = 1 Then
        ReverseArrayInPlace InputArray
    End If
End If

RecursionLevel = RecursionLevel - 1
QSortInPlace = True
End Function

Private Function QSortCompare(V1 As Variant, V2 As Variant, _
    Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QSortCompare
' This function is used in QSortInPlace to compare two elements. If
' V1 AND V2 are both numeric data types (integer, long, single, double)
' they are converted to Doubles and compared. If V1 and V2 are BOTH strings
' that contain numeric data, they are converted to Doubles and compared.
' If either V1 or V2 is a string and does NOT contain numeric data, both
' V1 and V2 are converted to Strings and compared with StrComp.
'
' The result is -1 if V1 < V2,
'                0 if V1 = V2
'                1 if V1 > V2
' For text comparisons, case sensitivity is controlled by CompareMode.
' If this is vbBinaryCompare, the result is case SENSITIVE. If this
' is omitted or any other value, the result is case INSENSITIVE.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim D1 As Double
Dim D2 As Double
Dim S1 As String
Dim S2 As String

Dim Compare As VbCompareMethod
''''''''''''''''''''''''''''''''''''''''''''''''
' Test CompareMode. Any value other than
' vbBinaryCompare will default to vbTextCompare.
''''''''''''''''''''''''''''''''''''''''''''''''
If CompareMode = vbBinaryCompare Or CompareMode = vbTextCompare Then
    Compare = CompareMode
Else
    Compare = vbTextCompare
End If
'''''''''''''''''''''''''''''''''''''''''''''''
' If either V1 or V2 is either an array or
' an Object, raise a error 13 - Type Mismatch.
'''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(V1) = True Or IsArray(V2) = True Then
    Err.Raise 13
    Exit Function
End If
If IsObject(V1) = True Or IsObject(V2) = True Then
    Err.Raise 13
    Exit Function
End If

If IsSimpleNumericType(V1) = True Then
    If IsSimpleNumericType(V2) = True Then
        '''''''''''''''''''''''''''''''''''''
        ' If BOTH V1 and V2 are numeric data
        ' types, then convert to Doubles and
        ' do an arithmetic compare and
        ' return the result.
        '''''''''''''''''''''''''''''''''''''
        D1 = CDbl(V1)
        D2 = CDbl(V2)
        If D1 = D2 Then
            QSortCompare = 0
            Exit Function
        End If
        If D1 < D2 Then
            QSortCompare = -1
            Exit Function
        End If
        If D1 > D2 Then
            QSortCompare = 1
            Exit Function
        End If
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''
' Either V1 or V2 was not numeric data type.
' Test whether BOTH V1 AND V2 are numeric
' strings. If BOTH are numeric, convert to
' Doubles and do a arithmetic comparison.
''''''''''''''''''''''''''''''''''''''''''''
If IsNumeric(V1) = True And IsNumeric(V2) = True Then
    D1 = CDbl(V1)
    D2 = CDbl(V2)
    If D1 = D2 Then
        QSortCompare = 0
        Exit Function
    End If
    If D1 < D2 Then
        QSortCompare = -1
        Exit Function
    End If
    If D1 > D2 Then
        QSortCompare = 1
        Exit Function
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''
' Either or both V1 and V2 was not numeric
' string. In this case, convert to Strings
' and use StrComp to compare.
''''''''''''''''''''''''''''''''''''''''''''''
S1 = CStr(V1)
S2 = CStr(V2)
QSortCompare = StrComp(S1, S2, Compare)

End Function



Private Function NumberOfArrayDimensions(Arr As Variant) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Integer
Dim Res As Integer
On Error Resume Next
' Loop, increasing the dimension index Ndx, until an error occurs.
' An error will occur when Ndx exceeds the number of dimension
' in the array. Return Ndx - 1.
Do
    Ndx = Ndx + 1
    Res = UBound(Arr, Ndx)
Loop Until Err.Number <> 0

NumberOfArrayDimensions = Ndx - 1

End Function
 
Private Function ReverseArrayInPlace(InputArray As Variant, _
    Optional NoAlerts As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayInPlace
' This procedure reverses the order of an array in place -- this is, the array variable
' in the calling procedure is sorted. An error will occur if InputArray is not an array,
 'if it is an empty, unallocated array, or if the number of dimensions is not 1.
'
' NOTE: Before calling the ReverseArrayInPlace procedure, consider if your needs can
' be met by simply reading the existing array in reverse order (Step -1). If so, you can save
' the overhead added to your application by calling this function.
'
' The function returns TRUE if the array was successfully reversed, or FALSE if
' an error occurred.
'
' If an error occurred, a message box is displayed indicating the error. To suppress
' the message box and simply return FALSE, set the NoAlerts parameter to TRUE.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long

''''''''''''''''''''''''''''''''
' Set the default return value.
''''''''''''''''''''''''''''''''
ReverseArrayInPlace = False

'''''''''''''''''''''''''''''''''
' Ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    If NoAlerts = False Then
        MsgBox "The InputArray parameter is not an array."
    End If
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(InputArray)
    Case 0
        '''''''''''''''''''''''''''''''''''''''''''
        ' Zero dimensions indicates an unallocated
        ' dynamic array.
        '''''''''''''''''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
        End If
        Exit Function
    Case 1
        '''''''''''''''''''''''''''''''''''''''''''
        ' We can reverse ONLY a single dimensional
        ' arrray.
        '''''''''''''''''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''''''''''''''''
        ' We can reverse ONLY a single dimensional
        ' arrray.
        '''''''''''''''''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        End If
        Exit Function

End Select

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure that we have only simple data types,
' not an array of objects or arrays.
'''''''''''''''''''''''''''''''''''''''''''''
If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
    If NoAlerts = False Then
        MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
            "ReverseArrayInPlace can reverse only arrays of simple data types."
        Exit Function
    End If
End If

Ndx2 = UBound(InputArray)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' loop from the LBound of InputArray to the midpoint of InputArray
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2)
    '''''''''''''''''''''''''''''''''
    'swap the elements
    '''''''''''''''''''''''''''''''''
    Temp = InputArray(Ndx)
    InputArray(Ndx) = InputArray(Ndx2)
    InputArray(Ndx2) = Temp
    '''''''''''''''''''''''''''''
    ' decrement the upper index
    '''''''''''''''''''''''''''''
    Ndx2 = Ndx2 - 1

Next Ndx
ReverseArrayInPlace = True
End Function

Private Function IsSimpleNumericType(V As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsSimpleNumericType
' This returns TRUE if V is one of the following data types:
'        vbBoolean
'        vbByte
'        vbCurrency
'        vbDate
'        vbDecimal
'        vbDouble
'        vbInteger
'        vbLong
'        vbSingle
'        vbVariant if it contains a numeric value
' It returns FALSE for any other data type, including any array
' or vbEmpty.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsSimpleDataType(V) = True Then
    Select Case VarType(V)
        Case vbBoolean, _
                vbByte, _
                vbCurrency, _
                vbDate, _
                vbDecimal, _
                vbDouble, _
                vbInteger, _
                vbLong, _
                vbSingle
            IsSimpleNumericType = True
        Case vbVariant
            If IsNumeric(V) = True Then
                IsSimpleNumericType = True
            Else
                IsSimpleNumericType = False
            End If
        Case Else
            IsSimpleNumericType = False
    End Select
Else
    IsSimpleNumericType = False
End If
End Function

Private Function IsSimpleDataType(V As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsSimpleDataType
' This function returns TRUE if V is one of the following
' variable types (as returned by the VarType function:
'    vbBoolean
'    vbByte
'    vbCurrency
'    vbDate
'    vbDecimal
'    vbDouble
'    vbEmpty
'    vbError
'    vbInteger
'    vbLong
'    vbNull
'    vbSingle
'    vbString
'    vbVariant
'
' It returns FALSE if V is any one of the following variable
' types:
'    vbArray
'    vbDataObject
'    vbObject
'    vbUserDefinedType
'    or if it is an array of any type.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test if V is an array. We can't just use VarType(V) = vbArray
' because the VarType of an array is vbArray + VarType(type
' of array element). E.g, the VarType of an Array of Longs is
' 8195 = vbArray + vbLong.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(V) = True Then
    IsSimpleDataType = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' We must also explicitly check whether V is an object, rather
' relying on VarType(V) to equal vbObject. The reason is that
' if V is an object and that object has a default proprety, VarType
' returns the data type of the default property. For example, if
' V is an Excel.Range object pointing to cell A1, and A1 contains
' 12345, VarType(V) would return vbDouble, the since Value is
' the default property of an Excel.Range object and the default
' numeric type of Value in Excel is Double. Thus, in order to
' prevent this type of behavior with default properties, we test
' IsObject(V) to see if V is an object.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsObject(V) = True Then
    IsSimpleDataType = False
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''
' Test the value returned by VarType.
'''''''''''''''''''''''''''''''''''''
Select Case VarType(V)
    Case vbArray, vbDataObject, vbObject, vbUserDefinedType
        '''''''''''''''''''''''
        ' not simple data types
        '''''''''''''''''''''''
        IsSimpleDataType = False
    Case Else
        ''''''''''''''''''''''''''''''''''''
        ' otherwise it is a simple data type
        ''''''''''''''''''''''''''''''''''''
        IsSimpleDataType = True
End Select

End Function

Public Function IsArrayAllocated(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllocated
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array has not been allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim N As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''
' If Arr is not an array, return FALSE and get out.
'''''''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    IsArrayAllocated = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Try to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occured.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
N = UBound(Arr, 1)
If Err.Number = 0 Then
    '''''''''''''''''''''''''''''''''''''
    ' No error. Array has been allocated.
    '''''''''''''''''''''''''''''''''''''
    IsArrayAllocated = True
Else
    '''''''''''''''''''''''''''''''''''''
    ' Error. Unallocated array.
    '''''''''''''''''''''''''''''''''''''
    IsArrayAllocated = False
End If

End Function


