Attribute VB_Name = "TestForm"
Private Function populatedForm(protected As Boolean) As Object
    Const vbext_ct_MSForm As Long = 3
    Const form_non_existant As Long = 424
    If protected Then
        'try to get existing form
        On Error Resume Next
        Dim errNum As Long
        Set populatedForm = VBA.UserForms.Add("TemplateForm")

        'check if form looks as it should (provided we got one as expected)
        If Err.Number = 0 Then If Not formIsCorrect(populatedForm) Then Err.Raise form_non_existant
        errNum = Err.Number
        On Error GoTo 0
        If Not (errNum = 0 Or errNum = form_non_existant) Then Err.Raise errNum 'uncaught error
    Else
        'create form
        With ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
            .Designer.Controls.Add("Forms.listbox.1").Name = "lst_1"
            .Designer.Controls.Add("Forms.commandbutton.1").Name = "cmd_1"
            .codeModule.InsertLines 1, "Private Sub cmd_1_Click()"
            .codeModule.InsertLines 2, "Me.Hide"
            .codeModule.InsertLines 3, "End Sub"
            Set populatedForm = VBA.UserForms.Add(.Name)
        End With
    End If
End Function

Sub testGet()
    Dim newForm As Object
    Const protected As Boolean = True
    Set newForm = populatedForm(protected)
    If Not newForm Is Nothing Then
        Dim formName As String
        formName = newForm.Name
        designForm newForm, ThisWorkbook
        newForm.Show
        If Not protected Then
            With ThisWorkbook.VBProject.VBComponents
                .Remove .Item(formName)
            End With
        End If
    Else
        Err.Description = "No listbox form could be found or created"
        Err.Raise 5
    End If
End Sub

Private Function formIsCorrect(form As Object) As Boolean
    On Error Resume Next
    With form.Controls
        If .Count <> 2 Then Err.Raise 5
        Set v = .Item("lst_1")
        Set r = .Item("cmd_1")
    End With
    formIsCorrect = Err.Number = 0
    On Error GoTo 0
End Function

Private Sub designForm(ByRef populatedForm As Object, ByVal callerBook As Workbook)
    'change overall appearence
    With populatedForm
        .Caption = "Select"
        .Width = 300
        .Height = 270
    End With

    'Change ListBox appearence
    With populatedForm.Controls("lst_1")
        .Top = 10
        .Left = 10
        .Width = 150
        .Height = 230
        .Font.size = 8
        .Font.Name = "Tahoma"
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectSunken
        .MultiSelect = fmMultiSelectMulti
    End With
    
    'Change CommandButton appearence
    With populatedForm.Controls("cmd_1")
        .Caption = "Choose"
        .Accelerator = "M"
        .Top = 10
        .Left = 200
        .Width = 66
        .Height = 20
        .Font.size = 8
        .Font.Name = "Tahoma"
        .BackStyle = fmBackStyleOpaque
    End With
    
    'populate listbox
    Dim codeItem As Object                            'VBComponent
    For Each codeItem In callerBook.VBProject.VBComponents
        If Not codeItem Is populatedForm Then populatedForm.lst_1.AddItem codeItem.Name
    Next cmp
    

End Sub
