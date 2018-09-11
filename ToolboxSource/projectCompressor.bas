Attribute VB_Name = "projectCompressor"
'Compressor module, compresses a load of files into strings to export
Option Explicit
'@Folder Compressor
Private Type codeItem
    extension As String
    module_name As String
    code_content() As String
End Type

Private Const TypeBinary As Long = 1
'@Ignore MultipleDeclarations
Private Const vbext_ct_StdModule As Long = 1, vbext_ct_ClassModule As Long = 2, vbext_ct_MSForm As Long = 3, vbext_ct_Document As Long = 100
Private Const fmBorderStyleSingle As Long = 1, fmSpecialEffectSunken As Long = 2, fmMultiSelectMulti As Long = 1, fmBackStyleOpaque As Long = 1
Private Const vbext_pp_none As Long = 0
Private Const invalid_argument_error As Long = 5

Public Sub CompressProjectFileSelector(Optional ByRef wb As Variant)

    Dim book As Workbook
    If IsMissing(wb) Then Set book = ActiveWorkbook Else Set book = wb 'or active workbook?
    
    Dim selections() As String
    On Error Resume Next
    selections = ListBoxChoice(book)
    If Err.Number <> 0 Then
        MsgBox Err.Description, Title:="Error getting selection"
        Exit Sub
    End If
    On Error GoTo 0
    
    If selections(1) = "None Selected" Then
        MsgBox "No modules selected to export"
    Else
        If CompressProject(book, selections) Then
            MsgBox "Compression succeeded"
        Else
            MsgBox "Compression unsuccessful, see immediate window for more info"
        End If
    End If
End Sub

Public Function CompressProject(ByRef wb As Workbook, ParamArray moduleNames()) As Boolean
    'Sub to convert selected files into self-extracting module
    'Input:
    '   filenames: array of strings based on names of modules in project
    If Not ProjectAccessible(wb) Then
        MsgBox "Access to VBA project is restricted, this won't work!"
        Exit Function
    ElseIf UBound(moduleNames) < 0 Then
        MsgBox "No module names passed!"
        Exit Function
    
    End If
    Dim filenames As Variant
    If IsArray(moduleNames(0)) Then filenames = moduleNames(0) Else filenames = moduleNames
    

    Dim codeItems() As codeItem
    Dim arraySt As Long, arrayEnd As Long, i As Long
    arraySt = LBound(filenames)
    arrayEnd = UBound(filenames)
    ReDim codeItems(arraySt To arrayEnd)
    
Debug.Print "Getting Definitions..."
    With wb.VBProject.VBComponents
        'loop through files compressing them int 64 bit strings
        For i = arraySt To arrayEnd
            codeItems(i) = ModuleDefinition(filenames(i), wb)
        Next i
    End With
Debug.Print , "Definitions saved"
    'write strings to skeleton file
Debug.Print "Writing file..."
    CompressProject = WriteSkeleton(codeItems, wb)
Debug.Print "Complete"
End Function

Private Function ListBoxChoice(ByVal wb As Workbook) As String()

    Dim newForm As Object
    Dim protected As Boolean
    protected = Not ProjectAccessible(ThisWorkbook)
    
     'This is to stop screen flashing while creating form
    Application.VBE.MainWindow.Visible = False
    
    Set newForm = populatedForm(protected)
    If newForm Is Nothing Then 'got something valid
    
        Err.Description = "No listbox form could be found or created"
        Err.Raise 5
        
    Else
        Dim formName As String
        formName = newForm.Name
        designForm newForm, wb 'populate form
        Application.VBE.MainWindow.Visible = True
        newForm.Show
        
        Dim selections() As String                   'hold output of list
        On Error GoTo noSelection
        ReDim selections(1 To newForm.lst_1.ListCount)
        On Error GoTo 0
    
        Dim selectedCount As Long
        selectedCount = 0
        Dim i As Long
        For i = 0 To newForm.lst_1.ListCount - 1
            If newForm.lst_1.Selected(i) = True Then
                selectedCount = selectedCount + 1
                selections(selectedCount) = newForm.lst_1.List(i)
            End If
        Next i
        
        On Error GoTo noSelection
        If selectedCount = 0 Then Err.Raise 0        'no selection
        On Error GoTo 0

        ReDim Preserve selections(1 To selectedCount)
    End If
    'Delete the form (Optional)
safeExit:
    If Not protected Then
        With ThisWorkbook.VBProject.VBComponents
            .Remove .item(formName)
        End With
    End If
    ListBoxChoice = selections

    Exit Function
noSelection:                                     'just don't assign anything
    ReDim selections(1 To 1)
    selections(1) = "None Selected"
    Resume safeExit
End Function

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

Private Function formIsCorrect(ByVal form As Object) As Boolean
    On Error Resume Next
    Dim v, r
    With form.Controls
        If .Count <> 2 Then Err.Raise 5
        Set v = .item("lst_1")
        Set r = .item("cmd_1")
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
    Dim formComponent As Object
    On Error Resume Next 'Temp form will have unique name and be in ThisWorkbook always
        Set formComponent = ThisWorkbook.VBProject.VBComponents(populatedForm.Name) 'may not exist
    On Error GoTo 0
    For Each codeItem In callerBook.VBProject.VBComponents
        If Not codeItem Is formComponent Then populatedForm.lst_1.AddItem codeItem.Name
    Next codeItem
End Sub

Private Function WriteSkeleton(ByRef codeItems() As codeItem, ByRef book As Workbook, Optional ByRef projectName As String = "myProject") As Boolean ' , Optional wb As Variant)
    Dim itemCount As Long
    itemCount = UBound(codeItems) - LBound(codeItems) + 1
    If itemCount < 1 Then Exit Function
    

    'create self-extracting module and set name

    Dim extractorModule As Object                'VBComponent
    Set extractorModule = book.VBProject.VBComponents.Add(vbext_ct_StdModule)
    On Error GoTo cleanExit
    WriteProjectName projectName, extractorModule 'avoid err if duplicate - changes
Debug.Print , "Project file added"
    'write code to module
    Dim codeInsertPoint As Long
    codeInsertPoint = FillModule(extractorModule.codeModule)(0) 'x coord
Debug.Print , "Project skeleton written"
    'ammend code with codeitems and killing line
    With extractorModule.codeModule

        .DeleteLines codeInsertPoint
        Dim singItem As codeItem
        Dim i As Long
        
        'loop through adding code definitions
        For i = LBound(codeItems) To UBound(codeItems)
            singItem = codeItems(i)
            If singItem.extension = "missing" Then
                Err.Description = printf("Warning: Module ""{0}"" cannot be found or is not supported for compression", singItem.module_name)
                Err.Raise invalid_argument_error
                itemCount = itemCount - 1
            Else
                Dim itemIndex As Long, content_lbound As Long, content_ubound As Long
                content_ubound = UBound(singItem.code_content)
                content_lbound = LBound(singItem.code_content)
                'code content array
                For itemIndex = UBound(singItem.code_content) To LBound(singItem.code_content) Step -1 'loop backwards
                    .InsertLines codeInsertPoint, printf(String(4, vbTab) & ".code_content({1}) = {0}", singItem.code_content(itemIndex), itemIndex)
                Next itemIndex
                'code content array definition
                .InsertLines codeInsertPoint, printf(String(4, vbTab) & "Redim .code_content({0} To {1})", content_lbound, content_ubound)
                'other code definition items
                .InsertLines codeInsertPoint, printf(String(4, vbTab) & ".module_name = ""{0}""", singItem.module_name)
                .InsertLines codeInsertPoint, printf(String(4, vbTab) & ".extension = ""{0}""", singItem.extension)
                .InsertLines codeInsertPoint, printf(String(3, vbTab) & "Case {0}", itemCount)
                itemCount = itemCount - 1
            End If
        Next i

        Dim killLine As Long                     'place for adding last bit of code to remove self-extractor
        .Find "{1}", killLine, 1, -1, -1
        .ReplaceLine killLine, Replace(.Lines(killLine, 1), "{1}", projectName)
Debug.Print , "Inserted killLine"
    End With
    WriteSkeleton = True
    Exit Function
    
cleanExit:
Debug.Print printf("Error writing to file: #{0} - {1}", Err.Number, Err.Description)
Debug.Print , IIf(RemoveModule(projectName, book), _
                  "Temp file cleared up successfully", _
                  "Could not remove temp file: """ & projectName & """")
    WriteSkeleton = False
    
End Function

Private Sub WriteProjectName(ByRef base As String, ByRef module As Object)
    Const rename_error As Long = 32813
    On Error Resume Next
    module.Name = base
    
    Dim suffix As String
    suffix = vbNullString
    Do While Err.Number = rename_error
        Err.Clear
        suffix = Val(suffix) + 1
        module.Name = base & suffix
    Loop
    
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then
        Err.Raise errNum
    Else
        base = base & suffix
    End If
    
End Sub

Private Function ModuleDefinition(ByVal moduleName As String, ByVal book As Workbook) As codeItem
    Dim codeModule As Object                     'VBComponent
    Dim result As codeItem
    On Error GoTo moduleMissing
    Set codeModule = book.VBProject.VBComponents(moduleName)
    On Error GoTo 0
    'get extension and name
    Select Case codeModule.Type
    Case vbext_ct_StdModule
        result.extension = ".bas"
    Case vbext_ct_ClassModule
        result.extension = ".cls"
    Case vbext_ct_Document
Debug.Print , printf("Warning: Module ""{0}"" has been converted to a standard class as document types are not fully supported", moduleName)
        result.extension = ".cls"
    Case vbext_ct_MSForm
        result.extension = ".frm"
    Case Else
        result.extension = "missing"
        result.module_name = moduleName
        ModuleDefinition = result
        Exit Function
    End Select
    
    result.module_name = codeModule.Name
    'save to temp path
    Dim tempPath As String
    tempPath = printf("{0}\{1}{2}", Environ$("temp"), result.module_name, result.extension)
    codeModule.Export tempPath
    On Error GoTo safeExit
    result.code_content = Chunkify(ToBase64(ReadBytes(tempPath))) 'encode and chunkify
    
safeExit:
    Kill tempPath
moduleMissing:
    ModuleDefinition = result
    If Err.Number <> 0 Then ModuleDefinition.extension = "missing"
    
End Function

Private Function printf(ByVal mask As String, ParamArray tokens()) As String
    'Debug.Print , " -> Formatting"; Len(tokens(0)); "chars into", """"; mask; """"
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    printf = mask
End Function

Private Function ProjectAccessible(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    With wb.VBProject
        ProjectAccessible = .Protection = vbext_pp_none
        ProjectAccessible = ProjectAccessible And Err.Number = 0
    End With
End Function

Private Function ReadBytes(ByVal file As String) As Byte()
    Dim inStream As Object
    ' ADODB stream object used
    Set inStream = CreateObject("ADODB.Stream")
    ' open with no arguments makes the stream an empty container
    inStream.Open
    inStream.Type = TypeBinary
    inStream.LoadFromFile (file)
    ReadBytes = inStream.Read()
End Function

Private Function Chunkify(ByVal base As String, Optional ByVal stringLength As Long = 900) As String()
    'splits a string at every stringLength charachters and delimits
    '1024 is max chars in a line
    Dim contentGroups() As String
    contentGroups = SplitString(base, stringLength * 10) 'Splits into arrays of 10 lines
    Dim i As Long
    For i = LBound(contentGroups) To UBound(contentGroups)
        contentGroups(i) = Join(SplitString(contentGroups(i), stringLength, quotations:=True), " & _" & vbCrLf)
    Next i
    Chunkify = contentGroups
End Function

Private Function SplitString(ByVal str As String, ByVal numOfChar As Long, Optional ByVal quotations As Boolean = False) As String()
    Dim result() As String
    Dim nCount As Long
    ReDim result((Len(str) - 1) \ numOfChar)
    Do While Len(str)
        result(nCount) = Left$(str, numOfChar)
        If quotations Then result(nCount) = """" & result(nCount) & """"
        str = Mid$(str, numOfChar + 1)
        nCount = nCount + 1
    Loop
    SplitString = result
End Function

Private Function ToBase64(ByRef data() As Byte) As String
    Dim b64(0 To 63) As Byte, str() As Byte, i As Long, j As Long, v As Long, n As Long
    n = UBound(data) - LBound(data) + 1
    If n Then Else Exit Function

    str = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    For i = 0 To 127 Step 2
        b64(i \ 2) = str(i)
    Next

    ReDim str(0 To ((n + 2) \ 3) * 8 - 1)

    For i = LBound(data) To UBound(data) - (n Mod 3) Step 3
        v = data(i) * 65536 + data(i + 1) * 256& + data(i + 2)
        str(j) = b64(v \ 262144)
        str(j + 2) = b64((v \ 4096) Mod 64)
        str(j + 4) = b64((v \ 64) Mod 64)
        str(j + 6) = b64(v Mod 64)
        j = j + 8
    Next

    If n Mod 3 = 2 Then
        v = data(n - 2) * 256& + data(n - 1)
        str(j) = b64((v \ 1024&) Mod 64)
        str(j + 2) = b64((v \ 16) Mod 64)
        str(j + 4) = b64((v * 4) Mod 64)
        str(j + 6) = 61                          ' = '
    ElseIf n Mod 3 = 1 Then
        v = data(n - 1)
        str(j) = b64(v \ 4 Mod 64)
        str(j + 2) = b64(v * 16 Mod 64)
        str(j + 4) = 61                          ' = '
        str(j + 6) = 61                          ' = '
    End If

    ToBase64 = str
End Function

Private Function RemoveModule(ByVal moduleName As String, ByRef book As Workbook) As Boolean
    On Error Resume Next
    With book.VBProject.VBComponents
        .Remove .item(moduleName)
    End With
    RemoveModule = Not (Err.Number = 9)
End Function

Private Function FillModule(ByVal codeSection As Object) As Long()
    With codeSection
        .InsertLines 1, "Option Explicit"
        .InsertLines 2, "Private Type codeItem"
        .InsertLines 3, "    extension As String"
        .InsertLines 4, "    module_name As String"
        .InsertLines 5, "    code_content() As String"
        .InsertLines 6, "End Type"
        .InsertLines 7, ""
        .InsertLines 8, "Private Const TypeBinary As Long = 1"
        .InsertLines 9, "Private Const vbext_pp_none As Long = 0"
        .InsertLines 10, "Private Const ForReading As Long = 1, ForWriting As Long = 2, ForAppending As Long = 8"
        .InsertLines 11, ""
        .InsertLines 12, "Private Function getCodeDefinition(ByVal itemNo As Long) As codeItem"
        .InsertLines 13, "    With getCodeDefinition"
        .InsertLines 14, "        Select Case itemNo"
        .InsertLines 15, "           '{0}"
        .InsertLines 16, "        Case Else"
        .InsertLines 17, "            .extension = ""missing"""
        .InsertLines 18, "        End Select"
        .InsertLines 19, "    End With"
        .InsertLines 20, "End Function"
        .InsertLines 21, ""
        .InsertLines 22, "Public Sub Extract()"
        .InsertLines 23, "    Dim wb As Workbook"
        .InsertLines 24, "    Set wb = ThisWorkbook"
        .InsertLines 25, "    Dim code_module As codeItem"
        .InsertLines 26, "    Dim savedPath As String, basePath As String"
        .InsertLines 27, "    Dim i As Long"
        .InsertLines 28, "   'check if vbproject accessible"
        .InsertLines 29, "    If Not ProjectAccessible(wb) Then"
        .InsertLines 30, "        MsgBox ""The VBA project cannot be accessed programmatically"""
        .InsertLines 31, "        Exit Sub"
        .InsertLines 32, "    End If"
        .InsertLines 33, "   'check if temp folder acessible"
        .InsertLines 34, "    i = 0"
        .InsertLines 35, "    basePath = Environ$(""Temp"") & ""\"""
        .InsertLines 36, "    Do While True"
        .InsertLines 37, "        i = i + 1"
        .InsertLines 38, "        code_module = getCodeDefinition(i)"
        .InsertLines 39, "        If code_module.extension = ""missing"" Then"
        .InsertLines 40, "            Exit Do"
        .InsertLines 41, "        Else"
        .InsertLines 42, "            savedPath = createFile(code_module, basePath)"
        .InsertLines 43, "            importFile savedPath, wb"
        .InsertLines 44, "            Kill savedPath"
        .InsertLines 45, "        End If"
        .InsertLines 46, "    Loop"
        .InsertLines 47, "    RemoveModule ""{1}"", wb"
        .InsertLines 48, "End Sub"
        .InsertLines 49, ""
        .InsertLines 50, "Private Function ProjectAccessible(ByVal wb As Workbook) As Boolean"
        .InsertLines 51, "    On Error Resume Next"
        .InsertLines 52, "    With wb.VBProject"
        .InsertLines 53, "        ProjectAccessible = .Protection = vbext_pp_none"
        .InsertLines 54, "        ProjectAccessible = ProjectAccessible And Err.Number = 0"
        .InsertLines 55, "    End With"
        .InsertLines 56, "End Function"
        .InsertLines 57, ""
        .InsertLines 58, "Private Function createFile(ByRef definition As codeItem, ByVal filePath As String) As String"
        .InsertLines 59, "    Dim newFileObj As Object"
        .InsertLines 60, "    Set newFileObj = CreateObject(""ADODB.Stream"")"
        .InsertLines 61, "    newFileObj.Type = TypeBinary"
        .InsertLines 62, "   'Open the stream and write binary data"
        .InsertLines 63, "    newFileObj.Open"
        .InsertLines 64, "   'create file from x64 string"
        .InsertLines 65, "    With definition"
        .InsertLines 66, "        Dim bytes() As Byte"
        .InsertLines 67, "        Dim fullPath As String"
        .InsertLines 68, "        fullPath = filePath & .module_name & .extension"
        .InsertLines 69, "        bytes = FromBase64(Join(.code_content))"
        .InsertLines 70, "        newFileObj.Write bytes"
        .InsertLines 71, "        newFileObj.SaveToFile fullPath, ForWriting"
        .InsertLines 72, "        createFile = fullPath"
        .InsertLines 73, "    End With"
        .InsertLines 74, "End Function"
        .InsertLines 75, ""
        .InsertLines 76, "Private Sub importFile(ByVal filePath As String, ByRef wb As Workbook)"
        .InsertLines 77, "    wb.VBProject.VBComponents.Import filePath"
        .InsertLines 78, "End Sub"
        .InsertLines 79, ""
        .InsertLines 80, "Private Function RemoveModule(ByVal moduleName As String, ByRef book As Workbook) As Boolean"
        .InsertLines 81, "    On Error Resume Next"
        .InsertLines 82, "    With book.VBProject.VBComponents"
        .InsertLines 83, "        .Remove .Item(moduleName)"
        .InsertLines 84, "    End With"
        .InsertLines 85, "    RemoveModule = Not (Err.Number = 9)"
        .InsertLines 86, "End Function"
        .InsertLines 87, ""
        .InsertLines 88, "Private Function FromBase64(ByVal Text As String) As Byte()"
        .InsertLines 89, "    Dim Out() As Byte"
        .InsertLines 90, "    Dim b64(0 To 255) As Byte, str() As Byte, i As Long, j As Long, v As Long, b0 As Long, b1 As Long, b2 As Long, b3 As Long"
        .InsertLines 91, "    Out = vbNullString"
        .InsertLines 92, "    If Len(Text) Then Else Exit Function"
        .InsertLines 93, ""
        .InsertLines 94, "    str = "" ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"""
        .InsertLines 95, "    For i = 2 To UBound(str) Step 2"
        .InsertLines 96, "        b64(str(i)) = i \ 2"
        .InsertLines 97, "    Next"
        .InsertLines 98, ""
        .InsertLines 99, "    ReDim Out(0 To ((Len(Text) + 3) \ 4) * 3 - 1)"
        .InsertLines 100, "    str = Text & String$(2, 0)"
        .InsertLines 101, ""
        .InsertLines 102, "    For i = 0 To UBound(str) - 7 Step 2"
        .InsertLines 103, "        b0 = b64(str(i))"
        .InsertLines 104, ""
        .InsertLines 105, "        If b0 Then"
        .InsertLines 106, "            b1 = b64(str(i + 2))"
        .InsertLines 107, "            b2 = b64(str(i + 4))"
        .InsertLines 108, "            b3 = b64(str(i + 6))"
        .InsertLines 109, "            v = b0 * 262144 + b1 * 4096& + b2 * 64& + b3 - 266305"
        .InsertLines 110, "            Out(j) = v \ 65536"
        .InsertLines 111, "            Out(j + 1) = (v \ 256&) Mod 256"
        .InsertLines 112, "            Out(j + 2) = v Mod 256"
        .InsertLines 113, "            j = j + 3"
        .InsertLines 114, "            i = i + 6"
        .InsertLines 115, "        End If"
        .InsertLines 116, "    Next"
        .InsertLines 117, ""
        .InsertLines 118, "    If b2 = 0 Then"
        .InsertLines 119, "        Out(j - 3) = (v + 65) \ 65536"
        .InsertLines 120, "        j = j - 2"
        .InsertLines 121, "    ElseIf b3 = 0 Then"
        .InsertLines 122, "        Out(j - 3) = (v + 1) \ 65536"
        .InsertLines 123, "        Out(j - 2) = ((v + 1) \ 256&) Mod 256"
        .InsertLines 124, "        j = j - 1"
        .InsertLines 125, "    End If"
        .InsertLines 126, ""
        .InsertLines 127, "    ReDim Preserve Out(j - 1)"
        .InsertLines 128, "    FromBase64 = Out"
        .InsertLines 129, "End Function"
Debug.Print , "Inserted skeleton"
        Dim result(0 To 1) As Long
        If .Find("{0}", result(0), result(1), -1, -1) Then 'search for point to insert lines
            FillModule = result
        Else
            result(0) = 0
            result(1) = 0
            FillModule = result
        End If
    End With
End Function
