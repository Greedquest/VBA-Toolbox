Attribute VB_Name = "projectCompressor"
'Compressor module, compresses a load of files into strings to export
Option Explicit
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


Private Function ListBoxChoice(ByVal wb As Workbook) As String()

    Dim myForm As Object
    Dim newButton As Object                      'MSForms.CommandButton
    Dim newListBox As Object                     'MSForms.ListBox

    'This is to stop screen flashing while creating form
    Application.VBE.MainWindow.Visible = False
    
    'Add to ThisWorkbook, not supplied workbook or VBE will crash
    Set myForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)

    'Create the User Form
    With myForm
        .Properties("Caption") = "Select"
        .Properties("Width") = 300
        .Properties("Height") = 270
    End With

    'Create ListBox
    Set newListBox = myForm.Designer.Controls.Add("Forms.listbox.1")
    With newListBox
        .Name = "lst_1"
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

    'Create CommandButton
    Set newButton = myForm.Designer.Controls.Add("Forms.commandbutton.1")
    With newButton
        .Name = "cmd_1"
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

    'Add code for Comand Button
    myForm.codeModule.InsertLines 1, "Private Sub cmd_1_Click()"
    myForm.codeModule.InsertLines 2, "Me.Hide"
    myForm.codeModule.InsertLines 3, "End Sub"
    
    'Show the form
    Dim finalForm As Object
    Set finalForm = VBA.UserForms.Add(myForm.Name)
    'populate list
    Dim cmp As Object                            'VBComponent
    For Each cmp In wb.VBProject.VBComponents
        If cmp.Name <> finalForm.Name Then finalForm.lst_1.AddItem cmp.Name
    Next cmp

    finalForm.Show
    Dim selections() As String                   'hold output of list
    On Error GoTo noSelection
    ReDim selections(1 To finalForm.lst_1.ListCount)
    On Error GoTo 0

    Dim selectedCount As Long
    selectedCount = 0
    Dim i As Long
    For i = 0 To finalForm.lst_1.ListCount - 1
        If finalForm.lst_1.Selected(i) = True Then
            selectedCount = selectedCount + 1
            selections(selectedCount) = finalForm.lst_1.List(i)
        End If
    Next i
    
    On Error GoTo noSelection
    If selectedCount = 0 Then Err.Raise 0 'no selection
    On Error GoTo 0

    ReDim Preserve selections(1 To selectedCount)
    'Delete the form (Optional)
safeExit:
    ThisWorkbook.VBProject.VBComponents.Remove myForm
    ListBoxChoice = selections

    Exit Function
noSelection:                                     'just don't assign anything
    ReDim selections(1 To 1)
    selections(1) = "None Selected"
    Resume safeExit
End Function

Public Sub CompressProjectFileSelector(Optional ByRef wb As Variant)

    Dim book As Workbook
    If IsMissing(wb) Then Set book = ActiveWorkbook Else Set book = wb 'or active workbook?
    
    Dim selections() As String
    selections = ListBoxChoice(book)
    
   
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
                Err.Description = Printf("Warning: Module ""{0}"" cannot be found or is not supported for compression", singItem.module_name)
                Err.Raise invalid_argument_error
                itemCount = itemCount - 1
            Else
                Dim itemIndex As Long, content_lbound As Long, content_ubound As Long
                content_ubound = UBound(singItem.code_content)
                content_lbound = LBound(singItem.code_content)
                'code content array
                For itemIndex = UBound(singItem.code_content) To LBound(singItem.code_content) Step -1 'loop backwards
                    .InsertLines codeInsertPoint, Printf(String(4, vbTab) & ".code_content({1}) = {0}", singItem.code_content(itemIndex), itemIndex)
                Next itemIndex
                'code content array definition
                .InsertLines codeInsertPoint, Printf(String(4, vbTab) & "Redim .code_content({0} To {1})", content_lbound, content_ubound)
                'other code definition items
                .InsertLines codeInsertPoint, Printf(String(4, vbTab) & ".module_name = ""{0}""", singItem.module_name)
                .InsertLines codeInsertPoint, Printf(String(4, vbTab) & ".extension = ""{0}""", singItem.extension)
                .InsertLines codeInsertPoint, Printf(String(3, vbTab) & "Case {0}", itemCount)
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
Debug.Print Printf("Error writing to file: #{0} - {1}", Err.Number, Err.Description)
Debug.Print , IIf(RemoveModule(projectName, book), _
                  "Temp file cleared up successfully", _
                  "Could not remove temp file: """ & projectName & """")
    WriteSkeleton = False
    
End Function

Private Sub WriteProjectName(ByRef base As String, ByRef module As Object)
    Const rename_error As Long = 32813
    On Error Resume Next
    module.Name = base
    
    Dim i As Variant
    i = vbNullString
    Do While Err.Number = rename_error
        Err.Clear
        i = i + 1
        module.Name = base & i
        
    Loop
    
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then
        Err.Raise errNum
    Else
        base = base & i
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
Debug.Print , Printf("Warning: Module ""{0}"" has been converted to a standard class as document types are not fully supported", moduleName)
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
    tempPath = Printf("{0}\{1}{2}", Environ$("temp"), result.module_name, result.extension)
    codeModule.Export tempPath
    On Error GoTo safeExit
    result.code_content = Chunkify(ToBase64(ReadBytes(tempPath))) 'encode and chunkify
    
safeExit:
    Kill tempPath
moduleMissing:
    ModuleDefinition = result
    If Err.Number <> 0 Then ModuleDefinition.extension = "missing"
    
End Function

Private Function Printf(ByVal mask As String, ParamArray tokens()) As String
    'Debug.Print , " -> Formatting"; Len(tokens(0)); "chars into", """"; mask; """"
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    Printf = mask
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
        .Remove .Item(moduleName)
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
        .InsertLines 8, "Private Const TypeBinary = 1, vbext_pp_none = 0"
        .InsertLines 9, "Private Const ForReading = 1, ForWriting = 2, ForAppending = 8"
        .InsertLines 10, ""
        .InsertLines 11, "Private Function getCodeDefinition(itemNo As Long) As codeItem"
        .InsertLines 12, "    With getCodeDefinition"
        .InsertLines 13, "        Select Case itemNo"
        .InsertLines 14, "            '{0}"
        .InsertLines 15, "        Case Else"
        .InsertLines 16, "            .extension = ""missing"""
        .InsertLines 17, "        End Select"
        .InsertLines 18, "    End With"
        .InsertLines 19, "End Function"
        .InsertLines 20, ""
        .InsertLines 21, "Public Sub Extract()"
        .InsertLines 22, "    Dim code_module As codeItem"
        .InsertLines 23, "    Dim savedPath As String, basePath As String"
        .InsertLines 24, "    Dim i As Long"
        .InsertLines 25, "    'check if vbproject accessible"
        .InsertLines 26, "    If Not project_accessible Then"
        .InsertLines 27, "        MsgBox ""The VBA project cannot be accessed programmatically"""
        .InsertLines 28, "        Exit Sub"
        .InsertLines 29, "    End If"
        .InsertLines 30, "    'check if temp folder acessible"
        .InsertLines 31, "    i = 0"
        .InsertLines 32, "    basePath = Environ$(""Temp"") & ""\"""
        .InsertLines 33, "    Do While True"
        .InsertLines 34, "        i = i + 1"
        .InsertLines 35, "        code_module = getCodeDefinition(i)"
        .InsertLines 36, "        If code_module.extension = ""missing"" Then"
        .InsertLines 37, "            Exit Do"
        .InsertLines 38, "        Else"
        .InsertLines 39, "            savedPath = createFile(code_module, basePath)"
        .InsertLines 40, "            importFile savedPath"
        .InsertLines 41, "            Kill savedPath"
        .InsertLines 42, "        End If"
        .InsertLines 43, "    Loop"
        .InsertLines 44, "    removemodule ""{1}"""
        .InsertLines 45, "End Sub"
        .InsertLines 46, ""
        .InsertLines 47, "Private Function project_accessible() As Boolean"
        .InsertLines 48, "    On Error Resume Next"
        .InsertLines 49, "    With thisworkbook.VBProject"
        .InsertLines 50, "        project_accessible = .Protection = vbext_pp_none"
        .InsertLines 51, "        project_accessible = project_accessible And Err.Number = 0"
        .InsertLines 52, "    End With"
        .InsertLines 53, "End Function"
        .InsertLines 54, ""
        .InsertLines 55, "Private Function createFile(definition As codeItem, filePath As Variant) As String"
        .InsertLines 56, "    Dim codeIndex As Long"
        .InsertLines 57, "    Dim newFileObj As Object"
        .InsertLines 58, "    Set newFileObj = CreateObject(""ADODB.Stream"")"
        .InsertLines 59, "    newFileObj.Type = TypeBinary"
        .InsertLines 60, "    'Open the stream and write binary data"
        .InsertLines 61, "    newFileObj.Open"
        .InsertLines 62, "    'create file from x64 string"
        .InsertLines 63, "    With definition"
        .InsertLines 64, "        Dim bytes() As Byte"
        .InsertLines 65, "        Dim fullPath As String"
        .InsertLines 66, "        fullPath = filePath & .module_name & .extension"
        .InsertLines 67, "        bytes = FromBase64(Join(.code_content))"
        .InsertLines 68, "        newFileObj.Write bytes"
        .InsertLines 69, "        newFileObj.SaveToFile fullPath, ForWriting"
        .InsertLines 70, "        createFile = fullPath"
        .InsertLines 71, "    End With"
        .InsertLines 72, "End Function"
        .InsertLines 73, ""
        .InsertLines 74, "Private Sub importFile(filePath As String)"
        .InsertLines 75, "    thisworkbook.VBProject.VBComponents.Import filePath"
        .InsertLines 76, "End Sub"
        .InsertLines 77, ""
        .InsertLines 78, "Private Function removemodule(moduleName As String) As Boolean"
        .InsertLines 79, "    On Error Resume Next"
        .InsertLines 80, "    With thisworkbook.VBProject.VBComponents"
        .InsertLines 81, "        .Remove .Item(moduleName)"
        .InsertLines 82, "    End With"
        .InsertLines 83, "    removemodule = Not (Err.Number = 9)"
        .InsertLines 84, "End Function"
        .InsertLines 85, ""
        .InsertLines 86, "Private Function FromBase64(Text As String) As Byte()"
        .InsertLines 87, "    Dim Out() As Byte"
        .InsertLines 88, "    Dim b64(0 To 255) As Byte, str() As Byte, i&, j&, v&, b0&, b1&, b2&, b3&"
        .InsertLines 89, "    Out = """""
        .InsertLines 90, "    If Len(Text) Then Else Exit Function"
        .InsertLines 91, ""
        .InsertLines 92, "    str = "" ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"""
        .InsertLines 93, "    For i = 2 To UBound(str) Step 2"
        .InsertLines 94, "        b64(str(i)) = i \ 2"
        .InsertLines 95, "    Next"
        .InsertLines 96, ""
        .InsertLines 97, "    ReDim Out(0 To ((Len(Text) + 3) \ 4) * 3 - 1)"
        .InsertLines 98, "    str = Text & String$(2, 0)"
        .InsertLines 99, ""
        .InsertLines 100, "    For i = 0 To UBound(str) - 7 Step 2"
        .InsertLines 101, "        b0 = b64(str(i))"
        .InsertLines 102, ""
        .InsertLines 103, "        If b0 Then"
        .InsertLines 104, "            b1 = b64(str(i + 2))"
        .InsertLines 105, "            b2 = b64(str(i + 4))"
        .InsertLines 106, "            b3 = b64(str(i + 6))"
        .InsertLines 107, "            v = b0 * 262144 + b1 * 4096& + b2 * 64& + b3 - 266305"
        .InsertLines 108, "            Out(j) = v \ 65536"
        .InsertLines 109, "            Out(j + 1) = (v \ 256&) Mod 256"
        .InsertLines 110, "            Out(j + 2) = v Mod 256"
        .InsertLines 111, "            j = j + 3"
        .InsertLines 112, "            i = i + 6"
        .InsertLines 113, "        End If"
        .InsertLines 114, "    Next"
        .InsertLines 115, ""
        .InsertLines 116, "    If b2 = 0 Then"
        .InsertLines 117, "        Out(j - 3) = (v + 65) \ 65536"
        .InsertLines 118, "        j = j - 2"
        .InsertLines 119, "    ElseIf b3 = 0 Then"
        .InsertLines 120, "        Out(j - 3) = (v + 1) \ 65536"
        .InsertLines 121, "        Out(j - 2) = ((v + 1) \ 256&) Mod 256"
        .InsertLines 122, "        j = j - 1"
        .InsertLines 123, "    End If"
        .InsertLines 124, ""
        .InsertLines 125, "    ReDim Preserve Out(j - 1)"
        .InsertLines 126, "    FromBase64 = Out"
        .InsertLines 127, "End Function"
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




