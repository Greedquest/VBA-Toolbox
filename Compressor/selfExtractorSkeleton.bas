Attribute VB_Name = "selfExtractorSkeleton"
Option Explicit
Private Type codeItem
    extension As String
    module_name As String
    code_content() As String
End Type

Private Const TypeBinary As Long = 1
Private Const vbext_pp_none As Long = 0
Private Const ForReading As Long = 1, ForWriting As Long = 2, ForAppending As Long = 8

Private Function getCodeDefinition(ByVal itemNo As Long) As codeItem
    With getCodeDefinition
        Select Case itemNo
            '{0}
        Case Else
            .extension = "missing"
        End Select
    End With
End Function

Public Sub Extract()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim code_module As codeItem
    Dim savedPath As String, basePath As String
    Dim i As Long
    'check if vbproject accessible
    If Not ProjectAccessible(wb) Then
        MsgBox "The VBA project cannot be accessed programmatically"
        Exit Sub
    End If
    'check if temp folder acessible
    i = 0
    basePath = Environ$("Temp") & "\"
    Do While True
        i = i + 1
        code_module = getCodeDefinition(i)
        If code_module.extension = "missing" Then
            Exit Do
        Else
            savedPath = createFile(code_module, basePath)
            importFile savedPath, wb
            Kill savedPath
        End If
    Loop
    RemoveModule "{1}", wb
End Sub

Private Function ProjectAccessible(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    With wb.VBProject
        ProjectAccessible = .Protection = vbext_pp_none
        ProjectAccessible = ProjectAccessible And Err.Number = 0
    End With
End Function

Private Function createFile(ByRef definition As codeItem, ByVal filePath As String) As String
    Dim newFileObj As Object
    Set newFileObj = CreateObject("ADODB.Stream")
    newFileObj.Type = TypeBinary
    'Open the stream and write binary data
    newFileObj.Open
    'create file from x64 string
    With definition
        Dim bytes() As Byte
        Dim fullPath As String
        fullPath = filePath & .module_name & .extension
        bytes = FromBase64(Join(.code_content))
        newFileObj.Write bytes
        newFileObj.SaveToFile fullPath, ForWriting
        createFile = fullPath
    End With
End Function

Private Sub importFile(ByVal filePath As String, ByRef wb As Workbook)
    wb.VBProject.VBComponents.Import filePath
End Sub

Private Function RemoveModule(ByVal moduleName As String, ByRef book As Workbook) As Boolean
    On Error Resume Next
    With book.VBProject.VBComponents
        .Remove .Item(moduleName)
    End With
    RemoveModule = Not (Err.Number = 9)
End Function

Private Function FromBase64(ByVal Text As String) As Byte()
    Dim Out() As Byte
    Dim b64(0 To 255) As Byte, str() As Byte, i As Long, j As Long, v As Long, b0 As Long, b1 As Long, b2 As Long, b3 As Long
    Out = vbNullString
    If Len(Text) Then Else Exit Function

    str = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    For i = 2 To UBound(str) Step 2
        b64(str(i)) = i \ 2
    Next

    ReDim Out(0 To ((Len(Text) + 3) \ 4) * 3 - 1)
    str = Text & String$(2, 0)

    For i = 0 To UBound(str) - 7 Step 2
        b0 = b64(str(i))

        If b0 Then
            b1 = b64(str(i + 2))
            b2 = b64(str(i + 4))
            b3 = b64(str(i + 6))
            v = b0 * 262144 + b1 * 4096& + b2 * 64& + b3 - 266305
            Out(j) = v \ 65536
            Out(j + 1) = (v \ 256&) Mod 256
            Out(j + 2) = v Mod 256
            j = j + 3
            i = i + 6
        End If
    Next

    If b2 = 0 Then
        Out(j - 3) = (v + 65) \ 65536
        j = j - 2
    ElseIf b3 = 0 Then
        Out(j - 3) = (v + 1) \ 65536
        Out(j - 2) = ((v + 1) \ 256&) Mod 256
        j = j - 1
    End If

    ReDim Preserve Out(j - 1)
    FromBase64 = Out
End Function
