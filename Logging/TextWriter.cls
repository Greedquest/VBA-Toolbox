VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Mathieu Guindon CR https://codereview.stackexchange.com/q/52185/146810"
'@Exposed
'@Folder("Logger.Utils.TextWriter")
'@ModuleDescription("Mathieu Guindon CR https://codereview.stackexchange.com/q/52185/146810")
'@PredeclaredID

'expose raised errors to clients:
Public Enum FileWriterError
    InvalidFileName = vbObjectError + 42
End Enum

'manage opened files in a Dictionary:
Private openedFiles As Dictionary

'skip dictionary lookup if only 1 file is opened:
Private quickWriteFile As Long

Option Explicit

Public Property Get OpenedFilesCount() As Integer
    OpenedFilesCount = openedFiles.Count
End Property

Public Function OpenFile(ByVal fileName As String, Optional ByVal overwrite As Boolean = False) As Boolean

    Dim fileNumber As Long
    fileNumber = GetFileNumber(fileName)

    'guard against opening a file that's already opened:
    If fileNumber <> FreeFile Then
        OpenFile = True
        Exit Function
    End If

    On Error GoTo Catch

    If overwrite Or Dir(fileName) = vbNullString Then
        Open fileName For Output As #fileNumber
    Else
        Open fileName For Append As #fileNumber
    End If

    openedFiles.Add fileName, fileNumber
    quickWriteFile = IIf(openedFiles.Count = 1, fileNumber, 0)

Catch:
    If Err.Number <> 0 Then
        Err.Clear
    End If

'    OpenFile = (openedFiles.ContainsKey(fileName))
    OpenFile = (openedFiles.Exists(fileName))

End Function

Public Sub WriteLine(ByVal data As String, Optional ByVal fileName As String = vbNullString)
    Dim fileNumber As Long
    Dim result As Boolean

    On Error GoTo Catch

    If CanWrite(fileName, fileNumber) Then
        Print #fileNumber, data
        result = True
    Else
        Err.Raise FileWriterError.InvalidFileName, TypeName(Me) & ".WriteLine", "Invalid filename."
    End If

Catch:
    If Err.Number <> 0 Then
        result = False
        'NOTE: patch
        'openedFiles.Remove fileNumber
        Patches.RemoveByValue fileNumber, openedFiles
        'End Patch
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If

End Sub

Private Function CanWrite(ByVal fileName As String, ByRef outFileNumber As Long) As Boolean
    Dim result As Boolean
    Dim fileNumber As Long
    Dim proceed As Boolean

    If quickWriteFile <> 0 And fileName = vbNullString Then
        fileNumber = quickWriteFile
        CanWrite = True
    Else
        CanWrite = openedFiles.TryGetValue(fileName, fileNumber)
    End If

    outFileNumber = fileNumber

End Function

Public Sub CloseFile(Optional ByVal fileName As String = vbNullString)
    If openedFiles.Count = 0 Then Exit Sub

    Dim fileNumber As Long
    fileNumber = GetFileNumber(fileName)

    If fileNumber <> FreeFile Then
        Close #fileNumber
        openedFiles.Remove fileNumber
        If fileNumber = quickWriteFile Then quickWriteFile = 0
    End If

    If openedFiles.Count = 1 Then quickWriteFile = openedFiles.Values.First

End Sub

Public Sub CloseAllFiles()
    'NOTE: Patch
'    Dim file As KeyValuePair
'    For Each file In openedFiles
'        Close #file.value
'    Next
'    openedFiles.Clear
'    quickWriteFile = 0
    Dim file As Variant
    For Each file In openedFiles.Items
        Close #file
    Next
    openedFiles.RemoveAll
    quickWriteFile = 0
    'End Patch
End Sub

Private Function GetFileNumber(ByVal fileName As String) As Long
    Dim result As Long

    If quickWriteFile <> 0 And fileName = vbNullString Then
        result = quickWriteFile

    ElseIf Not openedFiles.TryGetValue(fileName, result) Then
        result = FreeFile

    End If

    GetFileNumber = result
End Function

Public Function Create() As TextWriter
    Set Create = New TextWriter
End Function

Private Sub Class_Initialize()
    Set openedFiles = New Dictionary
End Sub

Private Sub Class_Terminate()
    CloseAllFiles
    Set openedFiles = Nothing
End Sub

