VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "codeReviewTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests")
Option Explicit

'''
'This class creates and runs a new multithread instance which runs clsHtmlWorker
'When each HTMLDocument is complete, the class scans it for e-mails
'''
Private WithEvents multiThreadGroup As clsMultiThread
Attribute multiThreadGroup.VB_VarHelpID = -1
'clsMultiThread is async so must be  declared separately (or in a doEvents loop)
Private Const REGEX_PATTERN As String = _
"(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|""(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*"")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4])|1[0-9]|[0-9]?[0-9]))\.){3}(?:(2(5[1-9]|[0-9])|1[0-5]|[0-4]?[0-9])|[0-9]*[0-9]:(?:[1-9]|\\[0-9])+)\])"

Public Sub run()
    'urls to check for emails are in a1:a10
    htmlRequestToUrls [a1:a10]
End Sub

Private Sub htmlRequestToUrls(urlCells As Range)

    Set multiThreadGroup = New clsMultiThread
    With multiThreadGroup
        .Size = urlCells                         'set iterable, here a load of urls
        Set .AsyncClass = New clsHtmlWorker      'set async worker
        .Execute                                 'run the group
    End With

End Sub

Private Sub multiThreadGroup_TaskComplete(returnVal As Variant, taskID As String, threadID As String)
    
    Dim rowI As Long, colI As Long
    rowI = Right(taskID, Len(taskID) - 4)
    
    If returnVal Is Nothing Then
        Cells(rowI, 2) = "Error in loading page"
    ElseIf TypeOf returnVal Is HTMLDocument Then
        Dim emailMatches() As String
        emailMatches = regexMatches(returnVal.body.innerText)
        If (Not emailMatches) = -1 Then
        'no emails on page
            Cells(rowI, 2) = "No e-mail matches"
        Else
            For colI = LBound(emailMatches) To UBound(emailMatches)
                Cells(rowI, colI + 2) = emailMatches(colI)
            Next colI
        End If
    Else                                         'nothing returned
        Cells(rowI, 2) = "Error in loading page"
    End If


End Sub

Private Function regexMatches(strInput As String) As String()

    Dim rMatch As Object
    Dim s As String
    Dim arrayMatches() As String
    Dim i As Long

    With CreateObject("VBScript.Regexp")
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = REGEX_PATTERN
        If .test(strInput) Then
            For Each rMatch In .Execute(strInput)
                ReDim Preserve arrayMatches(i)
                arrayMatches(i) = rMatch.value
                i = i + 1
            Next
        End If
    End With

    regexMatches = arrayMatches

End Function

