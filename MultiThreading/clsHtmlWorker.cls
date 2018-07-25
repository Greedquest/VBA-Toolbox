VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHtmlWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests")
Option Explicit
'''
'Basic worker object sends MSHTML GET request to webpage and returns an HTMLDocument or Nothing
'Requires reference to
' Microsoft HTML Object library (mshtml.tlb)
' Microsoft XML, v6.0 (msxml6.dll)
'''

Private httpRequest As MSXML2.XMLHTTP60

Implements IWorker

Private Type TWorker
    Events As IWorkerEvents
End Type

Private this As TWorker

Private Function IWorker_CreateNew() As IWorker
    Set IWorker_CreateNew = New clsHtmlWorker
End Function

Private Property Set IWorker_Events(RHS As IWorkerEvents)
    Set this.Events = RHS
End Property

Private Sub IWorker_Execute(Optional argument As Variant)
    Started                                      'raise event to thread handle
    'Do some task
    sendRequest argument
End Sub

'''
'Event raising
'''

Private Sub Started()
    If Not this.Events Is Nothing Then
        this.Events.Started
    End If
End Sub

Private Sub statusChange(ByVal statusText As String)
    If Not this.Events Is Nothing Then
        'status change is not fully implemented yet in clsMultiThread, I may get rid of it
        this.Events.statusChange statusText
    End If
End Sub

Private Sub Complete(Optional ByVal resultPage As HTMLDocument)
    If Not httpRequest Is Nothing Then Set httpRequest = Nothing
    If Not this.Events Is Nothing Then
        this.Events.Complete resultPage
    End If
End Sub

Private Sub sendRequest(ByVal url As String)
    '''
    'Sub to open a new XMLHTTP request at a given url
    'Also assigns OnReadyStateChange callback function to this class' default routine
    '''

    If httpRequest Is Nothing Then Set httpRequest = New MSXML2.XMLHTTP60

    With httpRequest
        'Assign callback function to handler class (by default property)
        .OnReadyStateChange = Me
        'open and send the request
        .Open "GET", url, True
        .send vbNullString
    End With

End Sub

Public Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
    '''
    'This is the default callback routine of the class
    '''
    With httpRequest

        statusChange .statusText

        If .ReadyState = 4 Then                  'loaded
            If .Status = 200 Then                'successful
                'mark complete and pass document
                Dim htmlDoc As HTMLDocument
                Set htmlDoc = New HTMLDocument
                htmlDoc.body.innerHTML = .responseText
                Complete htmlDoc
            Else                                 'unsuccessful
                Complete
            End If
        End If
    End With
End Sub

Private Sub Class_Terminate()
    If Not httpRequest Is Nothing Then Set httpRequest = Nothing
End Sub
