VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsThreadHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Main.Code")
Option Explicit


'THREAD HANDLE BASE PROPERTIES
Private eventHandle As clsHandleEvents           'Events module multithread set which handle belongs to. Called when handle state changes
Private taskID As String                         'holds the id of the current task
Private handleID As String                       'holds the id of this handle
Private handleArgs As Variant                    'holds any arguments that need to be passed to the task

'THREAD EVENTS
Private WithEvents workerEvents As IWorkerEvents
Attribute workerEvents.VB_VarHelpID = -1
Private workerObject As IWorker                  'interface to whatever worker may be passed to thread


Private Sub workerEvents_Complete(returnVal As Variant)
    eventHandle.NotifyComplete Me, returnVal
End Sub

Private Sub workerEvents_Started()
Debug.Print Me.Task; " started event was raised"
End Sub

Public Property Set Worker(ByVal workObj As IWorker)
    Set workerObject = workObj.CreateNew         'set worker to be a copy of the passed one
    Set workerEvents = New IWorkerEvents         'create event handler
    Set workerObject.Events = workerEvents       'pass it to the worker so it can listen in
End Property

Public Sub OpenHandle(multiThreadGroup As clsMultiThread, delegate As clsHandleEvents)
    'called when the handle is opened, sets the reference IDs of the string and the handle, as well as parent g
    Set eventHandle = delegate
    handleID = multiThreadGroup.NewHandleID
    eventHandle.NotifyThreadOpened (Name)
Debug.Print Name; " was opened"
End Sub

Public Sub Execute(Optional args As Variant)
Debug.Print Task; " executed on "; Name; " with "; IIf(IsMissing(args), "no arguments", "some arguments")

    workerObject.Execute args                    'run the event
End Sub

Public Property Get Task() As String
    Task = taskID
End Property

Public Property Let Task(val As String)
    taskID = val
Debug.Print Name; "'s task was set to "; taskID
End Property

Public Property Get Name() As String
    Name = handleID
End Property

Private Sub Class_Initialize()
Debug.Print "I'm made"
End Sub

Private Sub Class_Terminate()
    eventHandle.NotifyThreadClosed (Me.Name)
    Set eventHandle = Nothing
    Set workerObject = Nothing
End Sub

Private Sub workerEvents_StatusChange(statusVal As Variant)
'not yet implemented, probably unnecessary
End Sub
