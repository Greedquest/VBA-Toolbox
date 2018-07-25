VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMultiThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Main.Code")
Option Explicit

'''
'VBA class to run multiple asynchronous processes
'Interfaces directly with clsThreadHandle
'Requires references to:
'mscrolib.dll
'''

'THREAD GROUP SHAPE PROPERTIES
Private threadGroup As New Collection            'holds all the treads
Private maxThreads As Long                       'maximum number of threads that can be open
Private minThreads As Long                       '[minimum number of threads]
Private iterableQueue As mscorlib.Queue          'this item holds all the items from iterator set in queue
'replaces iterableGroup, newtaskindex, taskidset
Private iterableSize As Long                     'number of items in iterable group or
Private passesArguments As Boolean               'true if iterableGroup exists

'THREAD GROUP REFERENCES
Private WithEvents threadEvents As clsHandleEvents 'Event object to raise events from each thread handle
Attribute threadEvents.VB_VarHelpID = -1
Private workerClass As IWorker

'THREAD GROUP SETTINGS
Private autoQuitEnabled As Boolean               'dictates whether to quit on Complete event, should be false if daisychaining

'THREAD GROUP STATE PROPERTIES
Private openThreadCount As Long                  'number of threads/handles currently open
Private openTaskCount As Long                    'number of tasks running on those threads
Private closedTaskCount As Long                  'number of threads closed (failed and successful)
Private successfulTaskCount As Long              'number of threads completed sucessfully
Private newThreadIndex As Long                   'Iterator over handles (next new handle)
Private newTaskIndex As Long                     'Iterator over open tasks (next thread to be started)
Private taskIDset As Collection                  'Dictionary mapping taskIDs to iterableGroup location "REPLACE THIS. MERGE COLLECTION JUMBLES"
Private freeThreads As Collection                'holds all the free thread ids

'THREAD GROUP PERFORMANCE PROPERTIES
Private startTime As Date
'Private endTime As Date

'THREAD GROUP EVENTS
Public Event TaskComplete(returnVal As Variant, taskID As String, threadID As String) 'when a task is complete on a thread, maybe if failed
Public Event ThreadOpened(threadCount As Long, threadID As String) 'when a thread is opened, pass the new number of threads
Public Event ThreadClosed(threadCount As Long, threadID As String) 'when closed, pass closed thread ID
Public Event Complete(timeTaken As Date)         'when everything is (nearly) finished
Public Event Closed(timeTaken As Date)           'when entire group is closed
Public Event Opened(startTime As Date)           'when entire group is closed

'PRIVATE TYPES/ENUMS
Private Type Instruction                         'instruction on what to do next, and any necessary arguments that can be passed
    threadID As String
    instructionBody As InstructionType
End Type

Private Enum InstructionType
    mltCloseThread
    mltOpenThread
    mltSetTask
    mltDoNothing
    mltQuit
End Enum

Private Sub Class_Initialize()
    'Set defaults
    maxThreads = 5
    minThreads = 1
    newThreadIndex = 1
    newTaskIndex = 1
    autoQuitEnabled = True
    Set threadEvents = New clsHandleEvents
    Set taskIDset = New Collection
    Set freeThreads = New Collection
    startTime = Now
    RaiseEvent Opened(startTime)
    '''
    'Test space
    '''
End Sub

Private Sub threadEvents_Closed(threadID As String)
    RaiseEvent ThreadClosed(openThreadCount, threadID)
End Sub

Private Sub threadEvents_Opened(threadID As String)
    RaiseEvent ThreadOpened(openThreadCount, threadID)
End Sub

Private Sub threadEvents_Complete(obj As clsThreadHandle, returnVal As Variant) 'called when thread becomes free
    'DO NOT mark as free here
    RaiseEvent TaskComplete(returnVal, obj.Task, obj.Name) 'failed as boolean
    openTaskCount = openTaskCount - 1
    closedTaskCount = closedTaskCount + 1
    successfulTaskCount = successfulTaskCount + 1 'could be unsuccessful too though
    doInstructions obj.Name                      'pass object name so it can be marked free
    '    If failed Then
    '        failedTaskCount = failedTaskCount + 1
    '    Else
    '        successfulTaskCount = successfulTaskCount + 1
    '    End If
End Sub

Public Sub Execute()
    'check validity of user data, if valid, then execute task
    If iterableSize = 0 Then
        Err.Raise 5, Description:="You must set size argument to a non-zero value, or a non-empty iterable first"
    ElseIf workerClass Is Nothing Then
        Err.Raise 5, Description:="You must set the async class argument first"
    Else
        doInstructions
    End If
End Sub

Public Sub Quit()
    'Remove any references that would prevent proper closing
    'Default automatically called when openThreadCount = 0
    RaiseEvent Complete(Now - startTime)
    Set threadEvents = Nothing
End Sub

Private Sub doInstructions(Optional freeThreadID As String, Optional loopcount As Long = 1)
    Dim instructionVal As Instruction

    'mark thread free if applicable
    If freeThreadID <> vbNullString Then freeThread = freeThreadID

    'find out what to do
    instructionVal = getInstruction()
    'carry out instruction
    Select Case instructionVal.instructionBody
    Case InstructionType.mltCloseThread
        closeThread instructionVal.threadID
    Case InstructionType.mltOpenThread
        openThread
    Case InstructionType.mltSetTask
        Dim taskThread As clsThreadHandle
        Dim taskArguments As Variant
        Set taskThread = threadGroup(instructionVal.threadID)
        'assign task to thread
        assignTaskID (taskThread.Name)
        'get any arguments there may be
        'mark thread as busy

        BusyThread = taskThread.Name
        'iterate open tasks
        openTaskCount = openTaskCount + 1
        'execute task
        If passesArguments Then
            'pop appropriate item from queue
            Set taskArguments = iterableQueue.Dequeue
            taskThread.Execute taskArguments
        Else
            taskThread.Execute
        End If

    Case InstructionType.mltQuit
        'quit then do nothing
        Me.Quit
        instructionVal.instructionBody = mltDoNothing
    Case InstructionType.mltDoNothing
        'do nothing
    Case Else
        Err.Raise 5                              'invalid argument
    End Select

    'call self until no instruction
    If instructionVal.instructionBody <> mltDoNothing Then
Debug.Assert loopcount < maxThreads * 3 + 5      'max loop should be open all threads then run all tasks + a little
        doInstructions loopcount:=loopcount + 1  'watch for infinite loop
    End If

End Sub

Private Function getInstruction() As Instruction
    'function to determine what action to take next
    'called until do nothing returned
    'caller to doinstructions can specify a free thread in which case some parts skipped
    Dim results As Instruction                   'variable to hold instruction and any arguments

    Me.printState

    'Do we need to open or close threads?
    'Threads free? (threads open > tasks open):
    If openThreadCount > openTaskCount Then
        'Great we have a free thread, now use it or delete it (cos we have too many or no tasks remaining)
        If newTaskIndex > iterableSize Then      'already passed all tasks
            '[find] & close free thread
            results.instructionBody = mltCloseThread
            results.threadID = freeThread
        ElseIf openThreadCount <= maxThreads Then
            '[find] & use free thread (run a task on it)
            results.instructionBody = mltSetTask
            results.threadID = freeThread
        Else
            '[find] & close free thread
            results.instructionBody = mltCloseThread
            results.threadID = freeThread
        End If
    Else
        'No threads free, either open one (if not exceeding max, and there's a task left to put on it)
        'Or do nothing (can't close it if not free, shouldn't open new if no more tasks)
        If openThreadCount < maxThreads And newTaskIndex <= iterableSize Then
            results.instructionBody = mltOpenThread
        ElseIf openThreadCount = 0 And autoQuitEnabled Then
            results.instructionBody = mltQuit
        Else
            results.instructionBody = mltDoNothing
        End If
    End If

    getInstruction = results
End Function

Private Sub openThread()
    'opens a thread and assigns a task ID to it
    Dim newThread As New clsThreadHandle         'create new handle
    newThread.OpenHandle Me, threadEvents        'passes parent reference which allows handle to obtain thread ID
    threadGroup.Add newThread, newThread.Name    'add it to the group with a new id (set by itself)
    openThreadCount = openThreadCount + 1
    freeThread = newThread.Name                  'mark as free so task can be assigned to it

End Sub

Private Property Let freeThread(threadID As String)
    'NOT WORKING"""""
    'when a thread comes free, add it to the collection
    freeThreads.Add threadID, threadID
Debug.Print threadID; " marked as free; now"; freeThreads.Count; "threads are free"
End Property

Private Property Let BusyThread(threadID As String)
    'when a thread is not free or is closed, mark as busy by removing from free group
    On Error Resume Next                         'only remove ones what are there actually
    freeThreads.Remove threadID
Debug.Print threadID; " marked as busy"; IIf(Err.Number <> 0, ", but wasn't in free group", vbNullString)
End Property

Private Property Get freeThread() As String
    'gives up a free thread and adds it to the list
    freeThread = freeThreads(1)
    freeThreads.Remove (1)
End Property

Private Sub assignTaskID(threadID As String)
    '@Ignore WriteOnlyProperty
    'assigns task ID to thread
    'nb does NOT actually run the task (this is instruction stage still)
    Dim newThread As clsThreadHandle
    Set newThread = threadGroup(threadID)
    newThread.Task = NewTaskID
    Set newThread.Worker = AsyncClass
End Sub

Private Sub closeThread(threadID As String, Optional failed As Boolean = False)

    'close thread with appropriate id
    Dim oldThread As clsThreadHandle
    Set oldThread = threadGroup(threadID)
    'remove from all collections
    'taskIDset.Remove oldThread.Task remove from task id set if it was in there
    threadGroup.Remove oldThread.Name
    BusyThread = oldThread.Name                  'remove from free collection
    Set oldThread = Nothing
    'iterate counters
    openThreadCount = openThreadCount - 1
End Sub

Public Property Let Size(sizeFactor As Variant)
    'property of the thread group which dictates how many processes to run in total
    'size factor is either an iterable item, or an integer to dictate the size

    'Check if size factor is number
    If IsNumeric(sizeFactor) Then
        'If so, size is that
        iterableSize = CLng(sizeFactor)
        passesArguments = False                  'no argument to pass to thread, just run it a load of times

        'If not, *check if iterable
    ElseIf isIterable(sizeFactor) Then
        'If so, size is size of collection from extration
        Set iterableQueue = New Queue
        iterableSize = addIterableToQueue(sizeFactor, iterableQueue)
        passesArguments = True
    Else
        '[if not, raise error]
        Err.Raise 5                              'invalid argument
    End If

End Property

Public Sub IncreaseSize(sizeFactor As Variant)
    'method of threadGroup which adds more tasks to the queue, and immediately runs them
    'size factor is either an iterable item, or an integer to dictate the size

    'Check whether size is set yet
    If Me.Size = 0 Then
        Err.Raise 5, Description:="You must set Size before you can IncreaseSize"
    End If

    'check whether new data matches old type
    If IsNumeric(sizeFactor) Then
        If passesArguments Then
            Err.Raise 5, Description:="Size factor type doesn't match original type"
        Else
            'is numeric and was numeric, grand
            iterableSize = iterableSize + CLng(sizeFactor)
        End If
    ElseIf isIterable(sizeFactor) Then
        If passesArguments Then
            'was iterable and still is, great!
            Dim itemsAdded As Long
            itemsAdded = addIterableToQueue(sizeFactor, iterableQueue)
            iterableSize = iterableSize + itemsAdded
        Else
            'wasn't iterble, now is
            Err.Raise 5, Description:="Size factor type doesn't match original type"
        End If

    Else
        '[if not, raise error]
        Err.Raise 5                              'invalid argument
    End If
    Me.Execute
End Sub

Public Property Set AsyncClass(ByVal workObj As IWorker) 'Set the worker who carries out the tasks
    Set workerClass = workObj
End Property

Public Property Get AsyncClass() As IWorker
    Set AsyncClass = workerClass
End Property

Public Property Get Size() As Variant
    Size = iterableSize
End Property

Public Property Let autoQuit(ByVal value As Boolean)
    autoQuitEnabled = value
End Property

Public Property Get NewHandleID() As String
    NewHandleID = "Handle " & newThreadIndex
    newThreadIndex = newThreadIndex + 1          'use next one next time
End Property

Private Property Get NewTaskID() As String
    'generates new task, saves its ID to taskIDset, then bumps the task counter along one
    NewTaskID = "Task " & newTaskIndex
    taskIDset.Add newTaskIndex, NewTaskID        'add id to map
    newTaskIndex = newTaskIndex + 1
End Property

Private Sub Class_Terminate()
    'Set threadGroup = Nothing
Debug.Print "Terminating group"
    RaiseEvent Closed(Now - startTime)
End Sub

Public Sub printState()                          'for debugging
Debug.Print _
        "State:"; vbCrLf _
                  ; Space(5); "Threads open: "; openThreadCount; vbCrLf _
                  ; Space(5); "Threads in use: "; openTaskCount; vbCrLf _
                  ; Space(5); "Threads marked as free: "; freeThreads.Count; vbCrLf _
                  ; Space(5); "Tasks remaining: "; iterableSize - successfulTaskCount; vbCrLf _
                  ; Space(5); "Next task index: "; newTaskIndex
End Sub

