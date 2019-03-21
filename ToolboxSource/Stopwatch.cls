VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder Stopwatch

Private Type TStopWatch
    data As Object
    CurrentLabel As LabelTree
    Results As StopwatchResults
    FirstLabel As LabelTree
End Type

Private this As TStopWatch

Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Function MicroTimer() As Double
    'Accurate timing method - stackoverflow.com/a/7116928/6609896
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency

    MicroTimer = 0

    If cyFrequency = 0 Then getFrequency cyFrequency

    getTickCount cyTicks1

    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
End Function

Public Sub start()
    OpenLabel "Start"
End Sub

Public Sub Finish()
    CloseLabel
    Set this.Results = New StopwatchResults
    this.Results.LoadData this.data, this.FirstLabel
End Sub

Public Property Get Results() As StopwatchResults
    Set Results = this.Results
End Property

Public Sub OpenLabel(ByVal labelName As String)
    'Save time on arrival
    Dim clockTimes As New TimeInfo
    clockTimes.TimeIn = MicroTimer

    'Define new label, and make it a child of the current label
    Dim newNode As New LabelTree
    newNode.NodeName = labelName
    If Not this.CurrentLabel Is Nothing Then
        Set newNode.parentNode = this.CurrentLabel
        '1.2.1 format
        newNode.Location = this.CurrentLabel.Location & "." & this.CurrentLabel.ChildNodes.Count + 1
        this.CurrentLabel.ChildNodes.Add newNode, newNode.Location & newNode.NodeName
    Else
        newNode.Location = "1"
        Set this.FirstLabel = newNode
    End If
    Set this.CurrentLabel = newNode

    'Save time data to dictionary and return to execution
    Dim dictKey As String
    dictKey = newNode.Location & "_open"
    this.data.Add dictKey, clockTimes
    this.data(dictKey).TimeOut = MicroTimer
End Sub

Public Sub CloseLabel()
    'Save time on arrival
    Dim clockTimes As New TimeInfo
    clockTimes.TimeIn = MicroTimer

    'Save time data to dictionary and return to execution
    Dim dictKey As String
    dictKey = this.CurrentLabel.Location & "_close"
    this.data.Add dictKey, clockTimes

    'Close label by setting to parent
    Set this.CurrentLabel = this.CurrentLabel.parentNode
    this.data(dictKey).TimeOut = MicroTimer
End Sub

Public Sub Lap()
    'Save time on arrival
    Dim clockTimes As New TimeInfo
    clockTimes.TimeIn = MicroTimer

    'Define new label, and make it a child of the current label
    Dim newNode As New LabelTree
    newNode.Location = this.CurrentLabel.Location & "." & this.CurrentLabel.ChildNodes.Count + 1
    newNode.NodeName = "Lap" & this.CurrentLabel.ChildNodes.Count + 1 'this.CurrentLabel.NodeName & "_
    newNode.LabelType = stp_LapTime

    If this.CurrentLabel Is Nothing Then
        Err.Description = "No test is currently running to write lap data to"
        Err.Raise 5
    Else
        Set newNode.parentNode = this.CurrentLabel
        this.CurrentLabel.ChildNodes.Add newNode, newNode.NodeName
    End If


    'Save time data to dictionary and return to execution
    Dim dictKey As String
    dictKey = this.CurrentLabel.Location & "_" & newNode.NodeName
    this.data.Add dictKey, clockTimes
    this.data(dictKey).TimeOut = MicroTimer
End Sub

Private Sub Class_Initialize()
    Set this.data = CreateObject("Scripting.Dictionary")
End Sub


