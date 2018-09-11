VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopwatchResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'@Folder Stopwatch

Private Type TStopWatchResults
    TimeData As Object
    LabelData As LabelTree
End Type

Private this As TStopWatchResults

Public Sub LoadData(ByVal TimeData As Object, ByVal LabelData As LabelTree)
    Set this.LabelData = LabelData
    Set this.TimeData = TimeData
    writeTimes this.LabelData
End Sub

Public Property Get ToLabelTree() As LabelTree
    Set ToLabelTree = this.LabelData
End Property

Public Property Get RawData() As Object
    Set RawData = this.TimeData
End Property

Public Sub ToImmediateWindow()
'Prints time info to immediate window
    Dim resultsTree As LabelTree
    Set resultsTree = this.LabelData
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    flattenTree resultsTree, dict
    Debug.Print "Label name", "Time taken"
    Debug.Print String(35, "-")
    Dim value As Variant
    For Each value In dict.Keys
        Debug.Print value, dict(value)(0), dict(value)(1)
    Next value
End Sub

Private Sub flattenTree(ByVal treeItem As LabelTree, ByRef dict As Object, Optional ByVal depth As Long = 0)
'recursively converts a results tree to a dictionary of result keys
    dict.Add printf("{0} {1}", treeItem.Location, treeItem.NodeName), Array(treeItem.TimeSpent, treeItem.TimeWasted)
    If treeItem.ChildNodes.Count > 0 Then
        Dim item As Variant
        For Each item In treeItem.ChildNodes
            flattenTree item, dict, depth + 1
        Next
    End If
End Sub

Private Sub writeTimes(ByRef labelItem As LabelTree)
    'Recursively write absolute time data to time elapsed data

    Dim startTimes As TimeInfo
    Dim endTimes As TimeInfo

    setTimeStamps labelItem, startTimes, endTimes 'get timestamps from dictionary
    With labelItem
        If .ChildNodes.Count > 0 Then
            'has children, work out time spent for each then sum
            Dim childLabel As LabelTree
            Dim item As Variant

            For Each item In .ChildNodes         'recurse deeper
                Set childLabel = item
                writeTimes childLabel
                .TimeWasted = .TimeWasted + childLabel.TimeWasted 'add up child wasted time
            Next item
            .TimeSpent = endTimes.TimeIn - startTimes.TimeOut - .TimeWasted 'time diff - wasted time
            .TimeWasted = .TimeWasted + endTimes.TimeOut - endTimes.TimeIn + startTimes.TimeOut - startTimes.TimeIn
        Else                                     'No children
            If .LabelType = stp_LapTime Then
                .TimeWasted = endTimes.TimeOut - endTimes.TimeIn
            Else                                 'find time stamps for opening and closing label
                .TimeWasted = endTimes.TimeOut - endTimes.TimeIn + startTimes.TimeOut - startTimes.TimeIn
            End If
            .TimeSpent = endTimes.TimeIn - startTimes.TimeOut
        End If
    End With
End Sub

Private Sub setTimeStamps(ByVal labelItem As LabelTree, ByRef startTimes As TimeInfo, ByRef endTimes As TimeInfo)
    'writes timestamps byref
    With labelItem
        Dim startKey As String
        Dim endKey As String
        'location of timestamps in dictionary
        Select Case .LabelType
        Case stp_LapTime
            Dim keyBase As String
            keyBase = .parentNode.Location
            Dim lapNumber As Long
            lapNumber = Right$(.NodeName, Len(.NodeName) - 3)
            If lapNumber = 1 Then                'first lap, starts at
                startKey = printf("{0}_open", keyBase)
            Else
                startKey = printf("{0}_Lap{1}", keyBase, lapNumber - 1) 'start at prev lap, end here
            End If
            endKey = printf("{0}_Lap{1}", keyBase, lapNumber)
        Case Else
            startKey = printf("{0}_open", .Location)
            endKey = printf("{0}_close", .Location)
        End Select
        Set endTimes = this.TimeData(endKey)
        Set startTimes = this.TimeData(startKey)
    End With

End Sub

Private Function printf(ByVal mask As String, ParamArray tokens()) As String
'Format string with by substituting into mask - stackoverflow.com/a/17233834/6609896
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    printf = mask
End Function
