Attribute VB_Name = "BufferRDTests"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider
Private Buffer As clsBuffer
Private BufferEvents As BufferEventTests

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
    Set Buffer = New clsBuffer
    Set BufferEvents = New BufferEventTests
    Set BufferEvents.Buffer = Buffer
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
    Set Buffer = New clsBuffer
    BufferEvents.ClearCounts
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub TestAdding()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    Buffer.AddItems Array(1, New ArrayList)
    'Assert:
    
    Assert.AreEqual UBound(Buffer.AddedItems), "2", "Item not added correctly"
    Assert.AreEqual UBound(Buffer.AddedItems), "0", "Queue not cleared correctly"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub testAddingError()
    Const ExpectedError As Long = 5              'TODO Change to expected error number
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    Buffer.AddItems 1                            'add non iterable
    
Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'
'Sub testRemoval()
'    Dim testList As Object: Set testList = CreateObject("System.Collections.ArrayList")
'    For i = 1 To 5
'        testList.Add Cells(1, i)
'    Next i
'    Dim testObj As New clsBuffer
'    Dim item1 As Range: Set item1 = testList(1)  'byref, all point the same way
'    Dim item2 As Range: Set item2 = testList(2)
'    testObj.AddItems testList
'    testObj.RemoveItems Array(item1, item2)
'    Dim markedItems As Variant
'    markedItems = testObj.RemovalItems
'    For i = 1 To 2
'        testList.Remove markedItems(i)
'    Next i
'Debug.Assert testList.Contains(item1) = False
'Debug.Assert testList.Contains(item2) = False
'End Sub
'

'@TestMethod
Public Sub testAddingTrigger()
    On Error GoTo TestFail
    
    'Arrange:
    BufferEvents.Buffer.AddingTrigger = 2
    
    'Act:
    BufferEvents.Buffer.AddItems Array(1, 2, 3)

    'Assert:
    Assert.AreEqual "2", BufferEvents.AddedEventRaised

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

