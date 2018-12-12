Attribute VB_Name = "SourceWrapperRDTests"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider
Private sourceWrapper As SourceDataWrapper

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.PermissiveAssertClass
    Set Fakes = New Rubberduck.FakesProvider
 
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
    Set sourceWrapper = New SourceDataWrapper
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub TestAddingTwo()
    On Error GoTo TestFail
    
    'Arrange:

    Dim items1() As DummyGridItem, items2() As DummyGridItem
    items1() = getEmptyDummyClasses(1)
    items2() = getEmptyDummyClasses(3)
    
    'Act:
    Set sourceWrapper = New SourceDataWrapper
    With sourceWrapper
 
        .AddItems items1
        .AddItems items2
    End With
    
    With New FilterRunner                        'this triggers an event, but let's bypass it
        .SetFilterMode , , lstKeepAll
        .SetSortMode , lstNoSorting
        .FilterSourceToOutput sourceWrapper
    End With
    
    'Assert:
    With sourceWrapper
        Assert.AreEqual .AddedData.Count, "4", "Items weren't added correctly" 'add total of 4 items"
        Assert.AreEqual .AddedData.Count, "0", "Added data was not cleared as expected" 'Should wipe added data when read
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestRemoval()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dummyItems() As DummyGridItem
    dummyItems = getEmptyDummyClasses(5)
    sourceWrapper.AddItems dummyItems
    
    With New FilterRunner                        'adding triggers an event, let's assume that happened
        .SetFilterMode , , lstKeepAll
        .SetSortMode , lstNoSorting
        .FilterSourceToOutput sourceWrapper
    End With

    'Act:
    sourceWrapper.RemoveItems Array(dummyItems(3)) 'need to remove an iterable
        
    'Assert:
    With sourceWrapper.AddedData
        Assert.IsTrue .Count = 4                 '1 removed
        Assert.IsTrue .Contains(dummyItems(1))
        Assert.IsFalse .Contains(dummyItems(3))
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub testRemovalOfNonPresent()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dummyItems() As DummyGridItem
    dummyItems = getEmptyDummyClasses(5)
    Dim itemToRemove As DummyGridItem
    Set itemToRemove = getEmptyDummyClasses(1)(1)
    sourceWrapper.AddItems dummyItems
    
    With New FilterRunner                        'adding triggers an event, let's assume that happened
        .SetFilterMode , , lstKeepAll
        .SetSortMode , lstNoSorting
        .FilterSourceToOutput sourceWrapper
    End With

    'Act:
    sourceWrapper.RemoveItems Array(itemToRemove) 'need to remove an iterable
        
    'Assert:
    With sourceWrapper.AddedData
        Assert.IsTrue .Count = 5                 '1 removed
        Assert.IsTrue .Contains(dummyItems(1))
        Assert.IsFalse .Contains(itemToRemove)
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'
'Sub testFiltering()
'    Dim testitem As New clsSourceWrapper
'    testitem.SetFilterMode "Date"
'End Sub

'@TestMethod
Public Sub TestFilteringError()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim dummyItems() As DummyGridItem
    dummyItems = getEmptyDummyClasses(5)
    sourceWrapper.AddItems dummyItems
    
    'Act:
    With New FilterRunner                        'adding triggers an event, let's assume that happened
        '.setFilterMode
        .FilterSourceToOutput sourceWrapper
    End With

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

