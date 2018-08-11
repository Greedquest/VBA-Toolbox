Attribute VB_Name = "ShynchronizedListRDTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider
Private synchro As SynchronisedList
Private EventMonitor As SynchroEventsTest

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
    Set synchro = New SynchronisedList
    Set EventMonitor = New SynchroEventsTest
    Set EventMonitor.synchro = synchro
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
    items1() = getEmptyDummyClasses(2)
    items2() = getEmptyDummyClasses(3)
    
    'Act:
    With synchro
        .Add items1(1), items1(2)
        .Add items2
    End With
    
    'Assert:
    With synchro
        Assert.AreEqual "5", .SourceData.Count, "SourceData added incorrectly"
        Assert.AreEqual "5", .SourceData.Count, "SourceData affected by reading"
        Assert.AreEqual "5", .GridData.Count, "Default Grid behaviour not as expected"
    End With
    
    'check events
    With EventMonitor
        Assert.AreEqual "2", .OrderEventRaised   're-order on addition
        Assert.AreEqual "2", .LastChangeIndex    'assumes 1st 2 ordered, so now handle from 2 onwards
        Assert.AreEqual "0", .PropertiesEventRaised 'no ammendments made
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub testRemoval()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dummyItems() As DummyGridItem
    dummyItems = getEmptyDummyClasses(5)
    synchro.Add dummyItems
    
    'Act:
    synchro.Remove dummyItems(3)
        
    'Assert:
    With synchro
        Assert.AreEqual "4", .SourceData.Count
        Assert.AreEqual "4", .GridData.Count
        Assert.IsTrue .SourceData.Contains(dummyItems(1))
        Assert.IsFalse .SourceData.Contains(dummyItems(3))
        Assert.IsTrue .GridData.Contains(dummyItems(1))
        Assert.IsFalse .GridData.Contains(dummyItems(3))
    End With
    
    'Events check
    With EventMonitor
        Assert.AreEqual "2", .LastChangeIndex, "Change index not monitored as expected" 'since it is the 3rd item (now item 4 of original array) that has changed
        Assert.AreEqual "2", .OrderEventRaised   'one for adding, one for removing
        Assert.AreEqual "0", .PropertiesEventRaised
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestAmmedment()
    On Error GoTo TestFail
    'Arrange:

    Dim items() As DummyGridItem
    items() = getEmptyDummyClasses(5)
    synchro.Add items
    'Act:
    synchro.MarkAsAmmended items(3)              'raises 2 order events!
    
    'Assert:
    With synchro
        Assert.AreSame items(3), .GridData(4)    'last item since added to end of grid
        Assert.AreSame items(3), .SourceData(2)  '3rd item since no change
    End With
    
    With EventMonitor
        'Items first added, then ammended item removed and re-added
        Assert.AreEqual "3", .OrderEventRaised, "Number of re-ordering events is: " & .OrderEventRaised & ", not as expected" 're-order on addition but not when ammended
        Assert.AreEqual "4", .LastChangeIndex, "Last change index is: " & .LastChangeIndex & ", not as expected" 'no info, so need full re-order
        Assert.AreEqual "0", .PropertiesEventRaised 'no ammendments made
        Assert.SequenceEquals Array("0", "2", "4"), Array(.ChangeIndecies(1), .ChangeIndecies(2), .ChangeIndecies(3))
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestSorting()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Sorter As New CallByNameComparer
    Sorter.init "Name", VbGet
    
    'generate correct list
    Dim CorrectList As New ArrayList
    
    Dim i As Long
    For i = 1 To 5
        Dim itemToAdd: Set itemToAdd = New DummySortByNameItem
        CorrectList.Add itemToAdd
        synchro.Add itemToAdd
    Next i
    CorrectList.Sort_2 Sorter

    
    'Act:
    synchro.Sort Sorter                          'doubles grid data for some reason
    
    'Assert:

    Assert.SequenceEquals IterableToArray(CorrectList), IterableToArray(synchro.GridData.data)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestFilter()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Filterer As New CallByNameComparer
    Filterer.init "Name", VbGet
    
    Dim testClasses() As DummyGridItem           'should be named dummyItem1,2,3 etc.
    Dim filterAgainst As DummyGridItem
    testClasses = getEmptyDummyClasses()
    Set filterAgainst = getEmptyDummyClasses(1)(1) 'should also be named dummyItem1
    
    synchro.Add testClasses
    
    'Act:
    synchro.filter filterAgainst, Filterer       'auto filter mode is remove matching

    'Assert:
    'check if items with matching name are removed from  grid
    Assert.IsTrue synchro.SourceData.Contains(testClasses(1))
    Assert.IsFalse synchro.GridData.Contains(testClasses(1))
    
    'Act:
    Set filterAgainst = getEmptyDummyClasses(2)(2)
    synchro.filter filterAgainst, Filterer, lstKeepMatching

    'Assert:
    Assert.IsTrue synchro.GridData.Contains(testClasses(2))
    Assert.IsFalse synchro.GridData.Contains(testClasses(1))
    
    'Event assertions
    'expected an order change on addition
    'order change when filter
    'another order change on filter
    With EventMonitor
        Assert.AreEqual "3", .OrderEventRaised
        Assert.SequenceEquals Array("0", "0", "1"), IterableToArray(.ChangeIndecies)
    End With
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestAddingOne() 'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:

    Dim item As DummyGridItem
    Set item = getEmptyDummyClasses(2)(1)

    
    'Act:
    With synchro
        .Add item
    End With
    
    'Assert:
    With synchro
        Assert.AreSame item, .SourceData(0), "SourceData added incorrectly"
        Assert.AreEqual "1", .SourceData.Count, "SourceData affected by reading"
        Assert.AreEqual "1", .GridData.Count, "Default Grid behaviour not as expected"
    End With
    
    'check events
    With EventMonitor
        Assert.AreEqual "1", .OrderEventRaised   're-order on addition
        Assert.AreEqual "0", .LastChangeIndex    'assumes 1st 2 ordered, so now handle from 2 onwards
        Assert.AreEqual "0", .PropertiesEventRaised 'no ammendments made
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestAddingRange() 'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    synchro.Add Array(Range("a1:a5")) 'wrap iterables we don't want flattened
    'Assert:
    Assert.AreEqual "1", synchro.SourceData.Count, "Count incorrect"
    Assert.AreEqual "$A$1:$A$5", UCase(synchro.SourceData(0).Address), "Addresses don't line up"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


