Attribute VB_Name = "FilterListRDTests"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider
Private myList As FilterList
Private a() As New filterListComparerTest
Private Comparer As New propertyComparer

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    ReDim a(1 To 3)
    a(1).Value = 1
    a(2).Value = 7
    a(3).Value = 3
    Comparer.ComparisonProperty = "value"
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
    Set myList = New FilterList
    Dim item As Variant
    For Each item In a
        myList.Add item
    Next item
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub testSort()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    myList.Sort Comparer, lstSortAscending
    'Assert:
    
    Assert.AreEqual myList(0).Value, CLng(1), "First value '" & myList(0).Value & "' not equal to 1"
    Assert.AreEqual myList(1).Value, CLng(3), "Second value wrong"
    Assert.AreEqual myList(2).Value, CLng(7), "Third value wrong"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub testReverse()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    myList.Reverse
    'Assert:
    Assert.AreEqual myList(0).Value, "3"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestFilter()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    myList.Filter a(2), lstRemoveMatching, Comparer
    'Assert:
    Assert.IsFalse myList.Contains(a(2))
    Assert.IsTrue myList.Contains(a(1)) = True

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub testClone()
    'relies on all other passing tests
    On Error GoTo TestFail
    
    'Arrange:
    Dim clonedList As FilterList
    
    'Act:
    Set clonedList = myList.Clone
    myList.Remove a(3)
   
    'Assert:
    Assert.IsTrue clonedList.Contains(a(3)), "Cloned list doesn't contain the correct value"
    
    'Act:
    clonedList.Filter a(1), lstKeepMatching, Comparer
    
    'Assert:
    Assert.IsTrue myList.Contains(a(2)), "Original list doesn't contain the correct value"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub testIndex()
    'relies on all other passing tests
    On Error GoTo TestFail
    
    'Arrange:
    Dim val
    'Act:
    val = myList(0).Value
    'Assert:
    Assert.AreEqual a(1).Value, val
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub testIndexOf()                         'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim index As Long
    index = myList.IndexOf(a(2))
    'Assert:
    Assert.AreEqual "1", index
    Assert.AreEqual "-1", myList.IndexOf("blah") 'not in list

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
