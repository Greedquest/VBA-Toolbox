Attribute VB_Name = "codeReviewTestRunner"
'@Folder("Tests")
Private crTestClass As codeReviewTest

Public Sub runTest(Optional debugging As Boolean = False)
'Runs an email search over urls in A1:A10 of activesheet
'Returns all email address matches by
    Set crTestClass = New codeReviewTest
    If debugging Then Stop
    crTestClass.run
End Sub

Public Sub clickRun()
runTest False
End Sub
