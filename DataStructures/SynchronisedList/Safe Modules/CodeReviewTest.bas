Attribute VB_Name = "CodeReviewTest"
'@Folder("CodeReview")
Option Explicit

Sub showForm()
    Static runner As New FormRunner
    runner.init Worksheets("data").ListObjects("ExampleData")
End Sub
