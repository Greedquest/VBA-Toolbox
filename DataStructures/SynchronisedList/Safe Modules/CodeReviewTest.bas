Attribute VB_Name = "CodeReviewTest"
'@Folder("CodeReview")
Option Explicit

Sub showForm()
    Static runner As New FormRunner
    runner.init Worksheets("data").ListObjects("ExampleData")
End Sub

Sub compileSynchrolist()
Toolbox.CompressProject ThisWorkbook, "Filterlist" _
                                    , "FilterlistUtils" _
                                    , "ArraySupport" _
                                    , "FilterRunner" _
                                    , "SynchroListUtils" _
                                    , "ContentDataWrapper" _
                                    , "ListBuffer" _
                                    , "SourceDataWrapper" _
                                    , "SynchronisedList"
                                    
End Sub

