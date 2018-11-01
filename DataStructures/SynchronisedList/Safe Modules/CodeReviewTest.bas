Attribute VB_Name = "CodeReviewTest"
'@Folder("CodeReview")
Option Explicit

Sub showForm()
    Static runner As FormRunner
    Set runner = New FormRunner
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
                                    
Toolbox.CompressProject ThisWorkbook, "CodeReviewTest" _
                                    , "ExampleForm" _
                                    , "FormRunner" _
                                    , "dummyRange" _
                                    , "CallByNameComparer"
End Sub

