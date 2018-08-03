Attribute VB_Name = "PropertiesUtils"
'@Folder("Tests.Utils")


Function GetRndDate(dtStartDate As Date, dtEndDate As Date) As Date
    On Error GoTo Error_Handler
    Dim dtTmp                 As Date
 
    'Swap the dates if dtStartDate is after dtEndDate
    If dtStartDate > dtEndDate Then
        dtTmp = dtStartDate
        dtStartDate = dtEndDate
        dtEndDate = dtTmp
    End If
 
    Randomize
    GetRndDate = DateAdd("d", Int((DateDiff("d", dtStartDate, dtEndDate) + 1) * Rnd), dtStartDate)
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: GetRndDate" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, vbNullString, Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Function RandomString()
    'PURPOSE: Create a Randomized String of Characters
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    Const Length As Integer = 20
    Dim CharacterBank As Variant
    Dim X As Long
    Dim str As String

    'Test Length Input
    If Length < 1 Then
        MsgBox "Length variable must be greater than 0"
        Exit Function
    End If

    CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
                          "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
                          "y", "z")
  

    'Randomly Select Characters One-by-One
    For X = 1 To Length
        'Randomize
        str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
    Next X

    'Output Randomly Generated String
    RandomString = str

End Function


