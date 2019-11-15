Attribute VB_Name = "Module1"
Sub makeAGoodPoint()
    Dim a As Double
    a = 6.1
    
    Dim pA As Pointer
    Set pA = Pointer.Create(VarPtr(a), VarType(a), 1)
    
    
End Sub
