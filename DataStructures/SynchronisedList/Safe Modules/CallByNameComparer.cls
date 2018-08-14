VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallByNameComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("CodeReview.Utils")

'From PropertyComparer Class Implementing IComparer with ArrayList
'https://codereview.stackexchange.com/q/180937/146810

Implements mscorlib.IComparer

Private mArgs As Variant
Private mCallType As VbCallType
Private mProcName As String

Private Function IComparer_Compare(ByVal X As Variant, ByVal Y As Variant) As Long
    Dim x1 As Variant, y1 As Variant
    If Len(mProcName) = 0 Then
        x1 = X
        y1 = Y
    Else
        x1 = CallFunction(X)
        y1 = CallFunction(Y)
    End If
    If TypeName(x1) = "String" Then
        IComparer_Compare = StrComp(x1, y1, vbTextCompare)
    Else
        If x1 > y1 Then
            IComparer_Compare = 1
        ElseIf x1 < y1 Then
            IComparer_Compare = -1
        End If
    End If
End Function

Public Sub init(ProcName As String, CallType As VbCallType, ParamArray args())
    mProcName = ProcName
    mCallType = CallType
    mArgs = args
End Sub

' http://www.vbforums.com/showthread.php?405366-RESOLVED-Using-CallByName-with-variable-number-of-arguments
' Author: Joacim Andersson
Private Function CallFunction(Object As Variant)
    Select Case UBound(mArgs)
    Case -1
        CallFunction = CallByName(Object, mProcName, mCallType)
    Case 0
        CallFunction = CallByName(Object, mProcName, mCallType, val(mArgs(0)))
    Case 1
        CallFunction = CallByName(Object, mProcName, mCallType, val(mArgs(0)), val(mArgs(1)))
    Case 2
        CallFunction = CallByName(Object, mProcName, mCallType, val(mArgs(0)), val(mArgs(1)), val(mArgs(2)))
    Case 3
        CallFunction = CallByName(Object, mProcName, mCallType, val(mArgs(0)), val(mArgs(1)), val(mArgs(2)), val(mArgs(3)))
    Case 4
        CallFunction = CallByName(Object, mProcName, mCallType, val(mArgs(0)), val(mArgs(1)), val(mArgs(2)), val(mArgs(3)), val(mArgs(4)))
    End Select
End Function

