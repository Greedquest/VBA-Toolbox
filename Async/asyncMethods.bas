Attribute VB_Name = "asyncMethods"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal bytes As Long)

' read a value of any type from memory

Public Function Peek(ByVal address As Long, ByVal ValueType As VbVarType) As Variant
    Select Case ValueType
        Case vbByte
            Dim valueB As Byte
            CopyMemory valueB, ByVal address, 1
            Peek = valueB
        Case vbInteger
            Dim valueI As Integer
            CopyMemory valueI, ByVal address, 2
            Peek = valueI
        Case vbBoolean
            Dim valueBool As Boolean
            CopyMemory valueBool, ByVal address, 2
            Peek = valueBool
        Case vbLong
            Dim valueL As Long
            CopyMemory valueL, ByVal address, 4
            Peek = valueL
        Case vbSingle
            Dim valueS As Single
            CopyMemory valueS, ByVal address, 4
            Peek = valueS
        Case vbDouble
            Dim valueD As Double
            CopyMemory valueD, ByVal address, 8
            Peek = valueD
        Case vbCurrency
            Dim valueC As Currency
            CopyMemory valueC, ByVal address, 8
            Peek = valueC
        Case vbDate
            Dim valueDate As Date
            CopyMemory valueDate, ByVal address, 8
            Peek = valueDate
        Case vbVariant
            ' in this case we don't need an intermediate variable
            CopyMemory Peek, ByVal address, 16
        Case Else
            Err.Raise 1001, , "Unsupported data type"
    End Select

End Function
