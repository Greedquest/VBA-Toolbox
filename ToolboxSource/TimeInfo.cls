VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder Stopwatch

Private Type TTimeInfo
    TimeIn As Double
    TimeOut As Double
End Type

Private this As TTimeInfo

Public Property Get TimeIn() As Double
    TimeIn = this.TimeIn
End Property

Public Property Let TimeIn(ByVal value As Double)
    this.TimeIn = value
End Property

Public Property Get TimeOut() As Double
    TimeOut = this.TimeOut
End Property

Public Property Let TimeOut(ByVal value As Double)
    this.TimeOut = value
End Property

