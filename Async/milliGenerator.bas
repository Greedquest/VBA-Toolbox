Attribute VB_Name = "milliGenerator"
      Public Declare Function SetTimer Lib "user32" ( _
          ByVal HWnd As Long, ByVal nIDEvent As Long, _
          ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
      Public Declare Function KillTimer Lib "user32" ( _
          ByVal HWnd As Long, ByVal nIDEvent As Long) As Long

      Public TimerID As Long
      Public TimerSeconds As Single

      Sub StartTimer()
          TimerSeconds = 1 ' how often to "pop" the timer.
          TimerID = SetTimer(0&, 0&, TimerSeconds * 1000&, AddressOf TimerProc)
      End Sub

      Sub EndTimer()
          On Error Resume Next
          KillTimer 0&, TimerID
      End Sub

      Sub TimerProc(ByVal HWnd As Long, ByVal uMsg As Long, _
          ByVal nIDEvent As Long, ByVal dwTimer As Long)
          '
          ' The procedure is called by Windows. Put your
          ' timer-related code here.
          '
      End Sub

