Attribute VB_Name = "MInfoTip"
Option Explicit

Public CTipReference As CInfoTip

Public Sub tipTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
  CTipReference.Hide
End Sub

