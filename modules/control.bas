Attribute VB_Name = "control"
Public Sub cleanExit()
MsgBox "User Cancelled!", vbOKOnly, "User Cancellation"
End
End Sub


Public Sub prepAndCleanup(Optional bool As Boolean)
Application.DisplayAlerts = bool
Application.EnableEvents = bool
Application.ScreenUpdating = bool
If bool Then
    Application.Calculation = xlCalculationAutomatic
Else
    Application.Calculation = xlCalculationManual
End If
End Sub


