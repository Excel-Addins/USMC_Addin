Attribute VB_Name = "control"
'
'
'    This Module is for enabling modularity of the application
'    Primary Subs/Functions are administrative controls
'
'
'
'


Public Sub cleanExit()
MsgBox "User Cancelled!", vbOKOnly, "User Cancellation"
prepAndCleanup True
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


Public Function RegisterUDF(macroName As String, desc As String, cat As Byte)
'accepts description of function and category as byte for registration
Application.MacroOptions macro:=macroName, Description:=desc, Category:=cat, hasmenu:=True, menutext:=desc
End Function

