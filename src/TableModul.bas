Attribute VB_Name = "TableModul"
'@Folder "Modules.Utilities"

Public Sub disableUpdates()
With Application
    .ScreenUpdating = False
    statusCalc = .Calculation
    .Calculation = xlCalculationManual
    End With
    Application.EnableEvents = False
End Sub

Public Sub enableUpdates()
With Application
    .ScreenUpdating = True
    statusCalc = .Calculation
    .Calculation = xlCalculationAutomatic
    End With
    Application.EnableEvents = True
End Sub

