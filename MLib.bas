Attribute VB_Name = "MLib"
Option Explicit

Public Sub OptimizeWorkbook( _
    ByVal pEnable As Boolean _
)
    Application.StatusBar = False
    
    Application.EnableAnimations = Not pEnable
    Application.EnableEvents = Not pEnable
    Application.ScreenUpdating = Not pEnable
    Application.DisplayAlerts = Not pEnable
    Application.DisplayFormulaBar = Not pEnable
    Application.DisplayStatusBar = Not pEnable
    Application.Calculation = IIf(pEnable, xlCalculationManual, xlCalculationAutomatic)
End Sub

