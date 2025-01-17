''************************************************************************001 - Macro to call at the start of each macro ********************************************************
'Excel contains various settings that can slow down your macros. We want to turn off those settings before running a macro (which is example code 001),
'then restore them again after the macro finishes (which example code 002).
'Place at the start of the macro
Public calcMode As Long
Public pageBreakStatus As Boolean
Sub SettingsStartOfMacro()
With Application
    calcMode = .Calculation
    pageBreakStatus = ActiveSheet.DisplayPageBreaks
    'Turn calculation mode to manual
    .Calculation = xlCalculationManual
    'Turn off screen updating (i.e. no annoying screen flash)
    .ScreenUpdating = False
    'Alert windows will not be displayed
    .DisplayAlerts = False
End With
'Turn off page breaks on active sheet
ActiveSheet.DisplayPageBreaks = False
End Sub

 ''************************************************************************002 – Macro to call at the end of each macro ********************************************************
'To call this macro, include the following code at the end of your macro, directly before the End Sub statement.
'Call SettingsEndOfMacro
'Restores all the settings which were changed in the macro above.
Sub SettingsEndOfMacro()
With Application
    'Return calculation to automatic
    .Calculation = calcMode
    'Turn screen updating back on
    .ScreenUpdating = True
    'Enable alerts to be shown
    .DisplayAlerts = True
End With
'Reset page breaks on active sheet
ActiveSheet.DisplayPageBreaks = pageBreakStatus
End Sub
