Sub Autofit_Cells()
'
' Autofit_Cells Macro
' A macro to automatically adjust the cell width of all data on a worksheet to fit the data.
'
'

	Application.ScreenUpdating = False
	Application.Calculation = xlCalculationManual
	Application.EnableEvents = False

	Range("A1").Select
	Cells.Select
	Cells.EntireColumn.AutoFit
	Cells.EntireRow.AutoFit

	Application.ScreenUpdating = True
	Application.Calculation = xlCalculationAutomatic
	Application.EnableEvents = True

End Sub
