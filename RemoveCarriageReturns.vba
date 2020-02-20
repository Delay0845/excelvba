Sub RemoveCarriageReturns()
	Dim MyRange As Range

	Application.ScreenUpdating = False 'Turns off screen updates so your window isn't constantly trying to keep up with your commands; extremely helpful for performance improvements'
	Application.Calculation = xlCalculationManual
	Application.EnableEvents = False

	For Each MyRange In ActiveSheet.UsedRange
		If 0 < InStr(MyRange, Chr(10)) Then
			MyRange = Replace(MyRange, Chr(10), "")
		End If
	Next
	For Each MyRange In ActiveSheet.UsedRange
		If 0 < InStr(MyRange, Chr(13)) Then
			MyRange = Replace(MyRange, Chr(13), " ")
		End If
	Next

	Application.ScreenUpdating = True
	Application.Calculation = xlCalculationAutomatic
	Application.EnableEvents = True

End Sub
