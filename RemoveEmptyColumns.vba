Sub RemoveEmptyColumns()
'Remove_Empty_Columns Macro
' A macro that autmatically removes any empty columns (50 per run) in a worksheet.
'

'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Dim totalCol As Integer
    
    'totalCol = ActiveSheet.UsedRange.Columns.Count
    'MsgBox totalCol
    
    For i = 1 To 50
        Range("A1").Select
        Selection.End(xlToRight).Select
        ActiveCell.Offset(0, 1).Select
        ActiveCell.EntireColumn.Select
        'MsgBox WorksheetFunction.CountA(Selection)
        If WorksheetFunction.CountA(Selection) = 0 Then
            'MsgBox "Column Is Empty"
            Selection.EntireColumn.Select
            Selection.Delete Shift:=xlToLeft
        End If
    Next i
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
