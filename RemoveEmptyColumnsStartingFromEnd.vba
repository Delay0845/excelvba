Sub Remove_Empty_Columns()
'
'Remove_Empty_Columns Macro
' A macro that autmatically removes any empty columns in a worksheet.
'

'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Step1:  Declare your variables.
    Dim MyRange As Range
    Dim iCounter As Long
    
    Set MyRange = ActiveSheet.UsedRange 'Step 2:  Define the target Range.
    
    For iCounter = MyRange.Columns.Count To 1 Step -1 'Step 3:  Start reverse looping through the range.
        If Application.CountA(Columns(iCounter).EntireColumn) = 0 Then 'Step 4: If entire column is empty then delete it.
           Columns(iCounter).Delete
        End If
    Next iCounter 'Step 5: Increment the counter down
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
