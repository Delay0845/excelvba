Sub Save_{FileType/Name}()
'
'Save_{FileType/Name} Macro
' A macro to automatically save the current file with a value from the file itself (a date in this instance).
'

'
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    ActiveSheet.Shapes.SelectAll

    Selection.Delete
    ActiveWorkbook.SaveAs _
    "/Users/{user}/Documents/" & Format((Range("B3")), "yyyymmdd") & "{AdditionalFileNameInformation}.xlsx", FileFormat:=51
    'ActiveWorkbook.SendMail Recipients:="{EmailAddress}", Subject:="{Subject}" & Range("B3")
    Application.DisplayAlerts = True

End Sub
