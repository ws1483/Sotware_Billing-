Attribute VB_Name = "modMenuSave"
Public Sub SaveWorkbookNow()
    On Error GoTo Fail
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    MsgBox "Workbook saved.", vbInformation, "Save"
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    MsgBox "Save error: " & Err.Description, vbExclamation, "Save"
End Sub

