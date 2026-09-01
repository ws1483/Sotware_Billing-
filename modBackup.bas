Attribute VB_Name = "modBackup"
Option Explicit
' ============================================================================
' modBackup — SaveCopyAs a timestamped copy (OneDrive-safe), keep newest 3.
'   Navigation helpers included (full UI available — no hiding).
' ============================================================================
Private Const BACKUP_DIR As String = _
    "F:\One Drive\OneDrive\Documents\GreydataDental\Accounting\WeDental Billing\Backup"
Private Const KEEP_COUNT As Long = 3

Public Sub RunBackup()
    Dim stamp As String, destPath As String
    Dim fso As Object

    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(BACKUP_DIR) Then
        fso.CreateFolder BACKUP_DIR
        If Not fso.FolderExists(BACKUP_DIR) Then
            MsgBox "Backup folder not found/creatable:" & vbCrLf & BACKUP_DIR, _
                   vbExclamation, "Backup"
            Exit Sub
        End If
    End If

    ' save pending changes
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

    ' timestamped destination
    stamp = Format(Now, "yyyy-mm-dd_hhnn")
    destPath = BACKUP_DIR & "\" & "WeDental_Backup_" & stamp & ".xlsm"

    ' SaveCopyAs works for OneDrive-synced workbooks (no https path issue)
    ThisWorkbook.SaveCopyAs destPath

    ' prune old backups (keep newest KEEP_COUNT)
    PruneBackups fso

    MsgBox "Backup created:" & vbCrLf & destPath & vbCrLf & vbCrLf & _
           "Keeping the latest " & KEEP_COUNT & " backups.", vbInformation, "Backup"
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    MsgBox "Backup error: " & Err.Description & vbCrLf & vbCrLf & _
           "Target: " & destPath, vbExclamation, "Backup"
End Sub

Private Sub PruneBackups(fso As Object)
    Dim fld As Object, f As Object
    Dim names() As String, dates() As Double, n As Long, i As Long, j As Long
    Dim tS As String, tD As Double

    Set fld = fso.GetFolder(BACKUP_DIR)

    ' collect our backup files only
    n = 0
    ReDim names(1 To 1000)
    ReDim dates(1 To 1000)
    For Each f In fld.Files
        If LCase(f.Name) Like "wedental_backup_*.xlsm" Then
            n = n + 1
            names(n) = f.Path
            dates(n) = CDbl(f.DateLastModified)
        End If
    Next f
    If n <= KEEP_COUNT Then Exit Sub

    ' sort NEWEST first (descending by modified date)
    For i = 1 To n - 1
        For j = 1 To n - i
            If dates(j) < dates(j + 1) Then
                tD = dates(j): dates(j) = dates(j + 1): dates(j + 1) = tD
                tS = names(j): names(j) = names(j + 1): names(j + 1) = tS
            End If
        Next j
    Next i

    ' delete everything past KEEP_COUNT
    For i = KEEP_COUNT + 1 To n
        On Error Resume Next
        fso.DeleteFile names(i), True
        On Error GoTo 0
    Next i
End Sub

' ---- One-click "reset my view" safety button (single canonical copy) --------
Sub ShowRibbon()
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    On Error GoTo 0
    MsgBox "Full view restored.", vbInformation
End Sub

' ===== Navigation (full UI available; simply jump to a sheet) =================

' Go to the Menu sheet
Public Sub GoToMenu()
    NavTo "Menu"
End Sub

' Go to the Quote sheet
Public Sub GoToQuote()
    NavTo "Quote"
End Sub

' Go to the Invoice sheet
Public Sub GoToInvoice()
    NavTo "Invoice"
End Sub

' Core navigation helper: activate the target sheet (no UI hiding).
Private Sub NavTo(sheetName As String)
    Dim ws As Worksheet
    On Error GoTo Fail
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Activate
    ws.Range("A1").Select
    Exit Sub
Fail:
    MsgBox "Cannot navigate to '" & sheetName & "'." & vbCrLf & Err.Description, vbExclamation
End Sub



