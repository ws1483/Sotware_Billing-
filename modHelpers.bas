Attribute VB_Name = "modHelpers"
' ============================================================================
' modHelpers — shared PUBLIC helpers used across all modules.
'   Merge-proof cell ops, document numbering, audit logging, log lookup,
'   and CustID -> DrName reverse lookup (for Recall).
'   NOTE: all procedures are Public so other modules can call them.
' ============================================================================
Option Explicit

' ---- Clear a cell safely, even inside a merged range ----
Public Sub ClearCell(ws As Worksheet, addr As String)
    Dim ma As Range
    Set ma = ws.Range(addr).MergeArea
    If ma.Cells.Count > 1 Then
        ma.UnMerge
        ma.Cells(1, 1).ClearContents
        ma.Merge
    Else
        ma.ClearContents
    End If
End Sub

' ---- Set a formula safely, even inside a merged range ----
'   f uses ENGLISH COMMAS; Excel translates to locale separator (;)
Public Sub SetFormula(ws As Worksheet, addr As String, f As String)
    Dim ma As Range
    Set ma = ws.Range(addr).MergeArea
    If ma.Cells.Count > 1 Then
        ma.UnMerge
        ma.Cells(1, 1).Formula = f
        ma.Merge
    Else
        ma.Cells(1, 1).Formula = f
    End If
End Sub

' ---- Returns next document number string, increments the Settings counter ----
'   docType: "INV","QTE","CN"    dept: "WA"/"WD" (ignored for CN)
'   Format: INV-WA-0260 (prefix + UCase dept + 4-digit pad; grows past 9999)
Public Function NextDocNumber(dept As String, docType As String) As String
    Dim wsSet As Worksheet, cCell As String, n As Long, prefix As String
    Set wsSet = ThisWorkbook.Sheets("Settings")
    dept = UCase(Trim(dept))

    Select Case UCase(docType)
        Case "INV"
            prefix = "INV-" & dept & "-"
            cCell = IIf(dept = "WA", "B12", "B13")
        Case "QTE"
            prefix = "Q-" & dept & "-"
            cCell = IIf(dept = "WA", "B10", "B11")
        Case "CN"
            prefix = "CN-"
            cCell = "B14"
    End Select

    n = CLng(wsSet.Range(cCell).Value)
    NextDocNumber = prefix & Format(n, "0000")
    wsSet.Range(cCell).Value = n + 1
End Function

' ---- Find row of docNo in column A of a log sheet; 0 if not found ----
Public Function FindLogRow(ws As Worksheet, docNo As String) As Long
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If CStr(ws.Cells(i, "A").Value) = docNo Then
            FindLogRow = i
            Exit Function
        End If
    Next i
    FindLogRow = 0
End Function

' ---- Append one row to AuditLog (Timestamp,User,Action,DocNo,OldValue,NewValue,Comment) ----
Public Sub LogAudit(action As String, docNo As String, _
                    oldVal As String, newVal As String, comment As String)
    Dim wsA As Worksheet, r As Long
    Set wsA = ThisWorkbook.Sheets("AuditLog")
    r = wsA.Cells(wsA.Rows.Count, "A").End(xlUp).row + 1
    wsA.Cells(r, 1).Value = Now
    wsA.Cells(r, 2).Value = Environ$("Username")
    wsA.Cells(r, 3).Value = action
    wsA.Cells(r, 4).Value = docNo
    wsA.Cells(r, 5).Value = oldVal
    wsA.Cells(r, 6).Value = newVal
    wsA.Cells(r, 7).Value = comment
End Sub

' ---- Reverse lookup: CustID -> DrName (Customers col A -> col B, rows 2..) ----
Public Function CustIDToDrName(custID As String) As String
    Dim wsC As Worksheet, lastRow As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    lastRow = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If CStr(wsC.Cells(i, "A").Value) = custID Then
            CustIDToDrName = CStr(wsC.Cells(i, "B").Value)
            Exit Function
        End If
    Next i
    CustIDToDrName = ""      ' not found
End Function
' ---- Look up a doctor's BHF (Customers col H) by DrName (col B) ----
Public Function LookupBHFByDrName(drName As String) As String
    Dim wsC As Worksheet, last As Long, i As Long
    LookupBHFByDrName = ""
    If Trim(drName) = "" Then Exit Function
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If CStr(wsC.Cells(i, "B").Value) = drName Then         ' match DrName (col B)
            LookupBHFByDrName = CStr(wsC.Cells(i, "H").Value)   ' BHF = col H
            Exit Function
        End If
    Next i
End Function


