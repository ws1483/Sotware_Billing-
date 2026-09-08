Attribute VB_Name = "modHelpers"
' ============================================================================
' modHelpers ? shared PUBLIC helpers used across all modules.
'   PHASE 0 UPDATE:
'     - Unified, NORMALIZED FindLogRow (case/space-insensitive) ? replaces the
'       divergent copies (exact vs normalized) that existed across modules.
'     - Consolidated numeric/lookup helpers: Num, NrmID, DrNameToCustID,
'       CustIDToDrName, DoctorCredit, AddDoctorCredit.
'     - USERNAME casing fix.
'     - New number-collision / renumber support helpers.
'   NOTE: all procedures are Public so other modules can call them.
' ============================================================================
Option Explicit

' ============================ NORMALIZE ====================================
' Case + space insensitive key used for ALL doc-number / id comparisons.
Public Function NrmID(ByVal s As String) As String
    NrmID = UCase$(Replace(Trim$(s), " ", ""))
End Function



' ============================ MERGE-SAFE CELL OPS ==========================
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

' ============================ DOCUMENT NUMBERING ===========================
' Returns next document number string, increments the Settings counter.
'   docType: "INV","QTE","CN","MC"    dept: "WA"/"WD" (ignored for CN)
Public Function NextDocNumber(dept As String, docType As String) As String
    Dim wsSet As Worksheet, cCell As String, n As Long, prefix As String
    Set wsSet = ThisWorkbook.Sheets("Settings")
    dept = UCase$(Trim$(dept))

    Select Case UCase$(docType)
        Case "INV"
            prefix = "INV-" & dept & "-"
            cCell = IIf(dept = "WA", "B12", "B13")
        Case "QTE"
            prefix = "Q-" & dept & "-"
            cCell = IIf(dept = "WA", "B10", "B11")
        Case "CN"
            prefix = "CN-"
            cCell = "B14"
        Case "MC"
            prefix = "MC-INV-" & dept & "-"
            cCell = "B20"
    End Select

    n = CLng(wsSet.Range(cCell).value)
    NextDocNumber = prefix & Format(n, "0000")
    wsSet.Range(cCell).value = n + 1
End Function

' Generate a GUARANTEED-unique number for docType/dept by looping the counter
' until FindLogRow finds no clash in the given log. Prevents duplicate numbers.
Public Function NextUniqueDocNumber(wsLog As Worksheet, dept As String, docType As String) As String
    Dim candidate As String, guard As Long
    Do
        candidate = NextDocNumber(dept, docType)
        guard = guard + 1
        If guard > 100000 Then Err.Raise vbObjectError + 90, , "Could not allocate a unique number."
    Loop While FindLogRow(wsLog, candidate) > 0
    NextUniqueDocNumber = candidate
End Function

' ---- Find row of docNo in column A of a log sheet; 0 if not found ----
'   NORMALIZED match (case/space-insensitive) ? the single canonical version.
Public Function FindLogRow(ws As Worksheet, docNo As String) As Long
    Dim lastRow As Long, i As Long, key As String
    key = NrmID(docNo)
    If key = "" Then FindLogRow = 0: Exit Function
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If NrmID(CStr(ws.Cells(i, "A").value)) = key Then
            FindLogRow = i
            Exit Function
        End If
    Next i
    FindLogRow = 0
End Function

' ============================ OVERWRITE GUARD ==============================
' Safety net for save-time collisions. Returns True if it is SAFE to write
' docNo at expectedRow in wsLog; False (and shows a warning) if docNo already
' belongs to a DIFFERENT row (would overwrite another document).
'   expectedRow = 0 means "this is a NEW row" (docNo must not exist anywhere).
Public Function SafeToWriteNumber(wsLog As Worksheet, docNo As String, _
                                  ByVal expectedRow As Long) As Boolean
    Dim foundRow As Long
    foundRow = FindLogRow(wsLog, docNo)
    If foundRow = 0 Then
        SafeToWriteNumber = True
    ElseIf foundRow = expectedRow Then
        SafeToWriteNumber = True
    Else
        MsgBox "Number " & docNo & " already belongs to another document " & _
               "(row " & foundRow & ")." & vbCrLf & _
               "Save cancelled ? nothing was overwritten.", _
               vbCritical, "Duplicate number blocked"
        SafeToWriteNumber = False
    End If
End Function

' ============================ AUDIT =======================================
Public Sub LogAudit(action As String, docNo As String, _
                    oldVal As String, newVal As String, comment As String)
    Dim wsA As Worksheet, r As Long
    Set wsA = ThisWorkbook.Sheets("AuditLog")
    r = wsA.Cells(wsA.Rows.Count, "A").End(xlUp).row + 1
    wsA.Cells(r, 1).value = Now
    wsA.Cells(r, 2).value = Environ$("USERNAME")     ' PHASE 0: correct casing
    wsA.Cells(r, 3).value = action
    wsA.Cells(r, 4).value = docNo
    wsA.Cells(r, 5).value = oldVal
    wsA.Cells(r, 6).value = newVal
    wsA.Cells(r, 7).value = comment
End Sub

' ============================ CUSTOMER LOOKUPS ============================
Public Function CustIDToDrName(custID As String) As String
    Dim wsC As Worksheet, lastRow As Long, i As Long, key As String
    key = NrmID(custID)
    Set wsC = ThisWorkbook.Sheets("Customers")
    lastRow = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If NrmID(CStr(wsC.Cells(i, "A").value)) = key Then
            CustIDToDrName = CStr(wsC.Cells(i, "B").value)
            Exit Function
        End If
    Next i
    CustIDToDrName = ""
End Function

Public Function DrNameToCustID(drName As String) As String
    Dim wsC As Worksheet, lastRow As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    lastRow = wsC.Cells(wsC.Rows.Count, "B").End(xlUp).row
    For i = 2 To lastRow
        If CStr(wsC.Cells(i, "B").value) = drName Then
            DrNameToCustID = CStr(wsC.Cells(i, "A").value)
            Exit Function
        End If
    Next i
    DrNameToCustID = ""
End Function

Public Function LookupBHFByDrName(drName As String) As String
    Dim wsC As Worksheet, last As Long, i As Long
    LookupBHFByDrName = ""
    If Trim$(drName) = "" Then Exit Function
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If CStr(wsC.Cells(i, "B").value) = drName Then
            LookupBHFByDrName = CStr(wsC.Cells(i, "H").value)
            Exit Function
        End If
    Next i
End Function

Public Function DoctorCredit(custID As String) As Double
    Dim wsC As Worksheet, last As Long, i As Long, key As String
    key = NrmID(custID)
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsC.Cells(i, "A").value)) = key Then
            DoctorCredit = Num(wsC.Cells(i, "L").value)
            Exit Function
        End If
    Next i
End Function

Public Sub AddDoctorCredit(custID As String, addAmt As Double)
    Dim wsC As Worksheet, last As Long, i As Long, key As String
    key = NrmID(custID)
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsC.Cells(i, "A").value)) = key Then
            wsC.Cells(i, "L").value = Num(wsC.Cells(i, "L").value) + addAmt
            Exit Sub
        End If
    Next i
End Sub
Public Function Num(v As Variant) As Double
    If IsError(v) Then
        Num = 0
    ElseIf Trim$(CStr(v)) = "" Then
        Num = 0
    ElseIf IsNumeric(v) Then
        Num = CDbl(v)
    Else
        Num = 0
    End If
End Function
