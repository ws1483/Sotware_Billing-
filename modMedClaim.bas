Attribute VB_Name = "modMedClaim"
' modMedClaim ? SaveMedClaim / RecallMedClaim / autofill from Med Customers.
'   MC = medical-aid claim invoices. RecipientType always "patient",
'   Dept = WA/WD, number = MC-INV-<dept>-####  (counter Settings!B20).
'   Log = MedAidLog (cols per modConfig ML_*), lines = MedAidLines (LN_*).
'   Sheet layout mirrors the Invoice sheet (see cell map below).
' ============================================================================
Option Explicit

'' ---- Autofill the Med Claim header from Med Customers when Bill To (C6) set --
'   Med Customers (Image 13):
'     A=BillTo(key) B=Name C=Street D=Suburb E=City F=PostCode G=IDNom
'     H=Tel I=Email J=MedAidName K=CreditBalance L=MedAidNumber
'     M=DependantCode N=MainMember
'   Med Claim cells:
'     C7 Name, C8 PatientID, C9 Street, C10 City, C11 PostCode, C12 Email,
'     G8 MedAid(name), G9 MedNo(number), G10 MainMember, G11 DependantCode
Public Sub MedClaimAutofill(ws As Worksheet)
    Dim wsMC As Worksheet, key As String, last As Long, i As Long
    key = NrmID(CStr(ws.Range("C6").value))
    If key = "" Then Exit Sub

    On Error GoTo Clean
    Set wsMC = ThisWorkbook.Sheets("Med Customers")
    last = wsMC.Cells(wsMC.Rows.Count, "A").End(xlUp).row

    Application.EnableEvents = False
    For i = 2 To last
        If NrmID(CStr(wsMC.Cells(i, "A").value)) = key Then
            ws.Range("C7").value = wsMC.Cells(i, "B").value    ' Patient Name
            ws.Range("C8").value = wsMC.Cells(i, "G").value    ' Patient ID (ID Nom)
            ws.Range("C9").value = wsMC.Cells(i, "C").value    ' Street/Address
            ws.Range("C10").value = wsMC.Cells(i, "E").value   ' City
            ws.Range("C11").value = wsMC.Cells(i, "F").value   ' Post Code
            ws.Range("C12").value = wsMC.Cells(i, "I").value   ' Email
            ws.Range("G8").value = wsMC.Cells(i, "J").value    ' Med Aid (name)   J
            ws.Range("G9").value = wsMC.Cells(i, "L").value    ' Med No (number)  L
            ws.Range("G10").value = wsMC.Cells(i, "N").value   ' Main Member      N
            ws.Range("G11").value = wsMC.Cells(i, "M").value   ' Dependant Code   M
            Exit For
        End If
    Next i
Clean:
    Application.EnableEvents = True
End Sub

' ============================================================ SAVE MED CLAIM ==
Public Sub SaveMedClaim()
    Dim ws As Worksheet, wsLog As Worksheet, wsLines As Worksheet
    Dim docNo As String, dept As String
    Dim logRow As Long, isUpdate As Boolean
    Dim r As Long, ln As Long, lastLine As Long
    Dim oldTotal As String, newTotal As String
    Dim storedDept As String, existingPaid As Double, newBal As Double

    Set ws = ThisWorkbook.Sheets("Med Claim")
    Set wsLog = ThisWorkbook.Sheets("MedAidLog")
    Set wsLines = ThisWorkbook.Sheets("MedAidLines")

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 1. VALIDATE
    If Trim(ws.Range("C6").value) = "" Then
        MsgBox "Please select a Bill To (Med Customer) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Not IsDate(ws.Range("G6").value) Then
        MsgBox "Please enter a valid Invoice Date (G6) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Trim(ws.Range("D16").value) = "" Then
        MsgBox "Add at least one line item before saving.", vbExclamation: GoTo CleanExit
    End If

    dept = UCase(Trim(ws.Range("K1").value))
    If dept <> "WA" And dept <> "WD" Then dept = "WA"

    ' 2. DETERMINE NUMBER (with dept-change renumber handling)
    docNo = Trim(ws.Range("G7").value)
    If docNo = "" Then
        ' NEW doc ? unique number for this dept
        docNo = NextUniqueDocNumber(wsLog, dept, "MC")
        ws.Range("G7").value = docNo
        ws.Range("K4").value = docNo
        isUpdate = False
    Else
        logRow = FindLogRow(wsLog, docNo)
        isUpdate = (logRow > 0)
        If isUpdate Then
            ' dept-change detection
            storedDept = UCase(Trim(CStr(wsLog.Cells(logRow, ML_DEPT).value)))
            If storedDept <> "" And storedDept <> dept Then
                ' Hand off to the renumber engine, then STOP this save (engine re-saves).
                Application.EnableEvents = True: Application.ScreenUpdating = True
                RenumberOnDeptChange "MC", docNo, storedDept, dept
                Exit Sub
            End If
        End If
    End If

    ' 3. TARGET ROW + overwrite guard
    If isUpdate Then
        If Not SafeToWriteNumber(wsLog, docNo, logRow) Then GoTo CleanExit
        oldTotal = CStr(wsLog.Cells(logRow, ML_TOTAL).value)
        existingPaid = Num(wsLog.Cells(logRow, ML_PAID).value)
    Else
        If Not SafeToWriteNumber(wsLog, docNo, 0) Then GoTo CleanExit
        logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
        existingPaid = 0
    End If

    ' 4. HEADER
    With wsLog
        .Cells(logRow, ML_NO).value = docNo
        .Cells(logRow, ML_DEPT).value = dept
        .Cells(logRow, ML_RECIP).value = "patient"
        .Cells(logRow, ML_DATE).value = ws.Range("G6").value
        .Cells(logRow, ML_DUE).value = ws.Range("G6").value        ' MC due = doc date (no terms); adjust if needed
        .Cells(logRow, ML_PATIENT).value = ws.Range("C7").value
        .Cells(logRow, ML_APPLIANCE).value = ws.Range("C14").value
        .Cells(logRow, ML_SUBTOTAL).value = ws.Range("H35").value
        .Cells(logRow, ML_DISC).value = ws.Range("H34").value
        .Cells(logRow, ML_VAT).value = ws.Range("H36").value
        .Cells(logRow, ML_TOTAL).value = ws.Range("H37").value
        .Cells(logRow, ML_CUST).value = ws.Range("C6").value       ' Bill To (key)
        .Cells(logRow, ML_PATIENTID).value = ws.Range("C8").value  ' Patient ID
        .Cells(logRow, ML_MEDAID).value = ws.Range("G8").value
        .Cells(logRow, ML_MEDNO).value = ws.Range("G9").value
        .Cells(logRow, ML_MAINMEM).value = ws.Range("G10").value
        .Cells(logRow, ML_BHF).value = ws.Range("G13").value
        .Cells(logRow, ML_DOCTOR).value = ws.Range("G12").value
        .Cells(logRow, ML_DEPCODE).value = ws.Range("G11").value
        .Cells(logRow, ML_DISCPCT).value = ws.Range("C32").value
        .Cells(logRow, ML_DISCFIX).value = ws.Range("C33").value
        .Cells(logRow, ML_NOTES).value = ws.Range("A40").value     ' Notes (Phase 2 feature, wired here)

        If Not isUpdate Then
            .Cells(logRow, ML_PAID).value = 0
            .Cells(logRow, ML_CREATED).value = Now
        End If
        .Cells(logRow, ML_MODIFIED).value = Now

        ' Balance = Total - existing Paid ; Status re-derived (preserves payments)
        newBal = Num(ws.Range("H37").value) - existingPaid
        If newBal < 0 Then newBal = 0
        .Cells(logRow, ML_BALANCE).value = newBal
        If newBal <= 0.005 And existingPaid > 0 Then
            .Cells(logRow, ML_STATUS).value = "Paid"
        ElseIf existingPaid > 0 Then
            .Cells(logRow, ML_STATUS).value = "Part-Paid"
        Else
            .Cells(logRow, ML_STATUS).value = "Unpaid"
        End If
    End With
    newTotal = CStr(ws.Range("H37").value)

    ' 5. LINES (delete old for this docNo, re-write current)
    lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For r = lastLine To 2 Step -1
        If NrmID(CStr(wsLines.Cells(r, LN_DOCNO).value)) = NrmID(docNo) Then wsLines.Rows(r).Delete
    Next r
    ln = 0
    For r = 16 To 30
        If Trim(ws.Range("D" & r).value) <> "" Then
            ln = ln + 1
            lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row + 1
            With wsLines
                .Cells(lastLine, LN_DOCNO).value = docNo
                .Cells(lastLine, LN_LINENO).value = ln
                .Cells(lastLine, LN_CODE).value = ws.Range("B" & r).value
                .Cells(lastLine, LN_ZCODE).value = ws.Range("C" & r).value
                .Cells(lastLine, LN_DESC).value = ws.Range("D" & r).value
                .Cells(lastLine, LN_QTY).value = ws.Range("A" & r).value
                .Cells(lastLine, LN_EXCL).value = ws.Range("E" & r).value
                .Cells(lastLine, LN_VAT).value = ws.Range("F" & r).value
                .Cells(lastLine, LN_INCL).value = ws.Range("G" & r).value
                .Cells(lastLine, LN_TOTAL).value = ws.Range("H" & r).value
            End With
        End If
    Next r

    ' 6. AUDIT
    If isUpdate Then
        LogAudit "Update", docNo, "Total " & oldTotal, "Total " & newTotal, "Med claim updated"
    Else
        LogAudit "Save", docNo, "", "Total " & newTotal, "New med claim saved"
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' 7. CONFIRM + ASK CLEAR
    If gSuppressClearPrompt Then Exit Sub
    If MsgBox("Med Claim " & docNo & " saved." & vbCrLf & vbCrLf & _
              "Clear the sheet for a new claim?", vbQuestion + vbYesNo) = vbYes Then
        NewMedClaim
    End If
    Exit Sub

CleanExit:
    Application.EnableEvents = True: Application.ScreenUpdating = True: Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "SaveMedClaim error: " & Err.Description, vbExclamation
End Sub

' ========================================================== RECALL MED CLAIM ==
Public Sub RecallMedClaim(Optional ByVal docNoIn As String = "")
    Dim ws As Worksheet, wsLog As Worksheet, wsLines As Worksheet
    Dim docNo As String, lr As Long

    Set ws = ThisWorkbook.Sheets("Med Claim")
    Set wsLog = ThisWorkbook.Sheets("MedAidLog")
    Set wsLines = ThisWorkbook.Sheets("MedAidLines")

    If docNoIn <> "" Then
        docNo = Trim(docNoIn)
    Else
        docNo = Trim(InputBox("Enter Med Claim number to recall (e.g. MC-INV-WA-0001):", "Recall Med Claim"))
    End If
    If docNo = "" Then Exit Sub
    lr = FindLogRow(wsLog, docNo)
    If lr = 0 Then MsgBox "Med Claim " & docNo & " not found.", vbExclamation: Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    NewMedClaim

    ws.Range("K1").value = wsLog.Cells(lr, ML_DEPT).value          ' dept
    ws.Range("K2").value = "patient"
    ws.Range("C6").value = wsLog.Cells(lr, ML_CUST).value          ' Bill To
    ws.Range("C7").value = wsLog.Cells(lr, ML_PATIENT).value
    ws.Range("C8").value = wsLog.Cells(lr, ML_PATIENTID).value
    ws.Range("G8").value = wsLog.Cells(lr, ML_MEDAID).value
    ws.Range("G9").value = wsLog.Cells(lr, ML_MEDNO).value
    ws.Range("G10").value = wsLog.Cells(lr, ML_MAINMEM).value
    ws.Range("G11").value = wsLog.Cells(lr, ML_DEPCODE).value
    ws.Range("G12").value = wsLog.Cells(lr, ML_DOCTOR).value
    ws.Range("G13").value = wsLog.Cells(lr, ML_BHF).value
    ws.Range("C14").value = wsLog.Cells(lr, ML_APPLIANCE).value

    LoadMedClaimLines wsLines, ws, docNo

    ws.Range("C32").value = wsLog.Cells(lr, ML_DISCPCT).value
    ws.Range("C33").value = wsLog.Cells(lr, ML_DISCFIX).value
    ws.Range("A40").value = wsLog.Cells(lr, ML_NOTES).value        ' restore Notes

    ws.Range("G7").value = docNo
    ws.Range("K4").value = docNo
    ws.Range("G6").value = wsLog.Cells(lr, ML_DATE).value          ' date last

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Med Claim " & docNo & " recalled.", vbInformation
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "RecallMedClaim error: " & Err.Description, vbExclamation
End Sub

Private Sub LoadMedClaimLines(wsLines As Worksheet, ws As Worksheet, docNo As String)
    Dim lastRow As Long, i As Long, destRow As Long
    destRow = 16
    lastRow = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If NrmID(CStr(wsLines.Cells(i, LN_DOCNO).value)) = NrmID(docNo) Then
            If destRow > 30 Then Exit For
            ws.Range("A" & destRow).value = wsLines.Cells(i, LN_QTY).value
            ws.Range("D" & destRow).value = wsLines.Cells(i, LN_DESC).value
            ws.Range("G" & destRow).value = wsLines.Cells(i, LN_INCL).value
            destRow = destRow + 1
        End If
    Next i
End Sub

