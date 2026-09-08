Attribute VB_Name = "modSave"
' modSave ? SaveInvoice / SaveQuote / SaveCreditNote
'   FIX 1: honor gSuppressClearPrompt (set by ConvertQuoteToInvoice)
'   FIX 2: write med-aid fields on patient docs
'          Invoice T-X (20-24) | Quote Q-U (17-21)
'   PHASE 3: persist discount %/fixed
'          Invoice Y-Z (25-26) | Quote V-W (22-23)
'   3C: CreditNote validates + applies credit to source invoice; supports
'       Doctor AND Patient recipients (mirrors SaveInvoice).
' ============================================================================
Option Explicit
Public gSuppressClearPrompt As Boolean

Sub SaveInvoice()
    Dim ws As Worksheet, wsLog As Worksheet, wsLines As Worksheet
    Dim docNo As String, dept As String, isPatient As Boolean
    Dim logRow As Long, isUpdate As Boolean
    Dim r As Long, ln As Long, lastLine As Long
    Dim oldTotal As String, newTotal As String

    Set ws = ThisWorkbook.Sheets("Invoice")
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    Set wsLines = ThisWorkbook.Sheets("InvoiceLines")

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 1. VALIDATE
    If Trim(ws.Range("C6").value) = "" And LCase(Trim(CStr(ws.Range("K2").value))) <> "patient" Then
        MsgBox "Please select a Doctor/Practice (Bill To) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Not IsDate(ws.Range("G6").value) Then
        MsgBox "Please enter a valid Invoice Date (G6) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Trim(ws.Range("D16").value) = "" Then
        MsgBox "Add at least one line item before saving.", vbExclamation: GoTo CleanExit
    End If

    dept = UCase(Trim(ws.Range("K1").value))
    isPatient = (LCase(Trim(CStr(ws.Range("K2").value))) = "patient")

    ' 2. DETERMINE NUMBER
    docNo = Trim(ws.Range("G7").value)
    If docNo = "" Then
        Do
            docNo = NextDocNumber(dept, "INV")
        Loop While FindLogRow(wsLog, docNo) > 0
        ws.Range("G7").value = docNo
        ws.Range("K4").value = docNo
        isUpdate = False
    Else
        logRow = FindLogRow(wsLog, docNo)
        isUpdate = (logRow > 0)
    End If

    ' 3. TARGET ROW
    If isUpdate Then
        oldTotal = CStr(wsLog.Cells(logRow, "L").value)
    Else
        logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
    End If

    ' 4. HEADER
    With wsLog
        .Cells(logRow, 1).value = docNo
        .Cells(logRow, 2).value = dept
        .Cells(logRow, 3).value = ws.Range("K2").value
        .Cells(logRow, 4).value = ws.Range("G6").value
        .Cells(logRow, 5).value = ws.Range("G11").value
        .Cells(logRow, 6).value = IIf(isPatient, "", ws.Range("G8").value)
        .Cells(logRow, 7).value = ws.Range("F14").value
        .Cells(logRow, 8).value = ws.Range("C14").value
        .Cells(logRow, 9).value = ws.Range("H35").value
        .Cells(logRow, 10).value = ws.Range("H34").value
        .Cells(logRow, 11).value = ws.Range("H36").value
        .Cells(logRow, 12).value = ws.Range("H37").value
        .Cells(logRow, 13).value = ""
        If Not isUpdate Then
            .Cells(logRow, 14).value = 0
            .Cells(logRow, 15).value = ws.Range("H37").value
            .Cells(logRow, 16).value = "Unpaid"
            .Cells(logRow, 18).value = Now
        End If
        .Cells(logRow, 19).value = Now
        ' --- FIX 2: med-aid write-back (T-X = 20-24) ---
        If isPatient Then
            .Cells(logRow, 20).value = ws.Range("G8").value    ' T MedAid
            .Cells(logRow, 21).value = ws.Range("G9").value    ' U MedNo
            .Cells(logRow, 22).value = ws.Range("G10").value   ' V MainMember
            .Cells(logRow, 23).value = ws.Range("G12").value   ' W Doctor
            .Cells(logRow, 24).value = ws.Range("G13").value   ' X BHF
        Else
            .Range(.Cells(logRow, 20), .Cells(logRow, 24)).ClearContents
        End If
        ' --- PHASE 3: discount persistence (Y=25 %, Z=26 fixed) ---
        .Cells(logRow, 25).value = ws.Range("C32").value
        .Cells(logRow, 26).value = ws.Range("C33").value
    End With
    newTotal = CStr(ws.Range("H37").value)

    ' 5. LINES
    lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For r = lastLine To 2 Step -1
        If CStr(wsLines.Cells(r, "A").value) = docNo Then wsLines.Rows(r).Delete
    Next r
    ln = 0
    For r = 16 To 30
        If Trim(ws.Range("D" & r).value) <> "" Then
            ln = ln + 1
            lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row + 1
            With wsLines
                .Cells(lastLine, 1).value = docNo
                .Cells(lastLine, 2).value = ln
                .Cells(lastLine, 3).value = ws.Range("B" & r).value
                .Cells(lastLine, 4).value = ws.Range("C" & r).value
                .Cells(lastLine, 5).value = ws.Range("D" & r).value
                .Cells(lastLine, 6).value = ws.Range("A" & r).value
                .Cells(lastLine, 7).value = ws.Range("E" & r).value
                .Cells(lastLine, 8).value = ws.Range("F" & r).value
                .Cells(lastLine, 9).value = ws.Range("G" & r).value
                .Cells(lastLine, 10).value = ws.Range("H" & r).value
            End With
        End If
    Next r

    ' 6. AUDIT
    If isUpdate Then
        LogAudit "Update", docNo, "Total " & oldTotal, "Total " & newTotal, "Invoice updated"
    Else
        LogAudit "Save", docNo, "", "Total " & newTotal, "New invoice saved"
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' 7. CONFIRM + ASK CLEAR  (FIX 1: skip when suppressed)
    If gSuppressClearPrompt Then Exit Sub
    If MsgBox("Invoice " & docNo & " saved." & vbCrLf & vbCrLf & _
              "Clear the sheet for a new invoice?", vbQuestion + vbYesNo) = vbYes Then
        NewInvoice
    End If
    Exit Sub

CleanExit:
    Application.EnableEvents = True: Application.ScreenUpdating = True: Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "SaveInvoice error: " & Err.Description, vbExclamation
End Sub

' ============================================================================
Sub SaveQuote()
    Dim ws As Worksheet, wsLog As Worksheet, wsLines As Worksheet
    Dim docNo As String, dept As String, isPatient As Boolean
    Dim logRow As Long, isUpdate As Boolean
    Dim r As Long, ln As Long, lastLine As Long
    Dim oldTotal As String, newTotal As String

    Set ws = ThisWorkbook.Sheets("Quote")
    Set wsLog = ThisWorkbook.Sheets("QuoteLog")
    Set wsLines = ThisWorkbook.Sheets("QuoteLines")

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 1. VALIDATE
    If Trim(ws.Range("C6").value) = "" And LCase(Trim(CStr(ws.Range("K2").value))) <> "patient" Then
        MsgBox "Please select a Doctor/Practice (Bill To) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Not IsDate(ws.Range("G6").value) Then
        MsgBox "Please enter a valid Quote Date (G6) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Trim(ws.Range("D16").value) = "" Then
        MsgBox "Add at least one line item before saving.", vbExclamation: GoTo CleanExit
    End If

    dept = UCase(Trim(ws.Range("K1").value))
    isPatient = (LCase(Trim(CStr(ws.Range("K2").value))) = "patient")

    ' 2. DETERMINE NUMBER
    docNo = Trim(ws.Range("G7").value)
    If docNo = "" Then
        Do
            docNo = NextDocNumber(dept, "QTE")
        Loop While FindLogRow(wsLog, docNo) > 0
        ws.Range("G7").value = docNo
        ws.Range("K4").value = docNo
        isUpdate = False
    Else
        logRow = FindLogRow(wsLog, docNo)
        isUpdate = (logRow > 0)
    End If

    ' 3. TARGET HEADER ROW
    If isUpdate Then
        oldTotal = CStr(wsLog.Cells(logRow, "K").value)
    Else
        logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
    End If

    ' 4. WRITE HEADER
    With wsLog
        .Cells(logRow, 1).value = docNo
        .Cells(logRow, 2).value = dept
        .Cells(logRow, 3).value = ws.Range("K2").value
        .Cells(logRow, 4).value = ws.Range("G6").value
        .Cells(logRow, 5).value = IIf(isPatient, "", ws.Range("G8").value)
        .Cells(logRow, 6).value = ws.Range("F14").value
        .Cells(logRow, 7).value = ws.Range("C14").value
        .Cells(logRow, 8).value = ws.Range("H35").value
        .Cells(logRow, 9).value = ws.Range("H34").value
        .Cells(logRow, 10).value = ws.Range("H36").value
        .Cells(logRow, 11).value = ws.Range("H37").value
        .Cells(logRow, 12).value = ws.Range("K3").value
        If Not isUpdate Then
            .Cells(logRow, 13).value = "Saved"
            .Cells(logRow, 15).value = Now
        End If
        .Cells(logRow, 16).value = Now
        ' --- FIX 2: med-aid write-back (Q-U = 17-21) ---
        If isPatient Then
            .Cells(logRow, 17).value = ws.Range("G8").value    ' Q MedAid
            .Cells(logRow, 18).value = ws.Range("G9").value    ' R MedNo
            .Cells(logRow, 19).value = ws.Range("G10").value   ' S MainMember
            .Cells(logRow, 20).value = ws.Range("G11").value   ' T Doctor
            .Cells(logRow, 21).value = ws.Range("G12").value   ' U BHF
        Else
            .Range(.Cells(logRow, 17), .Cells(logRow, 21)).ClearContents
        End If
        ' --- PHASE 3: discount persistence (V=22 %, W=23 fixed) ---
        .Cells(logRow, 22).value = ws.Range("C32").value
        .Cells(logRow, 23).value = ws.Range("C33").value
    End With
    newTotal = CStr(ws.Range("H37").value)

    ' 5. WRITE LINES
    lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For r = lastLine To 2 Step -1
        If CStr(wsLines.Cells(r, "A").value) = docNo Then wsLines.Rows(r).Delete
    Next r
    ln = 0
    For r = 16 To 30
        If Trim(ws.Range("D" & r).value) <> "" Then
            ln = ln + 1
            lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row + 1
            With wsLines
                .Cells(lastLine, 1).value = docNo
                .Cells(lastLine, 2).value = ln
                .Cells(lastLine, 3).value = ws.Range("B" & r).value
                .Cells(lastLine, 4).value = ws.Range("C" & r).value
                .Cells(lastLine, 5).value = ws.Range("D" & r).value
                .Cells(lastLine, 6).value = ws.Range("A" & r).value
                .Cells(lastLine, 7).value = ws.Range("E" & r).value
                .Cells(lastLine, 8).value = ws.Range("F" & r).value
                .Cells(lastLine, 9).value = ws.Range("G" & r).value
                .Cells(lastLine, 10).value = ws.Range("H" & r).value
            End With
        End If
    Next r

    ' 6. AUDIT
    If isUpdate Then
        LogAudit "Update", docNo, "Total " & oldTotal, "Total " & newTotal, "Quote updated"
    Else
        LogAudit "Save", docNo, "", "Total " & newTotal, "New quote saved"
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' 7. CONFIRM + ASK CLEAR
    If gSuppressClearPrompt Then Exit Sub
    If MsgBox("Quote " & docNo & " saved." & vbCrLf & vbCrLf & _
              "Clear the sheet for a new quote?", vbQuestion + vbYesNo) = vbYes Then
        NewQuote
    End If
    Exit Sub

CleanExit:
    Application.EnableEvents = True: Application.ScreenUpdating = True: Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "SaveQuote error: " & Err.Description, vbExclamation
End Sub

' ============================================================================
Sub SaveCreditNote()
    Dim ws As Worksheet, wsLog As Worksheet, wsLines As Worksheet
    Dim docNo As String, isPatient As Boolean
    Dim logRow As Long, isUpdate As Boolean
    Dim r As Long, ln As Long, lastLine As Long
    Dim oldAmt As String, newAmt As String

    Set ws = ThisWorkbook.Sheets("CreditNote")
    Set wsLog = ThisWorkbook.Sheets("CreditNotes")
    Set wsLines = ThisWorkbook.Sheets("CreditNoteLines")

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 1. VALIDATE
    isPatient = (LCase(Trim(CStr(ws.Range("K2").value))) = "patient")
    If Not isPatient And Trim(ws.Range("C6").value) = "" Then
        MsgBox "Please select a Doctor/Practice (Bill To) before saving.", vbExclamation: GoTo CleanExit
    End If
    If isPatient And Trim(ws.Range("F14").value) = "" Then
        MsgBox "Please enter the Patient Name (F14) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Not IsDate(ws.Range("G6").value) Then
        MsgBox "Please enter a valid Credit Note Date (G6) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Trim(ws.Range("D16").value) = "" Then
        MsgBox "Add at least one line item before saving.", vbExclamation: GoTo CleanExit
    End If
    If Trim(ws.Range("C14").value) = "" Then
        MsgBox "Please select a Reason (C14) before saving.", vbExclamation: GoTo CleanExit
    End If

    ' --- 3C: validate source invoice (G11) exists ---
    If Not ValidateSourceInvoice(CStr(ws.Range("G11").value)) Then GoTo CleanExit

    ' 2. DETERMINE NUMBER
    docNo = Trim(ws.Range("G7").value)
    If docNo = "" Then
        Do
            docNo = NextDocNumber("", "CN")
        Loop While FindLogRow(wsLog, docNo) > 0
        ws.Range("G7").value = docNo
        ws.Range("K4").value = docNo
        isUpdate = False
    Else
        logRow = FindLogRow(wsLog, docNo)
        isUpdate = (logRow > 0)
    End If

    ' 3. TARGET ROW
    If isUpdate Then
        oldAmt = CStr(wsLog.Cells(logRow, "E").value)
    Else
        logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
    End If

    ' 4. WRITE HEADER (8 cols)
    With wsLog
        .Cells(logRow, 1).value = docNo
        .Cells(logRow, 2).value = ws.Range("G11").value           ' SourceInvNo
        .Cells(logRow, 3).value = ws.Range("G6").value            ' Date
        .Cells(logRow, 4).value = IIf(isPatient, "", _
             IIf(Trim(ws.Range("G8").value) <> "", ws.Range("G8").value, DrNameToCustID_Save(CStr(ws.Range("C6").value))))
        .Cells(logRow, 5).value = ws.Range("H37").value           ' Total
        .Cells(logRow, 6).value = ws.Range("H36").value           ' VAT
        .Cells(logRow, 7).value = ws.Range("C14").value           ' Reason
        If Not isUpdate Then
            .Cells(logRow, 8).value = "Issued"
        End If
    End With
    newAmt = CStr(ws.Range("H37").value)

    ' --- 3C: apply credit to source invoice (new CN only) ---
    If Not isUpdate Then
        Dim cnTot As Double
        cnTot = 0
        If IsNumeric(ws.Range("H37").value) Then cnTot = CDbl(ws.Range("H37").value)
        ApplyCreditToInvoice docNo, CStr(ws.Range("G11").value), cnTot, ws.Range("G6").value
    End If

    ' 5. WRITE LINES
    lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For r = lastLine To 2 Step -1
        If CStr(wsLines.Cells(r, "A").value) = docNo Then wsLines.Rows(r).Delete
    Next r
    ln = 0
    For r = 16 To 30
        If Trim(ws.Range("D" & r).value) <> "" Then
            ln = ln + 1
            lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row + 1
            With wsLines
                .Cells(lastLine, 1).value = docNo
                .Cells(lastLine, 2).value = ln
                .Cells(lastLine, 3).value = ws.Range("B" & r).value
                .Cells(lastLine, 4).value = ws.Range("C" & r).value
                .Cells(lastLine, 5).value = ws.Range("D" & r).value
                .Cells(lastLine, 6).value = ws.Range("A" & r).value
                .Cells(lastLine, 7).value = ws.Range("E" & r).value
                .Cells(lastLine, 8).value = ws.Range("F" & r).value
                .Cells(lastLine, 9).value = ws.Range("G" & r).value
                .Cells(lastLine, 10).value = ws.Range("H" & r).value
            End With
        End If
    Next r

    ' 6. AUDIT
    If isUpdate Then
        LogAudit "Update", docNo, "Amount " & oldAmt, "Amount " & newAmt, "Credit note updated"
    Else
        LogAudit "Save", docNo, "", "Amount " & newAmt, "New credit note saved"
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' 7. CONFIRM + ASK CLEAR
    If gSuppressClearPrompt Then Exit Sub
    If MsgBox("Credit Note " & docNo & " saved." & vbCrLf & vbCrLf & _
              "Clear the sheet for a new credit note?", vbQuestion + vbYesNo) = vbYes Then
        NewCreditNote
    End If
    Exit Sub

CleanExit:
    Application.EnableEvents = True: Application.ScreenUpdating = True: Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "SaveCreditNote error: " & Err.Description, vbExclamation
End Sub

' Local Customers name->ID lookup (self-contained; used by SaveCreditNote)
Private Function DrNameToCustID_Save(drName As String) As String
    Dim wsC As Worksheet, last As Long, i As Long
    If Trim(drName) = "" Then Exit Function
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "B").End(xlUp).row
    For i = 2 To last
        If CStr(wsC.Cells(i, "B").value) = drName Then DrNameToCustID_Save = CStr(wsC.Cells(i, "A").value): Exit Function
    Next i
End Function

