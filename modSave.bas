Attribute VB_Name = "modSave"
' ============================================================================
' modSave — SaveInvoice / SaveQuote / SaveCreditNote
'   FIX 1: honor gSuppressClearPrompt (set by ConvertQuoteToInvoice)
'   FIX 2: write med-aid fields on patient docs
'          Invoice T-X (20-24) | Quote Q-U (17-21)
'   PHASE 3: persist discount %/fixed
'          Invoice Y-Z (25-26) | Quote V-W (22-23)
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
    If Trim(ws.Range("C6").Value) = "" And LCase(Trim(CStr(ws.Range("K2").Value))) <> "patient" Then
        MsgBox "Please select a Doctor/Practice (Bill To) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Not IsDate(ws.Range("G6").Value) Then
        MsgBox "Please enter a valid Invoice Date (G6) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Trim(ws.Range("D16").Value) = "" Then
        MsgBox "Add at least one line item before saving.", vbExclamation: GoTo CleanExit
    End If

    dept = UCase(Trim(ws.Range("K1").Value))
    isPatient = (LCase(Trim(CStr(ws.Range("K2").Value))) = "patient")

    ' 2. DETERMINE NUMBER
    docNo = Trim(ws.Range("G7").Value)
    If docNo = "" Then
        Do
            docNo = NextDocNumber(dept, "INV")
        Loop While FindLogRow(wsLog, docNo) > 0
        ws.Range("G7").Value = docNo
        ws.Range("K4").Value = docNo
        isUpdate = False
    Else
        logRow = FindLogRow(wsLog, docNo)
        isUpdate = (logRow > 0)
    End If

    ' 3. TARGET ROW
    If isUpdate Then
        oldTotal = CStr(wsLog.Cells(logRow, "L").Value)
    Else
        logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
    End If

    ' 4. HEADER
    With wsLog
        .Cells(logRow, 1).Value = docNo
        .Cells(logRow, 2).Value = dept
        .Cells(logRow, 3).Value = ws.Range("K2").Value
        .Cells(logRow, 4).Value = ws.Range("G6").Value
        .Cells(logRow, 5).Value = ws.Range("G11").Value
        .Cells(logRow, 6).Value = IIf(isPatient, "", ws.Range("G8").Value)
        .Cells(logRow, 7).Value = ws.Range("F14").Value
        .Cells(logRow, 8).Value = ws.Range("C14").Value
        .Cells(logRow, 9).Value = ws.Range("H35").Value
        .Cells(logRow, 10).Value = ws.Range("H34").Value
        .Cells(logRow, 11).Value = ws.Range("H36").Value
        .Cells(logRow, 12).Value = ws.Range("H37").Value
        .Cells(logRow, 13).Value = ""
        If Not isUpdate Then
            .Cells(logRow, 14).Value = 0
            .Cells(logRow, 15).Value = ws.Range("H37").Value
            .Cells(logRow, 16).Value = "Unpaid"
            .Cells(logRow, 18).Value = Now
        End If
        .Cells(logRow, 19).Value = Now
        ' --- FIX 2: med-aid write-back (T-X = 20-24) ---
        If isPatient Then
            .Cells(logRow, 20).Value = ws.Range("G8").Value    ' T MedAid
            .Cells(logRow, 21).Value = ws.Range("G9").Value    ' U MedNo
            .Cells(logRow, 22).Value = ws.Range("G10").Value   ' V MainMember
            .Cells(logRow, 23).Value = ws.Range("G12").Value   ' W Doctor
            .Cells(logRow, 24).Value = ws.Range("G13").Value   ' X BHF
        Else
            .Range(.Cells(logRow, 20), .Cells(logRow, 24)).ClearContents
        End If
        ' --- PHASE 3: discount persistence (Y=25 %, Z=26 fixed) ---
        .Cells(logRow, 25).Value = ws.Range("C32").Value   ' Discount %  (decimal, e.g. 0.1)
        .Cells(logRow, 26).Value = ws.Range("C33").Value   ' Discount Fixed (R)
    End With
    newTotal = CStr(ws.Range("H37").Value)

    ' 5. LINES
    lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For r = lastLine To 2 Step -1
        If CStr(wsLines.Cells(r, "A").Value) = docNo Then wsLines.Rows(r).Delete
    Next r
    ln = 0
    For r = 16 To 30
        If Trim(ws.Range("D" & r).Value) <> "" Then
            ln = ln + 1
            lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row + 1
            With wsLines
                .Cells(lastLine, 1).Value = docNo
                .Cells(lastLine, 2).Value = ln
                .Cells(lastLine, 3).Value = ws.Range("B" & r).Value
                .Cells(lastLine, 4).Value = ws.Range("C" & r).Value
                .Cells(lastLine, 5).Value = ws.Range("D" & r).Value
                .Cells(lastLine, 6).Value = ws.Range("A" & r).Value
                .Cells(lastLine, 7).Value = ws.Range("E" & r).Value
                .Cells(lastLine, 8).Value = ws.Range("F" & r).Value
                .Cells(lastLine, 9).Value = ws.Range("G" & r).Value
                .Cells(lastLine, 10).Value = ws.Range("H" & r).Value
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
    If Trim(ws.Range("C6").Value) = "" And LCase(Trim(CStr(ws.Range("K2").Value))) <> "patient" Then
        MsgBox "Please select a Doctor/Practice (Bill To) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Not IsDate(ws.Range("G6").Value) Then
        MsgBox "Please enter a valid Quote Date (G6) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Trim(ws.Range("D16").Value) = "" Then
        MsgBox "Add at least one line item before saving.", vbExclamation: GoTo CleanExit
    End If

    dept = UCase(Trim(ws.Range("K1").Value))
    isPatient = (LCase(Trim(CStr(ws.Range("K2").Value))) = "patient")

    ' 2. DETERMINE NUMBER
    docNo = Trim(ws.Range("G7").Value)
    If docNo = "" Then
        Do
            docNo = NextDocNumber(dept, "QTE")
        Loop While FindLogRow(wsLog, docNo) > 0
        ws.Range("G7").Value = docNo
        ws.Range("K4").Value = docNo
        isUpdate = False
    Else
        logRow = FindLogRow(wsLog, docNo)
        isUpdate = (logRow > 0)
    End If

    ' 3. TARGET HEADER ROW
    If isUpdate Then
        oldTotal = CStr(wsLog.Cells(logRow, "K").Value)
    Else
        logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
    End If

    ' 4. WRITE HEADER
    With wsLog
        .Cells(logRow, 1).Value = docNo
        .Cells(logRow, 2).Value = dept
        .Cells(logRow, 3).Value = ws.Range("K2").Value
        .Cells(logRow, 4).Value = ws.Range("G6").Value
        .Cells(logRow, 5).Value = IIf(isPatient, "", ws.Range("G8").Value)
        .Cells(logRow, 6).Value = ws.Range("F14").Value
        .Cells(logRow, 7).Value = ws.Range("C14").Value
        .Cells(logRow, 8).Value = ws.Range("H35").Value
        .Cells(logRow, 9).Value = ws.Range("H34").Value
        .Cells(logRow, 10).Value = ws.Range("H36").Value
        .Cells(logRow, 11).Value = ws.Range("H37").Value
        .Cells(logRow, 12).Value = ws.Range("K3").Value
        If Not isUpdate Then
            .Cells(logRow, 13).Value = "Saved"
            .Cells(logRow, 15).Value = Now
        End If
        .Cells(logRow, 16).Value = Now
        ' --- FIX 2: med-aid write-back (Q-U = 17-21) ---
        If isPatient Then
            .Cells(logRow, 17).Value = ws.Range("G8").Value    ' Q MedAid
            .Cells(logRow, 18).Value = ws.Range("G9").Value    ' R MedNo
            .Cells(logRow, 19).Value = ws.Range("G10").Value   ' S MainMember
            .Cells(logRow, 20).Value = ws.Range("G11").Value   ' T Doctor
            .Cells(logRow, 21).Value = ws.Range("G12").Value   ' U BHF
        Else
            .Range(.Cells(logRow, 17), .Cells(logRow, 21)).ClearContents
        End If
        ' --- PHASE 3: discount persistence (V=22 %, W=23 fixed) ---
        .Cells(logRow, 22).Value = ws.Range("C32").Value   ' Discount %  (decimal)
        .Cells(logRow, 23).Value = ws.Range("C33").Value   ' Discount Fixed (R)
    End With
    newTotal = CStr(ws.Range("H37").Value)

    ' 5. WRITE LINES
    lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For r = lastLine To 2 Step -1
        If CStr(wsLines.Cells(r, "A").Value) = docNo Then wsLines.Rows(r).Delete
    Next r
    ln = 0
    For r = 16 To 30
        If Trim(ws.Range("D" & r).Value) <> "" Then
            ln = ln + 1
            lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row + 1
            With wsLines
                .Cells(lastLine, 1).Value = docNo
                .Cells(lastLine, 2).Value = ln
                .Cells(lastLine, 3).Value = ws.Range("B" & r).Value
                .Cells(lastLine, 4).Value = ws.Range("C" & r).Value
                .Cells(lastLine, 5).Value = ws.Range("D" & r).Value
                .Cells(lastLine, 6).Value = ws.Range("A" & r).Value
                .Cells(lastLine, 7).Value = ws.Range("E" & r).Value
                .Cells(lastLine, 8).Value = ws.Range("F" & r).Value
                .Cells(lastLine, 9).Value = ws.Range("G" & r).Value
                .Cells(lastLine, 10).Value = ws.Range("H" & r).Value
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

    ' 7. CONFIRM + ASK CLEAR  (FIX 1: skip when suppressed)
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
    Dim docNo As String
    Dim logRow As Long, isUpdate As Boolean
    Dim r As Long, ln As Long, lastLine As Long
    Dim oldAmt As String, newAmt As String

    Set ws = ThisWorkbook.Sheets("CreditNote")     ' FIX 3: no trailing space
    Set wsLog = ThisWorkbook.Sheets("CreditNotes")
    Set wsLines = ThisWorkbook.Sheets("CreditNoteLines")

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 1. VALIDATE
    If Trim(ws.Range("C6").Value) = "" Then
        MsgBox "Please select a Doctor/Practice (Bill To) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Not IsDate(ws.Range("G6").Value) Then
        MsgBox "Please enter a valid Credit Note Date (G6) before saving.", vbExclamation: GoTo CleanExit
    End If
    If Trim(ws.Range("D16").Value) = "" Then
        MsgBox "Add at least one line item before saving.", vbExclamation: GoTo CleanExit
    End If
    If Trim(ws.Range("C14").Value) = "" Then
        MsgBox "Please select a Reason (C14) before saving.", vbExclamation: GoTo CleanExit
    End If

    ' 2. DETERMINE NUMBER
    docNo = Trim(ws.Range("G7").Value)
    If docNo = "" Then
        Do
            docNo = NextDocNumber("", "CN")
        Loop While FindLogRow(wsLog, docNo) > 0
        ws.Range("G7").Value = docNo
        ws.Range("K4").Value = docNo
        isUpdate = False
    Else
        logRow = FindLogRow(wsLog, docNo)
        isUpdate = (logRow > 0)
    End If

    ' 3. TARGET ROW
    If isUpdate Then
        oldAmt = CStr(wsLog.Cells(logRow, "E").Value)
    Else
        logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
    End If

    ' 4. WRITE HEADER (8 cols)
    With wsLog
        .Cells(logRow, 1).Value = docNo
        .Cells(logRow, 2).Value = ws.Range("G11").Value
        .Cells(logRow, 3).Value = ws.Range("G6").Value
        .Cells(logRow, 4).Value = ws.Range("G8").Value
        .Cells(logRow, 5).Value = ws.Range("H37").Value
        .Cells(logRow, 6).Value = ws.Range("H36").Value
        .Cells(logRow, 7).Value = ws.Range("C14").Value
        If Not isUpdate Then
            .Cells(logRow, 8).Value = "Issued"
        End If
    End With
    newAmt = CStr(ws.Range("H37").Value)

    ' 5. WRITE LINES
    lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For r = lastLine To 2 Step -1
        If CStr(wsLines.Cells(r, "A").Value) = docNo Then wsLines.Rows(r).Delete
    Next r
    ln = 0
    For r = 16 To 30
        If Trim(ws.Range("D" & r).Value) <> "" Then
            ln = ln + 1
            lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row + 1
            With wsLines
                .Cells(lastLine, 1).Value = docNo
                .Cells(lastLine, 2).Value = ln
                .Cells(lastLine, 3).Value = ws.Range("B" & r).Value
                .Cells(lastLine, 4).Value = ws.Range("C" & r).Value
                .Cells(lastLine, 5).Value = ws.Range("D" & r).Value
                .Cells(lastLine, 6).Value = ws.Range("A" & r).Value
                .Cells(lastLine, 7).Value = ws.Range("E" & r).Value
                .Cells(lastLine, 8).Value = ws.Range("F" & r).Value
                .Cells(lastLine, 9).Value = ws.Range("G" & r).Value
                .Cells(lastLine, 10).Value = ws.Range("H" & r).Value
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



