Attribute VB_Name = "modRenumber"
' modRenumber ? dept-change renumber engine + revert-converted-quote.
'   RenumberOnDeptChange: when a saved INV/QTE/MC has its department corrected,
'     assign the next number in sequence for the NEW dept, move header + lines
'     + payments, delete the OLD record, audit, prompt-confirmed.
'   RevertQuoteToSaved: un-convert a quote; DELETES the (unpaid) linked invoice.
' ============================================================================
Option Explicit

' Called by Save* when stored dept <> sheet dept. docKind: "INV" | "QTE" | "MC"
Public Sub RenumberOnDeptChange(ByVal docKind As String, ByVal oldNo As String, _
                                ByVal oldDept As String, ByVal newDept As String)
    Dim cfg As DocConfig, wsLog As Worksheet, wsLines As Worksheet
    Dim lr As Long, newNo As String, i As Long, lastLine As Long
    Dim paid As Double

    cfg = GetDocConfig(docKind)
    Set wsLog = ThisWorkbook.Sheets(cfg.logName)
    Set wsLines = ThisWorkbook.Sheets(cfg.linesName)

    lr = FindLogRow(wsLog, oldNo)
    If lr = 0 Then MsgBox oldNo & " not found for renumber.", vbExclamation: Exit Sub

    ' QUOTE conversion conflict block
    If docKind = "QTE" Then
        If Trim(CStr(wsLog.Cells(lr, QL_CONVINV).value)) <> "" Then
            MsgBox "This quote is linked to invoice " & _
                   CStr(wsLog.Cells(lr, QL_CONVINV).value) & "." & vbCrLf & _
                   "Void or revert that invoice before changing the department.", _
                   vbExclamation, "Renumber blocked"
            Exit Sub
        End If
    End If

    ' Allocate the new number for the NEW dept
    newNo = NextUniqueDocNumber(wsLog, newDept, docKind)

    ' Confirm with the user
    Dim msg As String
    msg = "Department changed from " & oldDept & " to " & newDept & "." & vbCrLf & vbCrLf & _
          "This will reassign the number:" & vbCrLf & _
          "    " & oldNo & "  ->  " & newNo & vbCrLf & vbCrLf
    If cfg.colPaid > 0 Then msg = msg & "Any payments will be moved to the new number." & vbCrLf & vbCrLf
    msg = msg & "The old record will be deleted. Continue?"
    If MsgBox(msg, vbQuestion + vbYesNo, "Confirm renumber") <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 1. write the new number into the sheet + log header cell
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(cfg.sheetName)
    ws.Range("G7").value = newNo
    ws.Range("K4").value = newNo
    wsLog.Cells(lr, cfg.colNo).value = newNo
    wsLog.Cells(lr, cfg.colDept).value = newDept

    ' 2. move line rows to the new number
    lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastLine
        If NrmID(CStr(wsLines.Cells(i, LN_DOCNO).value)) = NrmID(oldNo) Then
            wsLines.Cells(i, LN_DOCNO).value = newNo
        End If
    Next i

    ' 3. move payments (invoices/MC only)
    If cfg.colPaid > 0 Then RepointPayments oldNo, newNo

    ' 4. audit + delete-old-not-needed (row was renamed in place, so no orphan)
    LogAudit "Renumber", oldNo, oldNo, newNo, docKind & " dept " & oldDept & "->" & newDept

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Renumbered to " & newNo & "." & vbCrLf & _
           "Re-saving the document now to sync all fields.", vbInformation

    ' 5. re-save so header totals/lines fully re-sync under the new number
    Select Case docKind
        Case "INV": SaveInvoice
        Case "QTE": SaveQuote
        Case "MC":  SaveMedClaim
    End Select
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "Renumber error: " & Err.Description, vbExclamation
End Sub

' Repoint all Payments rows from oldNo to newNo
Private Sub RepointPayments(ByVal oldNo As String, ByVal newNo As String)
    Dim wsP As Worksheet, last As Long, i As Long
    Set wsP = ThisWorkbook.Sheets("Payments")
    last = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsP.Cells(i, PY_INV).value)) = NrmID(oldNo) Then
            wsP.Cells(i, PY_INV).value = newNo
        End If
    Next i
End Sub

' ============================================================ REVERT QUOTE ====
' Un-converts a quote AND deletes its (unpaid) linked invoice.
'   Blocks if the linked invoice has Paid > 0.
Public Sub RevertQuoteToSaved(Optional ByVal quoteNoIn As String = "")
    Dim wsQ As Worksheet, wsI As Worksheet, wsIL As Worksheet
    Dim quoteNo As String, qr As Long, invNo As String, ir As Long
    Dim paid As Double

    Set wsQ = ThisWorkbook.Sheets("QuoteLog")
    Set wsIL = ThisWorkbook.Sheets("InvoiceLog")
    Set wsI = ThisWorkbook.Sheets("InvoiceLines")

    If quoteNoIn <> "" Then
        quoteNo = Trim(quoteNoIn)
    Else
        quoteNo = Trim(InputBox("Enter Quote number to revert (e.g. Q-WA-0012):", "Revert Quote"))
    End If
    If quoteNo = "" Then Exit Sub

    qr = FindLogRow(wsQ, quoteNo)
    If qr = 0 Then MsgBox "Quote " & quoteNo & " not found.", vbExclamation: Exit Sub

    invNo = Trim(CStr(wsQ.Cells(qr, QL_CONVINV).value))
    If invNo = "" Then
        MsgBox "Quote " & quoteNo & " is not converted ? nothing to revert.", vbInformation
        Exit Sub
    End If

    ' Safety gate: linked invoice must be unpaid
    ir = FindLogRow(wsIL, invNo)
    If ir > 0 Then
        paid = Num(wsIL.Cells(ir, IL_PAID).value)
        If paid > 0 Then
            MsgBox "Invoice " & invNo & " has payments recorded (R " & _
                   Format(paid, "#,##0.00") & ")." & vbCrLf & _
                   "Revert cancelled ? issue a credit note or handle the payment first.", _
                   vbExclamation, "Revert blocked"
            Exit Sub
        End If
    End If

    If MsgBox("This will DELETE invoice " & invNo & " (log + line items) and revert " & _
              "quote " & quoteNo & " to 'Saved' so it can be edited / re-converted." & vbCrLf & vbCrLf & _
              "This cannot be undone. Continue?", vbCritical + vbYesNo, "Revert Quote") <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' delete invoice lines + header
    If ir > 0 Then
        Dim last As Long, i As Long
        last = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).row
        For i = last To 2 Step -1
            If NrmID(CStr(wsI.Cells(i, LN_DOCNO).value)) = NrmID(invNo) Then wsI.Rows(i).Delete
        Next i
        wsIL.Rows(ir).Delete
    End If

    ' revert quote
    wsQ.Cells(qr, QL_STATUS).value = "Saved"
    wsQ.Cells(qr, QL_CONVINV).value = ""

    LogAudit "Revert", invNo, "", "DELETED", "Invoice deleted (quote " & quoteNo & " reverted)"
    LogAudit "Revert", quoteNo, "Converted", "Saved", "Quote reverted; invoice " & invNo & " deleted"

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    On Error Resume Next
    RefreshMenuSummary
    On Error GoTo 0

    MsgBox "Quote " & quoteNo & " reverted to Saved; invoice " & invNo & " deleted.", vbInformation
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "RevertQuoteToSaved error: " & Err.Description, vbExclamation
End Sub

