Attribute VB_Name = "modCreditLink"
Option Explicit
' ============================================================================
' modCreditLink ? 3C: Credit Note <-> Invoice linking.
'   - ValidateSourceInvoice: confirms a CN's source invoice exists.
'   - LoadCreditFromInvoice: pull dept/recipient/details + lines from source inv.
'   - ApplyCreditToInvoice: record CN as a credit (Payments row, "Credit Note"),
'     reduce invoice Paid/Balance/Status (same model as frmPayment).
'   - PickSourceInvoice: opens frmSearch in pick-for-CN mode (Flow C).
' ============================================================================

Private Function Num(v As Variant) As Double
    If IsError(v) Then Num = 0: Exit Function
    If Trim(CStr(v)) = "" Then Num = 0: Exit Function
    If IsNumeric(v) Then Num = CDbl(v) Else Num = 0
End Function

Private Function NrmID(s As String) As String
    NrmID = UCase(Replace(Trim(s), " ", ""))
End Function

' Returns InvoiceLog row of invNo, or 0 if not found
Public Function InvoiceRow(invNo As String) As Long
    Dim ws As Worksheet, last As Long, i As Long
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(ws.Cells(i, 1).value)) = NrmID(invNo) Then InvoiceRow = i: Exit Function
    Next i
End Function

' True if source invoice exists; shows message + returns False otherwise
Public Function ValidateSourceInvoice(invNo As String) As Boolean
    If Trim(invNo) = "" Then
        MsgBox "This credit note has no Source Invoice (G11)." & vbCrLf & _
               "Double-click G11 to pick the invoice this credit applies to.", vbExclamation
        Exit Function
    End If
    If InvoiceRow(invNo) = 0 Then
        MsgBox "Source Invoice '" & invNo & "' was not found in InvoiceLog." & vbCrLf & _
               "Check the number (G11) before saving.", vbExclamation
        Exit Function
    End If
    ValidateSourceInvoice = True
End Function

' Pull dept/recipient/details + line items from a source invoice into the CN sheet
Public Sub LoadCreditFromInvoice(ByVal srcInv As String)
    Dim ws As Worksheet, wsLog As Worksheet, wsLines As Worksheet
    Dim lr As Long, recip As String, dept As String, drName As String
    Dim last As Long, i As Long, dest As Long
    Set ws = ThisWorkbook.Sheets("CreditNote")
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    Set wsLines = ThisWorkbook.Sheets("InvoiceLines")

    lr = InvoiceRow(srcInv)
    If lr = 0 Then MsgBox "Invoice '" & srcInv & "' not found.", vbExclamation: Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Clean

    dept = UCase(Trim(CStr(wsLog.Cells(lr, 2).value)))     ' B = dept
    recip = LCase(Trim(CStr(wsLog.Cells(lr, 3).value)))    ' C = recipient type
    ws.Range("K1").value = dept
    SetRecipientType ws, recip                             ' switch CN layout (CN-aware)

    ws.Range("G11").value = srcInv                         ' keep source link

    If recip = "patient" Then
        ws.Range("C6").value = ""
        ws.Range("F14").value = wsLog.Cells(lr, 7).value   ' PatientName
        ws.Range("G8").value = wsLog.Cells(lr, 20).value   ' MedAid (T)
        ws.Range("G9").value = wsLog.Cells(lr, 21).value   ' MedNo (U)
        ws.Range("G10").value = wsLog.Cells(lr, 22).value  ' MainMember (V)
        ws.Range("G12").value = wsLog.Cells(lr, 23).value  ' Doctor (W)
        ws.Range("G13").value = wsLog.Cells(lr, 24).value  ' BHF (X)
    Else
        drName = CustIDToDrName(CStr(wsLog.Cells(lr, 6).value))  ' F = CustID -> name
        ws.Range("C6").value = drName
        ws.Range("F14").value = wsLog.Cells(lr, 7).value
    End If

    ' pre-load invoice line items (user trims what isn't credited)
    ws.Range("A16:A30").ClearContents
    ws.Range("D16:D30").ClearContents
    ws.Range("G16:G30").ClearContents
    dest = 16
    last = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsLines.Cells(i, "A").value)) = NrmID(srcInv) Then
            If dest > 30 Then Exit For
            ws.Range("A" & dest).value = wsLines.Cells(i, 6).value   ' Qty
            ws.Range("D" & dest).value = wsLines.Cells(i, 5).value   ' Description
            ws.Range("G" & dest).value = wsLines.Cells(i, 9).value   ' Price Incl
            dest = dest + 1
        End If
    Next i

    ws.Activate

Clean:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "LoadCreditFromInvoice error: " & Err.Description, vbExclamation
End Sub

' Apply a credit-note amount against its source invoice (called on NEW CN only)
' FIX: the portion of cnTotal that exceeds the invoice's outstanding balance
'      ("excess") used to be silently dropped (only a free-text audit note),
'      which destroyed real money. It is now routed to the customer's credit
'      store (doctor -> Customers!L via AddDoctorCredit; patient -> the
'      PatientCredits store via AddPatientCredit) so it is never lost and can
'      be surfaced on a future statement.
Public Sub ApplyCreditToInvoice(cnNo As String, srcInv As String, cnTotal As Double, payDate As Variant)
    Dim wsLog As Worksheet, wsPay As Worksheet, lr As Long
    Dim paid As Double, total As Double, bal As Double, applied As Double, excess As Double
    Dim payID As String, r As Long, recip As String, custID As String, patientName As String
    If cnTotal <= 0.005 Then Exit Sub

    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    Set wsPay = ThisWorkbook.Sheets("Payments")
    lr = InvoiceRow(srcInv)
    If lr = 0 Then Exit Sub

    bal = Num(wsLog.Cells(lr, 15).value): If bal < 0 Then bal = 0
    applied = cnTotal
    If applied > bal Then applied = bal
    excess = cnTotal - applied

    recip = LCase$(Trim$(CStr(wsLog.Cells(lr, 3).value)))
    custID = CStr(wsLog.Cells(lr, 6).value)
    patientName = CStr(wsLog.Cells(lr, 7).value)

    If applied <= 0.005 Then
        LogAudit "CreditNote", srcInv, "", "Credit " & Format(cnTotal, "0.00"), _
                 "CN " & cnNo & " exceeds balance; invoice already settled - full amount -> credit store"
    Else
        payID = NextPaymentIDp()
        r = wsPay.Cells(wsPay.Rows.Count, "A").End(xlUp).row + 1
        wsPay.Cells(r, 1).value = payID
        wsPay.Cells(r, 2).value = srcInv
        wsPay.Cells(r, 3).value = payDate
        wsPay.Cells(r, 4).value = applied
        wsPay.Cells(r, 5).value = "Credit Note"
        wsPay.Cells(r, 6).value = cnNo
        wsPay.Cells(r, 7).value = "Auto credit from " & cnNo

        paid = Num(wsLog.Cells(lr, 14).value) + applied
        total = Num(wsLog.Cells(lr, 12).value)
        bal = total - paid: If bal < 0 Then bal = 0
        wsLog.Cells(lr, 14).value = paid
        wsLog.Cells(lr, 15).value = bal
        If bal <= 0.005 Then
            wsLog.Cells(lr, 16).value = "Paid"
        ElseIf paid > 0 Then
            wsLog.Cells(lr, 16).value = "Part-Paid"
        Else
            wsLog.Cells(lr, 16).value = "Unpaid"
        End If
        wsLog.Cells(lr, 19).value = Now

        LogAudit "CreditNote", srcInv, "Bal " & Format(bal + applied, "0.00"), _
                 "Bal " & Format(bal, "0.00"), "CN " & cnNo & " applied (" & Format(applied, "0.00") & ")"
    End If

    If excess > 0.005 Then
        If recip = "doctor" Then
            AddDoctorCredit custID, excess
            LogAudit "CreditNote", srcInv, "", "Credit " & Format(excess, "0.00"), _
                     "CN " & cnNo & " excess over balance -> doctor credit store"
        Else
            AddPatientCredit patientName, excess
            LogAudit "CreditNote", srcInv, "", "Credit " & Format(excess, "0.00"), _
                     "CN " & cnNo & " excess over balance -> patient credit store"
        End If
    End If
End Sub

' Permanently delete a Credit Note and every entry it created:
'   - the Payments row(s) it posted against its source invoice (kind
'     "Credit Note", Ref = cnNo), then recompute that invoice's
'     Paid/Balance/Status from its remaining Payments rows (ReconcileDoc),
'   - any excess that spilled over into the doctor/patient credit store
'     (recomputed the same way ApplyCreditToInvoice split it: excess =
'     CN total - the amount actually posted as a Payments row),
'   - its own CreditNoteLines rows and CreditNotes log row.
' Requires typed confirmation, same pattern as modSearchVoid.VoidDocument.
Public Function VoidCreditNote(ByVal cnNo As String) As Boolean
    Dim wsLog As Worksheet, wsLines As Worksheet, wsPay As Worksheet, wsInvLog As Worksheet
    Dim lr As Long, invLr As Long, i As Long, last As Long
    Dim srcInv As String, cnTotal As Double, appliedSum As Double, excess As Double
    Dim recip As String, custID As String, patientName As String, typed As String

    VoidCreditNote = False
    If Trim(cnNo) = "" Then Exit Function

    Set wsLog = ThisWorkbook.Sheets("CreditNotes")
    Set wsLines = ThisWorkbook.Sheets("CreditNoteLines")
    Set wsPay = ThisWorkbook.Sheets("Payments")
    Set wsInvLog = ThisWorkbook.Sheets("InvoiceLog")

    lr = FindLogRow(wsLog, cnNo)
    If lr = 0 Then
        MsgBox cnNo & " was not found in CreditNotes.", vbExclamation
        Exit Function
    End If

    srcInv = Trim(CStr(wsLog.Cells(lr, 2).value))
    cnTotal = Num(wsLog.Cells(lr, 5).value)

    If MsgBox("Permanently DELETE credit note " & cnNo & " and all its associated entries?" & vbCrLf & _
              "This reverses its effect on source invoice " & srcInv & " (Payments row + balance/status)," & vbCrLf & _
              "reverses any amount it added to a doctor/patient credit store, and removes it from" & vbCrLf & _
              "the Credit Notes log and its line items. This cannot be undone.", _
              vbCritical + vbYesNo, "Confirm Delete Credit Note") = vbNo Then Exit Function

    typed = Trim(InputBox("To confirm, type the credit note number exactly:" & vbCrLf & cnNo, "Confirm Delete"))
    If typed = "" Then Exit Function
    If LCase(typed) <> LCase(cnNo) Then
        MsgBox "Typed value did not match. Delete cancelled.", vbExclamation
        Exit Function
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- reverse the Payments row(s) this CN posted against its source invoice ----
    appliedSum = 0
    last = wsPay.Cells(wsPay.Rows.Count, "A").End(xlUp).row
    For i = last To 2 Step -1
        If NrmID(CStr(wsPay.Cells(i, PY_REF).value)) = NrmID(cnNo) _
           And LCase$(Trim$(CStr(wsPay.Cells(i, PY_METHOD).value))) = "credit note" Then
            appliedSum = appliedSum + Num(wsPay.Cells(i, PY_AMT).value)
            wsPay.Rows(i).Delete
        End If
    Next i

    invLr = InvoiceRow(srcInv)
    If invLr > 0 Then
        ReconcileDoc "invoice", srcInv          ' recompute Paid/Balance/Status from what's left

        recip = LCase$(Trim$(CStr(wsInvLog.Cells(invLr, 3).value)))
        custID = CStr(wsInvLog.Cells(invLr, 6).value)
        patientName = CStr(wsInvLog.Cells(invLr, 7).value)

        ' ---- reverse any excess that spilled into a credit store ----
        excess = cnTotal - appliedSum
        If excess > 0.005 Then
            If recip = "doctor" Then
                AddDoctorCredit custID, -excess
                LogAudit "CreditNote", srcInv, "", "Credit -" & Format(excess, "0.00"), _
                         "CN " & cnNo & " deleted - reversed excess from doctor credit store"
            Else
                AddPatientCredit patientName, -excess
                LogAudit "CreditNote", srcInv, "", "Credit -" & Format(excess, "0.00"), _
                         "CN " & cnNo & " deleted - reversed excess from patient credit store"
            End If
        End If
    Else
        ' Source invoice no longer exists (voided separately) - the invoice
        ' side has nothing left to reverse. Warn if this CN had spilled an
        ' excess into a credit store, since we cannot tell which store
        ' (doctor/patient) without the invoice's recipient/name - flag for
        ' manual review rather than guessing.
        excess = cnTotal - appliedSum
        If excess > 0.005 Then
            LogAudit "CreditNote", cnNo, "", "", _
                     "Source invoice " & srcInv & " no longer exists; could not auto-reverse " & _
                     Format(excess, "0.00") & " credit-store amount from CN " & cnNo & " - check manually"
        End If
    End If

    LogAudit "Void", cnNo, "Amount " & Format(cnTotal, "0.00"), "DELETED", "Credit note voided (hard delete)"

    ' ---- remove the CN's own line items and log row ----
    last = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = last To 2 Step -1
        If NrmID(CStr(wsLines.Cells(i, 1).value)) = NrmID(cnNo) Then wsLines.Rows(i).Delete
    Next i
    wsLog.Rows(lr).Delete

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    On Error Resume Next
    RefreshMenuSummary
    On Error GoTo 0

    MsgBox cnNo & " has been permanently deleted." & vbCrLf & _
           IIf(invLr > 0, "Invoice " & srcInv & "'s balance/status has been reversed.", _
               "Source invoice " & srcInv & " no longer exists - nothing to reverse there."), vbInformation
    VoidCreditNote = True
    Exit Function
Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "VoidCreditNote error: " & Err.Description, vbExclamation
End Function

Private Function NextPaymentIDp() As String
    Dim wsSet As Worksheet, n As Long
    Set wsSet = ThisWorkbook.Sheets("Settings")
    n = CLng(wsSet.Range("B18").value)
    NextPaymentIDp = "PAY-" & Format(n, "0000")
    wsSet.Range("B18").value = n + 1
End Function

' Flow C: open frmSearch as an invoice picker for a new Credit Note
Public Sub PickSourceInvoice()
    frmSearch.ShowPickInvoiceForCN
End Sub

