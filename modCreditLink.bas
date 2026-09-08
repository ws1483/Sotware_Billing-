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
Public Sub ApplyCreditToInvoice(cnNo As String, srcInv As String, cnTotal As Double, payDate As Variant)
    Dim wsLog As Worksheet, wsPay As Worksheet, lr As Long
    Dim paid As Double, total As Double, bal As Double, applied As Double, payID As String, r As Long
    If cnTotal <= 0.005 Then Exit Sub

    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    Set wsPay = ThisWorkbook.Sheets("Payments")
    lr = InvoiceRow(srcInv)
    If lr = 0 Then Exit Sub

    bal = Num(wsLog.Cells(lr, 15).value): If bal < 0 Then bal = 0
    applied = cnTotal
    If applied > bal Then applied = bal
    If applied <= 0.005 Then
        LogAudit "CreditNote", srcInv, "", "Credit " & Format(cnTotal, "0.00"), _
                 "CN " & cnNo & " exceeds balance; invoice already settled - no balance change"
        Exit Sub
    End If

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
End Sub

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

