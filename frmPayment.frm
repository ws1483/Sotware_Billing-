VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPayment 
   Caption         =   "frmPayment"
   ClientHeight    =   8448.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10524
   OleObjectBlob   =   "frmPayment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ============================================================================
' frmPayment — Record Payment
'   PATCH: locale-safe numeric reads via Num() — prevents cent-truncation in
'          balances (which caused wrong FIFO allocation & R1 remainder rows).
' ============================================================================

Private Function NrmID(s As String) As String
    NrmID = UCase(Replace(Trim(s), " ", ""))
End Function

' Locale-safe numeric read (handles comma decimals; blank/error -> 0)
Private Function Num(v As Variant) As Double
    If IsError(v) Then Num = 0: Exit Function
    If Trim(CStr(v)) = "" Then Num = 0: Exit Function
    If IsNumeric(v) Then Num = CDbl(v) Else Num = 0
End Function

Private Sub Label5_Click()

End Sub

' ============================ FORM EVENTS ===================================
Private Sub UserForm_Initialize()
    txtDate.Value = Format(Date, "yyyy-mm-dd")
    LoadMethods
    optInvoice.Value = True
    lblTarget.Caption = "Invoice No:"
    PopulateTarget
End Sub

Private Sub optInvoice_Click()
    lblTarget.Caption = "Invoice No:"
    PopulateTarget
End Sub

Private Sub optDoctor_Click()
    lblTarget.Caption = "Doctor:"
    PopulateTarget
End Sub

Private Sub optPrivate_Click()
    lblTarget.Caption = "Patient:"
    PopulateTarget
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub LoadMethods()
    Dim wsR As Worksheet, last As Long, i As Long
    On Error GoTo EH
    Set wsR = ThisWorkbook.Sheets("Reference")
    cboMethod.Clear
    last = wsR.Cells(wsR.Rows.Count, "D").End(xlUp).row
    For i = 2 To last
        If Trim(wsR.Cells(i, "D").Value) <> "" Then cboMethod.AddItem wsR.Cells(i, "D").Value
    Next i
    If cboMethod.ListCount > 0 Then cboMethod.ListIndex = 0
    Exit Sub
EH:
    MsgBox "LoadMethods error: " & Err.Description, vbExclamation
End Sub

Private Sub PopulateTarget()
    Dim ws As Worksheet, last As Long, i As Long, nm As String
    Dim seen As Object
    On Error GoTo EH
    Set seen = CreateObject("Scripting.Dictionary")
    cboTarget.Clear
    lblInfo.Caption = ""

    If optInvoice.Value Then
        Set ws = ThisWorkbook.Sheets("InvoiceLog")
        last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        For i = 2 To last
            If UCase(CStr(ws.Cells(i, 16).Value)) <> "PAID" _
               And Trim(CStr(ws.Cells(i, 1).Value)) <> "" Then
                cboTarget.AddItem ws.Cells(i, 1).Value
            End If
        Next i

    ElseIf optDoctor.Value Then
        Set ws = ThisWorkbook.Sheets("Customers")
        last = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
        For i = 2 To last
            If Trim(ws.Cells(i, "B").Value) <> "" Then cboTarget.AddItem ws.Cells(i, "B").Value
        Next i

    Else
        Set ws = ThisWorkbook.Sheets("InvoiceLog")
        last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        For i = 2 To last
            If UCase(CStr(ws.Cells(i, 3).Value)) = "PATIENT" Then
                nm = Trim(CStr(ws.Cells(i, 7).Value))
                If nm <> "" And Not seen.Exists(UCase(nm)) Then
                    seen.Add UCase(nm), 1
                    cboTarget.AddItem nm
                End If
            End If
        Next i
    End If
    Exit Sub
EH:
    MsgBox "PopulateTarget error: " & Err.Description, vbExclamation
End Sub

Private Sub cboTarget_Change()
    Dim wsLog As Worksheet, lr As Long, cid As String
    On Error Resume Next
    If cboTarget.ListIndex < 0 Then Exit Sub
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    If optInvoice.Value Then
        lr = FindLogRow(wsLog, cboTarget.Value)
        If lr > 0 Then lblInfo.Caption = "Invoice balance: R " & _
            Format(Num(wsLog.Cells(lr, 15).Value), "#,##0.00")
    ElseIf optDoctor.Value Then
        cid = DrNameToCustID(cboTarget.Value)
        lblInfo.Caption = "Doctor credit: R " & Format(DoctorCredit(cid), "#,##0.00") & _
                          "    Outstanding: R " & Format(SumOutstandingByCust(cid), "#,##0.00")
    Else
        lblInfo.Caption = "Patient outstanding: R " & _
            Format(SumOutstandingByPatient(cboTarget.Value), "#,##0.00")
    End If
End Sub

' ================================ RECORD ====================================
Private Sub btnRecord_Click()
    Dim amt As Double, payDate As Variant, method As String
    Dim ref As String, notes As String, payID As String

    If cboTarget.ListIndex < 0 Then MsgBox "Select an invoice / doctor / patient.", vbExclamation: Exit Sub
    If Not IsNumeric(txtAmount.Value) Then MsgBox "Enter a valid amount.", vbExclamation: Exit Sub
    If Num(txtAmount.Value) <= 0 Then MsgBox "Amount must be greater than zero.", vbExclamation: Exit Sub
    If Not IsDate(txtDate.Value) Then MsgBox "Enter a valid date (yyyy-mm-dd).", vbExclamation: Exit Sub
    If Trim(cboMethod.Value) = "" Then MsgBox "Select a payment method.", vbExclamation: Exit Sub

    amt = Num(txtAmount.Value)
    payDate = CDate(txtDate.Value)
    method = cboMethod.Value
    ref = txtReference.Value
    notes = txtNotes.Value

    On Error GoTo Fail
    Application.ScreenUpdating = False
    payID = NextPaymentID()

    If optInvoice.Value Then
        PayOneInvoice cboTarget.Value, amt, payDate, method, ref, notes, payID
    ElseIf optDoctor.Value Then
        PaySweep "doctor", cboTarget.Value, amt, payDate, method, ref, notes, payID
    Else
        PaySweep "patient", cboTarget.Value, amt, payDate, method, ref, notes, payID
    End If

    Application.ScreenUpdating = True
    MsgBox "Payment " & payID & " recorded.", vbInformation
    Unload Me
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "RecordPayment error: " & Err.Description, vbExclamation
End Sub

Private Sub PayOneInvoice(invNo As String, amt As Double, payDate As Variant, _
                          method As String, ref As String, notes As String, payID As String)
    Dim wsLog As Worksheet, lr As Long, bal As Double
    Dim applied As Double, over As Double, cid As String
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    lr = FindLogRow(wsLog, invNo)
    If lr = 0 Then Err.Raise 513, , "Invoice " & invNo & " not found."

    cid = CStr(wsLog.Cells(lr, 6).Value)
    bal = Num(wsLog.Cells(lr, 15).Value): If bal < 0 Then bal = 0
    applied = amt: over = 0
    If amt > bal Then applied = bal: over = amt - bal

    WritePaymentRow payID, invNo, payDate, applied, method, ref, notes
    UpdateInvoicePaid wsLog, lr, applied
    LogAudit "Payment", invNo, "", "Paid " & Format(applied, "0.00"), "Pay " & payID

    If over > 0 Then HandleOverpay wsLog, lr, cid, over, payID
End Sub

Private Sub PaySweep(kind As String, targetVal As String, amt As Double, payDate As Variant, _
                     method As String, ref As String, notes As String, payID As String)
    Dim wsLog As Worksheet, last As Long, i As Long, j As Long, cid As String
    Dim rowsArr() As Long, dts() As Double, cnt As Long
    Dim tL As Long, tD As Double, remaining As Double, bal As Double, applied As Double
    Dim isMatch As Boolean
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row

    If kind = "doctor" Then cid = DrNameToCustID(targetVal)

    ReDim rowsArr(1 To last): ReDim dts(1 To last): cnt = 0
    For i = 2 To last
        isMatch = False
        If kind = "doctor" Then
            If NrmID(CStr(wsLog.Cells(i, 6).Value)) = NrmID(cid) _
               And UCase(CStr(wsLog.Cells(i, 3).Value)) = "DOCTOR" Then isMatch = True
        Else
            If UCase(Trim(CStr(wsLog.Cells(i, 7).Value))) = UCase(Trim(targetVal)) _
               And UCase(CStr(wsLog.Cells(i, 3).Value)) = "PATIENT" Then isMatch = True
        End If
        If isMatch And UCase(CStr(wsLog.Cells(i, 16).Value)) <> "PAID" Then
            cnt = cnt + 1
            rowsArr(cnt) = i
            dts(cnt) = CDbl(wsLog.Cells(i, 4).Value)
        End If
    Next i

    If cnt = 0 Then Err.Raise 513, , "No unpaid invoices found for " & targetVal

    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If dts(j) > dts(j + 1) Then
                tD = dts(j): dts(j) = dts(j + 1): dts(j + 1) = tD
                tL = rowsArr(j): rowsArr(j) = rowsArr(j + 1): rowsArr(j + 1) = tL
            End If
        Next j
    Next i

    remaining = amt
    For i = 1 To cnt
        If remaining <= 0.005 Then Exit For
        bal = Num(wsLog.Cells(rowsArr(i), 15).Value)
        If bal > 0.005 Then
            applied = IIf(remaining >= bal, bal, remaining)
            WritePaymentRow payID, CStr(wsLog.Cells(rowsArr(i), 1).Value), payDate, applied, method, ref, notes
            UpdateInvoicePaid wsLog, rowsArr(i), applied
            LogAudit "Payment", CStr(wsLog.Cells(rowsArr(i), 1).Value), "", _
                     "Paid " & Format(applied, "0.00"), "Sweep " & payID
            remaining = remaining - applied
        End If
    Next i

    If remaining > 0.005 Then
        If kind = "doctor" Then
            AddDoctorCredit cid, remaining
            LogAudit "Credit", cid, "", "Credit " & Format(remaining, "0.00"), "Overpay " & payID
            MsgBox "Overpayment of R " & Format(remaining, "#,##0.00") & _
                   " added to doctor credit.", vbInformation
        Else
            MsgBox "Overpayment of R " & Format(remaining, "#,##0.00") & _
                   " for private patient '" & targetVal & "'." & vbCrLf & _
                   "No credit store for private patients — please refund manually.", vbExclamation
            LogAudit "Overpay", targetVal, "", "Excess " & Format(remaining, "0.00"), _
                     "Private overpay - manual refund " & payID
        End If
    End If
End Sub

Private Sub HandleOverpay(wsLog As Worksheet, lr As Long, cid As String, over As Double, payID As String)
    If UCase(CStr(wsLog.Cells(lr, 3).Value)) = "DOCTOR" Then
        AddDoctorCredit cid, over
        LogAudit "Credit", CStr(wsLog.Cells(lr, 1).Value), "", _
                 "Credit " & Format(over, "0.00"), "Overpay " & payID
        MsgBox "Overpayment of R " & Format(over, "#,##0.00") & " added to doctor credit.", vbInformation
    Else
        MsgBox "Overpayment of R " & Format(over, "#,##0.00") & " on a private-patient invoice." & vbCrLf & _
               "No credit store for private patients — please refund manually.", vbExclamation
        LogAudit "Overpay", CStr(wsLog.Cells(lr, 1).Value), "", _
                 "Excess " & Format(over, "0.00"), "Private overpay " & payID
    End If
End Sub

' ============================ WRITERS / UPDATERS ============================
Private Sub WritePaymentRow(payID As String, invNo As String, payDate As Variant, _
                            amt As Double, method As String, ref As String, notes As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Payments")
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    ws.Cells(r, 1).Value = payID
    ws.Cells(r, 2).Value = invNo
    ws.Cells(r, 3).Value = payDate
    ws.Cells(r, 4).Value = amt
    ws.Cells(r, 5).Value = method
    ws.Cells(r, 6).Value = ref
    ws.Cells(r, 7).Value = notes
End Sub

Private Sub UpdateInvoicePaid(wsLog As Worksheet, lr As Long, applied As Double)
    Dim paid As Double, total As Double, bal As Double
    paid = Num(wsLog.Cells(lr, 14).Value) + applied
    total = Num(wsLog.Cells(lr, 12).Value)
    bal = total - paid: If bal < 0 Then bal = 0
    wsLog.Cells(lr, 14).Value = paid
    wsLog.Cells(lr, 15).Value = bal
    If bal <= 0.005 Then
        wsLog.Cells(lr, 16).Value = "Paid"
    ElseIf paid > 0 Then
        wsLog.Cells(lr, 16).Value = "Part-Paid"
    Else
        wsLog.Cells(lr, 16).Value = "Unpaid"
    End If
    wsLog.Cells(lr, 19).Value = Now
End Sub

' ================================ HELPERS ===================================
Private Function NextPaymentID() As String
    Dim wsSet As Worksheet, n As Long
    Set wsSet = ThisWorkbook.Sheets("Settings")
    n = CLng(wsSet.Range("B18").Value)
    NextPaymentID = "PAY-" & Format(n, "0000")
    wsSet.Range("B18").Value = n + 1
End Function

Private Function FindLogRow(wsLog As Worksheet, invNo As String) As Long
    Dim last As Long, i As Long
    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsLog.Cells(i, 1).Value)) = NrmID(invNo) Then FindLogRow = i: Exit Function
    Next i
End Function

Private Function DrNameToCustID(drName As String) As String
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "B").End(xlUp).row
    For i = 2 To last
        If CStr(wsC.Cells(i, "B").Value) = drName Then
            DrNameToCustID = CStr(wsC.Cells(i, "A").Value): Exit Function
        End If
    Next i
End Function

Private Function DoctorCredit(custID As String) As Double
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsC.Cells(i, "A").Value)) = NrmID(custID) Then
            DoctorCredit = Num(wsC.Cells(i, "L").Value): Exit Function
        End If
    Next i
End Function

Private Sub AddDoctorCredit(custID As String, addAmt As Double)
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsC.Cells(i, "A").Value)) = NrmID(custID) Then
            wsC.Cells(i, "L").Value = Num(wsC.Cells(i, "L").Value) + addAmt
            Exit Sub
        End If
    Next i
End Sub

Private Function SumOutstandingByCust(custID As String) As Double
    Dim ws As Worksheet, last As Long, i As Long, t As Double
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(ws.Cells(i, 6).Value)) = NrmID(custID) _
           And UCase(CStr(ws.Cells(i, 3).Value)) = "DOCTOR" Then
            t = t + Num(ws.Cells(i, 15).Value)
        End If
    Next i
    SumOutstandingByCust = t
End Function

Private Function SumOutstandingByPatient(pName As String) As Double
    Dim ws As Worksheet, last As Long, i As Long, t As Double
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If UCase(Trim(CStr(ws.Cells(i, 7).Value))) = UCase(Trim(pName)) _
           And UCase(CStr(ws.Cells(i, 3).Value)) = "PATIENT" Then
            t = t + Num(ws.Cells(i, 15).Value)
        End If
    Next i
    SumOutstandingByPatient = t
End Function



