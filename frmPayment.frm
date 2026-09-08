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
' frmPayment ? Record Payment
'   DOCS MODE (optDocs): merged, searchable, dept-filtered tick list of
'     Invoices + Med Claims. Pay-in-full: amount = sum of TICKED balances.
'   DOCTOR / PRIVATE SWEEP: editable-amount FIFO, now DEPT-FILTERED via cboDept.
'     cboTarget + cboDept + lstDocs all visible; lstDocs shows a read-only
'     preview of the dept's unpaid invoices; lblInfo shows dept outstanding.
'     Sweep reads InvoiceLog ONLY (med claims are paid in Docs mode).
' ============================================================================
Private mRowType() As String
Private mRowNo() As String
Private mRowBal() As Double
Private mRowCount As Long

Private Sub optInvoice_Click()
End Sub

Private Sub UserForm_Initialize()
    txtDate.value = Format(Date, "yyyy-mm-dd")
    LoadMethods

    cboDept.Clear
    cboDept.AddItem "All": cboDept.AddItem "WA": cboDept.AddItem "WD": cboDept.AddItem "MC"
    cboDept.value = "All"

    lstDocs.ColumnCount = 7
    lstDocs.ColumnHeads = False
    lstDocs.MultiSelect = fmMultiSelectMulti
    lstDocs.ColumnWidths = "80;120;110;65;70;70;60"

    optDocs.value = True
    ShowMode
    PopulateDocs
End Sub

' ---------------- mode switching ----------------
Private Sub optDocs_Click():    ShowMode: PopulateDocs: End Sub
Private Sub optDoctor_Click():  ShowMode: PopulateNames "doctor": End Sub
Private Sub optPrivate_Click(): ShowMode: PopulateNames "patient": End Sub

Private Sub ShowMode()
    Dim docsMode As Boolean, sweepMode As Boolean
    docsMode = optDocs.value
    sweepMode = (optDoctor.value Or optPrivate.value)

    lblTarget.Caption = IIf(docsMode, "Search Invoices / Med Claims:", _
                        IIf(optDoctor.value, "Doctor:", "Patient:"))

    ' Docs mode uses txtSearch; sweep mode does not
    txtSearch.Visible = docsMode
    ' cboDept + lstDocs are visible in BOTH docs and sweep modes
    cboDept.Visible = docsMode Or sweepMode
    lstDocs.Visible = docsMode Or sweepMode
    ' name picker only in sweep mode
    cboTarget.Visible = sweepMode

    ' amount box: locked auto-total in docs mode; editable in sweep
    txtAmount.Locked = docsMode
    txtAmount.BackColor = IIf(docsMode, &H8000000F, &H80000005)
    If docsMode Then txtAmount.value = "0.00"

    lblInfo.Caption = ""
End Sub

Private Sub txtSearch_Change(): If optDocs.value Then PopulateDocs
End Sub

Private Sub cboDept_Change()
    If optDocs.value Then
        PopulateDocs
    ElseIf optDoctor.value Or optPrivate.value Then
        RefreshSweepPreview
    End If
End Sub

' Recompute ticked total whenever selection changes (docs mode only)
Private Sub lstDocs_Change()
    If Not optDocs.value Then Exit Sub
    Dim i As Long, tot As Double, n As Long
    For i = 0 To lstDocs.ListCount - 1
        If lstDocs.Selected(i) Then
            tot = tot + mRowBal(i)
            n = n + 1
        End If
    Next i
    txtAmount.value = Format(tot, "0.00")
    lblInfo.Caption = n & " document(s) ticked ? total R " & Format(tot, "#,##0.00")
End Sub

' ---------------- populate the merged docs list (Docs mode) ----------------
Private Sub PopulateDocs()
    Dim term As String, deptF As String
    term = LCase(Trim(txtSearch.value))
    deptF = UCase(Trim(cboDept.value))

    lstDocs.Clear
    ReDim mRowType(0 To 5000)
    ReDim mRowNo(0 To 5000)
    ReDim mRowBal(0 To 5000)
    mRowCount = 0
    lblInfo.Caption = ""
    txtAmount.value = "0.00"

    If deptF = "ALL" Or deptF = "WA" Or deptF = "WD" Then
        AddDocsFromLog "InvoiceLog", "invoice", term, deptF
    End If
    If deptF = "ALL" Or deptF = "MC" Or deptF = "WA" Or deptF = "WD" Then
        AddDocsFromLog "MedAidLog", "medclaim", term, deptF
    End If
End Sub

Private Sub AddDocsFromLog(logName As String, docType As String, term As String, deptF As String)
    Dim ws As Worksheet, last As Long, i As Long
    Dim cNo As Long, cDept As Long, cRecip As Long, cDt As Long
    Dim cCust As Long, cPat As Long, cTotal As Long, cBal As Long, cStatus As Long
    Dim docNo As String, drBill As String, patient As String, status As String
    Dim total As Variant, bal As Double, dt As Variant, hay As String, dept As String
    Dim li As Long

    Set ws = ThisWorkbook.Sheets(logName)
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If last < 2 Then Exit Sub

    If docType = "medclaim" Then
        cNo = ML_NO: cDept = ML_DEPT: cRecip = ML_RECIP: cDt = ML_DATE
        cCust = ML_CUST: cPat = ML_PATIENT: cTotal = ML_TOTAL: cBal = ML_BALANCE: cStatus = ML_STATUS
    Else
        cNo = IL_NO: cDept = IL_DEPT: cRecip = IL_RECIP: cDt = IL_DATE
        cCust = IL_CUST: cPat = IL_PATIENT: cTotal = IL_TOTAL: cBal = IL_BALANCE: cStatus = IL_STATUS
    End If

    For i = 2 To last
        docNo = Trim(CStr(ws.Cells(i, cNo).value))
        If docNo = "" Then GoTo NextI

        status = Trim(CStr(ws.Cells(i, cStatus).value))
        If UCase(status) = "PAID" Then GoTo NextI

        dept = UCase(Trim(CStr(ws.Cells(i, cDept).value)))
        If deptF = "WA" Or deptF = "WD" Then
            If dept <> deptF Then GoTo NextI
        ElseIf deptF = "MC" Then
            If docType <> "medclaim" Then GoTo NextI
        End If

        If docType = "medclaim" Then
            drBill = Trim(CStr(ws.Cells(i, cCust).value))
        Else
            If UCase(CStr(ws.Cells(i, cRecip).value)) = "PATIENT" Then
                drBill = "(patient)"
            Else
                drBill = CustIDToDrName(CStr(ws.Cells(i, cCust).value))
            End If
        End If
        patient = Trim(CStr(ws.Cells(i, cPat).value))

        If term <> "" Then
            hay = LCase(docNo & " " & drBill & " " & patient)
            If InStr(hay, term) = 0 Then GoTo NextI
        End If

        total = ws.Cells(i, cTotal).value
        bal = Num(ws.Cells(i, cBal).value)
        dt = ws.Cells(i, cDt).value

        li = lstDocs.ListCount
        lstDocs.AddItem docNo
        lstDocs.List(li, 1) = drBill
        lstDocs.List(li, 2) = patient
        lstDocs.List(li, 3) = IIf(IsDate(dt), Format(dt, "yyyy-mm-dd"), "")
        lstDocs.List(li, 4) = IIf(IsNumeric(total), Format(total, "#,##0.00"), "")
        lstDocs.List(li, 5) = Format(bal, "#,##0.00")
        lstDocs.List(li, 6) = status

        mRowType(li) = docType
        mRowNo(li) = docNo
        mRowBal(li) = bal
        mRowCount = li + 1
NextI:
    Next i
End Sub

' ---------------- name selectors (doctor / private) ----------------
Private Sub PopulateNames(kind As String)
    Dim ws As Worksheet, last As Long, i As Long, nm As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    cboTarget.Clear
    lblInfo.Caption = ""
    txtAmount.value = ""
    lstDocs.Clear

    If kind = "doctor" Then
        Set ws = ThisWorkbook.Sheets("Customers")
        last = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
        For i = 2 To last
            If Trim(ws.Cells(i, "B").value) <> "" Then cboTarget.AddItem ws.Cells(i, "B").value
        Next i
    Else
        Set ws = ThisWorkbook.Sheets("InvoiceLog")
        last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        For i = 2 To last
            If UCase(CStr(ws.Cells(i, IL_RECIP).value)) = "PATIENT" Then
                nm = Trim(CStr(ws.Cells(i, IL_PATIENT).value))
                If nm <> "" And Not seen.Exists(UCase(nm)) Then
                    seen.Add UCase(nm), 1
                    cboTarget.AddItem nm
                End If
            End If
        Next i
    End If
End Sub

Private Sub cboTarget_Change()
    If cboTarget.ListIndex < 0 Then Exit Sub
    RefreshSweepPreview
End Sub

' Fill lstDocs with the selected doctor/patient's unpaid invoices for cboDept,
' and show the dept-filtered outstanding total in lblInfo. (Sweep = InvoiceLog only.)
Private Sub RefreshSweepPreview()
    Dim ws As Worksheet, last As Long, i As Long
    Dim kind As String, targetVal As String, cid As String, deptF As String
    Dim isMatch As Boolean, dept As String, docNo As String
    Dim drBill As String, patient As String, status As String
    Dim total As Variant, bal As Double, dt As Variant, li As Long, tot As Double

    If Not (optDoctor.value Or optPrivate.value) Then Exit Sub
    If cboTarget.ListIndex < 0 Then Exit Sub

    kind = IIf(optDoctor.value, "doctor", "patient")
    targetVal = cboTarget.value
    deptF = UCase(Trim(cboDept.value))
    If kind = "doctor" Then cid = DrNameToCustID(targetVal)

    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    lstDocs.Clear
    ReDim mRowType(0 To 5000)
    ReDim mRowNo(0 To 5000)
    ReDim mRowBal(0 To 5000)
    mRowCount = 0
    tot = 0

    For i = 2 To last
        docNo = Trim(CStr(ws.Cells(i, IL_NO).value))
        If docNo = "" Then GoTo NextI

        isMatch = False
        If kind = "doctor" Then
            If NrmID(CStr(ws.Cells(i, IL_CUST).value)) = NrmID(cid) _
               And UCase(CStr(ws.Cells(i, IL_RECIP).value)) = "DOCTOR" Then isMatch = True
        Else
            If UCase(Trim(CStr(ws.Cells(i, IL_PATIENT).value))) = UCase(Trim(targetVal)) _
               And UCase(CStr(ws.Cells(i, IL_RECIP).value)) = "PATIENT" Then isMatch = True
        End If
        If Not isMatch Then GoTo NextI

        ' dept filter (WA/WD). All = no dept limit. MC = no InvoiceLog match.
        dept = UCase(Trim(CStr(ws.Cells(i, IL_DEPT).value)))
        If deptF = "WA" Or deptF = "WD" Then
            If dept <> deptF Then GoTo NextI
        ElseIf deptF = "MC" Then
            GoTo NextI          ' sweep is InvoiceLog only
        End If

        status = Trim(CStr(ws.Cells(i, IL_STATUS).value))
        If UCase(status) = "PAID" Then GoTo NextI

        bal = Num(ws.Cells(i, IL_BALANCE).value)
        If bal <= 0.005 Then GoTo NextI

        drBill = IIf(kind = "doctor", CustIDToDrName(CStr(ws.Cells(i, IL_CUST).value)), "(patient)")
        patient = Trim(CStr(ws.Cells(i, IL_PATIENT).value))
        total = ws.Cells(i, IL_TOTAL).value
        dt = ws.Cells(i, IL_DATE).value

        li = lstDocs.ListCount
        lstDocs.AddItem docNo
        lstDocs.List(li, 1) = drBill
        lstDocs.List(li, 2) = patient
        lstDocs.List(li, 3) = IIf(IsDate(dt), Format(dt, "yyyy-mm-dd"), "")
        lstDocs.List(li, 4) = IIf(IsNumeric(total), Format(total, "#,##0.00"), "")
        lstDocs.List(li, 5) = Format(bal, "#,##0.00")
        lstDocs.List(li, 6) = status

        mRowType(li) = "invoice"
        mRowNo(li) = docNo
        mRowBal(li) = bal
        mRowCount = li + 1
        tot = tot + bal
NextI:
    Next i

    Dim credTxt As String
    If kind = "doctor" Then
        credTxt = "Doctor credit: R " & Format(DoctorCredit(cid), "#,##0.00") & "    "
    End If
    lblInfo.Caption = credTxt & "Outstanding (" & _
        IIf(deptF = "ALL", "All depts", deptF) & "): R " & Format(tot, "#,##0.00")
End Sub

Private Sub btnCancel_Click(): Unload Me
End Sub

Private Sub LoadMethods()
    Dim wsR As Worksheet, last As Long, i As Long
    On Error GoTo EH
    Set wsR = ThisWorkbook.Sheets("Reference")
    cboMethod.Clear
    last = wsR.Cells(wsR.Rows.Count, "D").End(xlUp).row
    For i = 2 To last
        If Trim(wsR.Cells(i, "D").value) <> "" Then cboMethod.AddItem wsR.Cells(i, "D").value
    Next i
    If cboMethod.ListCount > 0 Then cboMethod.ListIndex = 0
    Exit Sub
EH:
    MsgBox "LoadMethods error: " & Err.Description, vbExclamation
End Sub

' ================================ RECORD ====================================
Private Sub btnRecord_Click()
    Dim payDate As Variant, method As String, ref As String, notes As String, payID As String

    If Not IsDate(txtDate.value) Then MsgBox "Enter a valid date (yyyy-mm-dd).", vbExclamation: Exit Sub
    If Trim(cboMethod.value) = "" Then MsgBox "Select a payment method.", vbExclamation: Exit Sub
    payDate = CDate(txtDate.value)
    method = cboMethod.value
    ref = txtReference.value
    notes = txtNotes.value

    On Error GoTo Fail
    Application.ScreenUpdating = False

    If optDocs.value Then
        RecordDocsBatch payDate, method, ref, notes
    Else
        Dim amt As Double
        If Not IsNumeric(txtAmount.value) Then MsgBox "Enter a valid amount.", vbExclamation: GoTo CleanExit
        amt = Num(txtAmount.value)
        If amt <= 0 Then MsgBox "Amount must be greater than zero.", vbExclamation: GoTo CleanExit
        If cboTarget.ListIndex < 0 Then MsgBox "Select a doctor / patient.", vbExclamation: GoTo CleanExit
        payID = NextPaymentID()
        If optDoctor.value Then
            PaySweep "doctor", cboTarget.value, amt, payDate, method, ref, notes, payID, UCase(Trim(cboDept.value))
        Else
            PaySweep "patient", cboTarget.value, amt, payDate, method, ref, notes, payID, UCase(Trim(cboDept.value))
        End If
        MsgBox "Payment " & payID & " recorded.", vbInformation
    End If

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "RecordPayment error: " & Err.Description, vbExclamation
End Sub

' Docs batch: pay each TICKED doc in full; one shared PayID; mixed allowed
Private Sub RecordDocsBatch(payDate As Variant, method As String, ref As String, notes As String)
    Dim i As Long, n As Long, payID As String, tot As Double, docNo As String
    Dim dType As String, bal As Double, summary As String

    For i = 0 To lstDocs.ListCount - 1
        If lstDocs.Selected(i) Then n = n + 1: tot = tot + mRowBal(i)
    Next i
    If n = 0 Then MsgBox "Tick at least one document to pay.", vbExclamation: Exit Sub

    If MsgBox("Record payment of R " & Format(tot, "#,##0.00") & _
              " across " & n & " document(s), paying each in full?", _
              vbQuestion + vbYesNo, "Confirm Payment") = vbNo Then Exit Sub

    payID = NextPaymentID()
    For i = 0 To lstDocs.ListCount - 1
        If lstDocs.Selected(i) Then
            docNo = mRowNo(i)
            dType = mRowType(i)
            bal = mRowBal(i)
            If bal > 0.005 Then
                WritePaymentRow payID, docNo, payDate, bal, method, ref, notes
                ReconcileDoc dType, docNo
                LogAudit "Payment", docNo, "", "Paid " & Format(bal, "0.00"), "Batch " & payID
                summary = summary & "  ? " & docNo & "  R " & Format(bal, "#,##0.00") & vbCrLf
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "Payment " & payID & " recorded (R " & Format(tot, "#,##0.00") & "):" & vbCrLf & summary, vbInformation
    Unload Me
End Sub

' Doctor / Private FIFO sweep ? now DEPT-FILTERED (InvoiceLog only)
Private Sub PaySweep(kind As String, targetVal As String, amt As Double, payDate As Variant, _
                     method As String, ref As String, notes As String, payID As String, _
                     Optional deptFilt As String = "ALL")
    Dim wsLog As Worksheet, last As Long, i As Long, j As Long, cid As String
    Dim rowsArr() As Long, dts() As Double, cnt As Long
    Dim tL As Long, tD As Double, remaining As Double, bal As Double, applied As Double
    Dim isMatch As Boolean, dept As String
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
    If kind = "doctor" Then cid = DrNameToCustID(targetVal)

    ReDim rowsArr(1 To last): ReDim dts(1 To last): cnt = 0
    For i = 2 To last
        isMatch = False
        If kind = "doctor" Then
            If NrmID(CStr(wsLog.Cells(i, IL_CUST).value)) = NrmID(cid) _
               And UCase(CStr(wsLog.Cells(i, IL_RECIP).value)) = "DOCTOR" Then isMatch = True
        Else
            If UCase(Trim(CStr(wsLog.Cells(i, IL_PATIENT).value))) = UCase(Trim(targetVal)) _
               And UCase(CStr(wsLog.Cells(i, IL_RECIP).value)) = "PATIENT" Then isMatch = True
        End If

        ' dept filter
        If isMatch And (deptFilt = "WA" Or deptFilt = "WD") Then
            dept = UCase(Trim(CStr(wsLog.Cells(i, IL_DEPT).value)))
            If dept <> deptFilt Then isMatch = False
        ElseIf isMatch And deptFilt = "MC" Then
            isMatch = False          ' sweep is InvoiceLog only
        End If

        If isMatch And UCase(CStr(wsLog.Cells(i, IL_STATUS).value)) <> "PAID" Then
            cnt = cnt + 1: rowsArr(cnt) = i: dts(cnt) = CDbl(wsLog.Cells(i, IL_DATE).value)
        End If
    Next i
    If cnt = 0 Then Err.Raise 513, , "No unpaid invoices found for " & targetVal & _
        IIf(deptFilt = "WA" Or deptFilt = "WD", " (Dept " & deptFilt & ")", "") & "."

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
        bal = Num(wsLog.Cells(rowsArr(i), IL_BALANCE).value)
        If bal > 0.005 Then
            applied = IIf(remaining >= bal, bal, remaining)
            WritePaymentRow payID, CStr(wsLog.Cells(rowsArr(i), 1).value), payDate, applied, method, ref, notes
            ReconcileDoc "invoice", CStr(wsLog.Cells(rowsArr(i), 1).value)
            LogAudit "Payment", CStr(wsLog.Cells(rowsArr(i), 1).value), "", _
                     "Paid " & Format(applied, "0.00"), "Sweep " & payID
            remaining = remaining - applied
        End If
    Next i

    If remaining > 0.005 Then
        If kind = "doctor" Then
            AddDoctorCredit cid, remaining
            LogAudit "Credit", cid, "", "Credit " & Format(remaining, "0.00"), "Overpay " & payID
            MsgBox "Overpayment of R " & Format(remaining, "#,##0.00") & " added to doctor credit.", vbInformation
        Else
            MsgBox "Overpayment of R " & Format(remaining, "#,##0.00") & " for '" & targetVal & "'." & vbCrLf & _
                   "No credit store for private patients ? refund manually.", vbExclamation
            LogAudit "Overpay", targetVal, "", "Excess " & Format(remaining, "0.00"), "Private overpay " & payID
        End If
    End If
End Sub

' ============================ WRITERS / HELPERS ============================
Private Sub WritePaymentRow(payID As String, docNo As String, payDate As Variant, _
                            amt As Double, method As String, ref As String, notes As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Payments")
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    ws.Cells(r, PY_ID).value = payID
    ws.Cells(r, PY_INV).value = docNo
    ws.Cells(r, PY_DATE).value = payDate
    ws.Cells(r, PY_AMT).value = amt
    ws.Cells(r, PY_METHOD).value = method
    ws.Cells(r, PY_REF).value = ref
    ws.Cells(r, PY_NOTES).value = notes
End Sub

Private Function NextPaymentID() As String
    Dim wsSet As Worksheet, n As Long
    Set wsSet = ThisWorkbook.Sheets("Settings")
    n = CLng(wsSet.Range("B18").value)
    NextPaymentID = "PAY-" & Format(n, "0000")
    wsSet.Range("B18").value = n + 1
End Function

Private Function SumOutstandingByCust(custID As String) As Double
    Dim ws As Worksheet, last As Long, i As Long, t As Double
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(ws.Cells(i, IL_CUST).value)) = NrmID(custID) _
           And UCase(CStr(ws.Cells(i, IL_RECIP).value)) = "DOCTOR" Then
            t = t + Num(ws.Cells(i, IL_BALANCE).value)
        End If
    Next i
    SumOutstandingByCust = t
End Function

Private Function SumOutstandingByPatient(pName As String) As Double
    Dim ws As Worksheet, last As Long, i As Long, t As Double
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If UCase(Trim(CStr(ws.Cells(i, IL_PATIENT).value))) = UCase(Trim(pName)) _
           And UCase(CStr(ws.Cells(i, IL_RECIP).value)) = "PATIENT" Then
            t = t + Num(ws.Cells(i, IL_BALANCE).value)
        End If
    Next i
    SumOutstandingByPatient = t
End Function
Private Function CustIDToDrName(ByVal custID As String) As String
    Dim wsC As Worksheet, last As Long, i As Long, key As String
    key = UCase$(Replace(Trim$(custID), " ", ""))
    If key = "" Then Exit Function
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If UCase$(Replace(Trim$(CStr(wsC.Cells(i, 1).value)), " ", "")) = key Then
            CustIDToDrName = CStr(wsC.Cells(i, 2).value): Exit Function
        End If
    Next i
    CustIDToDrName = custID
End Function
