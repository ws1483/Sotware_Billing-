VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPaymentManage 
   Caption         =   "Manage Payments"
   ClientHeight    =   7224
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9228.001
   OleObjectBlob   =   "frmPaymentManage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPaymentManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ============================================================================
' frmPaymentManage ? view / edit / delete individual payments.
'   TOP: searchable, dept-filtered SINGLE-select list of Invoices + Med Claims.
'   MID: lstPays = all Payments rows for the selected doc (per-doc row only).
'   EDIT: amount(any positive)/date/method/ref/notes -> rewrite row ->
'         ReconcileDoc -> audit. DELETE: remove row -> ReconcileDoc -> audit.
'   Doctor-credit from original overpay is NOT auto-unwound (warn only).
' ============================================================================
Private mRowType() As String     ' per doc-list row: "invoice"|"medclaim"
Private mRowNo() As String        ' per doc-list row: doc number
Private mSelDocNo As String
Private mSelDocType As String
Private mPayRowSheet() As Long    ' per lstPays row: absolute Payments sheet row

Private Sub UserForm_Initialize()
    LoadMethods

    cboDept.Clear
    cboDept.AddItem "All": cboDept.AddItem "WA": cboDept.AddItem "WD": cboDept.AddItem "MC"
    cboDept.value = "All"

    lstDocs.ColumnCount = 7
    lstDocs.ColumnHeads = False
    lstDocs.MultiSelect = fmMultiSelectSingle
    lstDocs.ColumnWidths = "80;120;110;65;70;70;60"

    lstPays.ColumnCount = 6
    lstPays.ColumnHeads = False
    lstPays.ColumnWidths = "70;65;70;70;80;120"   ' PayID|Date|Amount|Method|Ref|Notes

    ClearEditFields
    lblDoc.Caption = "(select a document)"
    lblInfo.Caption = ""
    PopulateDocs
End Sub

Private Sub txtSearch_Change(): PopulateDocs
End Sub
Private Sub cboDept_Change():   PopulateDocs
End Sub

' ------------------- doc picker (single-select) -------------------
Private Sub PopulateDocs()
    Dim term As String, deptF As String
    term = LCase(Trim(txtSearch.value))
    deptF = UCase(Trim(cboDept.value))

    lstDocs.Clear
    ReDim mRowType(0 To 5000)
    ReDim mRowNo(0 To 5000)
    mSelDocNo = "": mSelDocType = ""
    lstPays.Clear
    ClearEditFields
    lblDoc.Caption = "(select a document)"
    lblInfo.Caption = ""

    If deptF = "ALL" Or deptF = "WA" Or deptF = "WD" Then AddDocs "InvoiceLog", "invoice", term, deptF
    If deptF = "ALL" Or deptF = "MC" Or deptF = "WA" Or deptF = "WD" Then AddDocs "MedAidLog", "medclaim", term, deptF
End Sub

Private Sub AddDocs(logName As String, docType As String, term As String, deptF As String)
    Dim ws As Worksheet, last As Long, i As Long
    Dim cNo As Long, cDept As Long, cRecip As Long, cDt As Long
    Dim cCust As Long, cPat As Long, cTotal As Long, cBal As Long, cStatus As Long
    Dim docNo As String, drBill As String, patient As String, status As String
    Dim total As Variant, bal As Double, dt As Variant, hay As String, dept As String, li As Long

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

        ' show docs that have any payment history OR a balance (i.e. not brand-new empty)
        status = Trim(CStr(ws.Cells(i, cStatus).value))

        dept = UCase(Trim(CStr(ws.Cells(i, cDept).value)))
        If deptF = "WA" Or deptF = "WD" Then
            If dept <> deptF Then GoTo NextI
        ElseIf deptF = "MC" Then
            If docType <> "medclaim" Then GoTo NextI
        End If

        If docType = "medclaim" Then
            drBill = Trim(CStr(ws.Cells(i, cCust).value))
        Else
            If LCase(Trim(CStr(ws.Cells(i, cRecip).value))) = "patient" Then
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
NextI:
    Next i
End Sub

Private Sub lstDocs_Click()
    If lstDocs.ListIndex < 0 Then Exit Sub
    mSelDocNo = mRowNo(lstDocs.ListIndex)
    mSelDocType = mRowType(lstDocs.ListIndex)
    lblDoc.Caption = "Payments for " & mSelDocNo & "  (" & mSelDocType & ")"
    PopulatePays
End Sub

' ------------------- payments for selected doc -------------------
Private Sub PopulatePays()
    Dim wsP As Worksheet, last As Long, i As Long, li As Long
    lstPays.Clear
    ClearEditFields
    ReDim mPayRowSheet(0 To 5000)
    If mSelDocNo = "" Then Exit Sub

    Set wsP = ThisWorkbook.Sheets("Payments")
    last = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsP.Cells(i, PY_INV).value)) = NrmID(mSelDocNo) Then
            li = lstPays.ListCount
            lstPays.AddItem CStr(wsP.Cells(i, PY_ID).value)
            lstPays.List(li, 1) = IIf(IsDate(wsP.Cells(i, PY_DATE).value), _
                                      Format(wsP.Cells(i, PY_DATE).value, "yyyy-mm-dd"), "")
            lstPays.List(li, 2) = Format(Num(wsP.Cells(i, PY_AMT).value), "#,##0.00")
            lstPays.List(li, 3) = CStr(wsP.Cells(i, PY_METHOD).value)
            lstPays.List(li, 4) = CStr(wsP.Cells(i, PY_REF).value)
            lstPays.List(li, 5) = CStr(wsP.Cells(i, PY_NOTES).value)
            mPayRowSheet(li) = i
        End If
    Next i

    lblInfo.Caption = lstPays.ListCount & " payment(s) for " & mSelDocNo
End Sub

Private Sub lstPays_Click()
    Dim wsP As Worksheet, r As Long
    If lstPays.ListIndex < 0 Then Exit Sub
    Set wsP = ThisWorkbook.Sheets("Payments")
    r = mPayRowSheet(lstPays.ListIndex)
    txtAmount.value = Format(Num(wsP.Cells(r, PY_AMT).value), "0.00")
    txtDate.value = IIf(IsDate(wsP.Cells(r, PY_DATE).value), _
                        Format(wsP.Cells(r, PY_DATE).value, "yyyy-mm-dd"), "")
    On Error Resume Next
    cboMethod.value = CStr(wsP.Cells(r, PY_METHOD).value)
    On Error GoTo 0
    txtRef.value = CStr(wsP.Cells(r, PY_REF).value)
    txtNotes.value = CStr(wsP.Cells(r, PY_NOTES).value)
End Sub

' ------------------- edit (save) -------------------
Private Sub btnSaveEdit_Click()
    Dim wsP As Worksheet, r As Long, payID As String
    Dim oldAmt As Double, newAmt As Double

    If mSelDocNo = "" Then MsgBox "Select a document first.", vbExclamation: Exit Sub
    If lstPays.ListIndex < 0 Then MsgBox "Select a payment row to edit.", vbExclamation: Exit Sub
    If Not IsNumeric(txtAmount.value) Then MsgBox "Enter a valid amount.", vbExclamation: Exit Sub
    newAmt = Num(txtAmount.value)
    If newAmt <= 0 Then MsgBox "Amount must be greater than zero.", vbExclamation: Exit Sub
    If Not IsDate(txtDate.value) Then MsgBox "Enter a valid date (yyyy-mm-dd).", vbExclamation: Exit Sub
    If Trim(cboMethod.value) = "" Then MsgBox "Select a payment method.", vbExclamation: Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set wsP = ThisWorkbook.Sheets("Payments")
    r = mPayRowSheet(lstPays.ListIndex)
    payID = CStr(wsP.Cells(r, PY_ID).value)
    oldAmt = Num(wsP.Cells(r, PY_AMT).value)

    ' warn if this payment likely created doctor credit (was an overpayment)
    WarnIfCreditLinked payID, oldAmt

    wsP.Cells(r, PY_DATE).value = CDate(txtDate.value)
    wsP.Cells(r, PY_AMT).value = newAmt
    wsP.Cells(r, PY_METHOD).value = cboMethod.value
    wsP.Cells(r, PY_REF).value = txtRef.value
    wsP.Cells(r, PY_NOTES).value = txtNotes.value

    ReconcileDoc mSelDocType, mSelDocNo
    LogAudit "PayEdit", mSelDocNo, "Amt " & Format(oldAmt, "0.00"), _
             "Amt " & Format(newAmt, "0.00"), "Edit " & payID

    Application.ScreenUpdating = True
    MsgBox "Payment " & payID & " updated.", vbInformation
    PopulatePays
    RefreshDocRow
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Edit error: " & Err.Description, vbExclamation
End Sub

' ------------------- delete -------------------
Private Sub btnDelete_Click()
    Dim wsP As Worksheet, r As Long, payID As String, amt As Double

    If mSelDocNo = "" Then MsgBox "Select a document first.", vbExclamation: Exit Sub
    If lstPays.ListIndex < 0 Then MsgBox "Select a payment row to delete.", vbExclamation: Exit Sub

    Set wsP = ThisWorkbook.Sheets("Payments")
    r = mPayRowSheet(lstPays.ListIndex)
    payID = CStr(wsP.Cells(r, PY_ID).value)
    amt = Num(wsP.Cells(r, PY_AMT).value)

    If MsgBox("Are you sure? This will delete " & payID & " (R " & Format(amt, "#,##0.00") & _
              ") and recompute the balance for " & mSelDocNo & ".", _
              vbQuestion + vbYesNo, "Delete Payment") = vbNo Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False

    WarnIfCreditLinked payID, amt

    wsP.Rows(r).Delete
    ReconcileDoc mSelDocType, mSelDocNo
    LogAudit "PayDelete", mSelDocNo, "Amt " & Format(amt, "0.00"), "", "Delete " & payID

    Application.ScreenUpdating = True
    MsgBox "Payment " & payID & " deleted.", vbInformation
    PopulatePays
    RefreshDocRow
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Delete error: " & Err.Description, vbExclamation
End Sub

Private Sub btnClose_Click(): Unload Me
End Sub

' ------------------- helpers -------------------
Private Sub ClearEditFields()
    txtAmount.value = "": txtDate.value = ""
    txtRef.value = "": txtNotes.value = ""
    If cboMethod.ListCount > 0 Then cboMethod.ListIndex = -1
End Sub

' Refresh the selected doc's balance/status shown in lstDocs after reconcile
Private Sub RefreshDocRow()
    Dim wsLog As Worksheet, lr As Long, balCol As Long, statCol As Long
    Dim idx As Long
    idx = lstDocs.ListIndex
    If idx < 0 Then Exit Sub
    If mSelDocType = "medclaim" Then
        Set wsLog = ThisWorkbook.Sheets("MedAidLog"): balCol = ML_BALANCE: statCol = ML_STATUS
    Else
        Set wsLog = ThisWorkbook.Sheets("InvoiceLog"): balCol = IL_BALANCE: statCol = IL_STATUS
    End If
    lr = FindLogRow(wsLog, mSelDocNo)
    If lr = 0 Then Exit Sub
    lstDocs.List(idx, 5) = Format(Num(wsLog.Cells(lr, balCol).value), "#,##0.00")
    lstDocs.List(idx, 6) = CStr(wsLog.Cells(lr, statCol).value)
End Sub

' Warn (do not auto-reverse) if this payID added doctor credit via an overpay audit
Private Sub WarnIfCreditLinked(payID As String, amt As Double)
    Dim wsA As Worksheet, last As Long, i As Long, hit As Boolean
    On Error Resume Next
    Set wsA = ThisWorkbook.Sheets("AuditLog")
    If wsA Is Nothing Then Exit Sub
    last = wsA.Cells(wsA.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        ' AuditLog note column typically last; scan any cell containing payID + 'Overpay'/'Credit'
        If InStr(1, CStr(wsA.Cells(i, wsA.UsedRange.Columns.Count).value), payID, vbTextCompare) > 0 _
           And (InStr(1, CStr(wsA.Cells(i, 1).value), "Credit", vbTextCompare) > 0 _
                Or InStr(1, CStr(wsA.Cells(i, 1).value), "Overpay", vbTextCompare) > 0) Then
            hit = True: Exit For
        End If
    Next i
    On Error GoTo 0
    If hit Then
        MsgBox "Note: " & payID & " is linked to a doctor-credit / overpayment entry." & vbCrLf & _
               "This change will NOT automatically adjust stored doctor credit." & vbCrLf & _
               "Please review the doctor's credit balance manually if needed.", vbExclamation
    End If
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
    Exit Sub
EH:
    MsgBox "LoadMethods error: " & Err.Description, vbExclamation
End Sub

