Attribute VB_Name = "modPayment"
Option Explicit
' ============================================================================
' modPayment ? launches Record Payment + (Phase 3) Manage Payments.
'   ReconcileDoc recomputes Paid/Balance/Status from remaining Payments rows.
' ============================================================================

Sub RecordPayment()
    frmPayment.Show
End Sub

Sub ManagePayments()
    frmPaymentManage.Show     ' Part B (built next)
End Sub

' docType: "invoice" | "medclaim"
Public Sub ReconcileDoc(ByVal docType As String, ByVal docNo As String)
    Dim wsLog As Worksheet, wsP As Worksheet
    Dim lr As Long, last As Long, i As Long
    Dim paidCol As Long, balCol As Long, statusCol As Long, totalCol As Long, modCol As Long
    Dim paid As Double, total As Double, bal As Double

    If docType = "medclaim" Then
        Set wsLog = ThisWorkbook.Sheets("MedAidLog")
        paidCol = ML_PAID: balCol = ML_BALANCE: statusCol = ML_STATUS
        totalCol = ML_TOTAL: modCol = ML_MODIFIED
    Else
        Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
        paidCol = IL_PAID: balCol = IL_BALANCE: statusCol = IL_STATUS
        totalCol = IL_TOTAL: modCol = IL_MODIFIED
    End If

    lr = FindLogRow(wsLog, docNo)
    If lr = 0 Then Exit Sub

    Set wsP = ThisWorkbook.Sheets("Payments")
    last = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).row
    paid = 0
    For i = 2 To last
        If NrmID(CStr(wsP.Cells(i, PY_INV).value)) = NrmID(docNo) Then
            paid = paid + Num(wsP.Cells(i, PY_AMT).value)
        End If
    Next i

    total = Num(wsLog.Cells(lr, totalCol).value)
    bal = total - paid: If bal < 0 Then bal = 0
    wsLog.Cells(lr, paidCol).value = paid
    wsLog.Cells(lr, balCol).value = bal
    If paid <= 0 Then
        wsLog.Cells(lr, statusCol).value = "Unpaid"
    ElseIf bal <= 0.005 Then
        wsLog.Cells(lr, statusCol).value = "Paid"
    Else
        wsLog.Cells(lr, statusCol).value = "Part-Paid"
    End If
    wsLog.Cells(lr, modCol).value = Now
End Sub
