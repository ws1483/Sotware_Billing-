Attribute VB_Name = "modMenu"
' modMenu ? dashboard: QUOTES (rows 20-44), INVOICES + MED CLAIMS (rows 46-68).
'   PHASE 1: Med Claims (MedAidLog) are folded into the Invoices section,
'   distinguished by their MC-INV- number. Uses modConfig column constants.
' ============================================================================
Private Const MENU_SHEET As String = "Menu"

Public Sub RefreshMenuSummary()
    Dim wsM As Worksheet, showType As String
    Set wsM = ThisWorkbook.Sheets(MENU_SHEET)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Clean

    wsM.Range("B19:H68").ClearContents
    wsM.Range("B19").value = "-- Quotes (latest 25) --"
    wsM.Range("B45").value = "-- Invoices / Med Claims (latest 25) --"
    wsM.Range("G45").value = "Paid"
    wsM.Range("H45").value = "Balance"

    showType = UCase(Trim(CStr(wsM.Range("G17").value)))      ' All/Invoice/Quote

    If showType <> "INVOICE" Then FillBlock wsM, "QuoteLog", 20, 44
    If showType <> "QUOTE" Then FillInvoiceBlock wsM, 46, 68
Clean:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ---- Quotes block (unchanged behaviour) ----
Private Sub FillBlock(wsM As Worksheet, logName As String, firstRow As Long, lastRow As Long)
    Dim ws As Worksheet, last As Long, i As Long, cap As Long
    Dim recipF As String, drF As String, deptF As String, dFrom As Variant, dTo As Variant
    Dim rowsArr() As Long, dts() As Double, cnt As Long, j As Long, tL As Long, tD As Double
    Dim cRecip As Long, colDate As Long, cNo As Long, cCust As Long, cDept As Long
    Dim cPat As Long, cTotal As Long, cStatus As Long
    Dim outRow As Long, rr As Long, dv As Double, drName As String

    Set ws = ThisWorkbook.Sheets(logName)
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If last < 2 Then Exit Sub

    drF = LCase(Trim(CStr(wsM.Range("E17").value)))
    recipF = LCase(Trim(CStr(wsM.Range("I17").value)))
    deptF = UCase(Trim(CStr(wsM.Range("C5").value)))
    dFrom = wsM.Range("B17").value
    dTo = wsM.Range("C17").value
    If recipF = "private" Then recipF = "patient"

    ' Quote column map
    cNo = QL_NO: cDept = QL_DEPT: cRecip = QL_RECIP: colDate = QL_DATE
    cCust = QL_CUST: cPat = QL_PATIENT: cTotal = QL_TOTAL: cStatus = QL_STATUS

    ReDim rowsArr(1 To last)
    ReDim dts(1 To last)
    cnt = 0
    For i = 2 To last
        If Trim(CStr(ws.Cells(i, cNo).value)) = "" Then GoTo NextI
        If deptF <> "" And deptF <> "ALL" Then
            If UCase(Trim(CStr(ws.Cells(i, cDept).value))) <> deptF Then GoTo NextI
        End If
        If recipF <> "" And recipF <> "all" Then
            If LCase(Trim(CStr(ws.Cells(i, cRecip).value))) <> recipF Then GoTo NextI
        End If
        If drF <> "" And drF <> "all" Then
            drName = LCase(Trim(CustIDToDrName(CStr(ws.Cells(i, cCust).value))))
            If drName <> drF Then GoTo NextI
        End If
        If IsDate(ws.Cells(i, colDate).value) Then
            dv = CDbl(ws.Cells(i, colDate).value)
            If IsDate(dFrom) Then If dv < CDbl(dFrom) Then GoTo NextI
            If IsDate(dTo) Then If dv > CDbl(dTo) Then GoTo NextI
        End If
        cnt = cnt + 1
        rowsArr(cnt) = i
        dts(cnt) = IIf(IsDate(ws.Cells(i, colDate).value), CDbl(ws.Cells(i, colDate).value), 0)
NextI:
    Next i
    If cnt = 0 Then Exit Sub

    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If dts(j) < dts(j + 1) Then
                tD = dts(j): dts(j) = dts(j + 1): dts(j + 1) = tD
                tL = rowsArr(j): rowsArr(j) = rowsArr(j + 1): rowsArr(j + 1) = tL
            End If
        Next j
    Next i

    cap = lastRow - firstRow + 1
    outRow = firstRow
    For i = 1 To cnt
        If i > cap Then Exit For
        rr = rowsArr(i)
        wsM.Cells(outRow, 2).value = ws.Cells(rr, cNo).value
        wsM.Cells(outRow, 3).value = CustIDToDrName(CStr(ws.Cells(rr, cCust).value))
        wsM.Cells(outRow, 4).value = ws.Cells(rr, cPat).value
        wsM.Cells(outRow, 5).value = ws.Cells(rr, cTotal).value
        wsM.Cells(outRow, 6).value = ws.Cells(rr, cStatus).value
        outRow = outRow + 1
    Next i
End Sub

' ---- Invoices + Med Claims block (union InvoiceLog + MedAidLog) ----
Private Sub FillInvoiceBlock(wsM As Worksheet, firstRow As Long, lastRow As Long)
    Dim recipF As String, drF As String, deptF As String, dFrom As Variant, dTo As Variant
    ' unioned arrays
    Dim srcLog() As String, srcRow() As Long, dts() As Double, cnt As Long
    Dim i As Long, j As Long, cap As Long, outRow As Long
    Dim tS As String, tL As Long, tD As Double

    drF = LCase(Trim(CStr(wsM.Range("E17").value)))
    recipF = LCase(Trim(CStr(wsM.Range("I17").value)))
    deptF = UCase(Trim(CStr(wsM.Range("C5").value)))
    dFrom = wsM.Range("B17").value
    dTo = wsM.Range("C17").value
    If recipF = "private" Then recipF = "patient"

    ReDim srcLog(1 To 20000): ReDim srcRow(1 To 20000): ReDim dts(1 To 20000)
    cnt = 0

    CollectInvoiceRows "InvoiceLog", drF, recipF, deptF, dFrom, dTo, srcLog, srcRow, dts, cnt
    CollectInvoiceRows "MedAidLog", drF, recipF, deptF, dFrom, dTo, srcLog, srcRow, dts, cnt

    If cnt = 0 Then Exit Sub

    ' sort newest first
    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If dts(j) < dts(j + 1) Then
                tD = dts(j): dts(j) = dts(j + 1): dts(j + 1) = tD
                tL = srcRow(j): srcRow(j) = srcRow(j + 1): srcRow(j + 1) = tL
                tS = srcLog(j): srcLog(j) = srcLog(j + 1): srcLog(j + 1) = tS
            End If
        Next j
    Next i

    cap = lastRow - firstRow + 1
    outRow = firstRow
    For i = 1 To cnt
        If i > cap Then Exit For
        RenderInvoiceRow wsM, outRow, srcLog(i), srcRow(i)
        outRow = outRow + 1
    Next i
End Sub

Private Sub CollectInvoiceRows(logName As String, drF As String, recipF As String, _
        deptF As String, dFrom As Variant, dTo As Variant, _
        ByRef srcLog() As String, ByRef srcRow() As Long, ByRef dts() As Double, ByRef cnt As Long)
    Dim ws As Worksheet, last As Long, i As Long
    Dim cNo As Long, cDept As Long, cRecip As Long, colDate As Long, cCust As Long, cPat As Long
    Dim dv As Double, drName As String, isMC As Boolean

    Set ws = ThisWorkbook.Sheets(logName)
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If last < 2 Then Exit Sub
    isMC = (logName = "MedAidLog")

    If isMC Then
        cNo = ML_NO: cDept = ML_DEPT: cRecip = ML_RECIP: colDate = ML_DATE
        cCust = ML_CUST: cPat = ML_PATIENT
    Else
        cNo = IL_NO: cDept = IL_DEPT: cRecip = IL_RECIP: colDate = IL_DATE
        cCust = IL_CUST: cPat = IL_PATIENT
    End If

    For i = 2 To last
        If Trim(CStr(ws.Cells(i, cNo).value)) = "" Then GoTo NextI
        ' dept filter: MC rows carry WA/WD in cDept, so C5=WA/WD still works; C5=MC shows MC only
        If deptF <> "" And deptF <> "ALL" Then
            If deptF = "MC" Then
                If Not isMC Then GoTo NextI
            Else
                If UCase(Trim(CStr(ws.Cells(i, cDept).value))) <> deptF Then GoTo NextI
            End If
        End If
        If recipF <> "" And recipF <> "all" Then
            If LCase(Trim(CStr(ws.Cells(i, cRecip).value))) <> recipF Then GoTo NextI
        End If
        If drF <> "" And drF <> "all" Then
            drName = LCase(Trim(CustIDToDrName(CStr(ws.Cells(i, cCust).value))))
            If drName <> drF Then GoTo NextI
        End If
        If IsDate(ws.Cells(i, colDate).value) Then
            dv = CDbl(ws.Cells(i, colDate).value)
            If IsDate(dFrom) Then If dv < CDbl(dFrom) Then GoTo NextI
            If IsDate(dTo) Then If dv > CDbl(dTo) Then GoTo NextI
        End If
        cnt = cnt + 1
        srcLog(cnt) = logName
        srcRow(cnt) = i
        dts(cnt) = IIf(IsDate(ws.Cells(i, colDate).value), CDbl(ws.Cells(i, colDate).value), 0)
NextI:
    Next i
End Sub

Private Sub RenderInvoiceRow(wsM As Worksheet, outRow As Long, logName As String, rr As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(logName)
    If logName = "MedAidLog" Then
        wsM.Cells(outRow, 2).value = ws.Cells(rr, ML_NO).value
        wsM.Cells(outRow, 3).value = ws.Cells(rr, ML_CUST).value       ' Bill To (patient key)
        wsM.Cells(outRow, 4).value = ws.Cells(rr, ML_PATIENT).value
        wsM.Cells(outRow, 5).value = ws.Cells(rr, ML_TOTAL).value
        wsM.Cells(outRow, 6).value = ws.Cells(rr, ML_STATUS).value
        wsM.Cells(outRow, 7).value = ws.Cells(rr, ML_PAID).value
        wsM.Cells(outRow, 8).value = ws.Cells(rr, ML_BALANCE).value
    Else
        wsM.Cells(outRow, 2).value = ws.Cells(rr, IL_NO).value
        wsM.Cells(outRow, 3).value = CustIDToDrName(CStr(ws.Cells(rr, IL_CUST).value))
        wsM.Cells(outRow, 4).value = ws.Cells(rr, IL_PATIENT).value
        wsM.Cells(outRow, 5).value = ws.Cells(rr, IL_TOTAL).value
        wsM.Cells(outRow, 6).value = ws.Cells(rr, IL_STATUS).value
        wsM.Cells(outRow, 7).value = ws.Cells(rr, IL_PAID).value
        wsM.Cells(outRow, 8).value = ws.Cells(rr, IL_BALANCE).value
    End If
End Sub

