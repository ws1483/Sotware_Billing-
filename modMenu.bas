Attribute VB_Name = "modMenu"
Option Explicit
' ============================================================================
' modMenu — dashboard: QUOTES first (rows 20-44), INVOICES second (rows 46-68).
'   Filters (row 17): B17 DateFrom, C17 DateTo, E17 Doctor(DrName/All),
'                     G17 Type(All/Invoice/Quote), I17 Recipient(All/Doctor/Private)
'   Dept filter: C5 (blank/ALL = show everything), matched vs log col B (Dept)
'   Output cols: B=DocNo C=Doctor D=Patient E=Total F=Status G=Paid H=Balance
' ============================================================================
Private Const MENU_SHEET As String = "Menu"

Public Sub RefreshMenuSummary()
    Dim wsM As Worksheet, showType As String
    Set wsM = ThisWorkbook.Sheets(MENU_SHEET)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Clean

    wsM.Range("B19:H68").ClearContents
    wsM.Range("B19").Value = "-- Quotes (latest 25) --"
    wsM.Range("B45").Value = "-- Invoices (latest 25) --"
    ' Paid/Balance headers only meaningful for invoices
    wsM.Range("G45").Value = "Paid"
    wsM.Range("H45").Value = "Balance"

    showType = UCase(Trim(CStr(wsM.Range("G17").Value)))      ' All/Invoice/Quote

    If showType <> "INVOICE" Then FillBlock wsM, "QuoteLog", 20, 44
    If showType <> "QUOTE" Then FillBlock wsM, "InvoiceLog", 46, 68
Clean:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub FillBlock(wsM As Worksheet, logName As String, firstRow As Long, lastRow As Long)
    Dim ws As Worksheet, last As Long, i As Long, cap As Long
    Dim recipF As String, drF As String, deptF As String, dFrom As Variant, dTo As Variant
    Dim rowsArr() As Long, dts() As Double, cnt As Long, j As Long, tL As Long, tD As Double
    Dim cRecip As Long, colDate As Long, cNo As Long, cCust As Long, cDept As Long
    Dim cPat As Long, cTotal As Long, cStatus As Long, cPaid As Long, cBal As Long
    Dim outRow As Long, rr As Long, dv As Double, drName As String

    Set ws = ThisWorkbook.Sheets(logName)
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If last < 2 Then Exit Sub

    ' ---- read filters ----
    drF = LCase(Trim(CStr(wsM.Range("E17").Value)))           ' doctor name or All/blank
    recipF = LCase(Trim(CStr(wsM.Range("I17").Value)))        ' all/doctor/private
    deptF = UCase(Trim(CStr(wsM.Range("C5").Value)))          ' dept or ALL/blank
    dFrom = wsM.Range("B17").Value
    dTo = wsM.Range("C17").Value

    ' Menu shows "Private" but log stores "patient"
    If recipF = "private" Then recipF = "patient"

    ' ---- column map per log ----
    If logName = "InvoiceLog" Then
        cNo = 1: cDept = 2: cRecip = 3: colDate = 4: cCust = 6: cPat = 7
        cTotal = 12: cStatus = 16: cPaid = 14: cBal = 15
    Else
        cNo = 1: cDept = 2: cRecip = 3: colDate = 4: cCust = 5: cPat = 6
        cTotal = 11: cStatus = 13: cPaid = 0: cBal = 0    ' quotes: no paid/balance
    End If

    ReDim rowsArr(1 To last)
    ReDim dts(1 To last)
    cnt = 0
    For i = 2 To last
        If Trim(CStr(ws.Cells(i, cNo).Value)) = "" Then GoTo NextI

        ' dept filter (C5)
        If deptF <> "" And deptF <> "ALL" Then
            If UCase(Trim(CStr(ws.Cells(i, cDept).Value))) <> deptF Then GoTo NextI
        End If

        ' recipient filter (I17)
        If recipF <> "" And recipF <> "all" Then
            If LCase(Trim(CStr(ws.Cells(i, cRecip).Value))) <> recipF Then GoTo NextI
        End If

        ' doctor filter (E17) — match resolved DrName
        If drF <> "" And drF <> "all" Then
            drName = LCase(Trim(CustIDToDrName(CStr(ws.Cells(i, cCust).Value))))
            If drName <> drF Then GoTo NextI
        End If

        ' date range filter (B17..C17)
        If IsDate(ws.Cells(i, colDate).Value) Then
            dv = CDbl(ws.Cells(i, colDate).Value)
            If IsDate(dFrom) Then
                If dv < CDbl(dFrom) Then GoTo NextI
            End If
            If IsDate(dTo) Then
                If dv > CDbl(dTo) Then GoTo NextI
            End If
        End If

        cnt = cnt + 1
        rowsArr(cnt) = i
        dts(cnt) = IIf(IsDate(ws.Cells(i, colDate).Value), CDbl(ws.Cells(i, colDate).Value), 0)
NextI:
    Next i
    If cnt = 0 Then Exit Sub

    ' sort NEWEST first (descending date)
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
        wsM.Cells(outRow, 2).Value = ws.Cells(rr, cNo).Value
        wsM.Cells(outRow, 3).Value = CustIDToDrName(CStr(ws.Cells(rr, cCust).Value))
        wsM.Cells(outRow, 4).Value = ws.Cells(rr, cPat).Value
        wsM.Cells(outRow, 5).Value = ws.Cells(rr, cTotal).Value
        wsM.Cells(outRow, 6).Value = ws.Cells(rr, cStatus).Value
        ' Paid / Balance — invoices only
        If cPaid > 0 Then
            wsM.Cells(outRow, 7).Value = ws.Cells(rr, cPaid).Value    ' G = Paid
            wsM.Cells(outRow, 8).Value = ws.Cells(rr, cBal).Value     ' H = Balance
        End If
        outRow = outRow + 1
    Next i
End Sub

