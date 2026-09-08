Attribute VB_Name = "modReport"
Option Explicit

Private Const RPT_SHEET As String = "Report"

Public Sub ShowReport()
    frmReport.Show
End Sub

Public Sub RunReport(ByVal rptType As String, ByVal mode As String, ByVal dFrom As Date, ByVal dTo As Date, ByVal dept As String)
    Dim ws As Worksheet, ans As VbMsgBoxResult
    Dim folder As String, fpath As String
    Dim t As String, m As String, d As String

    On Error GoTo Fail
    Application.ScreenUpdating = False

    t = UCase$(Trim$(rptType))
    m = UCase$(Trim$(mode))
    d = UCase$(Trim$(dept))
    If d = "" Then d = "ALL"

    Set ws = FreshReportSheet()
    HeaderBlock ws, t, m, dFrom, dTo, d

    Select Case t
        Case "VAT":         BuildVAT ws, m, dFrom, dTo, d
        Case "INVOICES":    BuildInvoices ws, m, dFrom, dTo, d
        Case "QUOTES":      BuildQuotes ws, m, dFrom, dTo, d
        Case "MEDCLAIMS":   BuildMedClaims ws, m, dFrom, dTo, d
        Case "CREDITNOTES": BuildCreditNotes ws, m, dFrom, dTo, d
        Case "SUMMARY":     BuildSummary ws, dFrom, dTo, d
        Case Else:          Err.Raise vbObjectError + 100, , "Unknown report type: " & rptType
    End Select

    ws.Columns("A:L").AutoFit
    ws.Range("A1").Select
    Application.ScreenUpdating = True

    ans = MsgBox("Report ready on screen." & vbCrLf & vbCrLf & "Create PDF now?", vbQuestion + vbYesNo, "Preview Report")
    If ans = vbYes Then
        folder = StmtPickFolder()
        If folder <> "" Then
            fpath = ExportReportPDF(ws, t, m, dFrom, dTo, d, folder)
            MsgBox "Report PDF saved:" & vbCrLf & fpath, vbInformation
        End If
    End If
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "RunReport error: " & Err.Description, vbExclamation
End Sub

Private Function FreshReportSheet() As Worksheet
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(RPT_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = RPT_SHEET
    ws.Cells.Font.Name = "Calibri"
    Set FreshReportSheet = ws
End Function

Private Sub HeaderBlock(ws As Worksheet, rptType As String, mode As String, dFrom As Date, dTo As Date, dept As String)
    ws.Cells(1, 1).value = "WE-DENTAL - " & UCase$(rptType) & " REPORT (" & UCase$(mode) & ")"
    ws.Cells(1, 1).Font.Size = 15
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).value = "Period: " & Format$(dFrom, "yyyy-mm-dd") & "  to  " & Format$(dTo, "yyyy-mm-dd") & "     Dept: " & UCase$(dept)
End Sub

Private Function InRange(v As Variant, dFrom As Date, dTo As Date) As Boolean
    If Not IsDate(v) Then Exit Function
    Dim dd As Date
    dd = CDate(v)
    InRange = (dd >= dFrom And dd <= dTo)
End Function

Private Sub SectionTitle(ws As Worksheet, rowN As Long, title As String)
    ws.Cells(rowN, 1).value = title
    ws.Cells(rowN, 1).Font.Bold = True
    ws.Range(ws.Cells(rowN, 1), ws.Cells(rowN, 6)).Interior.Color = RGB(31, 78, 121)
    ws.Range(ws.Cells(rowN, 1), ws.Cells(rowN, 6)).Font.Color = vbWhite
End Sub

Private Sub KeyVal(ws As Worksheet, rowN As Long, label As String, value As Variant)
    ws.Cells(rowN, 1).value = label
    ws.Cells(rowN, 2).value = value
    If IsNumeric(value) Then ws.Cells(rowN, 2).NumberFormat = "#,##0.00"
End Sub

Private Sub ColHead(ws As Worksheet, rowN As Long, headers As Variant)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(rowN, i + 1).value = headers(i)
        ws.Cells(rowN, i + 1).Font.Bold = True
        ws.Cells(rowN, i + 1).Interior.Color = RGB(220, 230, 241)
    Next i
End Sub

' ================= VAT =================

Private Sub BuildVAT(ws As Worksheet, mode As String, dFrom As Date, dTo As Date, deptCode As String)
    If UCase$(mode) = "FULL" Then
        VATFull ws, dFrom, dTo, deptCode
        Exit Sub
    End If

    Dim iVat As Double, mVat As Double, cVat As Double, r As Long

    iVat = VATInvoiceTotal(dFrom, dTo, deptCode)
    mVat = VATMedTotal(dFrom, dTo, deptCode)
    If deptCode = "ALL" Then cVat = VATCreditTotal(dFrom, dTo)

    r = 4
    SectionTitle ws, r, "VAT PAYABLE (period)": r = r + 1
    KeyVal ws, r, "Output VAT - Invoices", iVat: r = r + 1
    KeyVal ws, r, "Output VAT - Med Claims", mVat: r = r + 1
    KeyVal ws, r, "Less VAT - Credit Notes", -cVat: r = r + 1

    If deptCode <> "ALL" Then
        ws.Cells(r, 1).value = "Credit Notes have no dept; excluded"
        ws.Cells(r, 1).Font.Italic = True
        r = r + 1
    End If

    ws.Cells(r, 1).value = "VAT PAYABLE"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).value = iVat + mVat - IIf(deptCode = "ALL", cVat, 0)
    ws.Cells(r, 2).NumberFormat = "#,##0.00"
    ws.Cells(r, 2).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(220, 235, 220)
End Sub

Private Function VATInvoiceTotal(dFrom As Date, dTo As Date, deptCode As String) As Double
    Dim ws As Worksheet, lastRow As Long, i As Long, dp As String
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, IL_NO).value)) <> "" Then
            If InRange(ws.Cells(i, IL_DATE).value, dFrom, dTo) Then
                dp = UCase$(Trim$(CStr(ws.Cells(i, IL_DEPT).value)))
                If deptCode = "ALL" Or dp = deptCode Then
                    VATInvoiceTotal = VATInvoiceTotal + Num(ws.Cells(i, IL_VAT).value)
                End If
            End If
        End If
    Next i
End Function

Private Function VATMedTotal(dFrom As Date, dTo As Date, deptCode As String) As Double
    Dim ws As Worksheet, lastRow As Long, i As Long, dp As String
    Set ws = ThisWorkbook.Sheets("MedAidLog")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, ML_NO).value)) <> "" Then
            If InRange(ws.Cells(i, ML_DATE).value, dFrom, dTo) Then
                dp = UCase$(Trim$(CStr(ws.Cells(i, ML_DEPT).value)))
                If deptCode = "ALL" Then
                    VATMedTotal = VATMedTotal + Num(ws.Cells(i, ML_VAT).value)
                ElseIf (deptCode = "WA" Or deptCode = "WD") And dp = deptCode Then
                    VATMedTotal = VATMedTotal + Num(ws.Cells(i, ML_VAT).value)
                End If
            End If
        End If
    Next i
End Function

Private Function VATCreditTotal(dFrom As Date, dTo As Date) As Double
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("CreditNotes")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, 1).value)) <> "" Then
            If InRange(ws.Cells(i, 3).value, dFrom, dTo) Then
                VATCreditTotal = VATCreditTotal + Num(ws.Cells(i, 6).value)
            End If
        End If
    Next i
End Function

Private Sub VATFull(ws As Worksheet, dFrom As Date, dTo As Date, deptCode As String)
    Dim r As Long, tot As Double
    r = 4
    ColHead ws, r, Array("Doc No", "Date", "Dept", "Type", "Excl", "VAT")
    r = r + 1

    r = VATFullRows(ws, r, "InvoiceLog", IL_NO, IL_DATE, IL_DEPT, IL_SUBTOTAL, IL_VAT, "Inv", dFrom, dTo, deptCode, False, tot, 1)
    r = VATFullRows(ws, r, "MedAidLog", ML_NO, ML_DATE, ML_DEPT, ML_SUBTOTAL, ML_VAT, "MC", dFrom, dTo, deptCode, True, tot, 1)
    If deptCode = "ALL" Then
        r = VATFullRows(ws, r, "CreditNotes", 1, 3, 0, 5, 6, "CN", dFrom, dTo, deptCode, False, tot, -1)
    End If

    ws.Cells(r, 4).value = "VAT TOTAL"
    ws.Cells(r, 4).Font.Bold = True
    ws.Cells(r, 6).value = tot
    ws.Cells(r, 6).NumberFormat = "#,##0.00"
    ws.Cells(r, 6).Font.Bold = True
End Sub

Private Function VATFullRows(ws As Worksheet, startRow As Long, logName As String, cNo As Long, cDt As Long, cDept As Long, cExcl As Long, cVat As Long, tag As String, dFrom As Date, dTo As Date, deptCode As String, isMC As Boolean, ByRef tot As Double, sgn As Integer) As Long
    Dim src As Worksheet, lastRow As Long, i As Long
    Dim dp As String, vv As Double, ee As Double, r As Long

    Set src = ThisWorkbook.Sheets(logName)
    lastRow = src.Cells(src.Rows.Count, "A").End(xlUp).row
    r = startRow

    For i = 2 To lastRow
        If Trim$(CStr(src.Cells(i, cNo).value)) = "" Then GoTo NextI
        If Not InRange(src.Cells(i, cDt).value, dFrom, dTo) Then GoTo NextI

        dp = ""
        If cDept > 0 Then dp = UCase$(Trim$(CStr(src.Cells(i, cDept).value)))

        If deptCode <> "ALL" And cDept > 0 Then
            If isMC Then
                If (deptCode = "WA" Or deptCode = "WD") And dp <> deptCode Then GoTo NextI
            Else
                If dp <> deptCode Then GoTo NextI
            End If
        End If

        ee = Num(src.Cells(i, cExcl).value) * sgn
        vv = Num(src.Cells(i, cVat).value) * sgn

        ws.Cells(r, 1).value = src.Cells(i, cNo).value
        ws.Cells(r, 2).value = IIf(IsDate(src.Cells(i, cDt).value), Format$(src.Cells(i, cDt).value, "yyyy-mm-dd"), "")
        ws.Cells(r, 3).value = dp
        ws.Cells(r, 4).value = tag
        ws.Cells(r, 5).value = ee
        ws.Cells(r, 5).NumberFormat = "#,##0.00"
        ws.Cells(r, 6).value = vv
        ws.Cells(r, 6).NumberFormat = "#,##0.00"

        tot = tot + vv
        r = r + 1
NextI:
    Next i

    VATFullRows = r
End Function

' ================= INVOICES =================

Private Sub BuildInvoices(ws As Worksheet, mode As String, dFrom As Date, dTo As Date, deptCode As String)
    Dim src As Worksheet, lastRow As Long, i As Long, r As Long
    Dim dp As String, st As String, billTo As String, cnt As Long
    Dim sExcl As Double, sVat As Double, sTot As Double, sPaid As Double, sBal As Double
    Dim nU As Long, nP As Long, nPd As Long

    Set src = ThisWorkbook.Sheets("InvoiceLog")
    lastRow = src.Cells(src.Rows.Count, "A").End(xlUp).row
    r = 4

    If UCase$(mode) = "FULL" Then
        ColHead ws, r, Array("Inv No", "Date", "Dept", "Bill-To", "Patient", "Excl", "VAT", "Total", "Paid", "Balance", "Status")
        r = r + 1
    End If

    For i = 2 To lastRow
        If Trim$(CStr(src.Cells(i, IL_NO).value)) = "" Then GoTo NextI
        If Not InRange(src.Cells(i, IL_DATE).value, dFrom, dTo) Then GoTo NextI

        dp = UCase$(Trim$(CStr(src.Cells(i, IL_DEPT).value)))
        If deptCode <> "ALL" And dp <> deptCode Then GoTo NextI

        cnt = cnt + 1
        sExcl = sExcl + Num(src.Cells(i, IL_SUBTOTAL).value)
        sVat = sVat + Num(src.Cells(i, IL_VAT).value)
        sTot = sTot + Num(src.Cells(i, IL_TOTAL).value)
        sPaid = sPaid + Num(src.Cells(i, IL_PAID).value)
        sBal = sBal + Num(src.Cells(i, IL_BALANCE).value)

        st = UCase$(Trim$(CStr(src.Cells(i, IL_STATUS).value)))
        Select Case st
            Case "PAID": nPd = nPd + 1
            Case "PART-PAID": nP = nP + 1
            Case Else: nU = nU + 1
        End Select

        If UCase$(mode) = "FULL" Then
            If UCase$(CStr(src.Cells(i, IL_RECIP).value)) = "PATIENT" Then
                billTo = "(patient)"
            Else
                billTo = CustIDToDrName(CStr(src.Cells(i, IL_CUST).value))
            End If
            ws.Cells(r, 1).value = src.Cells(i, IL_NO).value
            ws.Cells(r, 2).value = Format$(src.Cells(i, IL_DATE).value, "yyyy-mm-dd")
            ws.Cells(r, 3).value = dp
            ws.Cells(r, 4).value = billTo
            ws.Cells(r, 5).value = src.Cells(i, IL_PATIENT).value
            ws.Cells(r, 6).value = Num(src.Cells(i, IL_SUBTOTAL).value)
            ws.Cells(r, 7).value = Num(src.Cells(i, IL_VAT).value)
            ws.Cells(r, 8).value = Num(src.Cells(i, IL_TOTAL).value)
            ws.Cells(r, 9).value = Num(src.Cells(i, IL_PAID).value)
            ws.Cells(r, 10).value = Num(src.Cells(i, IL_BALANCE).value)
            ws.Cells(r, 11).value = src.Cells(i, IL_STATUS).value
            ws.Range(ws.Cells(r, 6), ws.Cells(r, 10)).NumberFormat = "#,##0.00"
            r = r + 1
        End If
NextI:
    Next i

    If UCase$(mode) = "FULL" Then
        ws.Cells(r, 5).value = "TOTALS"
        ws.Cells(r, 5).Font.Bold = True
        ws.Cells(r, 6).value = sExcl
        ws.Cells(r, 7).value = sVat
        ws.Cells(r, 8).value = sTot
        ws.Cells(r, 9).value = sPaid
        ws.Cells(r, 10).value = sBal
        ws.Range(ws.Cells(r, 6), ws.Cells(r, 10)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(r, 6), ws.Cells(r, 10)).Font.Bold = True
    Else
        SectionTitle ws, r, "INVOICES - SUMMARY": r = r + 1
        KeyVal ws, r, "Count", cnt: r = r + 1
        KeyVal ws, r, "Subtotal (Excl)", sExcl: r = r + 1
        KeyVal ws, r, "VAT", sVat: r = r + 1
        KeyVal ws, r, "Total (Incl)", sTot: r = r + 1
        KeyVal ws, r, "Paid", sPaid: r = r + 1
        KeyVal ws, r, "Outstanding", sBal: r = r + 1
        ws.Cells(r, 1).value = "Unpaid / Part-Paid / Paid"
        ws.Cells(r, 2).value = nU & " / " & nP & " / " & nPd
    End If
End Sub

' ================= QUOTES =================

Private Sub BuildQuotes(ws As Worksheet, mode As String, dFrom As Date, dTo As Date, deptCode As String)
    Dim src As Worksheet, lastRow As Long, i As Long, r As Long
    Dim dp As String, cnt As Long, sTot As Double, nSaved As Long, nConv As Long, conv As Boolean

    Set src = ThisWorkbook.Sheets("QuoteLog")
    lastRow = src.Cells(src.Rows.Count, "A").End(xlUp).row
    r = 4

    If UCase$(mode) = "FULL" Then
        ColHead ws, r, Array("Quote No", "Date", "Dept", "Bill-To", "Total", "Status", "Conv Inv")
        r = r + 1
    End If

    For i = 2 To lastRow
        If Trim$(CStr(src.Cells(i, QL_NO).value)) = "" Then GoTo NextI
        If Not InRange(src.Cells(i, QL_DATE).value, dFrom, dTo) Then GoTo NextI

        dp = UCase$(Trim$(CStr(src.Cells(i, QL_DEPT).value)))
        If deptCode <> "ALL" And dp <> deptCode Then GoTo NextI

        cnt = cnt + 1
        sTot = sTot + Num(src.Cells(i, QL_TOTAL).value)
        conv = (Trim$(CStr(src.Cells(i, QL_CONVINV).value)) <> "")
        If conv Then nConv = nConv + 1 Else nSaved = nSaved + 1

        If UCase$(mode) = "FULL" Then
            ws.Cells(r, 1).value = src.Cells(i, QL_NO).value
            ws.Cells(r, 2).value = Format$(src.Cells(i, QL_DATE).value, "yyyy-mm-dd")
            ws.Cells(r, 3).value = dp
            ws.Cells(r, 4).value = CustIDToDrName(CStr(src.Cells(i, QL_CUST).value))
            ws.Cells(r, 5).value = Num(src.Cells(i, QL_TOTAL).value)
            ws.Cells(r, 5).NumberFormat = "#,##0.00"
            ws.Cells(r, 6).value = IIf(conv, "Converted", "Saved")
            ws.Cells(r, 7).value = src.Cells(i, QL_CONVINV).value
            r = r + 1
        End If
NextI:
    Next i

    If UCase$(mode) = "FULL" Then
        ws.Cells(r, 4).value = "TOTAL"
        ws.Cells(r, 4).Font.Bold = True
        ws.Cells(r, 5).value = sTot
        ws.Cells(r, 5).NumberFormat = "#,##0.00"
        ws.Cells(r, 5).Font.Bold = True
    Else
        SectionTitle ws, r, "QUOTES - SUMMARY": r = r + 1
        KeyVal ws, r, "Count", cnt: r = r + 1
        KeyVal ws, r, "Total Value", sTot: r = r + 1
        ws.Cells(r, 1).value = "Saved / Converted"
        ws.Cells(r, 2).value = nSaved & " / " & nConv
    End If
End Sub

' ================= MED CLAIMS =================

Private Sub BuildMedClaims(ws As Worksheet, mode As String, dFrom As Date, dTo As Date, deptCode As String)
    Dim src As Worksheet, lastRow As Long, i As Long, r As Long
    Dim dp As String, st As String, cnt As Long
    Dim sExcl As Double, sVat As Double, sTot As Double, sPaid As Double, sBal As Double
    Dim nSub As Long, nPd As Long, nU As Long, nOth As Long

    Set src = ThisWorkbook.Sheets("MedAidLog")
    lastRow = src.Cells(src.Rows.Count, "A").End(xlUp).row
    r = 4

    If UCase$(mode) = "FULL" Then
        ColHead ws, r, Array("MC No", "Date", "Dept", "Patient", "Doctor", "Excl", "VAT", "Total", "Paid", "Balance", "Status")
        r = r + 1
    End If

    For i = 2 To lastRow
        If Trim$(CStr(src.Cells(i, ML_NO).value)) = "" Then GoTo NextI
        If Not InRange(src.Cells(i, ML_DATE).value, dFrom, dTo) Then GoTo NextI

        dp = UCase$(Trim$(CStr(src.Cells(i, ML_DEPT).value)))
        If deptCode <> "ALL" Then
            If deptCode = "WA" Or deptCode = "WD" Then
                If dp <> deptCode Then GoTo NextI
            End If
        End If

        cnt = cnt + 1
        sExcl = sExcl + Num(src.Cells(i, ML_SUBTOTAL).value)
        sVat = sVat + Num(src.Cells(i, ML_VAT).value)
        sTot = sTot + Num(src.Cells(i, ML_TOTAL).value)
        sPaid = sPaid + Num(src.Cells(i, ML_PAID).value)
        sBal = sBal + Num(src.Cells(i, ML_BALANCE).value)

        st = UCase$(Trim$(CStr(src.Cells(i, ML_STATUS).value)))
        Select Case st
            Case "PAID": nPd = nPd + 1
            Case "SUBMITTED": nSub = nSub + 1
            Case "UNPAID", "": nU = nU + 1
            Case Else: nOth = nOth + 1
        End Select

        If UCase$(mode) = "FULL" Then
            ws.Cells(r, 1).value = src.Cells(i, ML_NO).value
            ws.Cells(r, 2).value = Format$(src.Cells(i, ML_DATE).value, "yyyy-mm-dd")
            ws.Cells(r, 3).value = dp
            ws.Cells(r, 4).value = src.Cells(i, ML_PATIENT).value
            ws.Cells(r, 5).value = src.Cells(i, ML_DOCTOR).value
            ws.Cells(r, 6).value = Num(src.Cells(i, ML_SUBTOTAL).value)
            ws.Cells(r, 7).value = Num(src.Cells(i, ML_VAT).value)
            ws.Cells(r, 8).value = Num(src.Cells(i, ML_TOTAL).value)
            ws.Cells(r, 9).value = Num(src.Cells(i, ML_PAID).value)
            ws.Cells(r, 10).value = Num(src.Cells(i, ML_BALANCE).value)
            ws.Cells(r, 11).value = src.Cells(i, ML_STATUS).value
            ws.Range(ws.Cells(r, 6), ws.Cells(r, 10)).NumberFormat = "#,##0.00"
            r = r + 1
        End If
NextI:
    Next i

    If UCase$(mode) = "FULL" Then
        ws.Cells(r, 5).value = "TOTALS"
        ws.Cells(r, 5).Font.Bold = True
        ws.Cells(r, 6).value = sExcl
        ws.Cells(r, 7).value = sVat
        ws.Cells(r, 8).value = sTot
        ws.Cells(r, 9).value = sPaid
        ws.Cells(r, 10).value = sBal
        ws.Range(ws.Cells(r, 6), ws.Cells(r, 10)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(r, 6), ws.Cells(r, 10)).Font.Bold = True
    Else
        SectionTitle ws, r, "MED CLAIMS - SUMMARY": r = r + 1
        KeyVal ws, r, "Count", cnt: r = r + 1
        KeyVal ws, r, "Subtotal (Excl)", sExcl: r = r + 1
        KeyVal ws, r, "VAT", sVat: r = r + 1
        KeyVal ws, r, "Total", sTot: r = r + 1
        KeyVal ws, r, "Paid", sPaid: r = r + 1
        KeyVal ws, r, "Outstanding", sBal: r = r + 1
        ws.Cells(r, 1).value = "Submitted / Paid / Unpaid / Other"
        ws.Cells(r, 2).value = nSub & " / " & nPd & " / " & nU & " / " & nOth
    End If
End Sub

' ================= CREDIT NOTES =================

Private Sub BuildCreditNotes(ws As Worksheet, mode As String, dFrom As Date, dTo As Date, deptCode As String)
    Dim src As Worksheet, lastRow As Long, i As Long, r As Long
    Dim cnt As Long, sTot As Double, sVat As Double

    Set src = ThisWorkbook.Sheets("CreditNotes")
    lastRow = src.Cells(src.Rows.Count, "A").End(xlUp).row
    r = 4

    If UCase$(mode) = "FULL" Then
        ColHead ws, r, Array("CN No", "Date", "Source Inv", "Total", "VAT", "Reason")
        r = r + 1
    End If

    For i = 2 To lastRow
        If Trim$(CStr(src.Cells(i, 1).value)) = "" Then GoTo NextI
        If Not InRange(src.Cells(i, 3).value, dFrom, dTo) Then GoTo NextI

        cnt = cnt + 1
        sTot = sTot + Num(src.Cells(i, 5).value)
        sVat = sVat + Num(src.Cells(i, 6).value)

        If UCase$(mode) = "FULL" Then
            ws.Cells(r, 1).value = src.Cells(i, 1).value
            ws.Cells(r, 2).value = Format$(src.Cells(i, 3).value, "yyyy-mm-dd")
            ws.Cells(r, 3).value = src.Cells(i, 2).value
            ws.Cells(r, 4).value = Num(src.Cells(i, 5).value)
            ws.Cells(r, 5).value = Num(src.Cells(i, 6).value)
            ws.Cells(r, 6).value = src.Cells(i, 7).value
            ws.Range(ws.Cells(r, 4), ws.Cells(r, 5)).NumberFormat = "#,##0.00"
            r = r + 1
        End If
NextI:
    Next i

    If UCase$(mode) = "FULL" Then
        ws.Cells(r, 3).value = "TOTALS"
        ws.Cells(r, 3).Font.Bold = True
        ws.Cells(r, 4).value = sTot
        ws.Cells(r, 5).value = sVat
        ws.Range(ws.Cells(r, 4), ws.Cells(r, 5)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(r, 4), ws.Cells(r, 5)).Font.Bold = True
    Else
        SectionTitle ws, r, "CREDIT NOTES - SUMMARY": r = r + 1
        KeyVal ws, r, "Count", cnt: r = r + 1
        KeyVal ws, r, "Total Credited", sTot: r = r + 1
        KeyVal ws, r, "VAT on Credits", sVat
    End If
End Sub

' ================= SUMMARY =================

Private Sub BuildSummary(ws As Worksheet, dFrom As Date, dTo As Date, deptCode As String)
    Dim r As Long, d As String, lines As Variant
    Dim iVat As Double, mVat As Double, cVat As Double

    d = deptCode
    r = 4

    lines = InvStats(dFrom, dTo, d)
    SectionTitle ws, r, "INVOICES": r = r + 1
    PutBlock ws, r, lines: r = NextFreeRow(ws)

    lines = QuoteStats(dFrom, dTo, d)
    SectionTitle ws, r, "QUOTES": r = r + 1
    PutBlock ws, r, lines: r = NextFreeRow(ws)

    lines = MCStats(dFrom, dTo, d)
    SectionTitle ws, r, "MED CLAIMS": r = r + 1
    PutBlock ws, r, lines: r = NextFreeRow(ws)

    lines = CNStats(dFrom, dTo)
    SectionTitle ws, r, "CREDIT NOTES": r = r + 1
    PutBlock ws, r, lines: r = NextFreeRow(ws)

    iVat = VATInvoiceTotal(dFrom, dTo, d)
    mVat = VATMedTotal(dFrom, dTo, d)
    If d = "ALL" Then cVat = VATCreditTotal(dFrom, dTo)

    SectionTitle ws, r, "VAT PAYABLE": r = r + 1
    ws.Cells(r, 1).value = "VAT PAYABLE (period)"
    ws.Cells(r, 2).value = iVat + mVat - IIf(d = "ALL", cVat, 0)
    ws.Cells(r, 2).NumberFormat = "#,##0.00"
    ws.Cells(r, 2).Font.Bold = True
End Sub

' ================= STATS =================

Private Function InvStats(dFrom As Date, dTo As Date, deptCode As String) As Variant
    Dim ws As Worksheet, lastRow As Long, i As Long, dp As String, st As String
    Dim cnt As Long, ex As Double, vt As Double, tt As Double, pd As Double, bl As Double
    Dim nU As Long, nP As Long, nPd As Long
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, IL_NO).value)) = "" Then GoTo NextI
        If Not InRange(ws.Cells(i, IL_DATE).value, dFrom, dTo) Then GoTo NextI
        dp = UCase$(Trim$(CStr(ws.Cells(i, IL_DEPT).value)))
        If deptCode <> "ALL" And dp <> deptCode Then GoTo NextI
        cnt = cnt + 1
        ex = ex + Num(ws.Cells(i, IL_SUBTOTAL).value)
        vt = vt + Num(ws.Cells(i, IL_VAT).value)
        tt = tt + Num(ws.Cells(i, IL_TOTAL).value)
        pd = pd + Num(ws.Cells(i, IL_PAID).value)
        bl = bl + Num(ws.Cells(i, IL_BALANCE).value)
        st = UCase$(Trim$(CStr(ws.Cells(i, IL_STATUS).value)))
        Select Case st
            Case "PAID": nPd = nPd + 1
            Case "PART-PAID": nP = nP + 1
            Case Else: nU = nU + 1
        End Select
NextI:
    Next i
    InvStats = Array("Count", cnt, "Excl", ex, "VAT", vt, "Total", tt, "Paid", pd, "Outstanding", bl, "Unpaid/Part/Paid", nU & " / " & nP & " / " & nPd)
End Function

Private Function QuoteStats(dFrom As Date, dTo As Date, deptCode As String) As Variant
    Dim ws As Worksheet, lastRow As Long, i As Long, dp As String, cnt As Long, tt As Double, sv As Long, cv As Long
    Set ws = ThisWorkbook.Sheets("QuoteLog")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, QL_NO).value)) = "" Then GoTo NextI
        If Not InRange(ws.Cells(i, QL_DATE).value, dFrom, dTo) Then GoTo NextI
        dp = UCase$(Trim$(CStr(ws.Cells(i, QL_DEPT).value)))
        If deptCode <> "ALL" And dp <> deptCode Then GoTo NextI
        cnt = cnt + 1
        tt = tt + Num(ws.Cells(i, QL_TOTAL).value)
        If Trim$(CStr(ws.Cells(i, QL_CONVINV).value)) <> "" Then cv = cv + 1 Else sv = sv + 1
NextI:
    Next i
    QuoteStats = Array("Count", cnt, "Total Value", tt, "Saved/Converted", sv & " / " & cv)
End Function

Private Function MCStats(dFrom As Date, dTo As Date, deptCode As String) As Variant
    Dim ws As Worksheet, lastRow As Long, i As Long, dp As String, st As String
    Dim cnt As Long, ex As Double, vt As Double, tt As Double, pd As Double, bl As Double
    Dim nSub As Long, nPd As Long, nU As Long, nOth As Long
    Set ws = ThisWorkbook.Sheets("MedAidLog")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, ML_NO).value)) = "" Then GoTo NextI
        If Not InRange(ws.Cells(i, ML_DATE).value, dFrom, dTo) Then GoTo NextI
        dp = UCase$(Trim$(CStr(ws.Cells(i, ML_DEPT).value)))
        If deptCode <> "ALL" Then
            If deptCode = "WA" Or deptCode = "WD" Then
                If dp <> deptCode Then GoTo NextI
            End If
        End If
        cnt = cnt + 1
        ex = ex + Num(ws.Cells(i, ML_SUBTOTAL).value)
        vt = vt + Num(ws.Cells(i, ML_VAT).value)
        tt = tt + Num(ws.Cells(i, ML_TOTAL).value)
        pd = pd + Num(ws.Cells(i, ML_PAID).value)
        bl = bl + Num(ws.Cells(i, ML_BALANCE).value)
        st = UCase$(Trim$(CStr(ws.Cells(i, ML_STATUS).value)))
        Select Case st
            Case "PAID": nPd = nPd + 1
            Case "SUBMITTED": nSub = nSub + 1
            Case "UNPAID", "": nU = nU + 1
            Case Else: nOth = nOth + 1
        End Select
NextI:
    Next i
    MCStats = Array("Count", cnt, "Excl", ex, "VAT", vt, "Total", tt, "Paid", pd, "Outstanding", bl, "Submitted/Paid/Unpaid/Other", nSub & " / " & nPd & " / " & nU & " / " & nOth)
End Function

Private Function CNStats(dFrom As Date, dTo As Date) As Variant
    Dim ws As Worksheet, lastRow As Long, i As Long, cnt As Long, tt As Double, vt As Double
    Set ws = ThisWorkbook.Sheets("CreditNotes")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, 1).value)) = "" Then GoTo NextI
        If Not InRange(ws.Cells(i, 3).value, dFrom, dTo) Then GoTo NextI
        cnt = cnt + 1
        tt = tt + Num(ws.Cells(i, 5).value)
        vt = vt + Num(ws.Cells(i, 6).value)
NextI:
    Next i
    CNStats = Array("Count", cnt, "Total Credited", tt, "VAT on Credits", vt)
End Function

Private Sub PutBlock(ws As Worksheet, startRow As Long, lines As Variant)
    Dim i As Long, r As Long
    r = startRow
    For i = LBound(lines) To UBound(lines) Step 2
        ws.Cells(r, 1).value = lines(i)
        ws.Cells(r, 2).value = lines(i + 1)
        If IsNumeric(lines(i + 1)) Then ws.Cells(r, 2).NumberFormat = "#,##0.00"
        r = r + 1
    Next i
End Sub

Private Function NextFreeRow(ws As Worksheet) As Long
    NextFreeRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 2
End Function

Private Function ExportReportPDF(ws As Worksheet, rptType As String, mode As String, dFrom As Date, dTo As Date, deptCode As String, folder As String) As String
    Dim fpath As String, fname As String
    If Right$(folder, 1) = "\" Or Right$(folder, 1) = "/" Then
        fpath = folder
    Else
        fpath = folder & Application.PathSeparator
    End If
    fname = UCase$(rptType) & "_" & UCase$(mode) & "_" & UCase$(deptCode) & "_" & Format$(dFrom, "yyyymmdd") & "-" & Format$(dTo, "yyyymmdd") & ".pdf"
    fpath = fpath & fname
    If UCase$(mode) = "FULL" Then ws.PageSetup.Orientation = xlLandscape Else ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fpath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    ExportReportPDF = fpath
End Function

' Locale-safe numeric read (blank/error -> 0). Local so this module is self-contained.
Private Function Num(v As Variant) As Double
    If IsError(v) Then Num = 0: Exit Function
    If Trim$(CStr(v)) = "" Then Num = 0: Exit Function
    If IsNumeric(v) Then Num = CDbl(v) Else Num = 0
End Function
' --- local shims: keep modReport self-contained ---

Private Function CustIDToDrName(ByVal custID As String) As String
    Dim wsC As Worksheet, last As Long, i As Long, key As String
    key = UCase$(Replace(Trim$(custID), " ", ""))
    If key = "" Then Exit Function
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets("Customers")
    If wsC Is Nothing Then CustIDToDrName = custID: Exit Function
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If UCase$(Replace(Trim$(CStr(wsC.Cells(i, 1).value)), " ", "")) = key Then
            CustIDToDrName = CStr(wsC.Cells(i, 2).value)
            Exit Function
        End If
    Next i
    CustIDToDrName = custID
End Function

Private Function StmtPickFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "Select folder to save report PDF"
    If fd.Show = -1 Then StmtPickFolder = fd.SelectedItems(1) Else StmtPickFolder = ""
End Function
