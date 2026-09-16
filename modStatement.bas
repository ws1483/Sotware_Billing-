Attribute VB_Name = "modStatement"
Option Explicit
' ===== Statement engine v3 + logging & recall =====================
' PATCHES:
'  - locale-safe numeric reads via Num() (comma decimals -> full cents)
'  - hardened ResetStatementSheet (never destroys data/log sheets)
'  - opening/aging/list balances are derived live from doc totals minus indexed payments
'  - undated payments are counted at the statement cutoff, logged to Debug.Print, and shown as blank-date '(undated)' ledger lines
'  - patient statements also fold MedAidLog claims into opening, ledger, aging, and balance lists
'  - unambiguous header dates (yyyy-mm-dd)
Private Const VAT_RATE As Double = 0.15
Private Const LINES_PER_PAGE As Long = 20
Private Const FIRST_LINE_ROW As Long = 13
Private Const TOTALS_ORIG_ROW As Long = 19
Private Const GAP_ROWS As Long = 1
Private Const FOOTER_LAST_ROW As Long = 36
Private Const TPL_SHEET As String = "StatementTpl"
Private Const OUT_SHEET As String = "Statement"
Private Const LOG_SHEET As String = "StatementLog"
Private mStmtInteractive As Boolean

Private Function NrmID(s As String) As String
    NrmID = UCase(Replace(Trim(s), " ", ""))
End Function

' Locale-safe numeric read (handles comma decimals; blank/error -> 0)
Private Function Num(v As Variant) As Double
    If IsError(v) Then Num = 0: Exit Function
    If Trim(CStr(v)) = "" Then Num = 0: Exit Function
    If IsNumeric(v) Then Num = CDbl(v) Else Num = 0
End Function

Private Function PaymentEffectiveDate(rawDate As Variant, fallbackDate As Date, ByRef wasUndated As Boolean) As Date
    If IsDate(rawDate) Then
        PaymentEffectiveDate = CDate(rawDate)
    Else
        PaymentEffectiveDate = fallbackDate
        wasUndated = True
    End If
End Function

Private Function BuildPaymentsIndex(wsPay As Worksheet, undatedAsOf As Date) As Object
    Dim idx As Object, entries As Collection, entry As Variant
    Dim lastP As Long, i As Long, docKey As String
    Dim payID As String, kind As String, pAmt As Double
    Dim effDate As Date, wasUndated As Boolean

    Set idx = CreateObject("Scripting.Dictionary")
    lastP = wsPay.Cells(wsPay.Rows.Count, 1).End(xlUp).row
    For i = 2 To lastP
        docKey = NrmID(CStr(wsPay.Cells(i, 2).Value))
        If docKey <> "" Then
            If Not idx.Exists(docKey) Then
                Set entries = New Collection
                idx.Add docKey, entries
            End If
            Set entries = idx(docKey)
            payID = CStr(wsPay.Cells(i, 1).Value)
            pAmt = Num(wsPay.Cells(i, 4).Value)
            kind = CStr(wsPay.Cells(i, 5).Value)
            wasUndated = False
            effDate = PaymentEffectiveDate(wsPay.Cells(i, 3).Value, undatedAsOf, wasUndated)
            If wasUndated Then
                Debug.Print "Undated payment treated as " & Format(undatedAsOf, "yyyy-mm-dd") & _
                            " for statement sums; Doc=" & CStr(wsPay.Cells(i, 2).Value) & _
                            " PayID=" & payID & " Amount=" & Format(pAmt, "0.00")
            End If
            entry = Array(CDbl(effDate), pAmt, kind, payID, wasUndated)
            entries.Add entry
        End If
    Next i
    Set BuildPaymentsIndex = idx
End Function

Private Function PaymentsAsOfIdx(idx As Object, docNo As String, cutoff As Date) As Double
    Dim entries As Collection, entry As Variant
    Dim docKey As String

    If idx Is Nothing Then Exit Function
    docKey = NrmID(docNo)
    If docKey = "" Then Exit Function
    If Not idx.Exists(docKey) Then Exit Function

    Set entries = idx(docKey)
    For Each entry In entries
        If CDbl(entry(0)) <= CDbl(cutoff) Then
            PaymentsAsOfIdx = PaymentsAsOfIdx + Num(entry(1))
        End If
    Next entry
End Function

Private Function PaymentsAsOf(wsPay As Worksheet, docNo As String, cutoff As Date) As Double
    Dim idx As Object
    Set idx = BuildPaymentsIndex(wsPay, cutoff)
    PaymentsAsOf = PaymentsAsOfIdx(idx, docNo, cutoff)
End Function

Private Function LiveOutstanding(docTotal As Double, idx As Object, docNo As String, cutoff As Date, _
                                 Optional floorAtZero As Boolean = False) As Double
    LiveOutstanding = docTotal - PaymentsAsOfIdx(idx, docNo, cutoff)
    If floorAtZero And LiveOutstanding < 0 Then LiveOutstanding = 0
End Function

Private Function PaymentLineLabel(wsPay As Worksheet, payRow As Long, wasUndated As Boolean) As String
    Dim labelText As String, payKind As String, docNo As String

    docNo = CStr(wsPay.Cells(payRow, 2).Value)
    payKind = UCase(Trim(CStr(wsPay.Cells(payRow, 5).Value)))
    If InStr(1, payKind, "CREDIT NOTE", vbTextCompare) > 0 Then
        labelText = "Credit Note - " & docNo
    Else
        labelText = "Payment - " & docNo
    End If
    If wasUndated Then labelText = labelText & " (undated)"
    PaymentLineLabel = labelText
End Function

Private Sub QueuePaymentEntry(wsPay As Worksheet, payRow As Long, dFrom As Date, dTo As Date, _
                              ByRef cnt As Long, ByRef rowsArr() As Long, ByRef dts() As Double, _
                              ByRef kinds() As String, ByRef undateds() As Boolean)
    Dim pDate As Date, wasUndated As Boolean

    wasUndated = False
    pDate = PaymentEffectiveDate(wsPay.Cells(payRow, 3).Value, dTo, wasUndated)
    If pDate >= dFrom And pDate <= dTo Then
        cnt = cnt + 1
        rowsArr(cnt) = payRow
        dts(cnt) = CDbl(pDate)
        kinds(cnt) = "PAY"
        undateds(cnt) = wasUndated
    End If
End Sub

Private Sub CheckAgingReconciliation(entityName As String, dFrom As Date, dTo As Date, _
                                     ageCur As Double, age30 As Double, age60 As Double, age90 As Double, _
                                     grossOutstanding As Double)
    Dim agingSum As Double, expectedGross As Double, msg As String

    agingSum = Round(ageCur + age30 + age60 + age90, 2)
    expectedGross = Round(grossOutstanding, 2)
    If Abs(agingSum - expectedGross) > 0.01 Then
        msg = "Statement aging mismatch for " & entityName & _
              " (" & Format(dFrom, "yyyy-mm-dd") & " to " & Format(dTo, "yyyy-mm-dd") & _
              "): Aging=" & Format(agingSum, "0.00") & _
              " GrossBeforeCredit=" & Format(expectedGross, "0.00")
        Debug.Print msg
        If mStmtInteractive Then MsgBox msg, vbExclamation, "Statement Reconciliation"
    End If
End Sub

Public Sub GenStatement()
    frmStatement.Show
End Sub

Public Function StmtPickFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select folder to save statement PDF(s)"
    If fd.Show = -1 Then StmtPickFolder = fd.SelectedItems(1) Else StmtPickFolder = ""
End Function

Private Function ResetStatementSheet() As Worksheet
    Dim tpl As Worksheet, ws As Worksheet, s As Worksheet

    ' 0. SAFETY: verify the template exists BEFORE deleting anything
    On Error Resume Next
    Set tpl = ThisWorkbook.Sheets(TPL_SHEET)
    On Error GoTo 0
    If tpl Is Nothing Then
        Err.Raise vbObjectError + 1, , _
            "Template sheet '" & TPL_SHEET & "' not found. Statement build aborted (no sheets deleted)."
    End If

    Application.DisplayAlerts = False

    ' 1. delete ONLY the output "Statement" and stray "StatementTpl (n)" copies.
    For Each s In ThisWorkbook.Worksheets
        If s.Name = OUT_SHEET Or (s.Name Like TPL_SHEET & " (*)") Then
            If s.Name <> TPL_SHEET Then
                On Error Resume Next
                s.Delete
                On Error GoTo 0
            End If
        End If
    Next s

    ' 2. copy the template to the end
    tpl.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    ' 3. rename the copy to "Statement" with clash guard
    On Error Resume Next
    ws.Name = OUT_SHEET
    On Error GoTo 0
    If ws.Name <> OUT_SHEET Then
        On Error Resume Next
        ThisWorkbook.Sheets(OUT_SHEET).Delete
        ws.Name = OUT_SHEET
        On Error GoTo 0
    End If

    ' 4. HARD STOP if the rename still failed
    If ws.Name <> OUT_SHEET Then
        Application.DisplayAlerts = True
        Err.Raise vbObjectError + 2, , _
            "Could not create the '" & OUT_SHEET & "' sheet (name clash). " & _
            "Delete any stray 'Statement' sheet and retry. No data sheets were modified."
    End If

    ws.Visible = xlSheetVisible
    Application.DisplayAlerts = True
    Set ResetStatementSheet = ws
End Function

' ============================ SINGLE (preview-first) =======================
Public Sub RunSingleStatement(drName As String, dFrom As Date, dTo As Date, dept As String)
    Dim ws As Worksheet, lastRow As Long, ans As VbMsgBoxResult, folder As String
    Dim custID As String, tInv As Double, tPaid As Double, cr As Double, bDue As Double
    On Error GoTo Fail
    mStmtInteractive = True
    Application.ScreenUpdating = False
    Set ws = BuildAndRender(drName, dFrom, dTo, dept, lastRow, _
                            custID, tInv, tPaid, cr, bDue)
    Application.ScreenUpdating = True
    mStmtInteractive = False
    If lastRow = 0 Then
        MsgBox "No transactions/outstanding invoices for " & drName & ".", vbInformation
        Exit Sub
    End If
    ws.Activate: ws.Range("A1").Select
    ans = MsgBox("Statement for " & drName & " is ready (on screen)." & vbCrLf & vbCrLf & _
                 "Create PDF now?", vbQuestion + vbYesNo, "Preview Statement")
    If ans = vbYes Then
        folder = StmtPickFolder()
        If folder <> "" Then
            Dim fpath As String
            fpath = ExportStatementPDF(ws, drName, dFrom, dTo, folder)
            LogStatement custID, drName, dept, dFrom, dTo, tInv, tPaid, cr, bDue, fpath
            MsgBox "Statement PDF saved and logged.", vbInformation
        End If
    End If
    Exit Sub
Fail:
    mStmtInteractive = False
    Application.ScreenUpdating = True
    MsgBox "RunSingleStatement error: " & Err.Description, vbExclamation
End Sub

' ============================ BATCH =======================================
Public Sub RunBatchStatements(dFrom As Date, dTo As Date, dept As String)
    Dim docs As Collection, i As Long, ws As Worksheet, lastRow As Long
    Dim drName As String, made As Long, folder As String
    Dim custID As String, tInv As Double, tPaid As Double, cr As Double, bDue As Double
    Dim fpath As String
    On Error GoTo Fail
    mStmtInteractive = False
    Set docs = StmtDoctorsWithBalance(dept, Date)
    If docs.Count = 0 Then MsgBox "No doctors with an outstanding balance.", vbInformation: Exit Sub
    If MsgBox(docs.Count & " doctor(s) have a balance. Generate a PDF for each?", _
              vbQuestion + vbYesNo, "Batch Statements") <> vbYes Then Exit Sub
    folder = StmtPickFolder()
    If folder = "" Then Exit Sub
    Application.ScreenUpdating = False
    made = 0
    For i = 1 To docs.Count
        drName = docs(i)
        Set ws = BuildAndRender(drName, dFrom, dTo, dept, lastRow, _
                                custID, tInv, tPaid, cr, bDue)
        If lastRow > 0 Then
            fpath = ExportStatementPDF(ws, drName, dFrom, dTo, folder)
            LogStatement custID, drName, dept, dFrom, dTo, tInv, tPaid, cr, bDue, fpath
            made = made + 1
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox made & " statement PDF(s) saved & logged to:" & vbCrLf & folder, vbInformation
    Exit Sub
Fail:
    mStmtInteractive = False
    Application.ScreenUpdating = True
    MsgBox "RunBatchStatements error: " & Err.Description, vbExclamation
End Sub

' ============================ LOGGING =====================================
Private Sub LogStatement(custID As String, drName As String, dept As String, _
                         dFrom As Date, dTo As Date, tInv As Double, tPaid As Double, _
                         cr As Double, bDue As Double, pdfPath As String)
    Dim wsL As Worksheet, r As Long, nextNo As Long, lastNo As String
    Set wsL = ThisWorkbook.Sheets(LOG_SHEET)
    r = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).row
    If r < 1 Then r = 1
    nextNo = 1
    If r >= 2 Then
        lastNo = CStr(wsL.Cells(r, "A").Value)
        If InStr(lastNo, "-") > 0 Then nextNo = Val(Mid(lastNo, InStr(lastNo, "-") + 1)) + 1
    End If
    r = r + 1
    wsL.Cells(r, "A").Value = "STMT-" & Format(nextNo, "0000")
    wsL.Cells(r, "B").Value = Now
    wsL.Cells(r, "B").NumberFormat = "yyyy/mm/dd hh:mm"
    wsL.Cells(r, "C").Value = custID
    wsL.Cells(r, "D").Value = drName
    wsL.Cells(r, "E").Value = dept
    wsL.Cells(r, "F").Value = dFrom
    wsL.Cells(r, "F").NumberFormat = "yyyy/mm/dd"
    wsL.Cells(r, "G").Value = dTo
    wsL.Cells(r, "G").NumberFormat = "yyyy/mm/dd"
    wsL.Cells(r, "H").Value = tInv
    wsL.Cells(r, "I").Value = tPaid
    wsL.Cells(r, "J").Value = cr
    wsL.Cells(r, "K").Value = bDue
    wsL.Cells(r, "L").Value = pdfPath
End Sub

' ============================ RECALL ======================================
Public Sub RecallStatement()
    Dim wsL As Worksheet, last As Long, i As Long, target As String, foundRow As Long
    Dim ans As VbMsgBoxResult, pdfPath As String
    Dim drName As String, dept As String, dFrom As Date, dTo As Date
    On Error GoTo Fail
    Set wsL = ThisWorkbook.Sheets(LOG_SHEET)

    target = InputBox("Enter the Statement No to recall (e.g. STMT-0001):", "Recall Statement")
    If Trim(target) = "" Then Exit Sub

    last = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).row
    foundRow = 0
    For i = 2 To last
        If NrmID(CStr(wsL.Cells(i, "A").Value)) = NrmID(target) Then foundRow = i: Exit For
    Next i
    If foundRow = 0 Then
        MsgBox "Statement '" & target & "' not found in the log.", vbExclamation
        Exit Sub
    End If

    drName = CStr(wsL.Cells(foundRow, "D").Value)
    dept = CStr(wsL.Cells(foundRow, "E").Value)
    dFrom = CDate(wsL.Cells(foundRow, "F").Value)
    dTo = CDate(wsL.Cells(foundRow, "G").Value)
    pdfPath = CStr(wsL.Cells(foundRow, "L").Value)

    ans = MsgBox("Found " & target & " for " & drName & "." & vbCrLf & vbCrLf & _
                 "YES = Re-open the saved PDF" & vbCrLf & _
                 "NO  = Re-generate the statement fresh" & vbCrLf & _
                 "CANCEL = Do nothing", vbQuestion + vbYesNoCancel, "Recall Statement")

    Select Case ans
        Case vbYes
            If pdfPath = "" Then
                MsgBox "No PDF path was logged for this statement.", vbExclamation
            ElseIf Dir(pdfPath) = "" Then
                MsgBox "The saved PDF no longer exists at:" & vbCrLf & pdfPath & vbCrLf & vbCrLf & _
                       "Use Re-generate instead.", vbExclamation
            Else
                ThisWorkbook.FollowHyperlink pdfPath
            End If
        Case vbNo
            If UCase(CStr(wsL.Cells(foundRow, "C").Value)) = "PATIENT" Then
                RunSinglePatientStatement drName, dFrom, dTo, dept
            Else
                RunSingleStatement drName, dFrom, dTo, dept
            End If
        Case vbCancel
            ' nothing
    End Select
    Exit Sub
Fail:
    MsgBox "RecallStatement error: " & Err.Description, vbExclamation
End Sub

' ============================ BUILD + RENDER ================================
Private Function BuildAndRender(drName As String, dFrom As Date, dTo As Date, _
                               dept As String, ByRef lastContentRow As Long, _
                               ByRef outCustID As String, ByRef outTotInv As Double, _
                               ByRef outTotPaid As Double, ByRef outCredit As Double, _
                               ByRef outBalDue As Double) As Worksheet
    Dim ws As Worksheet, wsLog As Worksheet, wsPay As Worksheet
    Dim payIdx As Object
    Dim custID As String, last As Long, i As Long, j As Long
    Dim opening As Double, storedCredit As Double
    Dim rowsArr() As Long, dts() As Double, kinds() As String, undateds() As Boolean, cnt As Long
    Dim tL As Long, tD As Double, tS As String, tB As Boolean
    Dim r As Long, running As Double
    Dim totalInv As Double, totalPaid As Double, openingMedInv As Double
    Dim ageCur As Double, age30 As Double, age60 As Double, age90 As Double
    Dim bal As Double, dueD As Date, days As Long
    Dim invDate As Date, invTotal As Double
    Dim lastP As Long, pInv As String, pAmt As Double, logRow As Long
    Dim totalLines As Long, availRows As Long, needRows As Long, off As Long
    Dim grossOut As Double, balDue As Double, excl As Double, vat As Double

    Set ws = ResetStatementSheet()
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    Set wsPay = ThisWorkbook.Sheets("Payments")
    Set payIdx = BuildPaymentsIndex(wsPay, dTo)
    custID = DrNameToCustIDp(drName)
    outCustID = custID
    storedCredit = DoctorCreditp(custID)
    RenderStatementHeader ws, drName, custID, dFrom, dTo, dept

    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
    lastP = wsPay.Cells(wsPay.Rows.Count, "A").End(xlUp).row
    ReDim rowsArr(1 To last + lastP + 10)
    ReDim dts(1 To last + lastP + 10)
    ReDim kinds(1 To last + lastP + 10)
    ReDim undateds(1 To last + lastP + 10)
    cnt = 0: opening = 0: openingMedInv = 0

    ' ---- invoices ----
    For i = 2 To last
        If NrmID(CStr(wsLog.Cells(i, 6).Value)) = NrmID(custID) _
           And LCase(Trim(CStr(wsLog.Cells(i, 3).Value))) = "doctor" _
           And DeptMatch(CStr(wsLog.Cells(i, 1).Value), dept) Then
            invDate = CDate(wsLog.Cells(i, 4).Value)
            invTotal = Num(wsLog.Cells(i, 12).Value)
            If invDate < dFrom Then
                opening = opening + LiveOutstanding(invTotal, payIdx, CStr(wsLog.Cells(i, 1).Value), dFrom - 1, True)
            ElseIf invDate <= dTo Then
                cnt = cnt + 1: rowsArr(cnt) = i: dts(cnt) = CDbl(invDate): kinds(cnt) = "INV"
            End If
        End If
    Next i

    ' ---- payments ----
    For i = 2 To lastP
        pInv = CStr(wsPay.Cells(i, 2).Value)
        logRow = FindLogRow(wsLog, pInv)
        If logRow > 0 Then
            If NrmID(CStr(wsLog.Cells(logRow, 6).Value)) = NrmID(custID) _
               And LCase(Trim(CStr(wsLog.Cells(logRow, 3).Value))) = "doctor" _
               And DeptMatch(pInv, dept) Then
                QueuePaymentEntry wsPay, i, dFrom, dTo, cnt, rowsArr, dts, kinds, undateds
            End If
        End If
    Next i

    If cnt = 0 And Abs(opening) < 0.005 And Abs(storedCredit) < 0.005 Then
        lastContentRow = 0: Set BuildAndRender = ws: Exit Function
    End If

    ' ---- sort oldest-first ----
    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If dts(j) > dts(j + 1) Then
                tD = dts(j): dts(j) = dts(j + 1): dts(j + 1) = tD
                tL = rowsArr(j): rowsArr(j) = rowsArr(j + 1): rowsArr(j + 1) = tL
                tS = kinds(j): kinds(j) = kinds(j + 1): kinds(j + 1) = tS
                tB = undateds(j): undateds(j) = undateds(j + 1): undateds(j + 1) = tB
            End If
        Next j
    Next i

    ' ==== SHIFT totals block DOWN first ====
    totalLines = 1 + cnt
    availRows = TOTALS_ORIG_ROW - FIRST_LINE_ROW
    needRows = totalLines + GAP_ROWS - availRows
    off = 0
    If needRows > 0 Then
        ws.Rows(TOTALS_ORIG_ROW & ":" & (TOTALS_ORIG_ROW + needRows - 1)).Insert Shift:=xlDown
        off = needRows
    End If

    ' ---- opening line ----
    r = FIRST_LINE_ROW
    running = opening
    ws.Cells(r, 1).Value = Format(dFrom, "dd/mm/yyyy")
    ws.Cells(r, 3).Value = "Balance Brought Forward"
    ws.Cells(r, 8).Value = running
    r = r + 1

    ' ---- ledger lines ----
    totalInv = 0: totalPaid = 0
    For i = 1 To cnt
        If kinds(i) = "INV" Then
            invTotal = Num(wsLog.Cells(rowsArr(i), 12).Value)
            running = running + invTotal
            totalInv = totalInv + invTotal
            ws.Cells(r, 1).Value = Format(CDate(wsLog.Cells(rowsArr(i), 4).Value), "dd/mm/yyyy")
            ws.Cells(r, 2).Value = wsLog.Cells(rowsArr(i), 1).Value
            ws.Cells(r, 3).Value = wsLog.Cells(rowsArr(i), 7).Value
            ws.Cells(r, 4).Value = wsLog.Cells(rowsArr(i), 8).Value
            ws.Cells(r, 5).Value = invTotal
            ws.Cells(r, 8).Value = running
        Else
            pAmt = Num(wsPay.Cells(rowsArr(i), 4).Value)
            running = running - pAmt
            totalPaid = totalPaid + pAmt
            If undateds(i) Then
                ws.Cells(r, 1).Value = ""
            Else
                ws.Cells(r, 1).Value = Format(CDate(wsPay.Cells(rowsArr(i), 3).Value), "dd/mm/yyyy")
            End If
            ws.Cells(r, 2).Value = wsPay.Cells(rowsArr(i), 1).Value
            ws.Cells(r, 3).Value = PaymentLineLabel(wsPay, rowsArr(i), undateds(i))
            ws.Cells(r, 6).Value = pAmt
            ws.Cells(r, 8).Value = running
        End If
        r = r + 1
    Next i

    ' ---- aging ----
    ageCur = 0: age30 = 0: age60 = 0: age90 = 0
    For i = 2 To last
        If NrmID(CStr(wsLog.Cells(i, 6).Value)) = NrmID(custID) _
           And LCase(Trim(CStr(wsLog.Cells(i, 3).Value))) = "doctor" _
           And DeptMatch(CStr(wsLog.Cells(i, 1).Value), dept) Then
            bal = LiveOutstanding(Num(wsLog.Cells(i, 12).Value), payIdx, CStr(wsLog.Cells(i, 1).Value), dTo)
            If bal > 0.005 Then
                If IsDate(wsLog.Cells(i, 5).Value) Then
                    dueD = CDate(wsLog.Cells(i, 5).Value)
                Else
                    dueD = CDate(wsLog.Cells(i, 4).Value)
                End If
                days = CLng(dTo - dueD)
                If days <= 0 Then
                    ageCur = ageCur + bal
                ElseIf days <= 30 Then
                    age30 = age30 + bal
                ElseIf days <= 60 Then
                    age60 = age60 + bal
                Else
                    age90 = age90 + bal
                End If
            End If
        End If
    Next i

    ' ---- totals ----
    totalInv = Round(totalInv, 2)
    totalPaid = Round(totalPaid, 2)
    ageCur = Round(ageCur, 2): age30 = Round(age30, 2)
    age60 = Round(age60, 2): age90 = Round(age90, 2)
    storedCredit = Round(storedCredit, 2)
    grossOut = Round(ageCur + age30 + age60 + age90, 2)
    balDue = Round(grossOut - storedCredit, 2)
    CheckAgingReconciliation drName, dFrom, dTo, ageCur, age30, age60, age90, grossOut
    If balDue >= 0 Then
        excl = Round(balDue / (1 + VAT_RATE), 2)
        vat = Round(balDue - excl, 2)
    Else
        ' Show statement credit entirely as VAT-exclusive; never present negative output VAT.
        excl = balDue
        vat = 0
    End If

    ws.Range("H" & (19 + off)).Value = totalInv
    ws.Range("H" & (20 + off)).Value = totalPaid
    ws.Range("H" & (21 + off)).Value = storedCredit
    ws.Range("H" & (22 + off)).Value = excl
    ws.Range("H" & (23 + off)).Value = vat
    ws.Range("H" & (24 + off)).Value = balDue
    ws.Range("G10").Value = balDue
    ws.Range("E" & (27 + off)).Value = ageCur
    ws.Range("F" & (27 + off)).Value = age30
    ws.Range("G" & (27 + off)).Value = age60
    ws.Range("H" & (27 + off)).Value = age90

    FormatStatement ws, r - 1, off

    outTotInv = totalInv
    outTotPaid = totalPaid
    outCredit = storedCredit
    outBalDue = balDue

    lastContentRow = FOOTER_LAST_ROW + off
    PaginateStatement ws, lastContentRow
    Set BuildAndRender = ws
End Function

' ============================ FORMATTING ===================================
Private Sub FormatStatement(ws As Worksheet, lastLineRow As Long, off As Long)
    ws.Range("G6").NumberFormat = "@"
    ws.Range("G6").Value = CStr(ws.Range("G6").Value)
    With ws.Range("E13:H" & lastLineRow)
        .NumberFormat = "R #,##0.00"
    End With
    ws.Columns("A").ColumnWidth = 13
    ws.Columns("B").ColumnWidth = 11.89
    ws.Columns("E").ColumnWidth = 12
    ws.Range("A13:A" & lastLineRow).HorizontalAlignment = xlLeft
    ws.Range("B13:B" & lastLineRow).HorizontalAlignment = xlLeft
    ws.Range("C13:C" & lastLineRow).HorizontalAlignment = xlLeft
    ws.Range("D13:D" & lastLineRow).HorizontalAlignment = xlCenter
    ws.Range("E13:E" & lastLineRow).HorizontalAlignment = xlRight
    ws.Range("F13:F" & lastLineRow).HorizontalAlignment = xlRight
    ws.Range("G13:G" & lastLineRow).HorizontalAlignment = xlRight
    ws.Range("H13:H" & lastLineRow).HorizontalAlignment = xlRight
End Sub

' ============================ HEADER =======================================
Private Sub RenderStatementHeader(ws As Worksheet, drName As String, custID As String, _
                                  dFrom As Date, dTo As Date, dept As String)
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    ws.Range("A3").Value = DeptStmtTitle(dept)
    ws.Range("B5").Value = drName
    ws.Range("G5").Value = Format(Date, "yyyy-mm-dd")
    ws.Range("G8").Value = Format(dFrom, "yyyy-mm-dd")
    ws.Range("G9").Value = Format(dTo, "yyyy-mm-dd")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsC.Cells(i, 1).Value)) = NrmID(custID) Then
            ws.Range("B6").Value = wsC.Cells(i, "C").Value
            ws.Range("B7").Value = wsC.Cells(i, "D").Value
            ws.Range("B8").Value = wsC.Cells(i, "E").Value
            ws.Range("B9").Value = wsC.Cells(i, "F").Value
            ws.Range("B10").Value = wsC.Cells(i, "G").Value
            ws.Range("B11").Value = wsC.Cells(i, "J").Value
            ws.Range("G6").NumberFormat = "@"
            ws.Range("G6").Value = CStr(wsC.Cells(i, "K").Value)
            ws.Range("G7").Value = wsC.Cells(i, "H").Value
            Exit For
        End If
    Next i
End Sub

' ============================ PAGINATION ===================================
Private Sub PaginateStatement(ws As Worksheet, lastContentRow As Long)
    Dim brk As Long
    ws.ResetAllPageBreaks
    ws.PageSetup.PrintTitleRows = "$1:$12"
    ws.PageSetup.PrintArea = "$A$1:$H$" & lastContentRow
    For brk = FIRST_LINE_ROW + LINES_PER_PAGE To lastContentRow Step LINES_PER_PAGE
        ws.HPageBreaks.Add Before:=ws.Rows(brk)
    Next brk
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
    End With
End Sub

' ============================ EXPORT =======================================
Private Function ExportStatementPDF(ws As Worksheet, drName As String, dFrom As Date, _
                                    dTo As Date, folder As String) As String
    Dim safeName As String, fpath As String
    safeName = Replace(Replace(Replace(drName, " ", "-"), "/", "-"), "\", "-")
    fpath = folder & Application.PathSeparator & "Statement_" & safeName & "_" & _
            Format(dFrom, "yyyy-mm-dd") & "_to_" & Format(dTo, "yyyy-mm-dd") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fpath, Quality:=xlQualityStandard
    ExportStatementPDF = fpath
End Function

' ============================ QUERIES / HELPERS ============================
Private Function FindLogRow(wsLog As Worksheet, invNo As String) As Long
    Dim last As Long, i As Long
    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsLog.Cells(i, 1).Value)) = NrmID(invNo) Then FindLogRow = i: Exit Function
    Next i
End Function

Private Function DeptMatch(invNo As String, dept As String) As Boolean
    If UCase(dept) = "ALL" Or dept = "" Then DeptMatch = True: Exit Function
    DeptMatch = (InStr(1, UCase(invNo), UCase(dept)) > 0)
End Function

Private Function FindMCRow(wsMC As Worksheet, docNo As String) As Long
    Dim last As Long, i As Long
    last = wsMC.Cells(wsMC.Rows.Count, ML_NO).End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsMC.Cells(i, ML_NO).Value)) = NrmID(docNo) Then
            FindMCRow = i
            Exit Function
        End If
    Next i
End Function

Private Function StmtDoctorsWithBalance(dept As String, cutoffDate As Date) As Collection
    Dim ws As Worksheet, wsPay As Worksheet, payIdx As Object
    Dim last As Long, i As Long, custID As String, docNo As String, key As String
    Dim seen As Object, totals As Object, order As New Collection, col As New Collection, drName As String
    Dim cutoff As Date

    cutoff = cutoffDate

    Set seen = CreateObject("Scripting.Dictionary")
    Set totals = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    Set wsPay = ThisWorkbook.Sheets("Payments")
    ' Batch selection uses an explicit live cutoff supplied by the caller.
    Set payIdx = BuildPaymentsIndex(wsPay, cutoff)

    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If LCase(Trim(CStr(ws.Cells(i, 3).Value))) = "doctor" _
           And DeptMatch(CStr(ws.Cells(i, 1).Value), dept) Then
            custID = CStr(ws.Cells(i, 6).Value)
            key = NrmID(custID)
            If key <> "" Then
                If Not seen.Exists(key) Then
                    seen.Add key, custID
                    totals.Add key, 0#
                    order.Add custID
                End If
                docNo = CStr(ws.Cells(i, 1).Value)
                totals(key) = CDbl(totals(key)) + LiveOutstanding(Num(ws.Cells(i, 12).Value), payIdx, docNo, cutoff, True)
            End If
        End If
    Next i

    For i = 1 To order.Count
        custID = CStr(order(i))
        key = NrmID(custID)
        If CDbl(totals(key)) - DoctorCreditp(custID) > 0.005 Then
            drName = CustIDToDrNamep(custID)
            If drName <> "" Then col.Add drName
        End If
    Next i
    Set StmtDoctorsWithBalance = col
End Function

Private Function DrNameToCustIDp(drName As String) As String
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "B").End(xlUp).row
    For i = 2 To last
        If CStr(wsC.Cells(i, "B").Value) = drName Then DrNameToCustIDp = CStr(wsC.Cells(i, "A").Value): Exit Function
    Next i
End Function

Private Function CustIDToDrNamep(custID As String) As String
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsC.Cells(i, "A").Value)) = NrmID(custID) Then CustIDToDrNamep = CStr(wsC.Cells(i, "B").Value): Exit Function
    Next i
End Function

Private Function DoctorCreditp(custID As String) As Double
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsC.Cells(i, "A").Value)) = NrmID(custID) Then DoctorCreditp = Num(wsC.Cells(i, "L").Value): Exit Function
    Next i
End Function

Private Function PatientCredit(patientName As String) As Double
    Dim wsP As Worksheet, last As Long, lastCol As Long, i As Long, creditCol As Long

    On Error Resume Next
    Set wsP = ThisWorkbook.Sheets("Patients")
    On Error GoTo 0
    If wsP Is Nothing Then Exit Function

    lastCol = wsP.Cells(1, wsP.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        Select Case UCase(Trim(CStr(wsP.Cells(1, i).Value)))
            Case "CREDIT", "PATIENT CREDIT"
                creditCol = i
                Exit For
        End Select
    Next i
    If creditCol = 0 Then Exit Function

    last = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsP.Cells(i, "A").Value)) = NrmID(patientName) _
           Or NrmID(CStr(wsP.Cells(i, "B").Value)) = NrmID(patientName) Then
            PatientCredit = Num(wsP.Cells(i, creditCol).Value)
            Exit Function
        End If
    Next i
End Function

' ============================================================================
' PATIENT STATEMENTS
' ============================================================================
Public Function PatientNamesList() As Collection
    Dim ws As Worksheet, wsMC As Worksheet, last As Long, lastMC As Long, i As Long, nm As String
    Dim seen As Object, col As New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If LCase(Trim(CStr(ws.Cells(i, 3).Value))) = "patient" Then
            nm = Trim(CStr(ws.Cells(i, 7).Value))
            If nm <> "" Then
                If Not seen.Exists(NrmID(nm)) Then
                    seen.Add NrmID(nm), 1
                    col.Add nm
                End If
            End If
        End If
    Next i
    On Error Resume Next
    Set wsMC = ThisWorkbook.Sheets("MedAidLog")
    On Error GoTo 0
    If Not wsMC Is Nothing Then
        lastMC = wsMC.Cells(wsMC.Rows.Count, ML_NO).End(xlUp).row
        For i = 2 To lastMC
            nm = Trim(CStr(wsMC.Cells(i, ML_PATIENT).Value))
            If nm <> "" Then
                If Not seen.Exists(NrmID(nm)) Then
                    seen.Add NrmID(nm), 1
                    col.Add nm
                End If
            End If
        Next i
    End If
    Set PatientNamesList = col
End Function

Private Function PatientsWithBalance(dept As String, cutoffDate As Date) As Collection
    Dim ws As Worksheet, wsMC As Worksheet, wsPay As Worksheet, payIdx As Object
    Dim last As Long, lastMC As Long, i As Long, nm As String, key As String, docNo As String
    Dim seen As Object, totals As Object, col As New Collection, order As New Collection
    Dim cutoff As Date

    cutoff = cutoffDate

    Set seen = CreateObject("Scripting.Dictionary")
    Set totals = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    Set wsPay = ThisWorkbook.Sheets("Payments")
    ' Batch selection uses an explicit live cutoff supplied by the caller.
    Set payIdx = BuildPaymentsIndex(wsPay, cutoff)
    On Error Resume Next
    Set wsMC = ThisWorkbook.Sheets("MedAidLog")
    On Error GoTo 0

    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If LCase(Trim(CStr(ws.Cells(i, 3).Value))) = "patient" _
           And DeptMatch(CStr(ws.Cells(i, 1).Value), dept) Then
            nm = Trim(CStr(ws.Cells(i, 7).Value))
            key = NrmID(nm)
            If key <> "" Then
                If Not seen.Exists(key) Then
                    seen.Add key, nm
                    totals.Add key, 0#
                    order.Add nm
                End If
                docNo = CStr(ws.Cells(i, 1).Value)
                totals(key) = CDbl(totals(key)) + LiveOutstanding(Num(ws.Cells(i, 12).Value), payIdx, docNo, cutoff, True)
            End If
        End If
    Next i

    If Not wsMC Is Nothing Then
        lastMC = wsMC.Cells(wsMC.Rows.Count, ML_NO).End(xlUp).row
        For i = 2 To lastMC
            If DeptMatch(CStr(wsMC.Cells(i, ML_NO).Value), dept) Then
                nm = Trim(CStr(wsMC.Cells(i, ML_PATIENT).Value))
                key = NrmID(nm)
                If key <> "" Then
                    If Not seen.Exists(key) Then
                        seen.Add key, nm
                        totals.Add key, 0#
                        order.Add nm
                    End If
                    docNo = CStr(wsMC.Cells(i, ML_NO).Value)
                    totals(key) = CDbl(totals(key)) + LiveOutstanding(Num(wsMC.Cells(i, ML_TOTAL).Value), payIdx, docNo, cutoff, True)
                End If
            End If
        Next i
    End If

    For i = 1 To order.Count
        nm = CStr(order(i))
        key = NrmID(nm)
        If CDbl(totals(key)) - PatientCredit(nm) > 0.005 Then col.Add nm
    Next i

    Set PatientsWithBalance = col
End Function

Public Sub RunSinglePatientStatement(patientName As String, dFrom As Date, dTo As Date, dept As String)
    Dim ws As Worksheet, lastRow As Long, ans As VbMsgBoxResult, folder As String
    Dim tInv As Double, tPaid As Double, bDue As Double
    On Error GoTo Fail
    mStmtInteractive = True
    Application.ScreenUpdating = False
    Set ws = BuildAndRenderPatient(patientName, dFrom, dTo, dept, lastRow, tInv, tPaid, bDue)
    Application.ScreenUpdating = True
    mStmtInteractive = False
    If lastRow = 0 Then
        MsgBox "No transactions/outstanding invoices for " & patientName & ".", vbInformation
        Exit Sub
    End If
    ws.Activate: ws.Range("A1").Select
    ans = MsgBox("Statement for " & patientName & " is ready (on screen)." & vbCrLf & vbCrLf & _
                 "Create PDF now?", vbQuestion + vbYesNo, "Preview Statement")
    If ans = vbYes Then
        folder = StmtPickFolder()
        If folder <> "" Then
            Dim fpath As String
            fpath = ExportStatementPDFp(ws, patientName, dFrom, dTo, folder)
            LogStatementPatient patientName, dept, dFrom, dTo, tInv, tPaid, bDue, fpath
            MsgBox "Statement PDF saved and logged.", vbInformation
        End If
    End If
    Exit Sub
Fail:
    mStmtInteractive = False
    Application.ScreenUpdating = True
    MsgBox "RunSinglePatientStatement error: " & Err.Description, vbExclamation
End Sub

Public Sub RunBatchPatientStatements(dFrom As Date, dTo As Date, dept As String)
    Dim pats As Collection, i As Long, ws As Worksheet, lastRow As Long
    Dim patientName As String, made As Long, folder As String
    Dim tInv As Double, tPaid As Double, bDue As Double, fpath As String
    On Error GoTo Fail
    mStmtInteractive = False
    Set pats = PatientsWithBalance(dept, Date)
    If pats.Count = 0 Then MsgBox "No patients with an outstanding balance.", vbInformation: Exit Sub
    If MsgBox(pats.Count & " patient(s) have a balance. Generate a PDF for each?", _
              vbQuestion + vbYesNo, "Batch Patient Statements") <> vbYes Then Exit Sub
    folder = StmtPickFolder()
    If folder = "" Then Exit Sub
    Application.ScreenUpdating = False
    made = 0
    For i = 1 To pats.Count
        patientName = pats(i)
        Set ws = BuildAndRenderPatient(patientName, dFrom, dTo, dept, lastRow, tInv, tPaid, bDue)
        If lastRow > 0 Then
            fpath = ExportStatementPDFp(ws, patientName, dFrom, dTo, folder)
            LogStatementPatient patientName, dept, dFrom, dTo, tInv, tPaid, bDue, fpath
            made = made + 1
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox made & " patient statement PDF(s) saved & logged to:" & vbCrLf & folder, vbInformation
    Exit Sub
Fail:
    mStmtInteractive = False
    Application.ScreenUpdating = True
    MsgBox "RunBatchPatientStatements error: " & Err.Description, vbExclamation
End Sub

Private Sub LogStatementPatient(patientName As String, dept As String, _
                                dFrom As Date, dTo As Date, tInv As Double, _
                                tPaid As Double, bDue As Double, pdfPath As String)
    Dim wsL As Worksheet, r As Long, nextNo As Long, lastNo As String
    Set wsL = ThisWorkbook.Sheets("StatementLog")
    r = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).row
    If r < 1 Then r = 1
    nextNo = 1
    If r >= 2 Then
        lastNo = CStr(wsL.Cells(r, "A").Value)
        If InStr(lastNo, "-") > 0 Then nextNo = Val(Mid(lastNo, InStr(lastNo, "-") + 1)) + 1
    End If
    r = r + 1
    wsL.Cells(r, "A").Value = "STMT-" & Format(nextNo, "0000")
    wsL.Cells(r, "B").Value = Now
    wsL.Cells(r, "B").NumberFormat = "yyyy/mm/dd hh:mm"
    wsL.Cells(r, "C").Value = "PATIENT"
    wsL.Cells(r, "D").Value = patientName
    wsL.Cells(r, "E").Value = dept
    wsL.Cells(r, "F").Value = dFrom: wsL.Cells(r, "F").NumberFormat = "yyyy/mm/dd"
    wsL.Cells(r, "G").Value = dTo:   wsL.Cells(r, "G").NumberFormat = "yyyy/mm/dd"
    wsL.Cells(r, "H").Value = tInv
    wsL.Cells(r, "I").Value = tPaid
    wsL.Cells(r, "J").Value = 0
    wsL.Cells(r, "K").Value = bDue
    wsL.Cells(r, "L").Value = pdfPath
End Sub

Private Function ExportStatementPDFp(ws As Worksheet, patientName As String, _
                                     dFrom As Date, dTo As Date, folder As String) As String
    Dim safeName As String, fpath As String
    safeName = Replace(Replace(Replace(patientName, " ", "-"), "/", "-"), "\", "-")
    fpath = folder & Application.PathSeparator & "Statement_" & safeName & "_" & _
            Format(dFrom, "yyyy-mm-dd") & "_to_" & Format(dTo, "yyyy-mm-dd") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fpath, Quality:=xlQualityStandard
    ExportStatementPDFp = fpath
End Function

Private Sub RenderPatientHeader(ws As Worksheet, patientName As String, _
                                dFrom As Date, dTo As Date, dept As String)
    Dim wsP As Worksheet, wsLog As Worksheet, wsMC As Worksheet, last As Long, i As Long
    Dim medAid As String, medNo As String

    ws.Range("A3").Value = DeptStmtTitle(dept)
    ws.Range("B5").Value = patientName
    ws.Range("B7").Value = patientName
    ws.Range("G5").Value = Format(Date, "yyyy-mm-dd")
    ws.Range("G8").Value = Format(dFrom, "yyyy-mm-dd")
    ws.Range("G9").Value = Format(dTo, "yyyy-mm-dd")

    On Error Resume Next
    Set wsP = ThisWorkbook.Sheets("Patients")
    On Error GoTo 0
    If Not wsP Is Nothing Then
        last = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).row
        For i = 2 To last
            If NrmID(CStr(wsP.Cells(i, "A").Value)) = NrmID(patientName) _
               Or NrmID(CStr(wsP.Cells(i, "B").Value)) = NrmID(patientName) Then
                ws.Range("B6").Value = wsP.Cells(i, "B").Value
                ws.Range("B8").Value = wsP.Cells(i, "C").Value
                ws.Range("B9").Value = wsP.Cells(i, "D").Value
                ws.Range("B10").Value = wsP.Cells(i, "E").Value
                ws.Range("B11").Value = wsP.Cells(i, "G").Value
                Exit For
            End If
        Next i
    End If

    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
    medAid = "": medNo = ""
    For i = 2 To last
        If LCase(Trim(CStr(wsLog.Cells(i, 3).Value))) = "patient" _
           And NrmID(CStr(wsLog.Cells(i, 7).Value)) = NrmID(patientName) Then
            medAid = CStr(wsLog.Cells(i, 20).Value)
            medNo = CStr(wsLog.Cells(i, 21).Value)
        End If
    Next i
    If medAid = "" And medNo = "" Then
        On Error Resume Next
        Set wsMC = ThisWorkbook.Sheets("MedAidLog")
        On Error GoTo 0
        If Not wsMC Is Nothing Then
            last = wsMC.Cells(wsMC.Rows.Count, ML_NO).End(xlUp).row
            For i = 2 To last
                If NrmID(CStr(wsMC.Cells(i, ML_PATIENT).Value)) = NrmID(patientName) Then
                    medAid = CStr(wsMC.Cells(i, ML_MEDAID).Value)
                    medNo = CStr(wsMC.Cells(i, ML_MEDNO).Value)
                    Exit For
                End If
            Next i
        End If
    End If

    ws.Range("F6").Value = "Medical Aid:"
    ws.Range("F7").Value = "Med Aid No:"
    ws.Range("G6").NumberFormat = "@"
    ws.Range("G6").Value = medAid
    ws.Range("G7").NumberFormat = "@"
    ws.Range("G7").Value = medNo
End Sub

Private Function BuildAndRenderPatient(patientName As String, dFrom As Date, dTo As Date, _
                                       dept As String, ByRef lastContentRow As Long, _
                                       ByRef outTotInv As Double, ByRef outTotPaid As Double, _
                                       ByRef outBalDue As Double) As Worksheet
    Dim ws As Worksheet, wsLog As Worksheet, wsPay As Worksheet, wsMC As Worksheet
    Dim payIdx As Object
    Dim last As Long, lastMC As Long, lastP As Long, i As Long, j As Long
    Dim opening As Double, storedCredit As Double
    Dim rowsArr() As Long, dts() As Double, kinds() As String, undateds() As Boolean, cnt As Long
    Dim tL As Long, tD As Double, tS As String, tB As Boolean
    Dim r As Long, running As Double
    Dim totalInv As Double, totalPaid As Double
    Dim ageCur As Double, age30 As Double, age60 As Double, age90 As Double
    Dim bal As Double, dueD As Date, days As Long
    Dim invDate As Date, invTotal As Double
    Dim pInv As String, pAmt As Double, logRow As Long, mcRow As Long
    Dim totalLines As Long, availRows As Long, needRows As Long, off As Long
    Dim grossOut As Double, balDue As Double, excl As Double, vat As Double

    Set ws = ResetStatementSheet()
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    Set wsPay = ThisWorkbook.Sheets("Payments")
    On Error Resume Next
    Set wsMC = ThisWorkbook.Sheets("MedAidLog")
    On Error GoTo 0
    Set payIdx = BuildPaymentsIndex(wsPay, dTo)
    storedCredit = PatientCredit(patientName)

    RenderPatientHeader ws, patientName, dFrom, dTo, dept

    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
    lastP = wsPay.Cells(wsPay.Rows.Count, "A").End(xlUp).row
    lastMC = 0
    If Not wsMC Is Nothing Then lastMC = wsMC.Cells(wsMC.Rows.Count, ML_NO).End(xlUp).row
    ReDim rowsArr(1 To last + lastP + lastMC + 10)
    ReDim dts(1 To last + lastP + lastMC + 10)
    ReDim kinds(1 To last + lastP + lastMC + 10)
    ReDim undateds(1 To last + lastP + lastMC + 10)
    cnt = 0: opening = 0

    ' ---- invoices (patient + name match) ----
    For i = 2 To last
        If LCase(Trim(CStr(wsLog.Cells(i, 3).Value))) = "patient" _
           And NrmID(CStr(wsLog.Cells(i, 7).Value)) = NrmID(patientName) _
           And DeptMatch(CStr(wsLog.Cells(i, 1).Value), dept) Then
            invDate = CDate(wsLog.Cells(i, 4).Value)
            invTotal = Num(wsLog.Cells(i, 12).Value)
            If invDate < dFrom Then
                opening = opening + LiveOutstanding(invTotal, payIdx, CStr(wsLog.Cells(i, 1).Value), dFrom - 1, True)
            ElseIf invDate <= dTo Then
                cnt = cnt + 1: rowsArr(cnt) = i: dts(cnt) = CDbl(invDate): kinds(cnt) = "INV"
            End If
        End If
    Next i

    ' ---- med claims ----
    If Not wsMC Is Nothing Then
        For i = 2 To lastMC
            If NrmID(CStr(wsMC.Cells(i, ML_PATIENT).Value)) = NrmID(patientName) _
               And DeptMatch(CStr(wsMC.Cells(i, ML_NO).Value), dept) Then
                invDate = CDate(wsMC.Cells(i, ML_DATE).Value)
                invTotal = Num(wsMC.Cells(i, ML_TOTAL).Value)
                If invDate < dFrom Then
                    bal = LiveOutstanding(invTotal, payIdx, CStr(wsMC.Cells(i, ML_NO).Value), dFrom - 1, True)
                    opening = opening + bal
                    openingMedInv = openingMedInv + bal
                ElseIf invDate <= dTo Then
                    cnt = cnt + 1: rowsArr(cnt) = i: dts(cnt) = CDbl(invDate): kinds(cnt) = "MC"
                End If
            End If
        Next i
    End If

    ' ---- payments ----
    For i = 2 To lastP
        pInv = CStr(wsPay.Cells(i, 2).Value)
        logRow = FindLogRow(wsLog, pInv)
        mcRow = 0
        If logRow = 0 And Not wsMC Is Nothing Then mcRow = FindMCRow(wsMC, pInv)
        If logRow > 0 Then
            If LCase(Trim(CStr(wsLog.Cells(logRow, 3).Value))) = "patient" _
               And NrmID(CStr(wsLog.Cells(logRow, 7).Value)) = NrmID(patientName) _
               And DeptMatch(pInv, dept) Then
                QueuePaymentEntry wsPay, i, dFrom, dTo, cnt, rowsArr, dts, kinds, undateds
            End If
        ElseIf mcRow > 0 Then
            If NrmID(CStr(wsMC.Cells(mcRow, ML_PATIENT).Value)) = NrmID(patientName) _
               And DeptMatch(pInv, dept) Then
                QueuePaymentEntry wsPay, i, dFrom, dTo, cnt, rowsArr, dts, kinds, undateds
            End If
        End If
    Next i

    If cnt = 0 And Abs(opening) < 0.005 And Abs(storedCredit) < 0.005 Then
        lastContentRow = 0: Set BuildAndRenderPatient = ws: Exit Function
    End If

    ' ---- sort oldest-first ----
    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If dts(j) > dts(j + 1) Then
                tD = dts(j): dts(j) = dts(j + 1): dts(j + 1) = tD
                tL = rowsArr(j): rowsArr(j) = rowsArr(j + 1): rowsArr(j + 1) = tL
                tS = kinds(j): kinds(j) = kinds(j + 1): kinds(j + 1) = tS
                tB = undateds(j): undateds(j) = undateds(j + 1): undateds(j + 1) = tB
            End If
        Next j
    Next i

    ' ==== shift totals block down ====
    totalLines = 1 + cnt
    availRows = TOTALS_ORIG_ROW - FIRST_LINE_ROW
    needRows = totalLines + GAP_ROWS - availRows
    off = 0
    If needRows > 0 Then
        ws.Rows(TOTALS_ORIG_ROW & ":" & (TOTALS_ORIG_ROW + needRows - 1)).Insert Shift:=xlDown
        off = needRows
    End If

    ' ---- opening line ----
    r = FIRST_LINE_ROW
    running = opening
    ws.Cells(r, 1).Value = Format(dFrom, "dd/mm/yyyy")
    ws.Cells(r, 3).Value = "Balance Brought Forward"
    ws.Cells(r, 8).Value = running
    r = r + 1

    ' ---- ledger lines ----
    totalInv = 0: totalPaid = 0
    For i = 1 To cnt
        Select Case kinds(i)
            Case "INV"
                invTotal = Num(wsLog.Cells(rowsArr(i), 12).Value)
                running = running + invTotal
                totalInv = totalInv + invTotal
                ws.Cells(r, 1).Value = Format(CDate(wsLog.Cells(rowsArr(i), 4).Value), "dd/mm/yyyy")
                ws.Cells(r, 2).Value = wsLog.Cells(rowsArr(i), 1).Value
                ws.Cells(r, 3).Value = wsLog.Cells(rowsArr(i), 7).Value
                ws.Cells(r, 4).Value = wsLog.Cells(rowsArr(i), 8).Value
                ws.Cells(r, 5).Value = invTotal
                ws.Cells(r, 8).Value = running
            Case "MC"
                invTotal = Num(wsMC.Cells(rowsArr(i), ML_TOTAL).Value)
                running = running + invTotal
                totalInv = totalInv + invTotal
                ws.Cells(r, 1).Value = Format(CDate(wsMC.Cells(rowsArr(i), ML_DATE).Value), "dd/mm/yyyy")
                ws.Cells(r, 2).Value = wsMC.Cells(rowsArr(i), ML_NO).Value
                ws.Cells(r, 3).Value = wsMC.Cells(rowsArr(i), ML_PATIENT).Value
                If Trim(CStr(wsMC.Cells(rowsArr(i), ML_APPLIANCE).Value)) <> "" Then
                    ws.Cells(r, 4).Value = "Med Claim - " & wsMC.Cells(rowsArr(i), ML_APPLIANCE).Value
                Else
                    ws.Cells(r, 4).Value = "Med Claim"
                End If
                ws.Cells(r, 5).Value = invTotal
                ws.Cells(r, 8).Value = running
            Case Else
                pAmt = Num(wsPay.Cells(rowsArr(i), 4).Value)
                running = running - pAmt
                totalPaid = totalPaid + pAmt
                If undateds(i) Then
                    ws.Cells(r, 1).Value = ""
                Else
                    ws.Cells(r, 1).Value = Format(CDate(wsPay.Cells(rowsArr(i), 3).Value), "dd/mm/yyyy")
                End If
                ws.Cells(r, 2).Value = wsPay.Cells(rowsArr(i), 1).Value
                ws.Cells(r, 3).Value = PaymentLineLabel(wsPay, rowsArr(i), undateds(i))
                ws.Cells(r, 6).Value = pAmt
                ws.Cells(r, 8).Value = running
        End Select
        r = r + 1
    Next i

    ' ---- aging ----
    ageCur = 0: age30 = 0: age60 = 0: age90 = 0
    For i = 2 To last
        If LCase(Trim(CStr(wsLog.Cells(i, 3).Value))) = "patient" _
           And NrmID(CStr(wsLog.Cells(i, 7).Value)) = NrmID(patientName) _
           And DeptMatch(CStr(wsLog.Cells(i, 1).Value), dept) Then
            bal = LiveOutstanding(Num(wsLog.Cells(i, 12).Value), payIdx, CStr(wsLog.Cells(i, 1).Value), dTo)
            If bal > 0.005 Then
                If IsDate(wsLog.Cells(i, 5).Value) Then
                    dueD = CDate(wsLog.Cells(i, 5).Value)
                Else
                    dueD = CDate(wsLog.Cells(i, 4).Value)
                End If
                days = CLng(dTo - dueD)
                If days <= 0 Then
                    ageCur = ageCur + bal
                ElseIf days <= 30 Then
                    age30 = age30 + bal
                ElseIf days <= 60 Then
                    age60 = age60 + bal
                Else
                    age90 = age90 + bal
                End If
            End If
        End If
    Next i
    If Not wsMC Is Nothing Then
        For i = 2 To lastMC
            If NrmID(CStr(wsMC.Cells(i, ML_PATIENT).Value)) = NrmID(patientName) _
               And DeptMatch(CStr(wsMC.Cells(i, ML_NO).Value), dept) Then
                bal = LiveOutstanding(Num(wsMC.Cells(i, ML_TOTAL).Value), payIdx, CStr(wsMC.Cells(i, ML_NO).Value), dTo)
                If bal > 0.005 Then
                    If IsDate(wsMC.Cells(i, ML_DUE).Value) Then
                        dueD = CDate(wsMC.Cells(i, ML_DUE).Value)
                    Else
                        dueD = CDate(wsMC.Cells(i, ML_DATE).Value)
                    End If
                    days = CLng(dTo - dueD)
                    If days <= 0 Then
                        ageCur = ageCur + bal
                    ElseIf days <= 30 Then
                        age30 = age30 + bal
                    ElseIf days <= 60 Then
                        age60 = age60 + bal
                    Else
                        age90 = age90 + bal
                    End If
                End If
            End If
        Next i
    End If

    ' ---- totals ----
    totalInv = totalInv + openingMedInv
    totalInv = Round(totalInv, 2): totalPaid = Round(totalPaid, 2)
    ageCur = Round(ageCur, 2): age30 = Round(age30, 2)
    age60 = Round(age60, 2): age90 = Round(age90, 2)
    storedCredit = Round(storedCredit, 2)
    grossOut = Round(ageCur + age30 + age60 + age90, 2)
    balDue = Round(grossOut - storedCredit, 2)
    CheckAgingReconciliation patientName, dFrom, dTo, ageCur, age30, age60, age90, grossOut
    If balDue >= 0 Then
        excl = Round(balDue / (1 + VAT_RATE), 2)
        vat = Round(balDue - excl, 2)
    Else
        ' Show statement credit entirely as VAT-exclusive; never present negative output VAT.
        excl = balDue
        vat = 0
    End If

    ws.Range("H" & (19 + off)).Value = totalInv
    ws.Range("H" & (20 + off)).Value = totalPaid
    ws.Range("H" & (21 + off)).Value = storedCredit
    ws.Range("H" & (22 + off)).Value = excl
    ws.Range("H" & (23 + off)).Value = vat
    ws.Range("H" & (24 + off)).Value = balDue
    ws.Range("G10").Value = balDue
    ws.Range("E" & (27 + off)).Value = ageCur
    ws.Range("F" & (27 + off)).Value = age30
    ws.Range("G" & (27 + off)).Value = age60
    ws.Range("H" & (27 + off)).Value = age90

    FormatStatementP ws, r - 1

    outTotInv = totalInv: outTotPaid = totalPaid: outBalDue = balDue
    lastContentRow = FOOTER_LAST_ROW + off
    PaginateStatementP ws, lastContentRow
    Set BuildAndRenderPatient = ws
End Function

Private Sub FormatStatementP(ws As Worksheet, lastLineRow As Long)
    With ws.Range("E13:H" & lastLineRow)
        .NumberFormat = "R #,##0.00"
    End With
    ws.Range("A13:A" & lastLineRow).HorizontalAlignment = xlLeft
    ws.Range("E13:H" & lastLineRow).HorizontalAlignment = xlRight
End Sub

Private Sub PaginateStatementP(ws As Worksheet, lastContentRow As Long)
    Dim brk As Long
    ws.ResetAllPageBreaks
    ws.PageSetup.PrintTitleRows = "$1:$12"
    ws.PageSetup.PrintArea = "$A$1:$H$" & lastContentRow
    For brk = 13 + 20 To lastContentRow Step 20
        ws.HPageBreaks.Add Before:=ws.Rows(brk)
    Next brk
    With ws.PageSetup
        .Orientation = xlPortrait: .Zoom = False: .FitToPagesWide = 1
    End With
End Sub

Private Function DeptStmtTitle(dept As String) As String
    Dim tWA As String, tWD As String
    On Error Resume Next
    tWA = CStr(ThisWorkbook.names("TitleWA").RefersToRange.Value)
    tWD = CStr(ThisWorkbook.names("TitleWD").RefersToRange.Value)
    On Error GoTo 0
    Select Case UCase(Trim(dept))
        Case "WA": DeptStmtTitle = tWA & " Statement"
        Case "WD": DeptStmtTitle = tWD & " Statement"
        Case Else: DeptStmtTitle = tWA & " / " & tWD & " Statement"
    End Select
End Function
