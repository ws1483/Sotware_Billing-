Attribute VB_Name = "modStatement"
Option Explicit
' ===== Statement engine v3 + logging & recall =====================
' PATCHES:
'  - locale-safe numeric reads via Num()
'  - hardened ResetStatementSheet
'  - OPENING BALANCE fix (pre-range outstanding -> opening, no double-count)
'  - unambiguous header dates (yyyy-mm-dd)
' PHASE 3B: patient statements fold in MED CLAIMS (MedAidLog) by patient name,
'           header autofills from "Med Customers"; A5="Main Member", A6="Patient Name".
Private Const VAT_RATE As Double = 0.15
Private Const LINES_PER_PAGE As Long = 20
Private Const FIRST_LINE_ROW As Long = 13
Private Const TOTALS_ORIG_ROW As Long = 19
Private Const GAP_ROWS As Long = 1
Private Const FOOTER_LAST_ROW As Long = 36
Private Const TPL_SHEET As String = "StatementTpl"
Private Const OUT_SHEET As String = "Statement"
Private Const LOG_SHEET As String = "StatementLog"

Private Function NrmID(s As String) As String
    NrmID = UCase(Replace(Trim(s), " ", ""))
End Function

' Locale-safe numeric read
Private Function Num(v As Variant) As Double
    If IsError(v) Then Num = 0: Exit Function
    If Trim(CStr(v)) = "" Then Num = 0: Exit Function
    If IsNumeric(v) Then Num = CDbl(v) Else Num = 0
End Function

Public Sub GenStatement()
    frmStatement.Show
End Sub

Public Function StmtPickFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "Select folder to save statement PDF(s)"
    If fd.Show = -1 Then StmtPickFolder = fd.SelectedItems(1) Else StmtPickFolder = ""
End Function

Private Function ResetStatementSheet() As Worksheet
    Dim tpl As Worksheet, ws As Worksheet, s As Worksheet
    On Error Resume Next
    Set tpl = ThisWorkbook.Sheets(TPL_SHEET)
    On Error GoTo 0
    If tpl Is Nothing Then
        Err.Raise vbObjectError + 1, , _
            "Template sheet '" & TPL_SHEET & "' not found. Statement build aborted (no sheets deleted)."
    End If

    Application.DisplayAlerts = False
    For Each s In ThisWorkbook.Worksheets
        If s.Name = OUT_SHEET Or (s.Name Like TPL_SHEET & " (*)") Then
            If s.Name <> TPL_SHEET Then
                On Error Resume Next
                s.Delete
                On Error GoTo 0
            End If
        End If
    Next s

    tpl.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    On Error Resume Next
    ws.Name = OUT_SHEET
    On Error GoTo 0
    If ws.Name <> OUT_SHEET Then
        On Error Resume Next
        ThisWorkbook.Sheets(OUT_SHEET).Delete
        ws.Name = OUT_SHEET
        On Error GoTo 0
    End If

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

' ============================ SINGLE (doctor) ==============================
Public Sub RunSingleStatement(drName As String, dFrom As Date, dTo As Date, dept As String)
    Dim ws As Worksheet, lastRow As Long, ans As VbMsgBoxResult, folder As String
    Dim custID As String, tInv As Double, tPaid As Double, cr As Double, bDue As Double
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set ws = BuildAndRender(drName, dFrom, dTo, dept, lastRow, custID, tInv, tPaid, cr, bDue)
    Application.ScreenUpdating = True
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
    Application.ScreenUpdating = True
    MsgBox "RunSingleStatement error: " & Err.Description, vbExclamation
End Sub

' ============================ BATCH (doctor) ===============================
Public Sub RunBatchStatements(dFrom As Date, dTo As Date, dept As String)
    Dim docs As Collection, i As Long, ws As Worksheet, lastRow As Long
    Dim drName As String, made As Long, folder As String
    Dim custID As String, tInv As Double, tPaid As Double, cr As Double, bDue As Double
    Dim fpath As String
    On Error GoTo Fail
    Set docs = StmtDoctorsWithBalance(dept)
    If docs.Count = 0 Then MsgBox "No doctors with an outstanding balance.", vbInformation: Exit Sub
    If MsgBox(docs.Count & " doctor(s) have a balance. Generate a PDF for each?", _
              vbQuestion + vbYesNo, "Batch Statements") <> vbYes Then Exit Sub
    folder = StmtPickFolder()
    If folder = "" Then Exit Sub
    Application.ScreenUpdating = False
    made = 0
    For i = 1 To docs.Count
        drName = docs(i)
        Set ws = BuildAndRender(drName, dFrom, dTo, dept, lastRow, custID, tInv, tPaid, cr, bDue)
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
        lastNo = CStr(wsL.Cells(r, "A").value)
        If InStr(lastNo, "-") > 0 Then nextNo = Val(Mid(lastNo, InStr(lastNo, "-") + 1)) + 1
    End If
    r = r + 1
    wsL.Cells(r, "A").value = "STMT-" & Format(nextNo, "0000")
    wsL.Cells(r, "B").value = Now
    wsL.Cells(r, "B").NumberFormat = "yyyy/mm/dd hh:mm"
    wsL.Cells(r, "C").value = custID
    wsL.Cells(r, "D").value = drName
    wsL.Cells(r, "E").value = dept
    wsL.Cells(r, "F").value = dFrom
    wsL.Cells(r, "F").NumberFormat = "yyyy/mm/dd"
    wsL.Cells(r, "G").value = dTo
    wsL.Cells(r, "G").NumberFormat = "yyyy/mm/dd"
    wsL.Cells(r, "H").value = tInv
    wsL.Cells(r, "I").value = tPaid
    wsL.Cells(r, "J").value = cr
    wsL.Cells(r, "K").value = bDue
    wsL.Cells(r, "L").value = pdfPath
End Sub

' ============================ RECALL ======================================
Public Sub RecallStatement()
    Dim wsL As Worksheet, last As Long, i As Long, Target As String, foundRow As Long
    Dim ans As VbMsgBoxResult, pdfPath As String
    Dim drName As String, dept As String, dFrom As Date, dTo As Date
    On Error GoTo Fail
    Set wsL = ThisWorkbook.Sheets(LOG_SHEET)

    Target = InputBox("Enter the Statement No to recall (e.g. STMT-0001):", "Recall Statement")
    If Trim(Target) = "" Then Exit Sub

    last = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).row
    foundRow = 0
    For i = 2 To last
        If NrmID(CStr(wsL.Cells(i, "A").value)) = NrmID(Target) Then foundRow = i: Exit For
    Next i
    If foundRow = 0 Then
        MsgBox "Statement '" & Target & "' not found in the log.", vbExclamation
        Exit Sub
    End If

    drName = CStr(wsL.Cells(foundRow, "D").value)
    dept = CStr(wsL.Cells(foundRow, "E").value)
    dFrom = CDate(wsL.Cells(foundRow, "F").value)
    dTo = CDate(wsL.Cells(foundRow, "G").value)
    pdfPath = CStr(wsL.Cells(foundRow, "L").value)

    ans = MsgBox("Found " & Target & " for " & drName & "." & vbCrLf & vbCrLf & _
                 "YES = Re-open the saved PDF" & vbCrLf & _
                 "NO  = Re-generate the statement fresh" & vbCrLf & _
                 "CANCEL = Do nothing", vbQuestion + vbYesNoCancel, "Recall Statement")

    Select Case ans
        Case vbYes
            If pdfPath = "" Then
                MsgBox "No PDF path was logged for this statement.", vbExclamation
            ElseIf dir(pdfPath) = "" Then
                MsgBox "The saved PDF no longer exists at:" & vbCrLf & pdfPath & vbCrLf & vbCrLf & _
                       "Use Re-generate instead.", vbExclamation
            Else
                ThisWorkbook.FollowHyperlink pdfPath
            End If
        Case vbNo
            If UCase(CStr(wsL.Cells(foundRow, "C").value)) = "PATIENT" Then
                RunSinglePatientStatement drName, dFrom, dTo, dept
            Else
                RunSingleStatement drName, dFrom, dTo, dept
            End If
        Case vbCancel
    End Select
    Exit Sub
Fail:
    MsgBox "RecallStatement error: " & Err.Description, vbExclamation
End Sub

' ============================ BUILD + RENDER (doctor) ======================
Private Function BuildAndRender(drName As String, dFrom As Date, dTo As Date, _
                               dept As String, ByRef lastContentRow As Long, _
                               ByRef outCustID As String, ByRef outTotInv As Double, _
                               ByRef outTotPaid As Double, ByRef outCredit As Double, _
                               ByRef outBalDue As Double) As Worksheet
    Dim ws As Worksheet, wsLog As Worksheet, wsPay As Worksheet
    Dim custID As String, last As Long, i As Long, j As Long
    Dim opening As Double, credit As Double
    Dim rowsArr() As Long, dts() As Double, kinds() As String, cnt As Long
    Dim tL As Long, tD As Double, tS As String
    Dim r As Long, running As Double
    Dim totalInv As Double, totalPaid As Double
    Dim ageCur As Double, age30 As Double, age60 As Double, age90 As Double
    Dim bal As Double, dueD As Date, days As Long
    Dim invDate As Date, invTotal As Double, invBal As Double
    Dim lastP As Long, pInv As String, pDate As Date, pAmt As Double, logRow As Long

    Set ws = ResetStatementSheet()
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    Set wsPay = ThisWorkbook.Sheets("Payments")
    custID = DrNameToCustIDp(drName)
    outCustID = custID
    credit = DoctorCreditp(custID)
    RenderStatementHeader ws, drName, custID, dFrom, dTo, dept

    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
    ReDim rowsArr(1 To (last + 50) * 2)
    ReDim dts(1 To (last + 50) * 2)
    ReDim kinds(1 To (last + 50) * 2)
    cnt = 0: opening = 0

    ' ---- invoices ----
    For i = 2 To last
        If NrmID(CStr(wsLog.Cells(i, 6).value)) = NrmID(custID) _
           And UCase(CStr(wsLog.Cells(i, 3).value)) = "DOCTOR" _
           And DeptMatch(CStr(wsLog.Cells(i, 1).value), dept) Then
            invDate = CDate(wsLog.Cells(i, 4).value)
            invBal = Num(wsLog.Cells(i, 15).value)
            If invDate < dFrom Then
                opening = opening + invBal
            ElseIf invDate <= dTo Then
                cnt = cnt + 1: rowsArr(cnt) = i: dts(cnt) = CDbl(invDate): kinds(cnt) = "INV"
            End If
        End If
    Next i

    ' ---- payments ----
    lastP = wsPay.Cells(wsPay.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastP
        pInv = CStr(wsPay.Cells(i, 2).value)
        logRow = FindLogRow(wsLog, pInv)
        If logRow > 0 Then
            If NrmID(CStr(wsLog.Cells(logRow, 6).value)) = NrmID(custID) _
               And UCase(CStr(wsLog.Cells(logRow, 3).value)) = "DOCTOR" _
               And DeptMatch(pInv, dept) Then
                pDate = CDate(wsPay.Cells(i, 3).value)
                If pDate >= dFrom And pDate <= dTo Then
                    cnt = cnt + 1: rowsArr(cnt) = i: dts(cnt) = CDbl(pDate): kinds(cnt) = "PAY"
                End If
            End If
        End If
    Next i

    If cnt = 0 And Abs(opening) < 0.005 And credit < 0.005 Then
        lastContentRow = 0: Set BuildAndRender = ws: Exit Function
    End If

    ' ---- sort oldest-first ----
    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If dts(j) > dts(j + 1) Then
                tD = dts(j): dts(j) = dts(j + 1): dts(j + 1) = tD
                tL = rowsArr(j): rowsArr(j) = rowsArr(j + 1): rowsArr(j + 1) = tL
                tS = kinds(j): kinds(j) = kinds(j + 1): kinds(j + 1) = tS
            End If
        Next j
    Next i

    Dim totalLines As Long, availRows As Long, needRows As Long, off As Long
    totalLines = 1 + cnt
    availRows = TOTALS_ORIG_ROW - FIRST_LINE_ROW
    needRows = totalLines + GAP_ROWS - availRows
    off = 0
    If needRows > 0 Then
        ws.Rows(TOTALS_ORIG_ROW & ":" & (TOTALS_ORIG_ROW + needRows - 1)).Insert Shift:=xlDown
        off = needRows
    End If

    r = FIRST_LINE_ROW
    running = opening
    ws.Cells(r, 1).value = Format(dFrom, "dd/mm/yyyy")
    ws.Cells(r, 3).value = "Balance Brought Forward"
    ws.Cells(r, 8).value = running
    r = r + 1

    totalInv = 0: totalPaid = 0
    For i = 1 To cnt
        If kinds(i) = "INV" Then
            invTotal = Num(wsLog.Cells(rowsArr(i), 12).value)
            running = running + invTotal
            totalInv = totalInv + invTotal
            ws.Cells(r, 1).value = Format(CDate(wsLog.Cells(rowsArr(i), 4).value), "dd/mm/yyyy")
            ws.Cells(r, 2).value = wsLog.Cells(rowsArr(i), 1).value
            ws.Cells(r, 3).value = wsLog.Cells(rowsArr(i), 7).value
            ws.Cells(r, 4).value = wsLog.Cells(rowsArr(i), 8).value
            ws.Cells(r, 5).value = invTotal
            ws.Cells(r, 8).value = running
        Else
            pAmt = Num(wsPay.Cells(rowsArr(i), 4).value)
            running = running - pAmt
            totalPaid = totalPaid + pAmt
            ws.Cells(r, 1).value = Format(CDate(wsPay.Cells(rowsArr(i), 3).value), "dd/mm/yyyy")
            ws.Cells(r, 2).value = wsPay.Cells(rowsArr(i), 1).value
            ws.Cells(r, 3).value = "Payment - " & wsPay.Cells(rowsArr(i), 2).value
            ws.Cells(r, 6).value = pAmt
            ws.Cells(r, 8).value = running
        End If
        r = r + 1
    Next i

    ' ---- aging ----
    ageCur = 0: age30 = 0: age60 = 0: age90 = 0
    For i = 2 To last
        If NrmID(CStr(wsLog.Cells(i, 6).value)) = NrmID(custID) _
           And UCase(CStr(wsLog.Cells(i, 3).value)) = "DOCTOR" _
           And DeptMatch(CStr(wsLog.Cells(i, 1).value), dept) Then
            bal = Num(wsLog.Cells(i, 15).value)
            If bal > 0.005 Then
                If IsDate(wsLog.Cells(i, 5).value) Then
                    dueD = CDate(wsLog.Cells(i, 5).value)
                Else
                    dueD = CDate(wsLog.Cells(i, 4).value)
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

    Dim grossOut As Double, balDue As Double, excl As Double, vat As Double
    grossOut = running: If grossOut < 0 Then grossOut = 0
    credit = Round(credit, 2)
    balDue = Round(grossOut - credit, 2): If balDue < 0 Then balDue = 0
    excl = Round(balDue / (1 + VAT_RATE), 2)
    vat = Round(balDue - excl, 2)
    totalInv = Round(totalInv, 2)
    totalPaid = Round(totalPaid, 2)
    ageCur = Round(ageCur, 2): age30 = Round(age30, 2)
    age60 = Round(age60, 2): age90 = Round(age90, 2)

    ws.Range("H" & (19 + off)).value = totalInv
    ws.Range("H" & (20 + off)).value = totalPaid
    ws.Range("H" & (21 + off)).value = credit
    ws.Range("H" & (22 + off)).value = excl
    ws.Range("H" & (23 + off)).value = vat
    ws.Range("H" & (24 + off)).value = balDue
    ws.Range("G10").value = balDue
    ws.Range("E" & (27 + off)).value = ageCur
    ws.Range("F" & (27 + off)).value = age30
    ws.Range("G" & (27 + off)).value = age60
    ws.Range("H" & (27 + off)).value = age90

    FormatStatement ws, r - 1, off

    outTotInv = totalInv
    outTotPaid = totalPaid
    outCredit = credit
    outBalDue = balDue

    lastContentRow = FOOTER_LAST_ROW + off
    PaginateStatement ws, lastContentRow
    Set BuildAndRender = ws
End Function

' ============================ FORMATTING (doctor) =========================
Private Sub FormatStatement(ws As Worksheet, lastLineRow As Long, off As Long)
    ws.Range("G6").NumberFormat = "@"
    ws.Range("G6").value = CStr(ws.Range("G6").value)
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

' ============================ HEADER (doctor) =============================
Private Sub RenderStatementHeader(ws As Worksheet, drName As String, custID As String, _
                                  dFrom As Date, dTo As Date, dept As String)
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    ws.Range("A3").value = DeptStmtTitle(dept)
    ' doctor mode keeps original labels
    ws.Range("A5").value = "Customer"
    ws.Range("B5").value = drName
    ws.Range("G5").value = Format(Date, "yyyy-mm-dd")
    ws.Range("G8").value = Format(dFrom, "yyyy-mm-dd")
    ws.Range("G9").value = Format(dTo, "yyyy-mm-dd")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsC.Cells(i, 1).value)) = NrmID(custID) Then
            ws.Range("B6").value = wsC.Cells(i, "C").value
            ws.Range("B7").value = wsC.Cells(i, "D").value
            ws.Range("B8").value = wsC.Cells(i, "E").value
            ws.Range("B9").value = wsC.Cells(i, "F").value
            ws.Range("B10").value = wsC.Cells(i, "G").value
            ws.Range("B11").value = wsC.Cells(i, "J").value
            ws.Range("G6").NumberFormat = "@"
            ws.Range("G6").value = CStr(wsC.Cells(i, "K").value)
            ws.Range("G7").value = wsC.Cells(i, "H").value
            Exit For
        End If
    Next i
End Sub

' ============================ PAGINATION (doctor) =========================
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

' ============================ EXPORT (doctor) =============================
Private Function ExportStatementPDF(ws As Worksheet, drName As String, dFrom As Date, _
                                    dTo As Date, folder As String) As String
    Dim safeName As String, fpath As String
    safeName = Replace(Replace(Replace(drName, " ", "-"), "/", "-"), "\", "-")
    fpath = folder & Application.PathSeparator & "Statement_" & safeName & "_" & _
            Format(dFrom, "yyyy-mm-dd") & "_to_" & Format(dTo, "yyyy-mm-dd") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fpath, Quality:=xlQualityStandard
    ExportStatementPDF = fpath
End Function

' ============================ QUERIES / HELPERS ===========================
Private Function FindLogRow(wsLog As Worksheet, invNo As String) As Long
    Dim last As Long, i As Long
    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsLog.Cells(i, 1).value)) = NrmID(invNo) Then FindLogRow = i: Exit Function
    Next i
End Function

' Find a MedAidLog row by MC number (col A); 0 if not found
Private Function FindMCRow(wsMC As Worksheet, mcNo As String) As Long
    Dim last As Long, i As Long
    last = wsMC.Cells(wsMC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsMC.Cells(i, 1).value)) = NrmID(mcNo) Then FindMCRow = i: Exit Function
    Next i
End Function

Private Function DeptMatch(invNo As String, dept As String) As Boolean
    If UCase(dept) = "ALL" Or dept = "" Then DeptMatch = True: Exit Function
    DeptMatch = (InStr(1, UCase(invNo), UCase(dept)) > 0)
End Function

Private Function StmtDoctorsWithBalance(dept As String) As Collection
    Dim ws As Worksheet, last As Long, i As Long, custID As String
    Dim seen As Object, col As New Collection, drName As String
    Set seen = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If UCase(CStr(ws.Cells(i, 3).value)) = "DOCTOR" _
           And DeptMatch(CStr(ws.Cells(i, 1).value), dept) Then
            If Num(ws.Cells(i, 15).value) > 0.005 Then
                custID = CStr(ws.Cells(i, 6).value)
                If Not seen.Exists(NrmID(custID)) Then
                    seen.Add NrmID(custID), 1
                    drName = CustIDToDrNamep(custID)
                    If drName <> "" Then col.Add drName
                End If
            End If
        End If
    Next i
    Set StmtDoctorsWithBalance = col
End Function

Private Function DrNameToCustIDp(drName As String) As String
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "B").End(xlUp).row
    For i = 2 To last
        If CStr(wsC.Cells(i, "B").value) = drName Then DrNameToCustIDp = CStr(wsC.Cells(i, "A").value): Exit Function
    Next i
End Function

Private Function CustIDToDrNamep(custID As String) As String
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsC.Cells(i, "A").value)) = NrmID(custID) Then CustIDToDrNamep = CStr(wsC.Cells(i, "B").value): Exit Function
    Next i
End Function

Private Function DoctorCreditp(custID As String) As Double
    Dim wsC As Worksheet, last As Long, i As Long
    Set wsC = ThisWorkbook.Sheets("Customers")
    last = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsC.Cells(i, "A").value)) = NrmID(custID) Then DoctorCreditp = Num(wsC.Cells(i, "L").value): Exit Function
    Next i
End Function

' ============================================================================
' PATIENT STATEMENTS (with MED CLAIMS folded in)
' ============================================================================
Public Function PatientNamesList() As Collection
    Dim ws As Worksheet, last As Long, i As Long, nm As String
    Dim seen As Object, col As New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If LCase(Trim(CStr(ws.Cells(i, 3).value))) = "patient" Then
            nm = Trim(CStr(ws.Cells(i, 7).value))
            If nm <> "" Then
                If Not seen.Exists(NrmID(nm)) Then
                    seen.Add NrmID(nm), 1
                    col.Add nm
                End If
            End If
        End If
    Next i
    ' also include patients that exist only as med claims
    Dim wsMC As Worksheet, lastMC As Long
    On Error Resume Next
    Set wsMC = ThisWorkbook.Sheets("MedAidLog")
    On Error GoTo 0
    If Not wsMC Is Nothing Then
        lastMC = wsMC.Cells(wsMC.Rows.Count, "A").End(xlUp).row
        For i = 2 To lastMC
            nm = Trim(CStr(wsMC.Cells(i, ML_PATIENT).value))
            If nm <> "" And Not seen.Exists(NrmID(nm)) Then
                seen.Add NrmID(nm), 1
                col.Add nm
            End If
        Next i
    End If
    Set PatientNamesList = col
End Function

Private Function PatientsWithBalance(dept As String) As Collection
    Dim ws As Worksheet, last As Long, i As Long, nm As String
    Dim seen As Object, col As New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("InvoiceLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If LCase(Trim(CStr(ws.Cells(i, 3).value))) = "patient" _
           And DeptMatch(CStr(ws.Cells(i, 1).value), dept) Then
            If Num(ws.Cells(i, 15).value) > 0.005 Then
                nm = Trim(CStr(ws.Cells(i, 7).value))
                If nm <> "" And Not seen.Exists(NrmID(nm)) Then
                    seen.Add NrmID(nm), 1
                    col.Add nm
                End If
            End If
        End If
    Next i
    ' med claims with outstanding balance
    Dim wsMC As Worksheet, lastMC As Long
    On Error Resume Next
    Set wsMC = ThisWorkbook.Sheets("MedAidLog")
    On Error GoTo 0
    If Not wsMC Is Nothing Then
        lastMC = wsMC.Cells(wsMC.Rows.Count, "A").End(xlUp).row
        For i = 2 To lastMC
            If Num(wsMC.Cells(i, ML_BALANCE).value) > 0.005 Then
                nm = Trim(CStr(wsMC.Cells(i, ML_PATIENT).value))
                If nm <> "" And Not seen.Exists(NrmID(nm)) Then
                    seen.Add NrmID(nm), 1
                    col.Add nm
                End If
            End If
        Next i
    End If
    Set PatientsWithBalance = col
End Function

Public Sub RunSinglePatientStatement(patientName As String, dFrom As Date, dTo As Date, dept As String)
    Dim ws As Worksheet, lastRow As Long, ans As VbMsgBoxResult, folder As String
    Dim tInv As Double, tPaid As Double, bDue As Double
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set ws = BuildAndRenderPatient(patientName, dFrom, dTo, dept, lastRow, tInv, tPaid, bDue)
    Application.ScreenUpdating = True
    If lastRow = 0 Then
        MsgBox "No transactions/outstanding items for " & patientName & ".", vbInformation
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
    Application.ScreenUpdating = True
    MsgBox "RunSinglePatientStatement error: " & Err.Description, vbExclamation
End Sub

Public Sub RunBatchPatientStatements(dFrom As Date, dTo As Date, dept As String)
    Dim pats As Collection, i As Long, ws As Worksheet, lastRow As Long
    Dim patientName As String, made As Long, folder As String
    Dim tInv As Double, tPaid As Double, bDue As Double, fpath As String
    On Error GoTo Fail
    Set pats = PatientsWithBalance(dept)
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
        lastNo = CStr(wsL.Cells(r, "A").value)
        If InStr(lastNo, "-") > 0 Then nextNo = Val(Mid(lastNo, InStr(lastNo, "-") + 1)) + 1
    End If
    r = r + 1
    wsL.Cells(r, "A").value = "STMT-" & Format(nextNo, "0000")
    wsL.Cells(r, "B").value = Now
    wsL.Cells(r, "B").NumberFormat = "yyyy/mm/dd hh:mm"
    wsL.Cells(r, "C").value = "PATIENT"
    wsL.Cells(r, "D").value = patientName
    wsL.Cells(r, "E").value = dept
    wsL.Cells(r, "F").value = dFrom: wsL.Cells(r, "F").NumberFormat = "yyyy/mm/dd"
    wsL.Cells(r, "G").value = dTo:   wsL.Cells(r, "G").NumberFormat = "yyyy/mm/dd"
    wsL.Cells(r, "H").value = tInv
    wsL.Cells(r, "I").value = tPaid
    wsL.Cells(r, "J").value = 0
    wsL.Cells(r, "K").value = bDue
    wsL.Cells(r, "L").value = pdfPath
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

' ---- PATIENT HEADER: autofill from "Med Customers", private-mode labels ----
Private Sub RenderPatientHeader(ws As Worksheet, patientName As String, _
                                dFrom As Date, dTo As Date, dept As String)
    Dim wsM As Worksheet, last As Long, i As Long, found As Boolean

    ws.Range("A3").value = DeptStmtTitle(dept)
    ws.Range("G5").value = Format(Date, "yyyy-mm-dd")
    ws.Range("G8").value = Format(dFrom, "yyyy-mm-dd")
    ws.Range("G9").value = Format(dTo, "yyyy-mm-dd")

    ' private-mode label changes
    ws.Range("A5").value = "Main Member"
    ws.Range("A6").value = "Patient Name"

    ' defaults if no Med Customers match
    ws.Range("B5").value = ""
    ws.Range("B6").value = patientName

    On Error Resume Next
    Set wsM = ThisWorkbook.Sheets("Med Customers")
    On Error GoTo 0
    If Not wsM Is Nothing Then
        last = wsM.Cells(wsM.Rows.Count, "B").End(xlUp).row
        For i = 2 To last
            If NrmID(CStr(wsM.Cells(i, "B").value)) = NrmID(patientName) Then
                ws.Range("B5").value = wsM.Cells(i, "N").value    ' Main Member
                ws.Range("B6").value = wsM.Cells(i, "B").value    ' Patient Name
                ws.Range("B7").value = wsM.Cells(i, "C").value    ' Street
                ws.Range("B8").value = wsM.Cells(i, "D").value    ' Suburb
                ws.Range("B9").value = wsM.Cells(i, "E").value    ' City
                ws.Range("B10").value = wsM.Cells(i, "F").value   ' PostCode
                ws.Range("B11").value = wsM.Cells(i, "H").value   ' Tel
                ws.Range("G6").NumberFormat = "@"
                ws.Range("G6").value = CStr(wsM.Cells(i, "J").value)  ' med aid name
                ws.Range("G7").NumberFormat = "@"
                ws.Range("G7").value = CStr(wsM.Cells(i, "L").value)  ' med aid number
                found = True
                Exit For
            End If
        Next i
    End If

    ws.Range("F6").value = "Medical Aid:"
    ws.Range("F7").value = "Med Aid No:"

    If Not found Then
        Dim wsLog As Worksheet, medAid As String, medNo As String
        Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
        last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
        For i = 2 To last
            If LCase(Trim(CStr(wsLog.Cells(i, 3).value))) = "patient" _
               And NrmID(CStr(wsLog.Cells(i, 7).value)) = NrmID(patientName) Then
                medAid = CStr(wsLog.Cells(i, 20).value)
                medNo = CStr(wsLog.Cells(i, 21).value)
            End If
        Next i
        ws.Range("G6").NumberFormat = "@": ws.Range("G6").value = medAid
        ws.Range("G7").NumberFormat = "@": ws.Range("G7").value = medNo
    End If
End Sub

Private Function BuildAndRenderPatient(patientName As String, dFrom As Date, dTo As Date, _
                                       dept As String, ByRef lastContentRow As Long, _
                                       ByRef outTotInv As Double, ByRef outTotPaid As Double, _
                                       ByRef outBalDue As Double) As Worksheet
    Dim ws As Worksheet, wsLog As Worksheet, wsPay As Worksheet, wsMC As Worksheet
    Dim last As Long, lastMC As Long, i As Long, j As Long
    Dim opening As Double
    Dim rowsArr() As Long, dts() As Double, kinds() As String, cnt As Long
    Dim tL As Long, tD As Double, tS As String
    Dim r As Long, running As Double
    Dim totalInv As Double, totalPaid As Double
    Dim ageCur As Double, age30 As Double, age60 As Double, age90 As Double
    Dim bal As Double, dueD As Date, days As Long
    Dim invDate As Date, invTotal As Double, invBal As Double
    Dim lastP As Long, pInv As String, pDate As Date, pAmt As Double, logRow As Long
    Dim mcDate As Date, mcBal As Double, mcRow As Long

    Set ws = ResetStatementSheet()
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    Set wsPay = ThisWorkbook.Sheets("Payments")
    On Error Resume Next
    Set wsMC = ThisWorkbook.Sheets("MedAidLog")
    On Error GoTo 0

    RenderPatientHeader ws, patientName, dFrom, dTo, dept

    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
    lastMC = 0
    If Not wsMC Is Nothing Then lastMC = wsMC.Cells(wsMC.Rows.Count, "A").End(xlUp).row
    ReDim rowsArr(1 To (last + lastMC + 100) * 2)
    ReDim dts(1 To (last + lastMC + 100) * 2)
    ReDim kinds(1 To (last + lastMC + 100) * 2)
    cnt = 0: opening = 0

    ' ---- invoices (patient + name match) ----
    For i = 2 To last
        If LCase(Trim(CStr(wsLog.Cells(i, 3).value))) = "patient" _
           And NrmID(CStr(wsLog.Cells(i, 7).value)) = NrmID(patientName) _
           And DeptMatch(CStr(wsLog.Cells(i, 1).value), dept) Then
            invDate = CDate(wsLog.Cells(i, 4).value)
            invBal = Num(wsLog.Cells(i, 15).value)
            If invDate < dFrom Then
                opening = opening + invBal
            ElseIf invDate <= dTo Then
                cnt = cnt + 1: rowsArr(cnt) = i: dts(cnt) = CDbl(invDate): kinds(cnt) = "INV"
            End If
        End If
    Next i

    ' ---- MED CLAIMS (patient name match) ----
    If Not wsMC Is Nothing Then
        For i = 2 To lastMC
            If NrmID(CStr(wsMC.Cells(i, ML_PATIENT).value)) = NrmID(patientName) Then
                If IsDate(wsMC.Cells(i, ML_DATE).value) Then
                    mcDate = CDate(wsMC.Cells(i, ML_DATE).value)
                    mcBal = Num(wsMC.Cells(i, ML_BALANCE).value)
                    If mcDate < dFrom Then
                        opening = opening + mcBal
                    ElseIf mcDate <= dTo Then
                        cnt = cnt + 1: rowsArr(cnt) = i
                        dts(cnt) = CDbl(mcDate): kinds(cnt) = "MC"
                    End If
                End If
            End If
        Next i
    End If

    ' ---- invoice payments ----
    lastP = wsPay.Cells(wsPay.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastP
        pInv = CStr(wsPay.Cells(i, 2).value)
        logRow = FindLogRow(wsLog, pInv)
        If logRow > 0 Then
            If LCase(Trim(CStr(wsLog.Cells(logRow, 3).value))) = "patient" _
               And NrmID(CStr(wsLog.Cells(logRow, 7).value)) = NrmID(patientName) _
               And DeptMatch(pInv, dept) Then
                pDate = CDate(wsPay.Cells(i, 3).value)
                If pDate >= dFrom And pDate <= dTo Then
                    cnt = cnt + 1: rowsArr(cnt) = i: dts(cnt) = CDbl(pDate): kinds(cnt) = "PAY"
                End If
            End If
        End If
    Next i

    ' ---- med-claim payments (Payments.InvNo -> MedAidLog MC of this patient) ----
    If Not wsMC Is Nothing Then
        For i = 2 To lastP
            pInv = CStr(wsPay.Cells(i, 2).value)
            mcRow = FindMCRow(wsMC, pInv)
            If mcRow > 0 Then
                If NrmID(CStr(wsMC.Cells(mcRow, ML_PATIENT).value)) = NrmID(patientName) Then
                    pDate = CDate(wsPay.Cells(i, 3).value)
                    If pDate >= dFrom And pDate <= dTo Then
                        cnt = cnt + 1: rowsArr(cnt) = i
                        dts(cnt) = CDbl(pDate): kinds(cnt) = "PAY"
                    End If
                End If
            End If
        Next i
    End If

    If cnt = 0 And Abs(opening) < 0.005 Then
        lastContentRow = 0: Set BuildAndRenderPatient = ws: Exit Function
    End If

    ' ---- sort oldest-first ----
    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If dts(j) > dts(j + 1) Then
                tD = dts(j): dts(j) = dts(j + 1): dts(j + 1) = tD
                tL = rowsArr(j): rowsArr(j) = rowsArr(j + 1): rowsArr(j + 1) = tL
                tS = kinds(j): kinds(j) = kinds(j + 1): kinds(j + 1) = tS
            End If
        Next j
    Next i

    Dim totalLines As Long, availRows As Long, needRows As Long, off As Long
    totalLines = 1 + cnt
    availRows = 19 - 13
    needRows = totalLines + 1 - availRows
    off = 0
    If needRows > 0 Then
        ws.Rows(19 & ":" & (19 + needRows - 1)).Insert Shift:=xlDown
        off = needRows
    End If

    r = 13
    running = opening
    ws.Cells(r, 1).value = Format(dFrom, "dd/mm/yyyy")
    ws.Cells(r, 3).value = "Balance Brought Forward"
    ws.Cells(r, 8).value = running
    r = r + 1

    totalInv = 0: totalPaid = 0
    For i = 1 To cnt
        If kinds(i) = "INV" Then
            invTotal = Num(wsLog.Cells(rowsArr(i), 12).value)
            running = running + invTotal
            totalInv = totalInv + invTotal
            ws.Cells(r, 1).value = Format(CDate(wsLog.Cells(rowsArr(i), 4).value), "dd/mm/yyyy")
            ws.Cells(r, 2).value = wsLog.Cells(rowsArr(i), 1).value
            ws.Cells(r, 3).value = wsLog.Cells(rowsArr(i), 7).value
            ws.Cells(r, 4).value = wsLog.Cells(rowsArr(i), 8).value
            ws.Cells(r, 5).value = invTotal
            ws.Cells(r, 8).value = running
        ElseIf kinds(i) = "MC" Then
            invTotal = Num(wsMC.Cells(rowsArr(i), ML_TOTAL).value)
            running = running + invTotal
            totalInv = totalInv + invTotal
            ws.Cells(r, 1).value = Format(CDate(wsMC.Cells(rowsArr(i), ML_DATE).value), "dd/mm/yyyy")
            ws.Cells(r, 2).value = wsMC.Cells(rowsArr(i), ML_NO).value
            ws.Cells(r, 3).value = "Med Claim - " & wsMC.Cells(rowsArr(i), ML_PATIENT).value
            ws.Cells(r, 5).value = invTotal
            ws.Cells(r, 8).value = running
        Else
            pAmt = Num(wsPay.Cells(rowsArr(i), 4).value)
            running = running - pAmt
            totalPaid = totalPaid + pAmt
            ws.Cells(r, 1).value = Format(CDate(wsPay.Cells(rowsArr(i), 3).value), "dd/mm/yyyy")
            ws.Cells(r, 2).value = wsPay.Cells(rowsArr(i), 1).value
            ws.Cells(r, 3).value = "Payment - " & wsPay.Cells(rowsArr(i), 2).value
            ws.Cells(r, 6).value = pAmt
            ws.Cells(r, 8).value = running
        End If
        r = r + 1
    Next i

    ' ---- aging: invoices ----
    ageCur = 0: age30 = 0: age60 = 0: age90 = 0
    For i = 2 To last
        If LCase(Trim(CStr(wsLog.Cells(i, 3).value))) = "patient" _
           And NrmID(CStr(wsLog.Cells(i, 7).value)) = NrmID(patientName) _
           And DeptMatch(CStr(wsLog.Cells(i, 1).value), dept) Then
            bal = Num(wsLog.Cells(i, 15).value)
            If bal > 0.005 Then
                If IsDate(wsLog.Cells(i, 5).value) Then
                    dueD = CDate(wsLog.Cells(i, 5).value)
                Else
                    dueD = CDate(wsLog.Cells(i, 4).value)
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

    ' ---- aging: med claims ----
    If Not wsMC Is Nothing Then
        For i = 2 To lastMC
            If NrmID(CStr(wsMC.Cells(i, ML_PATIENT).value)) = NrmID(patientName) Then
                bal = Num(wsMC.Cells(i, ML_BALANCE).value)
                If bal > 0.005 Then
                    If IsDate(wsMC.Cells(i, ML_DUE).value) Then
                        dueD = CDate(wsMC.Cells(i, ML_DUE).value)
                    ElseIf IsDate(wsMC.Cells(i, ML_DATE).value) Then
                        dueD = CDate(wsMC.Cells(i, ML_DATE).value)
                    Else
                        GoTo NextMCAge
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
NextMCAge:
        Next i
    End If

    ' ---- totals (credit = 0) ----
    Dim grossOut As Double, balDue As Double, excl As Double, vat As Double
    Const VATR As Double = 0.15
    grossOut = running: If grossOut < 0 Then grossOut = 0
    balDue = Round(grossOut, 2): If balDue < 0 Then balDue = 0
    excl = Round(balDue / (1 + VATR), 2)
    vat = Round(balDue - excl, 2)
    totalInv = Round(totalInv, 2): totalPaid = Round(totalPaid, 2)
    ageCur = Round(ageCur, 2): age30 = Round(age30, 2)
    age60 = Round(age60, 2): age90 = Round(age90, 2)

    ws.Range("H" & (19 + off)).value = totalInv
    ws.Range("H" & (20 + off)).value = totalPaid
    ws.Range("H" & (21 + off)).value = 0
    ws.Range("H" & (22 + off)).value = excl
    ws.Range("H" & (23 + off)).value = vat
    ws.Range("H" & (24 + off)).value = balDue
    ws.Range("G10").value = balDue
    ws.Range("E" & (27 + off)).value = ageCur
    ws.Range("F" & (27 + off)).value = age30
    ws.Range("G" & (27 + off)).value = age60
    ws.Range("H" & (27 + off)).value = age90

    FormatStatementP ws, r - 1

    outTotInv = totalInv: outTotPaid = totalPaid: outBalDue = balDue
    lastContentRow = 36 + off
    PaginateStatementP ws, lastContentRow
    Set BuildAndRenderPatient = ws
End Function

Private Sub FormatStatementP(ws As Worksheet, lastLineRow As Long)
    With ws.Range("E13:H" & lastLineRow)
        .NumberFormat = "R #,##0.00"
    End With
    ws.Range("A13:A" & lastLineRow).HorizontalAlignment = xlLeft
    ws.Range("E13:H" & lastLineRow).HorizontalAlignment = xlRight

    ' widen header label/value columns so labels aren't cut off
    ws.Columns("A").ColumnWidth = 15
    ws.Columns("B").ColumnWidth = 22
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
    tWA = CStr(ThisWorkbook.names("TitleWA").RefersToRange.value)
    tWD = CStr(ThisWorkbook.names("TitleWD").RefersToRange.value)
    On Error GoTo 0
    Select Case UCase(Trim(dept))
        Case "WA": DeptStmtTitle = tWA & " Statement"
        Case "WD": DeptStmtTitle = tWD & " Statement"
        Case Else: DeptStmtTitle = tWA & " / " & tWD & " Statement"
    End Select
End Function

