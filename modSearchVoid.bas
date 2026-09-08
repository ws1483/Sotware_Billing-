Attribute VB_Name = "modSearchVoid"
' modSearchVoid ? launches frmSearch; hard-deletes (voids) a doc.
'   PHASE 1: added "medclaim" support (MedAidLog / MedAidLines).
' ============================================================================

Public Sub ShowSearchForm()
    frmSearch.Show
End Sub

Public Function VoidDocument(ByVal docNo As String, ByVal docType As String) As Boolean
    Dim wsLog As Worksheet, wsLines As Worksheet
    Dim lr As Long, paid As Double, srcQuote As String, typed As String
    Dim paidCol As Long
    VoidDocument = False
    If docNo = "" Then Exit Function

    Select Case docType
        Case "invoice"
            Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
            Set wsLines = ThisWorkbook.Sheets("InvoiceLines")
            paidCol = IL_PAID
        Case "medclaim"
            Set wsLog = ThisWorkbook.Sheets("MedAidLog")
            Set wsLines = ThisWorkbook.Sheets("MedAidLines")
            paidCol = ML_PAID
        Case Else   ' quote
            Set wsLog = ThisWorkbook.Sheets("QuoteLog")
            Set wsLines = ThisWorkbook.Sheets("QuoteLines")
            paidCol = 0
    End Select

    lr = FindLogRow(wsLog, docNo)
    If lr = 0 Then MsgBox docNo & " not found.", vbExclamation: Exit Function

    ' block paid invoices / med claims
    If paidCol > 0 Then
        paid = Num(wsLog.Cells(lr, paidCol).value)
        If paid > 0 Then
            MsgBox "This document has payments recorded (R" & Format(paid, "#,##0.00") & ")." & vbCrLf & _
                   "You cannot void it ? issue a credit note instead.", vbExclamation, "Void blocked"
            Exit Function
        End If
    End If

    If MsgBox("Permanently DELETE " & docNo & " and all its line items?" & vbCrLf & _
              "This removes it from all logs, statements and reports. This cannot be undone.", _
              vbCritical + vbYesNo, "Confirm Void") = vbNo Then Exit Function

    typed = Trim(InputBox("To confirm, type the document number exactly:" & vbCrLf & docNo, "Confirm Void"))
    If typed = "" Then Exit Function
    If LCase(typed) <> LCase(docNo) Then
        MsgBox "Typed value did not match. Void cancelled.", vbExclamation: Exit Function
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' source quote (invoice only) BEFORE deleting
    srcQuote = ""
    If docType = "invoice" Then srcQuote = Trim(CStr(wsLog.Cells(lr, IL_SRCQUOTE).value))

    LogAudit "Void", docNo, "", "DELETED", docType & " voided (hard delete)"

    DeleteLinesFor wsLines, docNo
    wsLog.Rows(lr).Delete

    ' revert source quote to re-convertible
    If srcQuote <> "" Then
        Dim wsQLog As Worksheet, qr As Long
        Set wsQLog = ThisWorkbook.Sheets("QuoteLog")
        qr = FindLogRow(wsQLog, srcQuote)
        If qr > 0 Then
            wsQLog.Cells(qr, QL_STATUS).value = "Saved"
            wsQLog.Cells(qr, QL_CONVINV).value = ""
            LogAudit "Void", srcQuote, "Converted", "Saved", "Reverted after invoice " & docNo & " voided"
        End If
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    On Error Resume Next
    RefreshMenuSummary
    On Error GoTo 0

    MsgBox docNo & " has been permanently voided.", vbInformation
    VoidDocument = True
    Exit Function
Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "VoidDocument error: " & Err.Description, vbExclamation
End Function

Private Sub DeleteLinesFor(wsLines As Worksheet, docNo As String)
    Dim last As Long, i As Long
    last = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = last To 2 Step -1
        If NrmID(CStr(wsLines.Cells(i, LN_DOCNO).value)) = NrmID(docNo) Then
            wsLines.Rows(i).Delete
        End If
    Next i
End Sub

' Open a logged statement's saved PDF. Falls back to re-generate.
Public Sub OpenStatementPDF(ByVal stmtNo As String)
    Dim ws As Worksheet, last As Long, i As Long, foundRow As Long
    Dim pdfPath As String, drName As String

    Set ws = ThisWorkbook.Sheets("StatementLog")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    foundRow = 0
    For i = 2 To last
        If NrmID(CStr(ws.Cells(i, 1).value)) = NrmID(stmtNo) Then foundRow = i: Exit For
    Next i
    If foundRow = 0 Then MsgBox "Statement " & stmtNo & " not found.", vbExclamation: Exit Sub

    pdfPath = Trim(CStr(ws.Cells(foundRow, 12).value))
    drName = CStr(ws.Cells(foundRow, 4).value)

    If pdfPath = "" Then
        MsgBox "No PDF path was logged for " & stmtNo & ".", vbExclamation
    ElseIf dir(pdfPath) = "" Then
        If MsgBox("The saved PDF no longer exists at:" & vbCrLf & pdfPath & vbCrLf & vbCrLf & _
                  "Re-generate this statement now?", vbExclamation + vbYesNo, "PDF missing") = vbYes Then
            RunSingleStatement drName, CDate(ws.Cells(foundRow, 6).value), _
                               CDate(ws.Cells(foundRow, 7).value), CStr(ws.Cells(foundRow, 5).value)
        End If
    Else
        ThisWorkbook.FollowHyperlink pdfPath
    End If
End Sub

