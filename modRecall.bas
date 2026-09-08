Attribute VB_Name = "modRecall"
Option Explicit
' ============================================================================
' modRecall ? RecallQuote / RecallInvoice / RecallCreditNote
'   (RecallMedClaim lives in modMedClaim ? do NOT duplicate it here.)
'   Resets sheet, applies saved Recipient layout, reverse-looks-up DrName,
'   loads lines + med-aid fields, sets G7 for UPDATE. Date = SAVED date, last.
'   RecallCreditNote re-derives dept/recipient from the stored SourceInvNo.
' ============================================================================

Private Sub LoadLines(wsLines As Worksheet, ws As Worksheet, docNo As String)
    Dim lastRow As Long, i As Long, destRow As Long
    destRow = 16
    lastRow = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If CStr(wsLines.Cells(i, "A").value) = docNo Then
            If destRow > 30 Then Exit For
            ws.Range("A" & destRow).value = wsLines.Cells(i, 6).value   ' Qty
            ws.Range("D" & destRow).value = wsLines.Cells(i, 5).value   ' Description
            ws.Range("G" & destRow).value = wsLines.Cells(i, 9).value   ' Price Incl
            destRow = destRow + 1
        End If
    Next i
End Sub

' ============================================================ RECALL INVOICE ==
Sub RecallInvoice(Optional ByVal docNoIn As String = "")
    Dim ws As Worksheet, wsLog As Worksheet, wsLines As Worksheet
    Dim docNo As String, lr As Long, drName As String, recip As String

    Set ws = ThisWorkbook.Sheets("Invoice")
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    Set wsLines = ThisWorkbook.Sheets("InvoiceLines")

    If docNoIn <> "" Then
        docNo = Trim(docNoIn)
    Else
        docNo = Trim(InputBox("Enter Invoice number to recall (e.g. INV-WA-0260):", "Recall Invoice"))
    End If
    If docNo = "" Then Exit Sub
    lr = FindLogRow(wsLog, docNo)
    If lr = 0 Then MsgBox "Invoice " & docNo & " not found.", vbExclamation: Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    NewInvoice

    recip = LCase(Trim(CStr(wsLog.Cells(lr, 3).value)))
    SetRecipientType ws, recip

    If recip = "patient" Then
        ws.Range("C6").value = ""
        ws.Range("F14").value = wsLog.Cells(lr, 7).value
        ws.Range("G8").value = wsLog.Cells(lr, 20).value
        ws.Range("G9").value = wsLog.Cells(lr, 21).value
        ws.Range("G10").value = wsLog.Cells(lr, 22).value
        ws.Range("G12").value = wsLog.Cells(lr, 23).value
        ws.Range("G13").value = wsLog.Cells(lr, 24).value
    Else
        drName = CustIDToDrName(CStr(wsLog.Cells(lr, 6).value))
        ws.Range("C6").value = drName
        ws.Range("F14").value = wsLog.Cells(lr, 7).value
    End If

    ws.Range("C14").value = wsLog.Cells(lr, 8).value
    LoadLines wsLines, ws, docNo

    ws.Range("C32").value = wsLog.Cells(lr, 25).value
    ws.Range("C33").value = wsLog.Cells(lr, 26).value

    ws.Range("G7").value = docNo
    ws.Range("K4").value = docNo
    ws.Range("G6").value = wsLog.Cells(lr, 4).value

    ShowInvoiceWatermark

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Invoice " & docNo & " recalled.", vbInformation
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "RecallInvoice error: " & Err.Description, vbExclamation
End Sub

' ============================================================== RECALL QUOTE ==
Sub RecallQuote(Optional ByVal docNoIn As String = "")
    Dim ws As Worksheet, wsLog As Worksheet, wsLines As Worksheet
    Dim docNo As String, lr As Long, drName As String, recip As String

    Set ws = ThisWorkbook.Sheets("Quote")
    Set wsLog = ThisWorkbook.Sheets("QuoteLog")
    Set wsLines = ThisWorkbook.Sheets("QuoteLines")

    If docNoIn <> "" Then
        docNo = Trim(docNoIn)
    Else
        docNo = Trim(InputBox("Enter Quote number to recall (e.g. Q-WA-0160):", "Recall Quote"))
    End If
    If docNo = "" Then Exit Sub
    lr = FindLogRow(wsLog, docNo)
    If lr = 0 Then MsgBox "Quote " & docNo & " not found.", vbExclamation: Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    NewQuote

    recip = LCase(Trim(CStr(wsLog.Cells(lr, 3).value)))
    SetRecipientType ws, recip

    If recip = "patient" Then
        ws.Range("C6").value = ""
        ws.Range("F14").value = wsLog.Cells(lr, 6).value
        ws.Range("G8").value = wsLog.Cells(lr, 17).value
        ws.Range("G9").value = wsLog.Cells(lr, 18).value
        ws.Range("G10").value = wsLog.Cells(lr, 19).value
        ws.Range("G11").value = wsLog.Cells(lr, 20).value
        ws.Range("G12").value = wsLog.Cells(lr, 21).value
    Else
        drName = CustIDToDrName(CStr(wsLog.Cells(lr, 5).value))
        ws.Range("C6").value = drName
        ws.Range("F14").value = wsLog.Cells(lr, 6).value
    End If

    ws.Range("C14").value = wsLog.Cells(lr, 7).value
    LoadLines wsLines, ws, docNo

    ws.Range("C32").value = wsLog.Cells(lr, 22).value
    ws.Range("C33").value = wsLog.Cells(lr, 23).value

    ws.Range("G7").value = docNo
    ws.Range("K4").value = docNo
    ws.Range("G6").value = wsLog.Cells(lr, 4).value

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Quote " & docNo & " recalled.", vbInformation
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "RecallQuote error: " & Err.Description, vbExclamation
End Sub

' ========================================================= RECALL CREDITNOTE ==
Sub RecallCreditNote(Optional ByVal docNoIn As String = "")
    Dim ws As Worksheet, wsLog As Worksheet, wsLines As Worksheet
    Dim docNo As String, lr As Long, srcInv As String

    Set ws = ThisWorkbook.Sheets("CreditNote")
    Set wsLog = ThisWorkbook.Sheets("CreditNotes")
    Set wsLines = ThisWorkbook.Sheets("CreditNoteLines")

    If docNoIn <> "" Then
        docNo = Trim(docNoIn)
    Else
        docNo = Trim(InputBox("Enter Credit Note number to recall (e.g. CN-0001):", "Recall Credit Note"))
    End If
    If docNo = "" Then Exit Sub
    lr = FindLogRow(wsLog, docNo)
    If lr = 0 Then MsgBox "Credit Note " & docNo & " not found.", vbExclamation: Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    NewCreditNote

    srcInv = Trim(CStr(wsLog.Cells(lr, 2).value))       ' B = SourceInvNo
    If srcInv <> "" And InvoiceRow(srcInv) > 0 Then
        LoadCreditFromInvoice srcInv
    Else
        ws.Range("C6").value = CustIDToDrName(CStr(wsLog.Cells(lr, 4).value))
    End If

    ws.Range("G11").value = srcInv
    ws.Range("C14").value = wsLog.Cells(lr, 7).value     ' Reason
    ws.Range("G7").value = docNo
    ws.Range("K4").value = docNo
    ws.Range("G6").value = wsLog.Cells(lr, 3).value      ' Date

    LoadCNOwnLines wsLines, ws, docNo

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Credit Note " & docNo & " recalled.", vbInformation
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "RecallCreditNote error: " & Err.Description, vbExclamation
End Sub

' Reload a credit note's OWN saved lines (overrides invoice lines)
Private Sub LoadCNOwnLines(wsLines As Worksheet, ws As Worksheet, docNo As String)
    Dim last As Long, i As Long, dest As Long, anyFound As Boolean
    last = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If CStr(wsLines.Cells(i, "A").value) = docNo Then anyFound = True: Exit For
    Next i
    If Not anyFound Then Exit Sub

    ws.Range("A16:A30").ClearContents
    ws.Range("D16:D30").ClearContents
    ws.Range("G16:G30").ClearContents
    dest = 16
    For i = 2 To last
        If CStr(wsLines.Cells(i, "A").value) = docNo Then
            If dest > 30 Then Exit For
            ws.Range("A" & dest).value = wsLines.Cells(i, 6).value
            ws.Range("D" & dest).value = wsLines.Cells(i, 5).value
            ws.Range("G" & dest).value = wsLines.Cells(i, 9).value
            dest = dest + 1
        End If
    Next i
End Sub
