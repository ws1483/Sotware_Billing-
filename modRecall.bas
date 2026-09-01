Attribute VB_Name = "modRecall"
Option Explicit
' ============================================================================
' modRecall — RecallQuote / RecallInvoice / RecallCreditNote
'   Resets sheet, applies saved Recipient layout, reverse-looks-up DrName,
'   loads lines + med-aid fields, sets G7 for UPDATE. Date = SAVED date, last.
'   RecallQuote/RecallInvoice accept optional docNoIn (from Menu double-click);
'   when blank they prompt via InputBox (manual recall).
'   1B FIX: LoadLines restores Price INCL (col 9 -> G); E/F/H derive.
'   PHASE 3: restore discount %/fixed into C32/C33.
' ============================================================================

Private Sub LoadLines(wsLines As Worksheet, ws As Worksheet, docNo As String)
    Dim lastRow As Long, i As Long, destRow As Long
    destRow = 16
    lastRow = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If CStr(wsLines.Cells(i, "A").Value) = docNo Then
            If destRow > 30 Then Exit For
            ws.Range("A" & destRow).Value = wsLines.Cells(i, 6).Value   ' Qty
            ws.Range("D" & destRow).Value = wsLines.Cells(i, 5).Value   ' Description
            ws.Range("G" & destRow).Value = wsLines.Cells(i, 9).Value   ' Price Incl (authoritative; E/F/H derive)
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
        docNo = Trim(docNoIn)                                ' from Menu double-click
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

    recip = LCase(Trim(CStr(wsLog.Cells(lr, 3).Value)))     ' C=RecipientType
    SetRecipientType ws, recip

    If recip = "patient" Then
        ws.Range("C6").Value = ""                            ' manual name (not a CustID lookup)
        ws.Range("F14").Value = wsLog.Cells(lr, 7).Value     ' PatientName
        ws.Range("G8").Value = wsLog.Cells(lr, 20).Value     ' T MedAid
        ws.Range("G9").Value = wsLog.Cells(lr, 21).Value     ' U MedNo
        ws.Range("G10").Value = wsLog.Cells(lr, 22).Value    ' V MainMember
        ws.Range("G12").Value = wsLog.Cells(lr, 23).Value    ' W Doctor
        ws.Range("G13").Value = wsLog.Cells(lr, 24).Value    ' X BHF
    Else
        drName = CustIDToDrName(CStr(wsLog.Cells(lr, 6).Value))
        ws.Range("C6").Value = drName
        ws.Range("F14").Value = wsLog.Cells(lr, 7).Value
    End If

    ws.Range("C14").Value = wsLog.Cells(lr, 8).Value         ' ApplianceType (H=8)
    LoadLines wsLines, ws, docNo

    ' --- PHASE 3: restore discount inputs (Y=25 %, Z=26 fixed) ---
    ws.Range("C32").Value = wsLog.Cells(lr, 25).Value        ' Discount %
    ws.Range("C33").Value = wsLog.Cells(lr, 26).Value        ' Discount Fixed

    ws.Range("G7").Value = docNo
    ws.Range("K4").Value = docNo
    ws.Range("G6").Value = wsLog.Cells(lr, 4).Value          ' DATE LAST (InvDate D=4)

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
        docNo = Trim(docNoIn)                                ' from Menu double-click
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

    recip = LCase(Trim(CStr(wsLog.Cells(lr, 3).Value)))      ' C=RecipientType
    SetRecipientType ws, recip

    If recip = "patient" Then
        ws.Range("C6").Value = ""
        ws.Range("F14").Value = wsLog.Cells(lr, 6).Value      ' PatientName (F=6)
        ws.Range("G8").Value = wsLog.Cells(lr, 17).Value      ' Q MedAid
        ws.Range("G9").Value = wsLog.Cells(lr, 18).Value      ' R MedNo
        ws.Range("G10").Value = wsLog.Cells(lr, 19).Value     ' S MainMember
        ws.Range("G11").Value = wsLog.Cells(lr, 20).Value     ' T Doctor
        ws.Range("G12").Value = wsLog.Cells(lr, 21).Value     ' U BHF
    Else
        drName = CustIDToDrName(CStr(wsLog.Cells(lr, 5).Value))
        ws.Range("C6").Value = drName
        ws.Range("F14").Value = wsLog.Cells(lr, 6).Value
    End If

    ws.Range("C14").Value = wsLog.Cells(lr, 7).Value          ' ApplianceType (G=7)
    LoadLines wsLines, ws, docNo

    ' --- PHASE 3: restore discount inputs (V=22 %, W=23 fixed) ---
    ws.Range("C32").Value = wsLog.Cells(lr, 22).Value         ' Discount %
    ws.Range("C33").Value = wsLog.Cells(lr, 23).Value         ' Discount Fixed

    ws.Range("G7").Value = docNo
    ws.Range("K4").Value = docNo
    ws.Range("G6").Value = wsLog.Cells(lr, 4).Value           ' DATE LAST (QuoteDate D=4)

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Quote " & docNo & " recalled.", vbInformation
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "RecallQuote error: " & Err.Description, vbExclamation
End Sub

' ========================================================= RECALL CREDITNOTE ==
Sub RecallCreditNote()
    Dim ws As Worksheet, wsLog As Worksheet, wsLines As Worksheet
    Dim docNo As String, lr As Long, drName As String

    Set ws = ThisWorkbook.Sheets("CreditNote")
    Set wsLog = ThisWorkbook.Sheets("CreditNotes")
    Set wsLines = ThisWorkbook.Sheets("CreditNoteLines")

    docNo = Trim(InputBox("Enter Credit Note number to recall (e.g. CN-0001):", "Recall Credit Note"))
    If docNo = "" Then Exit Sub
    lr = FindLogRow(wsLog, docNo)
    If lr = 0 Then MsgBox "Credit Note " & docNo & " not found.", vbExclamation: Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    NewCreditNote

    drName = CustIDToDrName(CStr(wsLog.Cells(lr, 4).Value))   ' CustID D=4
    ws.Range("C6").Value = drName

    ws.Range("G11").Value = wsLog.Cells(lr, 2).Value          ' SourceInvNo (B=2)
    ws.Range("C14").Value = wsLog.Cells(lr, 7).Value          ' Reason (G=7)

    LoadLines wsLines, ws, docNo

    ws.Range("G7").Value = docNo
    ws.Range("K4").Value = docNo
    ws.Range("G6").Value = wsLog.Cells(lr, 3).Value           ' DATE LAST (Date C=3)

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Credit Note " & docNo & " recalled.", vbInformation
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "RecallCreditNote error: " & Err.Description, vbExclamation
End Sub

