Attribute VB_Name = "modConvert"
Option Explicit
' ============================================================================
' ConvertQuoteToInvoice — button on the QUOTE sheet.
'   1B FIX: copy line Price INCL (G), not Excl (E) — E/F/H derive.
'   PHASE 3 FIX: copy discounts from C32/C33 (not K5/K6).
'   Copies C14/D14 ApplianceType header across.
' ============================================================================

Private Const QL_STATUS_COL   As Long = 13    ' M = Status
Private Const QL_CONVINV_COL  As Long = 14    ' N = ConvertedInvNo
Private Const IL_SRCQUOTE_COL As Long = 17    ' Q = SourceQuoteNo

Public Sub ConvertQuoteToInvoice()
    Dim wsQ As Worksheet, wsI As Worksheet
    Dim wsQLog As Worksheet, wsILog As Worksheet
    Dim quoteNo As String, qrow As Long, newInv As String
    Dim existingConv As String, irow As Long

    Set wsQ = ThisWorkbook.Sheets("Quote")
    Set wsI = ThisWorkbook.Sheets("Invoice")
    Set wsQLog = ThisWorkbook.Sheets("QuoteLog")
    Set wsILog = ThisWorkbook.Sheets("InvoiceLog")

    ' ---- 1. validate current quote ----
    quoteNo = Trim(CStr(wsQ.Range("G7").Value))
    If quoteNo = "" Then
        MsgBox "No saved quote is displayed. Recall or Save a quote first.", vbExclamation
        Exit Sub
    End If

    qrow = FindLogRow(wsQLog, quoteNo)
    If qrow = 0 Then
        MsgBox "Quote " & quoteNo & " not found in QuoteLog. Save it first.", vbExclamation
        Exit Sub
    End If

    ' ---- 2. block re-conversion ----
    existingConv = Trim(CStr(wsQLog.Cells(qrow, QL_CONVINV_COL).Value))
    If existingConv <> "" Then
        MsgBox "This quote was already converted to " & existingConv & ".", vbExclamation
        Exit Sub
    End If

    If MsgBox("Convert quote " & quoteNo & " to a new invoice?", _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- 3. start a clean invoice, then copy quote data across ----
    NewInvoice                                   ' resets Invoice sheet + restores formulas

    ' header fields (same layout on both sheets)
    wsI.Range("C6").Value = wsQ.Range("C6").Value      ' Bill-To / DrName
    wsI.Range("C14").Value = wsQ.Range("C14").Value    ' Appliance Type
    wsI.Range("D14").Value = wsQ.Range("D14").Value    ' Appliance Type (2nd cell / merged partner)
    wsI.Range("F14").Value = wsQ.Range("F14").Value    ' Patient Name

    ' line items — A=Qty, D=Description, G=Price Incl (authoritative; E/F/H derive)
    wsI.Range("A16:A30").Value = wsQ.Range("A16:A30").Value   ' Qty
    wsI.Range("D16:D30").Value = wsQ.Range("D16:D30").Value   ' Description
    wsI.Range("G16:G30").Value = wsQ.Range("G16:G30").Value   ' Price Incl (1B authoritative)

    ' options: dept, recipient, discounts (C32 %, C33 fixed)
    CopyIfExists wsQ, wsI, "K1"     ' Dept
    CopyIfExists wsQ, wsI, "K2"     ' Recipient
    CopyIfExists wsQ, wsI, "C32"    ' Discount %
    CopyIfExists wsQ, wsI, "C33"    ' Discount fixed

    Application.EnableEvents = True     ' allow sheet formulas to recalc
    Application.Calculate

    ' ---- 4. save the invoice (assigns new INV number) ----
    gSuppressClearPrompt = True
    SaveInvoice
    gSuppressClearPrompt = False

    newInv = Trim(CStr(wsI.Range("G7").Value))
    If newInv = "" Then Err.Raise 513, , "Invoice number was not assigned by SaveInvoice."

    ' ---- 5. link both logs ----
    irow = FindLogRow(wsILog, newInv)
    If irow > 0 Then wsILog.Cells(irow, IL_SRCQUOTE_COL).Value = quoteNo   ' Q SourceQuoteNo

    wsQLog.Cells(qrow, QL_STATUS_COL).Value = "Converted"                  ' M Status
    wsQLog.Cells(qrow, QL_CONVINV_COL).Value = newInv                      ' N ConvertedInvNo

    LogAudit "Convert", quoteNo, "", "-> " & newInv, "Quote converted to invoice"
    LogAudit "Convert", newInv, "", "<- " & quoteNo, "Invoice created from quote"

    Application.ScreenUpdating = True
    wsI.Activate

    MsgBox "Quote " & quoteNo & " converted to invoice " & newInv & "." & vbCrLf & _
           "Choose where to save the PDF next.", vbInformation

    ' ---- 6. launch PDF export for the new invoice ----
    ExportPDF
    Exit Sub

Fail:
    gSuppressClearPrompt = False        ' always reset the flag on error
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "ConvertQuoteToInvoice error: " & Err.Description, vbExclamation
End Sub

' ---- copy a cell only if it exists / has a value ----
Private Sub CopyIfExists(src As Worksheet, dst As Worksheet, addr As String)
    On Error Resume Next
    dst.Range(addr).Value = src.Range(addr).Value
    On Error GoTo 0
End Sub



