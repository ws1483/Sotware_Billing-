Attribute VB_Name = "modReset"
' ============================================================================
' modReset — NewQuote / NewInvoice / NewCreditNote
'   Clears user inputs and restores all formulas. Uses modHelpers.
' ============================================================================
Option Explicit

' ------------------------------------------------------------------ INVOICE --
Sub NewInvoice()
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Invoice")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanExit

    ' 1. CLEAR INPUTS
    ClearCell ws, "C6"
    ClearCell ws, "H8"
    ClearCell ws, "C32"
    ClearCell ws, "C33"
    ClearCell ws, "C14"
    ClearCell ws, "F14"
    ClearCell ws, "G6"
    ws.Range("G6").Value = Date          ' Option B: auto-fill today (frozen value)
    ws.Range("A16:A30").ClearContents
    ws.Range("D16:D30").ClearContents
    ws.Range("E16:E30").ClearContents
    ClearCell ws, "A40"
    With ws.Range("A40")
        .Value = "NOTE: ": .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ClearCell ws, "G7"

        ' 2. RESTORE FORMULAS  (ZeroVAT removed, ROUND added, /100 removed)
    SetFormula ws, "G11", "=IF(G6="""","""",IF(DAY(G6)>27,EOMONTH(G6,1),EOMONTH(G6,0)))"
        SetFormula ws, "G8", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$A$2:$A$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
        For r = 16 To 30
        SetFormula ws, "B" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$A$2:$A$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "C" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$C$2:$C$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "E" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "/(1+VATRate),2)))"
        SetFormula ws, "F" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "-$E" & r & ",2)))"
        ws.Range("G" & r).ClearContents      ' G is an INPUT (Pricelist fill / manual) — no formula
        SetFormula ws, "H" & r, "=IF($D" & r & "="""","""",IF($A" & r & "="""","""",ROUND($G" & r & "*$A" & r & ",2)))"
    Next r
        SetFormula ws, "H34", "=ROUND(IF(N(I_DiscFixedCell)>0,N(I_DiscFixedCell),SUM($H$16:$H$30)*N(I_DiscPctCell)),2)"
    SetFormula ws, "H37", "=ROUND(SUM($H$16:$H$30)-I_DiscAmt,2)"
    SetFormula ws, "H36", "=ROUND(H37-ROUND(H37/(1+VATRate),2),2)"
    SetFormula ws, "H35", "=ROUND(H37-I_Vat,2)"
    ClearCell ws, "G7"        ' (existing)
    RemoveInvoiceWatermark    ' <-- clears any PAID stamp on new invoice

    ws.Activate: ws.Range("C6").Select
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "NewInvoice error: " & Err.Description, vbExclamation
End Sub

' -------------------------------------------------------------- CREDIT NOTE --
Sub NewCreditNote()
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("CreditNote")   ' NOTE: trailing space in tab name

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanExit

    ClearCell ws, "C6"
    ClearCell ws, "H8"
    ClearCell ws, "C14"
    ClearCell ws, "F14"
    ClearCell ws, "G6"
    ws.Range("G6").Value = Date          ' Option B: auto-fill today (frozen value)
    ClearCell ws, "G11"
    ws.Range("A16:A30").ClearContents
    ws.Range("D16:D30").ClearContents
    ws.Range("E16:E30").ClearContents
    ClearCell ws, "A40"
    With ws.Range("A40")
        .Value = "NOTE: ": .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ClearCell ws, "G7"

            For r = 16 To 30
        SetFormula ws, "B" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$A$2:$A$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "C" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$C$2:$C$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "E" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "/(1+VATRate),2)))"
        SetFormula ws, "F" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "-$E" & r & ",2)))"
        ws.Range("G" & r).ClearContents      ' G is an INPUT (Pricelist fill / manual) — no formula
        SetFormula ws, "H" & r, "=IF($D" & r & "="""","""",IF($A" & r & "="""","""",ROUND($G" & r & "*$A" & r & ",2)))"
    Next r
        SetFormula ws, "H37", "=ROUND(SUM($H$16:$H$30),2)"
    SetFormula ws, "H36", "=ROUND(H37-ROUND(H37/(1+VATRate),2),2)"
    SetFormula ws, "H35", "=ROUND(H37-H36,2)"
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "NewCreditNote error: " & Err.Description, vbExclamation
End Sub

' -------------------------------------------------------------------- QUOTE --
' NOTE: Paste your existing tested NewQuote here unchanged (it already works).
' It uses the same modHelpers (ClearCell/SetFormula), so no edits needed.
' -------------------------------------------------------------------- QUOTE --
' Add to modReset. Reconstructed from reference doc.
'   Has discount block, uses Q_ named ranges, plain "Quote" title, no due date.
Sub NewQuote()
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Quote")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanExit

    ' 1. CLEAR INPUTS
    ClearCell ws, "C6"
    ClearCell ws, "H8"
    ClearCell ws, "C32"
    ClearCell ws, "C33"
    ClearCell ws, "C14"
    ClearCell ws, "F14"
    ClearCell ws, "G6"
    ws.Range("G6").Value = Date          ' Option B: auto-fill today (frozen value)
    ws.Range("A16:A30").ClearContents
    ws.Range("D16:D30").ClearContents
    ws.Range("E16:E30").ClearContents
    ClearCell ws, "A40"
    With ws.Range("A40")
        .Value = "NOTE: ": .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ClearCell ws, "G7"

        ' 2. RESTORE FORMULAS  (ZeroVAT removed, ROUND added, /100 removed)
            SetFormula ws, "G8", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$A$2:$A$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
        For r = 16 To 30
        SetFormula ws, "B" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$A$2:$A$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "C" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$C$2:$C$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "E" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "/(1+VATRate),2)))"
        SetFormula ws, "F" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "-$E" & r & ",2)))"
        ws.Range("G" & r).ClearContents      ' G is an INPUT (Pricelist fill / manual) — no formula
        SetFormula ws, "H" & r, "=IF($D" & r & "="""","""",IF($A" & r & "="""","""",ROUND($G" & r & "*$A" & r & ",2)))"
    Next r
        SetFormula ws, "H34", "=ROUND(IF(N(Q_DiscFixedCell)>0,N(Q_DiscFixedCell),SUM($H$16:$H$30)*N(Q_DiscPctCell)),2)"
    SetFormula ws, "H37", "=ROUND(SUM($H$16:$H$30)-Q_DiscAmt,2)"
    SetFormula ws, "H36", "=ROUND(H37-ROUND(H37/(1+VATRate),2),2)"
    SetFormula ws, "H35", "=ROUND(H37-Q_Vat,2)"
    ws.Activate: ws.Range("C6").Select
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "NewQuote error: " & Err.Description, vbExclamation
End Sub

