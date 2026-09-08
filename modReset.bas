Attribute VB_Name = "modReset"
' modReset ? NewQuote / NewInvoice / NewCreditNote / NewMedClaim
'   Clears user inputs and restores all formulas. Uses modHelpers.
'   PHASE 1: added NewMedClaim (mirrors NewInvoice; patient-style layout).
' ============================================================================
Option Explicit

' ------------------------------------------------------------------ INVOICE --
Sub NewInvoice()
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Invoice")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanExit

    ClearCell ws, "C6"
    ClearCell ws, "H8"
    ClearCell ws, "C32"
    ClearCell ws, "C33"
    ClearCell ws, "C14"
    ClearCell ws, "F14"
    ClearCell ws, "G6"
    ws.Range("G6").value = Date
    ws.Range("A16:A30").ClearContents
    ws.Range("D16:D30").ClearContents
    ws.Range("E16:E30").ClearContents
    ClearCell ws, "A40"
    With ws.Range("A40")
        .value = "NOTE: ": .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ClearCell ws, "G7"

    SetFormula ws, "G11", "=IF(G6="""","""",IF(DAY(G6)>27,EOMONTH(G6,1),EOMONTH(G6,0)))"
    SetFormula ws, "G8", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$A$2:$A$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
    For r = 16 To 30
        SetFormula ws, "B" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$A$2:$A$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "C" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$C$2:$C$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "E" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "/(1+VATRate),2)))"
        SetFormula ws, "F" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "-$E" & r & ",2)))"
        ws.Range("G" & r).ClearContents
        SetFormula ws, "H" & r, "=IF($D" & r & "="""","""",IF($A" & r & "="""","""",ROUND($G" & r & "*$A" & r & ",2)))"
    Next r
    SetFormula ws, "H34", "=ROUND(IF(N(I_DiscFixedCell)>0,N(I_DiscFixedCell),SUM($H$16:$H$30)*N(I_DiscPctCell)),2)"
    SetFormula ws, "H37", "=ROUND(SUM($H$16:$H$30)-I_DiscAmt,2)"
    SetFormula ws, "H36", "=ROUND(H37-ROUND(H37/(1+VATRate),2),2)"
    SetFormula ws, "H35", "=ROUND(H37-I_Vat,2)"
    ClearCell ws, "G7"
    RemoveInvoiceWatermark

    ws.Activate: ws.Range("C6").Select
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "NewInvoice error: " & Err.Description, vbExclamation
End Sub

' -------------------------------------------------------------- CREDIT NOTE --
Sub NewCreditNote()
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("CreditNote")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanExit

    ClearCell ws, "C6"
    ClearCell ws, "H8"
    ClearCell ws, "C14"
    ClearCell ws, "F14"
    ClearCell ws, "G6"
    ws.Range("G6").value = Date
    ClearCell ws, "G11"
    ws.Range("A16:A30").ClearContents
    ws.Range("D16:D30").ClearContents
    ws.Range("E16:E30").ClearContents
    ClearCell ws, "A40"
    With ws.Range("A40")
        .value = "NOTE: ": .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ClearCell ws, "G7"

    For r = 16 To 30
        SetFormula ws, "B" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$A$2:$A$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "C" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$C$2:$C$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "E" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "/(1+VATRate),2)))"
        SetFormula ws, "F" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "-$E" & r & ",2)))"
        ws.Range("G" & r).ClearContents
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
Sub NewQuote()
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Quote")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanExit

    ClearCell ws, "C6"
    ClearCell ws, "H8"
    ClearCell ws, "C32"
    ClearCell ws, "C33"
    ClearCell ws, "C14"
    ClearCell ws, "F14"
    ClearCell ws, "G6"
    ws.Range("G6").value = Date
    ws.Range("A16:A30").ClearContents
    ws.Range("D16:D30").ClearContents
    ws.Range("E16:E30").ClearContents
    ClearCell ws, "A40"
    With ws.Range("A40")
        .value = "NOTE: ": .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ClearCell ws, "G7"

    SetFormula ws, "G8", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$A$2:$A$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
    For r = 16 To 30
        SetFormula ws, "B" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$A$2:$A$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "C" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$C$2:$C$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "E" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "/(1+VATRate),2)))"
        SetFormula ws, "F" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "-$E" & r & ",2)))"
        ws.Range("G" & r).ClearContents
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

' ---------------------------------------------------------------- MED CLAIM --
' Patient-style layout (Image 5). Autofill from Med Customers on C6 change.
' Med-aid header fields (G8..G13) are inputs; discount C32/C33; note A40.
Sub NewMedClaim()
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Med Claim")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanExit

    ' 1. CLEAR INPUTS
    ClearCell ws, "C6"
    ClearCell ws, "C7": ClearCell ws, "C8": ClearCell ws, "C9"
    ClearCell ws, "C10": ClearCell ws, "C11": ClearCell ws, "C12"
    ClearCell ws, "G8": ClearCell ws, "G9": ClearCell ws, "G10"
    ClearCell ws, "G11": ClearCell ws, "G12": ClearCell ws, "G13"
    ClearCell ws, "C14"
    ClearCell ws, "F14"
    ClearCell ws, "C32"
    ClearCell ws, "C33"
    ClearCell ws, "G6"
    ws.Range("G6").value = Date
    ws.Range("A16:A30").ClearContents
    ws.Range("D16:D30").ClearContents
    ws.Range("E16:E30").ClearContents
    ClearCell ws, "A40"
    With ws.Range("A40")
        .value = "NOTE: ": .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ClearCell ws, "G7"

    ' 2. RESTORE FORMULAS
    For r = 16 To 30
        SetFormula ws, "B" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$A$2:$A$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "C" & r, "=IF($D" & r & "="""","""",IFERROR(INDEX(Pricelist!$C$2:$C$300,MATCH($D" & r & ",Pricelist!$B$2:$B$300,0)),""""))"
        SetFormula ws, "E" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "/(1+VATRate),2)))"
        SetFormula ws, "F" & r, "=IF($D" & r & "="""","""",IF($G" & r & "="""","""",ROUND($G" & r & "-$E" & r & ",2)))"
        ws.Range("G" & r).ClearContents
        SetFormula ws, "H" & r, "=IF($D" & r & "="""","""",IF($A" & r & "="""","""",ROUND($G" & r & "*$A" & r & ",2)))"
    Next r

    ' Totals ? uses MC_ named ranges (confirmed to exist)
    SetFormula ws, "H34", "=ROUND(IF(N(MC_DiscFixedCell)>0,N(MC_DiscFixedCell),SUM($H$16:$H$30)*N(MC_DiscPctCell)),2)"
    SetFormula ws, "H37", "=ROUND(SUM($H$16:$H$30)-MC_DiscAmt,2)"
    SetFormula ws, "H36", "=ROUND(H37-ROUND(H37/(1+VATRate),2),2)"
    SetFormula ws, "H35", "=ROUND(H37-MC_Vat,2)"

    ws.Activate: ws.Range("C6").Select
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "NewMedClaim error: " & Err.Description, vbExclamation
End Sub
