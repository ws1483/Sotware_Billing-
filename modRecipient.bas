Attribute VB_Name = "modRecipient"
Option Explicit
' ============================================================================
' modRecipient — Doctor/Private dynamic top-section for Quote & Invoice.
'   Selector = K2 (Q_Recip / I_Recip), values "doctor" / "patient".
'   Labels col F, values merged G:H (stored in G). CreditNote = doctor-only.
' ============================================================================

Public Sub SetRecipientType(ws As Worksheet, ByVal mode As String)
    Dim isInv As Boolean
    mode = LCase(Trim(mode))
    If mode <> "patient" Then mode = "doctor"
    isInv = (ws.Name = "Invoice")

    Application.EnableEvents = False
    On Error GoTo Clean
    ws.Range("K2").Value = mode
    If mode = "doctor" Then ApplyDoctor ws, isInv Else ApplyPrivate ws, isInv
Clean:
    Application.EnableEvents = True
End Sub

Private Sub ApplyDoctor(ws As Worksheet, isInv As Boolean)
    RestoreLeftBlockFormulas ws
    ws.Range("F8").Value = "Cust ID:"
    ws.Range("F9").Value = "Vat No:"
    ws.Range("F10").Value = "BHF No:"
    If isInv Then ws.Range("F11").Value = "Due Date:" Else ClearCell ws, "F11"
    ClearCell ws, "F12": ClearCell ws, "F13"

    SetFormula ws, "G8", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$A$2:$A$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
    SetFormula ws, "G9", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$K$2:$K$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
    SetFormula ws, "G10", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$H$2:$H$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
    If isInv Then
        SetFormula ws, "G11", "=IF(G6="""","""",IF(DAY(G6)>27,EOMONTH(G6,1),EOMONTH(G6,0)))"
    Else
        ClearCell ws, "G11"
    End If
    ClearCell ws, "G12": ClearCell ws, "G13"
    ClearValidation ws.Range("G9"): ClearValidation ws.Range("G10"): ClearValidation ws.Range("G13")
End Sub

Private Sub ApplyPrivate(ws As Worksheet, isInv As Boolean)
    Dim drCell As String, bhfCell As String
    ClearLeftBlockToManual ws
    ws.Range("F8").Value = "Med Aid:"
    ws.Range("F9").Value = "Med No:"
    ws.Range("F10").Value = "Main Mem:"
    If isInv Then
        ws.Range("F11").Value = "Due Date:"
        ws.Range("F12").Value = "Doctor:"
        ws.Range("F13").Value = "BHF:"
        drCell = "G12": bhfCell = "G13"
        SetFormula ws, "G11", "=IF(G6="""","""",IF(DAY(G6)>27,EOMONTH(G6,1),EOMONTH(G6,0)))"
    Else
        ws.Range("F11").Value = "Doctor:"
        ws.Range("F12").Value = "BHF:"
        ClearCell ws, "F13"
        drCell = "G11": bhfCell = "G12"
        ClearCell ws, "G13"
    End If
    ClearCell ws, "G8": ClearCell ws, "G9": ClearCell ws, "G10"
    ClearCell ws, drCell: ClearCell ws, bhfCell
    AddListValidation ws.Range("G8"), "MedAidList"
    AddListValidation ws.Range(drCell), "CustNameList"
End Sub

Private Sub RestoreLeftBlockFormulas(ws As Worksheet)
    SetFormula ws, "C7", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$C$2:$C$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
    SetFormula ws, "C8", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$D$2:$D$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
    SetFormula ws, "C9", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$E$2:$E$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
    SetFormula ws, "C10", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$F$2:$F$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
    SetFormula ws, "C11", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$G$2:$G$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
    SetFormula ws, "C12", "=IF($C$6="""","""",IFERROR(INDEX(Customers!$J$2:$J$1000,MATCH($C$6,Customers!$B$2:$B$1000,0)),""""))"
End Sub

Private Sub ClearLeftBlockToManual(ws As Worksheet)
    Dim r As Long
    For r = 7 To 12
        ClearCell ws, "C" & r
    Next r
End Sub

Private Sub AddListValidation(rng As Range, listName As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=" & listName
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub

Private Sub ClearValidation(rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
End Sub



