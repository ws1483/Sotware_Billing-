Attribute VB_Name = "modFormulas"
Option Explicit
' ============================================================================
' modFormulas — line-item Price INCL (col G) auto-fill from Pricelist.
'   G is the ENTERED/authoritative price (manual override allowed).
'   E (Excl) and F (VAT) are derived by sheet formulas from G.
'   Pricelist: A=Code B=Description(key) C=ZCode G=NewPriceIncl H=VATable
' ============================================================================
Private Const PRICELIST_SHEET As String = "Pricelist"

Public Sub FillLinePrice(ws As Worksheet, ByVal r As Long)
    Dim desc As String, wsP As Worksheet, m As Variant
    If r < 16 Or r > 30 Then Exit Sub
    desc = Trim(CStr(ws.Cells(r, 4).Value))
    If desc = "" Then ws.Cells(r, 7).ClearContents: Exit Sub   ' clear G (Incl)

    On Error GoTo Fail
    Set wsP = ThisWorkbook.Sheets(PRICELIST_SHEET)
    m = Application.match(desc, wsP.Range("B2:B300"), 0)
    If IsError(m) Then
        ws.Cells(r, 7).ClearContents
    Else
        ' col 7 = G = NewPriceIncl on Pricelist; write to G on the doc sheet
        ws.Cells(r, 7).Value = Round(CDbl(wsP.Cells(CLng(m) + 1, 7).Value), 2)
    End If
    Exit Sub
Fail:
End Sub

Public Sub RefillAllLinePrices(ws As Worksheet)
    Dim r As Long
    Application.EnableEvents = False
    For r = 16 To 30
        FillLinePrice ws, r
    Next r
    Application.EnableEvents = True
End Sub

