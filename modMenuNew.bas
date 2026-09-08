Attribute VB_Name = "modMenuNew"
Option Explicit
' ============================================================================
' modMenuNew ? Menu "New Quote" / "New Invoice" / "New Med Claim" launchers.
'   Reads Menu!C5 (Dept: WA/WD/MC/All) and Menu!E5 (Recipient: Doctor/Private),
'   creates a fresh doc, applies Dept -> K1 and Recipient -> K2, then activates.
'   PHASE 1: added MC handling. When dept = MC, "New Invoice" / "New Med Claim"
'   opens the Med Claim sheet (patient recipient forced).
' ============================================================================
Private Const MENU_SHEET As String = "Menu"

' ---- read + normalize the Menu Department (WA/WD/MC; All/blank -> WA) ----
Private Function MenuDept() As String
    Dim d As String
    d = UCase(Trim(CStr(ThisWorkbook.Sheets(MENU_SHEET).Range("C5").value)))
    If d <> "WA" And d <> "WD" And d <> "MC" Then d = "WA"
    MenuDept = d
End Function

' ---- read + translate the Menu Recipient (Doctor->doctor, Private->patient) ----
Private Function MenuRecip() As String
    Dim r As String
    r = LCase(Trim(CStr(ThisWorkbook.Sheets(MENU_SHEET).Range("E5").value)))
    Select Case r
        Case "private", "patient": MenuRecip = "patient"
        Case Else:                 MenuRecip = "doctor"
    End Select
End Function

' MC uses WA/WD in the number; if the Menu dept is "MC" (a mode, not a billing
' dept), fall back to WA for the number segment unless E5 encodes WA/WD.
Private Function McBillingDept() As String
    Dim d As String
    d = UCase(Trim(CStr(ThisWorkbook.Sheets(MENU_SHEET).Range("C5").value)))
    If d = "WA" Or d = "WD" Then McBillingDept = d Else McBillingDept = "WA"
End Function

' ============================================================ NEW QUOTE ======
Public Sub MenuNewQuote()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Quote")
    NewQuote
    Application.EnableEvents = False
    ws.Range("K1").value = MenuDept
    Application.EnableEvents = True
    ws.Range("K2").value = MenuRecip
    ws.Activate
End Sub

' =========================================================== NEW INVOICE =====
Public Sub MenuNewInvoice()
    ' If dept mode is MC, route to a new Med Claim instead of a normal invoice.
    If UCase(Trim(CStr(ThisWorkbook.Sheets(MENU_SHEET).Range("C5").value))) = "MC" Then
        MenuNewMedClaim
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Invoice")
    NewInvoice
    Application.EnableEvents = False
    ws.Range("K1").value = MenuDept
    Application.EnableEvents = True
    ws.Range("K2").value = MenuRecip
    ws.Activate
End Sub

' ========================================================= NEW MED CLAIM =====
Public Sub MenuNewMedClaim()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Med Claim")
    NewMedClaim
    Application.EnableEvents = False
    ws.Range("K1").value = McBillingDept       ' WA/WD segment for the number
    ws.Range("K2").value = "patient"           ' MC is always patient-based
    Application.EnableEvents = True
    ws.Activate
End Sub
