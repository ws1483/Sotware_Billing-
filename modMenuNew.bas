Attribute VB_Name = "modMenuNew"
Option Explicit
' ============================================================================
' modMenuNew — Menu "New Quote" / "New Invoice" launchers.
'   Reads Menu!C5 (Dept: WA/WD/All) and Menu!E5 (Recipient: Doctor/Private),
'   creates a fresh doc, applies Dept -> K1 and Recipient -> K2 (which fires
'   the sheet's Worksheet_Change -> SetRecipientType), then activates it.
'   Assign these to the Menu "New Quote" / "New Invoice" buttons.
' ============================================================================
Private Const MENU_SHEET As String = "Menu"

' ---- read + normalize the Menu Department (defaults to WA if All/blank) ----
Private Function MenuDept() As String
    Dim d As String
    d = UCase(Trim(CStr(ThisWorkbook.Sheets(MENU_SHEET).Range("C5").Value)))
    If d <> "WA" And d <> "WD" Then d = "WA"     ' All/blank -> default WA
    MenuDept = d
End Function

' ---- read + translate the Menu Recipient (Doctor->doctor, Private->patient) ----
Private Function MenuRecip() As String
    Dim r As String
    r = LCase(Trim(CStr(ThisWorkbook.Sheets(MENU_SHEET).Range("E5").Value)))
    Select Case r
        Case "private", "patient": MenuRecip = "patient"
        Case Else:                 MenuRecip = "doctor"   ' Doctor/blank -> doctor
    End Select
End Function

' ============================================================ NEW QUOTE ======
Public Sub MenuNewQuote()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Quote")

    NewQuote                                   ' reset sheet + restore formulas

    Application.EnableEvents = False
    ws.Range("K1").Value = MenuDept            ' Dept -> K1
    Application.EnableEvents = True

    ws.Range("K2").Value = MenuRecip           ' fires Worksheet_Change -> SetRecipientType
    ws.Activate
End Sub

' =========================================================== NEW INVOICE =====
Public Sub MenuNewInvoice()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Invoice")

    NewInvoice                                 ' reset sheet + restore formulas

    Application.EnableEvents = False
    ws.Range("K1").Value = MenuDept            ' Dept -> K1
    Application.EnableEvents = True

    ws.Range("K2").Value = MenuRecip           ' fires Worksheet_Change -> SetRecipientType
    ws.Activate
End Sub

