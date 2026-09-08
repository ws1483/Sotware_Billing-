VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' ---- date pickers on filter cells B17 (From) / C17 (To) ----
    If Not Application.Intersect(Target, Me.Range("B17,C17")) Is Nothing Then
        Cancel = True
        Dim d As Variant
        d = PickDate(Target.value)
        If IsDate(d) Then
            Application.EnableEvents = False
            Target.value = d
            Target.NumberFormat = "yyyy/mm/dd"
            Application.EnableEvents = True
            RefreshMenuSummary
        End If
        Exit Sub
    End If

    ' ---- double-click Doc Number (col B) to open the doc ----
    If Target.Column <> 2 Then Exit Sub
    Dim n As String
    n = Trim(CStr(Target.value))
    If n = "" Then Exit Sub
    If InStr(n, "--") > 0 Then Exit Sub                     ' ignore marker rows

    If Target.row >= 20 And Target.row <= 44 Then           ' Quotes section
        Cancel = True
        RecallQuote n
    ElseIf Target.row >= 46 And Target.row <= 68 Then       ' Invoices / Med Claims
        Cancel = True
        If InStr(1, UCase(n), "MC-INV-") > 0 Then
            RecallMedClaim n
        Else
            RecallInvoice n
        End If
    End If
End Sub
