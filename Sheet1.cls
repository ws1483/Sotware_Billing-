VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)
    ' auto-refresh when any filter cell changes (incl. dept C5)
    If Not Application.Intersect(target, Me.Range("B17,C17,E17,G17,I17,C5")) Is Nothing Then
        RefreshMenuSummary
    End If
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
    ' ---- date pickers on filter cells B17 (From) / C17 (To) ----
    If Not Application.Intersect(target, Me.Range("B17,C17")) Is Nothing Then
        Cancel = True
        Dim d As Variant
        d = PickDate(target.Value)
        If IsDate(d) Then
            Application.EnableEvents = False
            target.Value = d
            target.NumberFormat = "yyyy/mm/dd"
            Application.EnableEvents = True
            RefreshMenuSummary
        End If
        Exit Sub
    End If

    ' ---- double-click Doc Number (col B) to open the doc ----
    If target.Column <> 2 Then Exit Sub
    Dim n As String
    n = Trim(CStr(target.Value))
    If n = "" Then Exit Sub
    If InStr(n, "--") > 0 Then Exit Sub                     ' ignore marker rows

    If target.row >= 20 And target.row <= 44 Then           ' Quotes section
        Cancel = True
        RecallQuote n
    ElseIf target.row >= 46 And target.row <= 68 Then       ' Invoices section
        Cancel = True
        RecallInvoice n
    End If
End Sub

