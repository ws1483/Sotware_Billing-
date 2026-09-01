Attribute VB_Name = "modWatermark"
Option Explicit
' ============================================================================
' Invoice PAID watermark — diagonal, semi-transparent light-red WordArt
' across the line-item area (rows 18-28). Invoice sheet only.
' ============================================================================

Private Const WM_NAME As String = "wmPaid"

' ---- Show stamp only when the invoice is fully Paid ----
Public Sub ShowInvoiceWatermark()
    Dim ws As Worksheet, status As String
    Set ws = ThisWorkbook.Sheets("Invoice")

    RemoveInvoiceWatermark
    status = InvoiceStatus(CStr(ws.Range("G7").Value))
    If UCase(status) = "PAID" Then AddPaidStamp ws
End Sub

' ---- Remove the stamp (safe if none exists) ----
Public Sub RemoveInvoiceWatermark()
    Dim ws As Worksheet, shp As Shape
    Set ws = ThisWorkbook.Sheets("Invoice")
    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.Name = WM_NAME Then shp.Delete
    Next shp
    On Error GoTo 0
End Sub

' ---- Build the diagonal WordArt stamp centered over rows 18-28 ----
Private Sub AddPaidStamp(ws As Worksheet)
    Dim shp As Shape, area As Range
    Dim l As Single, t As Single, w As Single, h As Single

    Set area = ws.Range("A18:H28")
    l = area.Left: t = area.Top: w = area.Width: h = area.Height

    Set shp = ws.Shapes.AddTextEffect( _
        PresetTextEffect:=msoTextEffect1, Text:="PAID", _
        FontName:="Arial Black", FontSize:=96, _
        FontBold:=msoTrue, FontItalic:=msoFalse, _
        Left:=l, Top:=t)

    With shp
        .Name = WM_NAME
        .LockAspectRatio = msoFalse
        .Width = w * 0.9
        .Height = h * 0.8
        .Left = l + (w - .Width) / 2
        .Top = t + (h - .Height) / 2
        .Rotation = -45
        .Fill.ForeColor.RGB = RGB(230, 60, 60)
        .Fill.Transparency = 0.7
        .Line.Visible = msoFalse
        .Placement = xlMove
    End With

    shp.ZOrder msoSendToBack
End Sub

' ---- Look up an invoice's Status from InvoiceLog (col P=16) ----
Private Function InvoiceStatus(invNo As String) As String
    Dim wsLog As Worksheet, lr As Long
    If Trim(invNo) = "" Then Exit Function
    Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
    lr = FindLogRow(wsLog, invNo)
    If lr > 0 Then InvoiceStatus = CStr(wsLog.Cells(lr, 16).Value)
End Function


