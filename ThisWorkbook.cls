VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    ThisWorkbook.Sheets("Menu").Activate
    RefreshMenuSummary
    ' HideRibbon removed — full workbook UI stays available
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh.Name = "Menu" Then RefreshMenuSummary
    On Error GoTo 0
End Sub

