VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDayBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Public Frm As Object          ' frmCalendar instance

Private Sub Btn_Click()
    If Len(Btn.Caption) > 0 Then
        Frm.PickDay CInt(Btn.Caption)
    End If
End Sub
