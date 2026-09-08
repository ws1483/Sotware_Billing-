VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "UserForm1"
   ClientHeight    =   4380
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4152
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Controls: cboType, optSummary, optFull, txtFrom, txtTo (dbl-click PickDate),
'           cboDept, btnGenerate, btnCancel

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Initialize()
    cboType.Clear
    cboType.AddItem "VAT"
    cboType.AddItem "Invoices"
    cboType.AddItem "Quotes"
    cboType.AddItem "Med Claims"
    cboType.AddItem "Credit Notes"
    cboType.AddItem "Full Summary"
    cboType.value = "Full Summary"

    cboDept.Clear
    cboDept.AddItem "All": cboDept.AddItem "WA": cboDept.AddItem "WD": cboDept.AddItem "MC"
    cboDept.value = "All"

    optSummary.value = True
    txtFrom.value = Format(DateSerial(Year(Date), Month(Date), 1), "yyyy-mm-dd")
    txtTo.value = Format(Date, "yyyy-mm-dd")
End Sub

Private Sub txtFrom_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim d As Variant: d = PickDate(txtFrom.value)
    If IsDate(d) Then txtFrom.value = Format(d, "yyyy-mm-dd")
End Sub
Private Sub txtTo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim d As Variant: d = PickDate(txtTo.value)
    If IsDate(d) Then txtTo.value = Format(d, "yyyy-mm-dd")
End Sub

Private Sub btnGenerate_Click()
    Dim dFrom As Date, dTo As Date, rt As String, md As String
    If Not IsDate(txtFrom.value) Then MsgBox "Enter a valid From date.", vbExclamation: Exit Sub
    If Not IsDate(txtTo.value) Then MsgBox "Enter a valid To date.", vbExclamation: Exit Sub
    dFrom = CDate(txtFrom.value): dTo = CDate(txtTo.value)
    If dTo < dFrom Then MsgBox "'To' cannot be before 'From'.", vbExclamation: Exit Sub

    Select Case cboType.value
        Case "VAT": rt = "VAT"
        Case "Invoices": rt = "INVOICES"
        Case "Quotes": rt = "QUOTES"
        Case "Med Claims": rt = "MEDCLAIMS"
        Case "Credit Notes": rt = "CREDITNOTES"
        Case Else: rt = "SUMMARY"
    End Select
    md = IIf(optFull.value, "FULL", "SUMMARY")

    Me.Hide
    RunReport rt, md, dFrom, dTo, cboDept.value
    Unload Me
End Sub

Private Sub btnCancel_Click(): Unload Me
End Sub
