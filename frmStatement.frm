VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStatement 
   Caption         =   "frmStatement"
   ClientHeight    =   5796
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6780
   OleObjectBlob   =   "frmStatement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub OptionButton2_Click()

End Sub

Private Sub UserForm_Initialize()
    ' department list
    cboDept.Clear
    cboDept.AddItem "All": cboDept.AddItem "WA": cboDept.AddItem "WD"
    cboDept.ListIndex = 0

    ' recipient mode default = Doctor
    optRecDoctor.Value = True
    LoadRecipientList False           ' load doctors into cboDoctor

    ' dates
    txtFrom.Value = Format(DateSerial(Year(Date), Month(Date), 1), "yyyy-mm-dd")
    txtTo.Value = Format(Date, "yyyy-mm-dd")

    ' single vs batch default = Single
    optSingle.Value = True
    cboDoctor.Enabled = True
End Sub

' ---------------- Single / Batch toggle ----------------
Private Sub optSingle_Click()
    cboDoctor.Enabled = True
End Sub
Private Sub optBatch_Click()
    cboDoctor.Enabled = False
End Sub

' ---------------- Doctor / Patient toggle ----------------
Private Sub optRecDoctor_Click()
    LoadRecipientList False
End Sub
Private Sub optRecPatient_Click()
    LoadRecipientList True
End Sub

' Populate cboDoctor with either doctors (Customers) or patients (InvoiceLog)
Private Sub LoadRecipientList(asPatient As Boolean)
    Dim wsC As Worksheet, last As Long, i As Long, pats As Collection
    cboDoctor.Clear
    If asPatient Then
        Set pats = PatientNamesList()                 ' from modStatement
        For i = 1 To pats.Count
            cboDoctor.AddItem pats(i)
        Next i
    Else
        Set wsC = ThisWorkbook.Sheets("Customers")
        last = wsC.Cells(wsC.Rows.Count, "B").End(xlUp).row
        For i = 2 To last
            If Trim(CStr(wsC.Cells(i, "B").Value)) <> "" Then _
                cboDoctor.AddItem wsC.Cells(i, "B").Value
        Next i
    End If
    cboDoctor.ListIndex = -1
End Sub

' ---------------- Date pickers (double-click) ----------------
Private Sub txtFrom_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim d As Variant
    d = PickDate(txtFrom.Value)
    If IsDate(d) Then txtFrom.Value = Format(d, "yyyy-mm-dd")
End Sub

Private Sub txtTo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim d As Variant
    d = PickDate(txtTo.Value)
    If IsDate(d) Then txtTo.Value = Format(d, "yyyy-mm-dd")
End Sub

' ---------------- Buttons ----------------
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnGenerate_Click()
    Dim dFrom As Date, dTo As Date, dept As String

    ' validate dates
    If Not IsDate(txtFrom.Value) Then MsgBox "Enter a valid From date (yyyy-mm-dd).", vbExclamation: Exit Sub
    If Not IsDate(txtTo.Value) Then MsgBox "Enter a valid To date (yyyy-mm-dd).", vbExclamation: Exit Sub
    dFrom = CDate(txtFrom.Value): dTo = CDate(txtTo.Value)
    If dTo < dFrom Then MsgBox "'To' cannot be before 'From'.", vbExclamation: Exit Sub

    dept = cboDept.Value: If dept = "" Then dept = "All"

    ' single mode requires a selection
    If optSingle.Value And cboDoctor.ListIndex < 0 Then
        MsgBox "Select a " & IIf(optRecPatient.Value, "patient", "doctor") & ".", vbExclamation
        Exit Sub
    End If

    Me.Hide

    If optRecPatient.Value Then
        ' ----- PATIENT statements -----
        If optSingle.Value Then
            RunSinglePatientStatement cboDoctor.Value, dFrom, dTo, dept
        Else
            RunBatchPatientStatements dFrom, dTo, dept
        End If
    Else
        ' ----- DOCTOR statements -----
        If optSingle.Value Then
            RunSingleStatement cboDoctor.Value, dFrom, dTo, dept
        Else
            RunBatchStatements dFrom, dTo, dept
        End If
    End If

    Unload Me
End Sub



