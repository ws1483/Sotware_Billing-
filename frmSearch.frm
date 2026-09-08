VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "Search / Recall / Void"
   ClientHeight    =   5436
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8844.001
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' frmSearch � search across Invoices + Quotes + Credit Notes + Med Claims + Statements.
' ListBox columns: DocNo | Type | Doctor | Patient | Date | Total | Status
'
' TWO MODES:
'   Normal            -> Open recalls the document.
'   Pick-for-CN mode  -> Open returns the selected INVOICE to a new Credit Note
'                        via LoadCreditFromInvoice (set by ShowPickInvoiceForCN).

Private mPickForCN As Boolean

Public Sub ShowPickInvoiceForCN()
    mPickForCN = True
    Me.Show
End Sub

Private Sub UserForm_Initialize()
    cboType.AddItem "All"
    cboType.AddItem "Invoice"
    cboType.AddItem "Quote"
    cboType.AddItem "Credit Note"
    cboType.AddItem "Med Claim"
    cboType.AddItem "Statement"

    cboDoctor.AddItem "All"
    Dim wsC As Worksheet, lastR As Long, i As Long, nm As String
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets("Customers")
    If Not wsC Is Nothing Then
        lastR = wsC.Cells(wsC.Rows.Count, "B").End(xlUp).row
        For i = 2 To lastR
            nm = Trim(CStr(wsC.Cells(i, 2).value))
            If nm <> "" Then cboDoctor.AddItem nm
        Next i
    End If
    On Error GoTo 0
    cboDoctor.value = "All"

    If mPickForCN Then
        cboType.value = "Invoice"
        cboType.Enabled = False
        Me.Caption = "Pick Source Invoice for Credit Note"
    Else
        cboType.value = "All"
        cboType.Enabled = True
        Me.Caption = "Search / Recall / Void"
    End If

    lstResults.ColumnCount = 7
    RunSearch
End Sub

Private Sub txtSearch_Change():   RunSearch: End Sub
Private Sub cboDoctor_Change():   RunSearch: End Sub
Private Sub cboType_Change():     RunSearch: End Sub

Private Sub RunSearch()
    Dim term As String, docFilt As String, typeFilt As String
    term = LCase(Trim(txtSearch.value))
    docFilt = LCase(Trim(cboDoctor.value))
    typeFilt = LCase(Trim(cboType.value))

    lstResults.Clear

    If mPickForCN Then
        FillFromLog "InvoiceLog", "Invoice", term, docFilt
        Exit Sub
    End If

    If typeFilt = "all" Or typeFilt = "quote" Then FillFromLog "QuoteLog", "Quote", term, docFilt
    If typeFilt = "all" Or typeFilt = "invoice" Then FillFromLog "InvoiceLog", "Invoice", term, docFilt
    If typeFilt = "all" Or typeFilt = "credit note" Then FillFromCreditNotes term, docFilt
    If typeFilt = "all" Or typeFilt = "med claim" Then FillFromMedClaims term, docFilt
    If typeFilt = "all" Or typeFilt = "statement" Then FillFromStatements term, docFilt
End Sub

Private Sub FillFromLog(logName As String, docType As String, term As String, docFilt As String)
    Dim ws As Worksheet, last As Long, i As Long
    Dim cNo As Long, colDate As Long, cCust As Long, cPat As Long
    Dim cTotal As Long, cStatus As Long
    Dim docNo As String, drName As String, patient As String, status As String
    Dim total As Variant, dt As Variant, hay As String

    Set ws = ThisWorkbook.Sheets(logName)
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If last < 2 Then Exit Sub

    If logName = "InvoiceLog" Then
        cNo = 1: colDate = 4: cCust = 6: cPat = 7: cTotal = 12: cStatus = 16
    Else
        cNo = 1: colDate = 4: cCust = 5: cPat = 6: cTotal = 11: cStatus = 13
    End If

    For i = 2 To last
        docNo = Trim(CStr(ws.Cells(i, cNo).value))
        If docNo = "" Then GoTo NextI

        drName = CustIDToDrName(CStr(ws.Cells(i, cCust).value))
        patient = Trim(CStr(ws.Cells(i, cPat).value))
        status = Trim(CStr(ws.Cells(i, cStatus).value))
        total = ws.Cells(i, cTotal).value
        dt = ws.Cells(i, colDate).value

        If docFilt <> "" And docFilt <> "all" Then
            If LCase(drName) <> docFilt Then GoTo NextI
        End If

        If term <> "" Then
            hay = LCase(docNo & " " & patient & " " & drName)
            If InStr(hay, term) = 0 Then GoTo NextI
        End If

        With lstResults
            .AddItem docNo
            .List(.ListCount - 1, 1) = docType
            .List(.ListCount - 1, 2) = drName
            .List(.ListCount - 1, 3) = patient
            .List(.ListCount - 1, 4) = IIf(IsDate(dt), Format(dt, "yyyy-mm-dd"), "")
            .List(.ListCount - 1, 5) = IIf(IsNumeric(total), Format(total, "#,##0.00"), "")
            .List(.ListCount - 1, 6) = status
        End With
NextI:
    Next i
End Sub

' Credit Notes: A=CNNo, B=SrcInv, C=Date, D=CustID, E=Total, F=VAT, G=Reason, H=Status
Private Sub FillFromCreditNotes(term As String, docFilt As String)
    Dim ws As Worksheet, last As Long, i As Long
    Dim cnNo As String, srcInv As String, drName As String, status As String
    Dim total As Variant, dt As Variant, hay As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CreditNotes")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If last < 2 Then Exit Sub

    For i = 2 To last
        cnNo = Trim(CStr(ws.Cells(i, 1).value))
        If cnNo = "" Then GoTo NextI

        srcInv = Trim(CStr(ws.Cells(i, 2).value))
        drName = CustIDToDrName(CStr(ws.Cells(i, 4).value))
        status = Trim(CStr(ws.Cells(i, 8).value))
        total = ws.Cells(i, 5).value
        dt = ws.Cells(i, 3).value

        If docFilt <> "" And docFilt <> "all" Then
            If LCase(drName) <> docFilt Then GoTo NextI
        End If

        If term <> "" Then
            hay = LCase(cnNo & " " & srcInv & " " & drName)
            If InStr(hay, term) = 0 Then GoTo NextI
        End If

        With lstResults
            .AddItem cnNo
            .List(.ListCount - 1, 1) = "Credit Note"
            .List(.ListCount - 1, 2) = drName
            .List(.ListCount - 1, 3) = "src: " & srcInv
            .List(.ListCount - 1, 4) = IIf(IsDate(dt), Format(dt, "yyyy-mm-dd"), "")
            .List(.ListCount - 1, 5) = IIf(IsNumeric(total), Format(total, "#,##0.00"), "")
            .List(.ListCount - 1, 6) = status
        End With
NextI:
    Next i
End Sub

' Med Claims (MedAidLog): NO=1 Date=4 Patient=6 Total=11 Doctor=21(U) Status=27(AA)
Private Sub FillFromMedClaims(term As String, docFilt As String)
    Dim ws As Worksheet, last As Long, i As Long
    Dim mcNo As String, patient As String, drName As String, status As String
    Dim total As Variant, dt As Variant, hay As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("MedAidLog")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If last < 2 Then Exit Sub

    For i = 2 To last
        mcNo = Trim(CStr(ws.Cells(i, ML_NO).value))
        If mcNo = "" Then GoTo NextI

        patient = Trim(CStr(ws.Cells(i, ML_PATIENT).value))
        drName = Trim(CStr(ws.Cells(i, ML_DOCTOR).value))     ' U = Treating Doctor
        status = Trim(CStr(ws.Cells(i, ML_STATUS).value))
        total = ws.Cells(i, ML_TOTAL).value
        dt = ws.Cells(i, ML_DATE).value

        If docFilt <> "" And docFilt <> "all" Then
            If LCase(drName) <> docFilt Then GoTo NextI
        End If

        If term <> "" Then
            hay = LCase(mcNo & " " & patient & " " & drName)
            If InStr(hay, term) = 0 Then GoTo NextI
        End If

        With lstResults
            .AddItem mcNo
            .List(.ListCount - 1, 1) = "Med Claim"
            .List(.ListCount - 1, 2) = drName
            .List(.ListCount - 1, 3) = patient
            .List(.ListCount - 1, 4) = IIf(IsDate(dt), Format(dt, "yyyy-mm-dd"), "")
            .List(.ListCount - 1, 5) = IIf(IsNumeric(total), Format(total, "#,##0.00"), "")
            .List(.ListCount - 1, 6) = status
        End With
NextI:
    Next i
End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnOpen_Click
End Sub

Private Sub btnOpen_Click()
    If lstResults.ListIndex < 0 Then
        MsgBox "Select a document first.", vbExclamation: Exit Sub
    End If
    Dim docNo As String, docType As String
    docNo = CStr(lstResults.List(lstResults.ListIndex, 0))
    docType = LCase(CStr(lstResults.List(lstResults.ListIndex, 1)))

    If mPickForCN Then
        If docType <> "invoice" Then
            MsgBox "Please select an Invoice.", vbExclamation: Exit Sub
        End If
        mPickForCN = False
        Unload Me
        LoadCreditFromInvoice docNo
        Exit Sub
    End If

    Unload Me
    Select Case docType
        Case "invoice":     RecallInvoice docNo
        Case "quote":       RecallQuote docNo
        Case "credit note": RecallCreditNote docNo
        Case "med claim":   RecallMedClaim docNo
        Case "statement":   OpenStatementPDF docNo
    End Select
End Sub

Private Sub btnVoid_Click()
    If mPickForCN Then Exit Sub
    If lstResults.ListIndex < 0 Then
        MsgBox "Select a document first.", vbExclamation: Exit Sub
    End If
    Dim docNo As String, docType As String
    docNo = CStr(lstResults.List(lstResults.ListIndex, 0))
    docType = LCase(CStr(lstResults.List(lstResults.ListIndex, 1)))
    If docType = "statement" Then
        MsgBox "Statements cannot be voided (they are reprints). Void the underlying invoices instead.", vbExclamation
        Exit Sub
    End If
    If docType = "credit note" Then
        If VoidCreditNote(docNo) Then RunSearch
        Exit Sub
    End If
    If docType = "med claim" Then
        MsgBox "Med Claims cannot be voided here.", vbExclamation
        Exit Sub
    End If
    If VoidDocument(docNo, docType) Then RunSearch
End Sub

Private Sub btnClose_Click()
    mPickForCN = False
    Unload Me
End Sub

Private Sub FillFromStatements(term As String, docFilt As String)
    Dim ws As Worksheet, last As Long, i As Long
    Dim stmtNo As String, drName As String, dept As String
    Dim total As Variant, dt As Variant, hay As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("StatementLog")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If last < 2 Then Exit Sub

    For i = 2 To last
        stmtNo = Trim(CStr(ws.Cells(i, 1).value))
        If stmtNo = "" Then GoTo NextI

        drName = Trim(CStr(ws.Cells(i, 4).value))
        dept = Trim(CStr(ws.Cells(i, 5).value))
        total = ws.Cells(i, 11).value
        dt = ws.Cells(i, 2).value

        If docFilt <> "" And docFilt <> "all" Then
            If LCase(drName) <> docFilt Then GoTo NextI
        End If

        If term <> "" Then
            hay = LCase(stmtNo & " " & drName)
            If InStr(hay, term) = 0 Then GoTo NextI
        End If

        With lstResults
            .AddItem stmtNo
            .List(.ListCount - 1, 1) = "Statement"
            .List(.ListCount - 1, 2) = drName
            .List(.ListCount - 1, 3) = "(" & dept & ")"
            .List(.ListCount - 1, 4) = IIf(IsDate(dt), Format(dt, "yyyy-mm-dd"), "")
            .List(.ListCount - 1, 5) = IIf(IsNumeric(total), Format(total, "#,##0.00"), "")
            .List(.ListCount - 1, 6) = "Statement"
        End With
NextI:
    Next i
End Sub

