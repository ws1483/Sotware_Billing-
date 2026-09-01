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
' frmSearch — live search across Invoices + Quotes; Open or Void a doc.
' ListBox columns: DocNo | Type | Doctor | Patient | Date | Total | Status

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Initialize()
    ' Type dropdown
    cboType.AddItem "All"
    cboType.AddItem "Invoice"
    cboType.AddItem "Quote"
    cboType.AddItem "Statement"
    cboType.Value = "All"

    ' Doctor dropdown from Customers sheet (col B = DrName; adjust if different)
    cboDoctor.AddItem "All"
    Dim wsC As Worksheet, lastR As Long, i As Long, nm As String
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets("Customers")
    If Not wsC Is Nothing Then
        lastR = wsC.Cells(wsC.Rows.Count, "B").End(xlUp).row
        For i = 2 To lastR
            nm = Trim(CStr(wsC.Cells(i, 2).Value))
            If nm <> "" Then cboDoctor.AddItem nm
        Next i
    End If
    On Error GoTo 0
    cboDoctor.Value = "All"

    lstResults.ColumnCount = 7
    RunSearch
End Sub

Private Sub txtSearch_Change():   RunSearch: End Sub
Private Sub cboDoctor_Change():   RunSearch: End Sub
Private Sub cboType_Change():     RunSearch: End Sub

Private Sub RunSearch()
    Dim term As String, docFilt As String, typeFilt As String
    term = LCase(Trim(txtSearch.Value))
    docFilt = LCase(Trim(cboDoctor.Value))
    typeFilt = LCase(Trim(cboType.Value))

    lstResults.Clear

    If typeFilt = "all" Or typeFilt = "quote" Then FillFromLog "QuoteLog", "Quote", term, docFilt
    If typeFilt = "all" Or typeFilt = "invoice" Then FillFromLog "InvoiceLog", "Invoice", term, docFilt
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
        docNo = Trim(CStr(ws.Cells(i, cNo).Value))
        If docNo = "" Then GoTo NextI

        drName = CustIDToDrName(CStr(ws.Cells(i, cCust).Value))
        patient = Trim(CStr(ws.Cells(i, cPat).Value))
        status = Trim(CStr(ws.Cells(i, cStatus).Value))
        total = ws.Cells(i, cTotal).Value
        dt = ws.Cells(i, colDate).Value

        ' doctor filter
        If docFilt <> "" And docFilt <> "all" Then
            If LCase(drName) <> docFilt Then GoTo NextI
        End If

        ' text search (contains): docNo + patient + doctor
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
    Unload Me
    Select Case docType
        Case "invoice":   RecallInvoice docNo
        Case "quote":     RecallQuote docNo
        Case "statement": OpenStatementPDF docNo
    End Select
End Sub

Private Sub btnVoid_Click()
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
    If VoidDocument(docNo, docType) Then RunSearch
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
Private Sub FillFromStatements(term As String, docFilt As String)
    Dim ws As Worksheet, last As Long, i As Long
    Dim stmtNo As String, drName As String, dept As String, status As String
    Dim total As Variant, dt As Variant, hay As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("StatementLog")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If last < 2 Then Exit Sub

    For i = 2 To last
        stmtNo = Trim(CStr(ws.Cells(i, 1).Value))       ' A StmtNo
        If stmtNo = "" Then GoTo NextI

        drName = Trim(CStr(ws.Cells(i, 4).Value))        ' D Doctor
        dept = Trim(CStr(ws.Cells(i, 5).Value))          ' E Dept
        total = ws.Cells(i, 11).Value                    ' K BalanceDue
        dt = ws.Cells(i, 2).Value                        ' B DateGenerated

        ' doctor filter
        If docFilt <> "" And docFilt <> "all" Then
            If LCase(drName) <> docFilt Then GoTo NextI
        End If

        ' text search (contains): stmtNo + doctor
        If term <> "" Then
            hay = LCase(stmtNo & " " & drName)
            If InStr(hay, term) = 0 Then GoTo NextI
        End If

        With lstResults
            .AddItem stmtNo
            .List(.ListCount - 1, 1) = "Statement"
            .List(.ListCount - 1, 2) = drName
            .List(.ListCount - 1, 3) = "(" & dept & ")"      ' show dept in patient col
            .List(.ListCount - 1, 4) = IIf(IsDate(dt), Format(dt, "yyyy-mm-dd"), "")
            .List(.ListCount - 1, 5) = IIf(IsNumeric(total), Format(total, "#,##0.00"), "")
            .List(.ListCount - 1, 6) = "Statement"
        End With
NextI:
    Next i
End Sub

