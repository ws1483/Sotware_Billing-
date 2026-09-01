Attribute VB_Name = "modExportPDF"
Option Explicit
' ============================================================================
' ExportPDF — exports the ACTIVE document sheet (Quote/Invoice/CreditNote)
' to a PDF. User chooses the save location via a Save-As dialog.
' Suggested filename: [DocNo]_[PatientName].pdf (editable in the dialog).
' Opens the PDF after export. Overwrite is handled by the Save dialog.
'
' PRINT INTEGRITY FIX: before export we lock the print area to $A$1:$H$48 and
' force FitToPagesTall=1, so long wrapped descriptions can't push the footer /
' banking block onto a second page.
' ============================================================================

Public Sub ExportPDF()
    Dim ws As Worksheet, docNo As String, patient As String
    Dim suggested As String, chosen As Variant

    Set ws = ActiveSheet

    ' 1. only allow document sheets
    Select Case ws.Name
        Case "Quote", "Invoice", "CreditNote"
            ' ok
        Case Else
            MsgBox "Go to a Quote, Invoice, or Credit Note sheet before exporting.", vbExclamation
            Exit Sub
    End Select

    ' 2. gather naming pieces
    docNo = Trim(CStr(ws.Range("G7").Value))
    patient = Trim(CStr(ws.Range("F14").Value))
    If docNo = "" Then
        MsgBox "This document has no number yet — please Save it first.", vbExclamation
        Exit Sub
    End If
    If patient = "" Then patient = "NoName"

    ' 3. build suggested filename (sanitised)
    suggested = CleanName(docNo & "_" & patient) & ".pdf"

    ' 4. Save-As dialog — user picks folder + confirms/edits name
    chosen = Application.GetSaveAsFilename( _
        InitialFileName:=suggested, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save PDF As")

    If chosen = False Then Exit Sub          ' user cancelled

    ' 5. lock layout to ONE page (prevents footer spilling to page 2)
    FitDocToOnePage ws, "$A$1:$H$48"

    ' 6. export the sheet's print area to PDF
    On Error GoTo Fail
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=CStr(chosen), _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    LogAudit "ExportPDF", docNo, "", CStr(chosen), "Exported " & ws.Name
    Exit Sub
Fail:
    MsgBox "ExportPDF error: " & Err.Description, vbExclamation
End Sub

' ============================================================================
' Force a document sheet onto ONE page so tall wrapped rows can't push the
' footer/banking block off the printable area.
' ============================================================================
Private Sub FitDocToOnePage(ws As Worksheet, printAreaRef As String)
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = printAreaRef            ' locked boundary
        .Orientation = xlPortrait
        .Zoom = False                        ' MUST be False for FitToPages to work
        .FitToPagesWide = 1
        .FitToPagesTall = 1                  ' the key line: squeeze onto 1 page
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .CenterHorizontally = True
    End With
    Application.PrintCommunication = True
End Sub

' ---- strip characters illegal in Windows filenames ----
Private Function CleanName(s As String) As String
    Dim bad As Variant, ch As Variant, out As String
    out = s
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        out = Replace(out, ch, "-")
    Next ch
    out = Application.WorksheetFunction.Trim(out)
    CleanName = out
End Function
' ============================================================================
' PrintDoc — prints the ACTIVE document sheet (Quote/Invoice/CreditNote).
' Applies the same one-page fit as ExportPDF so long descriptions never push
' the footer/banking block onto a second page. Uses the default printer.
' ============================================================================
Public Sub PrintDoc()
    Dim ws As Worksheet, docNo As String

    Set ws = ActiveSheet

    ' 1. only allow document sheets
    Select Case ws.Name
        Case "Quote", "Invoice", "CreditNote"
            ' ok
        Case Else
            MsgBox "Go to a Quote, Invoice, or Credit Note sheet before printing.", vbExclamation
            Exit Sub
    End Select

    ' 2. must be a saved doc
    docNo = Trim(CStr(ws.Range("G7").Value))
    If docNo = "" Then
        MsgBox "This document has no number yet — please Save it first.", vbExclamation
        Exit Sub
    End If

    ' 3. confirm
    If MsgBox("Print " & ws.Name & " " & docNo & " now?", _
              vbQuestion + vbYesNo, "Print") <> vbYes Then Exit Sub

    ' 4. lock layout to ONE page (same as PDF export)
    FitDocToOnePage ws, "$A$1:$H$48"

    ' 5. send to printer
    On Error GoTo Fail
    ws.PrintOut Copies:=1, Collate:=True

    LogAudit "Print", docNo, "", "", "Printed " & ws.Name
    Exit Sub
Fail:
    MsgBox "PrintDoc error: " & Err.Description, vbExclamation
End Sub


