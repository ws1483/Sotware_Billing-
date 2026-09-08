VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim isect As Range, c As Range

    ' Bill To (C6) changed -> autofill patient + med-aid details
    If Not Application.Intersect(Target, Me.Range("C6")) Is Nothing Then
        MedClaimAutofill Me
        Exit Sub
    End If

    ' Treating Doctor (G12) picked -> auto-fill BHF (G13)
    If Not Application.Intersect(Target, Me.Range("G12")) Is Nothing Then
        Dim bhf As String
        bhf = LookupBHFByDrName(CStr(Me.Range("G12").value))
        Application.EnableEvents = False
        ' write to the merge-area anchor so merged G13:H13 accepts it
        Me.Range("G13").MergeArea.Cells(1, 1).value = bhf
        Application.EnableEvents = True
        Exit Sub
    End If

    ' price auto-fill on description entry (D16:D30)
    Set isect = Application.Intersect(Target, Me.Range("D16:D30"))
    If isect Is Nothing Then Exit Sub
    Application.EnableEvents = False
    On Error GoTo Clean
    For Each c In isect.Cells
        FillLinePrice Me, c.row
    Next c
Clean:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' date picker on document date G6
    If Not Application.Intersect(Target, Me.Range("G6")) Is Nothing Then
        Cancel = True
        Dim d As Variant
        d = PickDate(Target.value)
        If IsDate(d) Then
            Application.EnableEvents = False
            Target.value = d
            Application.EnableEvents = True
        End If
    End If
End Sub
