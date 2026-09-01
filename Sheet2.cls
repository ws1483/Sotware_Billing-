VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal target As Range)
    Dim isect As Range, c As Range

    ' recipient toggle
    If Not Application.Intersect(target, Me.Range("K2")) Is Nothing Then
        SetRecipientType Me, CStr(Me.Range("K2").Value)
        Exit Sub
    End If

    ' patient-mode: Doctor picked (G11) -> auto-fill BHF (G12)
    If Not Application.Intersect(target, Me.Range("G11")) Is Nothing Then
        If LCase(Trim(CStr(Me.Range("K2").Value))) = "patient" Then
            Application.EnableEvents = False
            Me.Range("G12").Value = LookupBHFByDrName(CStr(Me.Range("G11").Value))
            Application.EnableEvents = True
        End If
        Exit Sub
    End If

    ' price auto-fill
    Set isect = Application.Intersect(target, Me.Range("D16:D30"))
    If isect Is Nothing Then Exit Sub
    Application.EnableEvents = False
    On Error GoTo Clean
    For Each c In isect.Cells
        FillLinePrice Me, c.row
    Next c
Clean:
    Application.EnableEvents = True
End Sub

