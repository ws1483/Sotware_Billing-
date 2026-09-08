Attribute VB_Name = "modMedClaimSetup_auto"
Option Explicit
' ============================================================================
' Med Claim dropdowns ? merge-safe, using the CONFIRMED named ranges:
'   Description = PriceDesc   Doctor = CustNameList   Appliance = ApplianceList
' ============================================================================
Public Sub SetupMedClaimDropdowns()
    Dim wsM As Worksheet, r As Long
    Set wsM = ThisWorkbook.Sheets("Med Claim")

    ApplyListSafe wsM.Range("C14"), "=ApplianceList"     ' Appliance Type
    ApplyListSafe wsM.Range("G12"), "=CustNameList"      ' Treating Doctor
    For r = 16 To 30
        ApplyListSafe wsM.Range("D" & r), "=PriceDesc"   ' Descriptions
    Next r

    MsgBox "Med Claim dropdowns installed.", vbInformation
End Sub

Private Sub ApplyListSafe(rng As Range, formula1 As String)
    Dim anchor As Range
    Set anchor = rng.MergeArea.Cells(1, 1)
    On Error Resume Next
    anchor.Validation.Delete
    On Error GoTo 0
    On Error Resume Next
    anchor.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, formula1:=formula1
    anchor.Validation.IgnoreBlank = True
    anchor.Validation.InCellDropdown = True
    If Err.Number <> 0 Then
        MsgBox "Could not apply list to " & anchor.Address & vbCrLf & _
               "Source: " & formula1 & vbCrLf & "Error: " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error GoTo 0
End Sub

