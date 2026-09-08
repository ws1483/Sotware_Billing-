Attribute VB_Name = "modNameAudit"
Option Explicit
' Lists every defined name whose RefersTo points to an external URL/workbook.
Public Sub AuditBrokenNames()
    Dim nm As Name, msg As String, cnt As Long
    For Each nm In ThisWorkbook.names
        On Error Resume Next
        Dim rt As String
        rt = nm.RefersTo
        On Error GoTo 0
        If InStr(1, rt, "http", vbTextCompare) > 0 Or InStr(rt, "[") > 0 Then
            cnt = cnt + 1
            msg = msg & nm.Name & "  ->  " & rt & vbCrLf
        End If
    Next nm
    If cnt = 0 Then
        MsgBox "No broken (external) names found.", vbInformation
    Else
        MsgBox cnt & " broken name(s) point to external URLs/workbooks:" & vbCrLf & vbCrLf & msg, vbExclamation
    End If
End Sub
