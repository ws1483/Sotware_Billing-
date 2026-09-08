Attribute VB_Name = "modNameRepair"
Option Explicit
' ============================================================================
' Repairs named ranges whose RefersTo points to the external OneDrive copy
' (WeDental_Ver 3 Test.xlsm). Strips the external workbook prefix so each name
' points to the SAME sheet/range inside THIS workbook.
'   RUN AuditBrokenNames FIRST to see them; run this to fix; audit again to verify.
' ============================================================================
Public Sub RepairBrokenNames()
    Dim nm As Name, rt As String, fixed As String
    Dim cnt As Long, failed As Long, report As String

    For Each nm In ThisWorkbook.names
        rt = ""
        On Error Resume Next
        rt = nm.RefersTo
        On Error GoTo 0

        If InStr(1, rt, "http", vbTextCompare) > 0 Or InStr(rt, "[") > 0 Then
            fixed = StripExternal(rt)
            If fixed <> "" And fixed <> rt Then
                On Error Resume Next
                nm.RefersTo = fixed
                If Err.Number <> 0 Then
                    failed = failed + 1
                    report = report & "FAILED: " & nm.Name & " -> " & fixed & vbCrLf
                    Err.Clear
                Else
                    cnt = cnt + 1
                End If
                On Error GoTo 0
            End If
        End If
    Next nm

    MsgBox "Repaired " & cnt & " name(s)." & IIf(failed > 0, vbCrLf & failed & " failed:" & vbCrLf & report, ""), _
           IIf(failed > 0, vbExclamation, vbInformation)
End Sub

' Turn ='https://.../[WeDental_Ver 3 Test.xlsm]Reference'!$A:$A  into  =Reference!$A:$A
Private Function StripExternal(ByVal rt As String) As String
    Dim s As String, p1 As Long, p2 As Long
    s = rt
    ' remove everything from the first quote-URL up to and including the ]workbook]
    If InStr(s, "[") > 0 And InStr(s, "]") > 0 Then
        p1 = InStr(s, "=")            ' keep leading =
        ' find the "]" that closes the [workbook.xlsm]
        p2 = InStr(s, "]")
        ' rebuild:  "=" & everything AFTER the "]"
        s = "=" & Mid$(s, p2 + 1)
    End If
    ' clean stray quotes:  ='Reference'!  -> =Reference!
    s = Replace(s, "='", "=")
    s = Replace(s, "'!", "!")
    StripExternal = s
End Function

