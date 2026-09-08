Attribute VB_Name = "ModMenuButtons"
Option Explicit
' ============================================================================
' modMenuButtons ? no-argument wrappers so these are assignable to buttons
' and always visible in the Alt+F8 macro list.
' ============================================================================

Public Sub BtnRecallMedClaim()
    RecallMedClaim            ' no arg -> prompts via InputBox
End Sub

Public Sub BtnRevertQuote()
    RevertQuoteToSaved        ' no arg -> prompts via InputBox
End Sub

Public Sub BtnNewMedClaim()
    MenuNewMedClaim
End Sub

Public Sub BtnSaveMedClaim()
    SaveMedClaim
End Sub

Public Sub BtnRecallInvoice()
    RecallInvoice             ' existing; wrapper for consistency
End Sub

Public Sub BtnRecallQuote()
    RecallQuote
End Sub
