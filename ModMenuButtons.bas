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
Public Sub BtnNewCreditNote()
    MenuNewCreditNote
End Sub

Public Sub BtnPaymentOptions()
    On Error GoTo Failed
    VBA.UserForms.Add("frmPaymentOptions").Show
    Exit Sub
Failed:
    MsgBox "Payment Options is unavailable. Run BuildPaymentOptionsForm once in this workbook. " & _
           Err.Description, vbExclamation, "Payment Options"
End Sub

Public Sub BtnDepositsStatement()
    RunDepositsStatement
End Sub

Public Sub BtnInstallmentsStatement()
    RunInstallmentsStatement
End Sub

Public Sub BtnPaymentPlanStatement()
    RunSinglePlanStatement
End Sub

Public Sub BtnSavePaymentPlanPDFs()
    Dim id As String
    id = Trim$(InputBox("Enter the Plan ID:", "Save payment invoices"))
    If id <> "" Then modPaymentOptions.SavePlanInvoicePDFs id
End Sub
