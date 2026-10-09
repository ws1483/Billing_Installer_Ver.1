Attribute VB_Name = "modPaymentOptionsTests"
Option Explicit

' Import optionally; run TestPaymentOptions from the VBE on a workbook copy.
Public Sub TestPaymentOptions()
    Dim p As PaymentPlanResult, i As Long, sum As Currency, n As Long
    p = CalcPaymentPlan(10, 20000, 0.5, 2, 4.345)
    Debug.Assert p.TreatmentWeeks = 20
    Debug.Assert p.InstallmentCount = 5
    Debug.Assert p.Deposit = 10000
    Debug.Assert p.MonthlyAmount = 2000
    For i = 0 To p.InstallmentCount
        sum = sum + p.Schedule(i)
    Next i
    Debug.Assert sum = 20000

    p = CalcPaymentPlan(3, 100, 0.5, 2, 4.345)
    Debug.Assert p.InstallmentCount = 1
    AssertSchedule p, 100

    p = CalcPaymentPlan(10, 100.01, 0.5, 2, 4.345)
    AssertSchedule p, 100.01
    p = CalcPaymentPlan(1, 100, 0, 2, 4.345)
    Debug.Assert p.InstallmentCount = 1
    Debug.Assert p.Deposit = 0
    AssertSchedule p, 100
    p = CalcPaymentPlan(10, 100, 1, 2, 4.345)
    Debug.Assert p.Deposit = 100
    AssertSchedule p, 100
    p = CalcPaymentPlan(100, 0.01, 0, 2, 4.345)
    AssertSchedule p, 0.01
    For n = 1 To 100
        p = CalcPaymentPlan(n, CCur(n) / 100, 0.5, 2, 4.345)
        AssertSchedule p, CCur(n) / 100
    Next n
    AssertInvalid -1, 100, 0.5, 2, 4.345
    AssertInvalid 10, -100, 0.5, 2, 4.345
    AssertInvalid 10, 100, 1.1, 2, 4.345
    AssertInvalid 10, 100, 0.5, 0, 4.345
    AssertInvalid 10, 100, 0.5, 2, 0
    MsgBox "Payment Options calculation assertions completed.", vbInformation
End Sub

Private Sub AssertSchedule(p As PaymentPlanResult, expected As Currency)
    Dim i As Long, sum As Currency
    Debug.Assert p.InstallmentCount >= 1
    For i = 0 To p.InstallmentCount
        Debug.Assert p.Schedule(i) >= 0
        Debug.Assert p.Schedule(i) * 100 = Fix(p.Schedule(i) * 100)
        sum = sum + p.Schedule(i)
    Next i
    Debug.Assert sum = expected
End Sub

Private Sub AssertInvalid(aligners As Double, total As Currency, pct As Double, weeks As Double, divisor As Double)
    Dim p As PaymentPlanResult, raised As Boolean
    On Error GoTo ExpectedError
    p = CalcPaymentPlan(aligners, total, pct, weeks, divisor)
    Debug.Assert False
    Exit Sub
ExpectedError:
    raised = (Err.Number <> 0)
    Debug.Assert raised
End Sub

' Build first; run on a workbook copy with trusted VBA project access.
Public Sub TestPaymentOptionsForm()
    Dim component As Object, form As Object, name As Variant, source As String
    On Error GoTo Failed
    Set component = ThisWorkbook.VBProject.VBComponents("frmPaymentOptions")
    Debug.Assert component.Type = 3
    Set form = component.Designer
    For Each name In Array("poSource", "poRecipient", "poDept", "poAligners", "poTotal", _
                          "poPlan", "poPreview", "poCustomer", "poGenerate", "poCancelPlan", _
                          "poAmend", "poDeposits", "poInstallments", "poSingle", "poMail", "poClose")
        Debug.Assert form.Controls(CStr(name)).Name = CStr(name)
    Next name
    Debug.Assert form.BackColor = &H00B0872E&
    Debug.Assert form.ForeColor = &H80000012&
    Debug.Assert form.Font.Name = "Tahoma"
    Debug.Assert form.BorderStyle = 0
    Debug.Assert form.Cycle = 0
    Debug.Assert form.KeepScrollBarsVisible = 3
    Debug.Assert form.SpecialEffect = 0
    Debug.Assert component.Properties("StartUpPosition") = 1
    Debug.Assert component.Properties("ShowModal") = True
    Debug.Assert TypeName(form.Controls("poPreview")) = "TextBox"
    Debug.Assert form.Controls("poPreview").MultiLine
    Debug.Assert form.Controls("poPreview").ScrollBars = 3
    source = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
    Debug.Assert Left$(source, Len("Option Explicit")) = "Option Explicit"
    Debug.Assert InStr(source, "Begin {") = 0
    For Each name In Array("CalcPaymentPlan", "GeneratePlanInvoices", "CancelPaymentPlan", _
                          "AmendPaymentPlan", "RunDepositsStatement", "RunInstallmentsStatement", _
                          "RunSinglePlanStatement", "MailInstallmentInvoice")
        Debug.Assert InStr(source, "modPaymentOptions." & CStr(name)) > 0
    Next name
    Set form = VBA.UserForms.Add("frmPaymentOptions")
    Debug.Assert form.Controls("poRecipient").ListCount = 3
    Debug.Assert form.Controls("poDept").ListCount = 4
    form.Controls("poAligners").Value = 10
    form.Controls("poTotal").Value = 20000
    Debug.Assert InStr(form.Controls("poPlan").Value, "deposit;") > 0
    Debug.Assert InStr(form.Controls("poPreview").Value, "Deposit") > 0
    Debug.Assert InStr(form.Controls("poPreview").Value, "Month 1") > 0
    form.Controls("poTotal").Value = 0
    Debug.Assert Left$(form.Controls("poPlan").Value, 8) = "Preview:"
    Debug.Assert form.Controls("poPreview").Value = ""
    Unload form
    MsgBox "Payment Options form assertions completed.", vbInformation
    Exit Sub
Failed:
    MsgBox "Form test error: " & Err.Description & vbCrLf & _
           "Enable trusted VBA project access and run BuildPaymentOptionsForm first.", vbExclamation
    On Error Resume Next
    If Not form Is Nothing Then Unload form
End Sub
