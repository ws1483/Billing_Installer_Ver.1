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
