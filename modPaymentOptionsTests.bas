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
    AssertInvalid 1000000, 100, 0.5, 2, 4.345
    TestRequestNumbering
    TestPaymentAllocation
    TestSavedQuoteHeader
    MsgBox "Payment Options calculation, numbering, allocation and saved-source assertions completed.", vbInformation
End Sub

Private Sub TestRequestNumbering()
    Debug.Assert PaymentPlanBaseID("INV-WA-0260") = "PP-WA-0260"
    Debug.Assert PaymentPlanBaseID("INV-WD-0260") = "PP-WD-0260"
    Debug.Assert PaymentPlanBaseID("MC-INV-WA-0260") = "PP-MC-WA-0260"
    Debug.Assert PaymentRequestRef("PP-WA-0260", 0) = "PP-WA-0260-DEP"
    Debug.Assert PaymentRequestRef("PP-WA-0260", 1) = "PP-WA-0260-01"
    Debug.Assert PaymentRequestRef("PP-WA-0260-R02", 1) = "PP-WA-0260-R02-01"
    Debug.Assert PaymentRequestRef("PP-WA-0260", 100) = "PP-WA-0260-100"
    Debug.Assert PaymentPlanRevisionID("INV-WA-0260", 1) = "PP-WA-0260"
    Debug.Assert PaymentPlanRevisionID("INV-WA-0260", 2) = "PP-WA-0260-R02"
    Debug.Assert PaymentPlanRevisionID("INV-WA-0260", 3) = "PP-WA-0260-R03"
    Debug.Assert PaymentRequestWithinMailWindow(1, DateSerial(2026, 2, 8), DateSerial(2026, 2, 1), 7)
    Debug.Assert Not PaymentRequestWithinMailWindow(1, DateSerial(2026, 2, 9), DateSerial(2026, 2, 1), 7)
    Debug.Assert PaymentRequestWithinMailWindow(1, DateSerial(2026, 1, 31), DateSerial(2026, 2, 1), 7)
End Sub

Private Sub TestPaymentAllocation()
    Dim amounts(0 To 3) As Currency, paid(0 To 3) As Currency, excess As Currency
    amounts(0) = 50: amounts(1) = 20: amounts(2) = 20: amounts(3) = 10
    excess = AllocatePlanAmount(amounts, paid, 30)
    Debug.Assert excess = 0 And paid(0) = 30 And paid(1) = 0
    excess = AllocatePlanAmount(amounts, paid, 25)
    Debug.Assert excess = 0 And paid(0) = 50 And paid(1) = 5
    excess = AllocatePlanAmount(amounts, paid, 15, 2)
    Debug.Assert excess = 0 And paid(1) = 5 And paid(2) = 15
    excess = AllocatePlanAmount(amounts, paid, 50)
    Debug.Assert excess = 20
    Debug.Assert paid(0) + paid(1) + paid(2) + paid(3) = 100
    Debug.Assert paid(3) = 10
    ' Reversal / edit is a full rebuild, not an additive allocation.
    Erase paid
    excess = AllocatePlanAmount(amounts, paid, 20)
    Debug.Assert paid(0) = 20 And paid(1) = 0 And paid(2) = 0
    Erase paid
    amounts(0) = 0
    excess = AllocatePlanAmount(amounts, paid, 20)
    Debug.Assert paid(0) = 0 And paid(1) = 20
    AssertInvalidAllocation
End Sub

Private Sub AssertInvalidAllocation()
    Dim amounts(0 To 1) As Currency, paid(0 To 1) As Currency, excess As Currency
    On Error GoTo ExpectedError
    amounts(0) = 10: amounts(1) = 10
    excess = AllocatePlanAmount(amounts, paid, -1)
    Debug.Assert False
    Exit Sub
ExpectedError:
    Debug.Assert Err.Number <> 0
End Sub

Private Sub TestSavedQuoteHeader()
    Dim source(1 To 1, 1 To 24) As Variant, result As Variant, due As Date, issued As Date
    source(1, QL_NO) = "Q-WA-0260": source(1, QL_DEPT) = "WA"
    source(1, QL_RECIP) = "doctor": source(1, QL_CUST) = "C001"
    source(1, QL_DATE) = DateSerial(2026, 1, 10): source(1, QL_PATIENT) = "Original patient"
    source(1, QL_APPLIANCE) = "Original appliance"
    source(1, QL_SUBTOTAL) = 1000: source(1, QL_DISC) = 100
    source(1, QL_VAT) = 135: source(1, QL_TOTAL) = 1035
    source(1, QL_KIND) = "Saved kind": source(1, QL_DISCPCT) = 0.1: source(1, QL_DISCFIX) = 0
    source(1, QL_MEDAID) = "Original aid": source(1, QL_MEDNO) = "M001"
    source(1, QL_MAINMEM) = "Original member": source(1, QL_DOCTOR) = "Original doctor"
    source(1, QL_BHF) = "BHF001": source(1, QL_NOTES) = "Original saved notes"
    due = DateSerial(2026, 2, 10)
    issued = DateSerial(2026, 1, 20)
    result = modConvert.PaymentPlanQuoteHeader(source, "INV-WA-0300", due, issued)
    Debug.Assert result(1, IL_NO) = "INV-WA-0300" And result(1, IL_SRCQUOTE) = "Q-WA-0260"
    Debug.Assert result(1, IL_DATE) = issued And result(1, IL_DUE) = due
    Debug.Assert source(1, QL_DATE) = DateSerial(2026, 1, 10)
    Debug.Assert result(1, IL_SUBTOTAL) = 1000 And result(1, IL_DISC) = 100
    Debug.Assert result(1, IL_VAT) = 135 And result(1, IL_TOTAL) = 1035
    Debug.Assert result(1, IL_BALANCE) = 1035 And result(1, IL_PAID) = 0
    Debug.Assert result(1, IL_PATIENT) = source(1, QL_PATIENT)
    Debug.Assert result(1, IL_APPLIANCE) = source(1, QL_APPLIANCE)
    Debug.Assert result(1, IL_MEDAID) = source(1, QL_MEDAID)
    Debug.Assert result(1, IL_MEDNO) = source(1, QL_MEDNO)
    Debug.Assert result(1, IL_MAINMEM) = source(1, QL_MAINMEM)
    Debug.Assert result(1, IL_DOCTOR) = source(1, QL_DOCTOR)
    Debug.Assert result(1, IL_BHF) = source(1, QL_BHF)
    Debug.Assert result(1, IL_NOTES) = source(1, QL_NOTES)
    Debug.Assert result(1, IL_DISCPCT) = 0.1 And result(1, IL_DISCFIX) = 0
    Debug.Assert source(1, QL_NO) = "Q-WA-0260" And source(1, QL_TOTAL) = 1035
    source(1, QL_SUBTOTAL) = 900
    result = modConvert.PaymentPlanQuoteHeader(source, "INV-WA-0301", due, issued)
    Debug.Assert result(1, IL_SUBTOTAL) = 900 And result(1, IL_DISC) = 100 And result(1, IL_TOTAL) = 1035
    source(1, QL_SUBTOTAL) = 1000
    source(1, QL_RECIP) = "medclaim"
    source(1, QL_VAT) = 0: source(1, QL_TOTAL) = 900
    result = modConvert.PaymentPlanQuoteHeader(source, "INV-WA-0301", due, issued)
    Debug.Assert result(1, IL_RECIP) = "patient" And result(1, IL_VAT) = 0 And result(1, IL_TOTAL) = 900
    source(1, QL_TOTAL) = -1
    AssertInvalidQuoteHeader source, due
End Sub

Private Sub AssertInvalidQuoteHeader(ByVal source As Variant, ByVal due As Date)
    Dim result As Variant
    On Error GoTo ExpectedError
    result = modConvert.PaymentPlanQuoteHeader(source, "INV-WA-0302", due)
    Debug.Assert False
    Exit Sub
ExpectedError:
    Debug.Assert Err.Number <> 0
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
