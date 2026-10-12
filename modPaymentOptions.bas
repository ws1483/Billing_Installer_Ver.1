Attribute VB_Name = "modPaymentOptions"
Option Explicit

Private mWriting As Boolean
Private mRefreshing As Boolean

Public Function CalcPaymentPlan(ByVal aligners As Double, ByVal total As Currency, _
        Optional ByVal depositPercent As Double = -1, Optional ByVal weeksPerAligner As Double = -1, _
        Optional ByVal weeksPerMonth As Double = -1) As PaymentPlanResult
    Dim p As PaymentPlanResult, cents As Variant, depositCents As Variant
    Dim remainingCents As Variant, monthlyCents As Variant, term As Double, i As Long
    If total <= 0 Or aligners <= 0 Then Err.Raise vbObjectError + 710, , "Total and aligners must be positive."
    If aligners <> Fix(aligners) Or aligners > 2147483647# Then _
        Err.Raise vbObjectError + 710, , "Aligners must be a positive whole number."
    If depositPercent = -1 Then depositPercent = GetDepositPercent()
    If weeksPerAligner = -1 Then weeksPerAligner = GetWeeksPerAligner()
    If weeksPerMonth = -1 Then weeksPerMonth = GetWeeksPerMonth()
    If depositPercent < 0 Or depositPercent > 1 Then Err.Raise vbObjectError + 711, , "Deposit percent must be between 0 and 1."
    If weeksPerAligner <= 0 Or weeksPerMonth <= 0 Then Err.Raise vbObjectError + 712, , "Weeks settings must be positive."
    cents = Fix(CDec(total) * 100 + CDec(0.5))
    If cents < 1 Then Err.Raise vbObjectError + 713, , "Total must be at least one cent."
    depositCents = Fix(CDec(cents) * CDec(depositPercent) + CDec(0.5))
    remainingCents = cents - depositCents
    p.TreatmentWeeks = CDbl(aligners) * weeksPerAligner
    term = p.TreatmentWeeks / weeksPerMonth
    If term > 1200 Then Err.Raise vbObjectError + 714, , "Payment term exceeds 100 years."
    p.InstallmentCount = CLng(Round(term, 0))
    If p.InstallmentCount < 1 Then p.InstallmentCount = 1
    monthlyCents = Fix(CDec(remainingCents) / p.InstallmentCount + CDec(0.5))
    If monthlyCents * (p.InstallmentCount - 1) > remainingCents Then _
        monthlyCents = Fix(CDec(remainingCents) / p.InstallmentCount)
    p.Deposit = depositCents / 100
    p.MonthlyAmount = monthlyCents / 100
    ReDim p.Schedule(0 To p.InstallmentCount)
    p.Schedule(0) = p.Deposit
    For i = 1 To p.InstallmentCount
        p.Schedule(i) = p.MonthlyAmount
    Next i
    p.Schedule(p.InstallmentCount) = (remainingCents - monthlyCents * (p.InstallmentCount - 1)) / 100
    CalcPaymentPlan = p
End Function

Private Function SettingNumber(ByVal key As String, ByVal fallback As Double, _
        ByVal minimum As Double, ByVal maximum As Double, Optional ByVal wholeNumber As Boolean = False) As Double
    Dim ws As Worksheet, r As Long, found As Long, v As Variant
    Set ws = ThisWorkbook.Worksheets("Settings")
    For r = 30 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If NrmID(CStr(ws.Cells(r, 1).Value)) = NrmID(key) Then
            If found > 0 Then Err.Raise vbObjectError + 715, , "Duplicate Settings key: " & key
            found = r
        End If
    Next r
    If found = 0 Then
        found = Application.Max(30, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1, _
            ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 1)
        ws.Cells(found, 1).Value = key
        ws.Cells(found, 2).Value = fallback
    End If
    v = ws.Cells(found, 2).Value
    If IsError(v) Then Err.Raise vbObjectError + 716, , "Invalid setting: " & key
    If Not IsNumeric(v) Or Trim$(CStr(v)) = "" Then Err.Raise vbObjectError + 716, , "Invalid setting: " & key
    SettingNumber = CDbl(v)
    If SettingNumber < minimum Or SettingNumber > maximum Then Err.Raise vbObjectError + 716, , "Out-of-range setting: " & key
    If wholeNumber Then
        If SettingNumber <> Fix(SettingNumber) Then Err.Raise vbObjectError + 716, , "Setting must be whole: " & key
    End If
End Function

Public Function GetDepositPercent(Optional ByVal fallback As Double = 0.5) As Double
    GetDepositPercent = SettingNumber("DepositPercent", fallback, 0, 1)
End Function
Public Function GetWeeksPerAligner(Optional ByVal fallback As Double = 2) As Double
    GetWeeksPerAligner = SettingNumber("WeeksPerAligner", fallback, 0, 1E+100)
    If GetWeeksPerAligner <= 0 Then Err.Raise vbObjectError + 716, , "WeeksPerAligner must be positive."
End Function
Public Function GetWeeksPerMonth(Optional ByVal fallback As Double = 4.345) As Double
    GetWeeksPerMonth = SettingNumber("WeeksPerMonth", fallback, 0, 1E+100)
    If GetWeeksPerMonth <= 0 Then Err.Raise vbObjectError + 716, , "WeeksPerMonth must be positive."
End Function
Public Function GetReminderLeadDays(Optional ByVal fallback As Long = 7) As Long
    GetReminderLeadDays = CLng(SettingNumber("ReminderLeadDays", fallback, 0, 2147483647#, True))
End Function

Public Function EnsurePaymentPlansSheet() As Worksheet
    Dim ws As Worksheet, headers As Variant, i As Long, populated As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PaymentPlans")
    On Error GoTo 0
    headers = Array("PlanID", "SourceDocNo", "RecipientType", "CustID", "Dept", _
        "Total", "Deposit", "InstallmentNo", "InstallmentCount", "DueDate", "Amount", _
        "InvoiceNo", "Generated", "Status", "Reminded", "CreatedAt", "ModifiedAt", _
        "PatientName", "PDFPath", "Aligners", "Paid", "Balance", _
        "RequestRef", "SchemaVersion", "SendState", "DepositPercent")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "PaymentPlans"
    End If
    populated = (LastPlanRow(ws) > 1)
    For i = 0 To UBound(headers)
        If CStr(ws.Cells(1, i + 1).Value) <> CStr(headers(i)) Then
            If populated Or Trim$(CStr(ws.Cells(1, i + 1).Value)) <> "" Then _
                Err.Raise vbObjectError + 717, , "Incompatible PaymentPlans layout. Legacy financial plans are not supported."
            ws.Cells(1, i + 1).Value = headers(i)
        End If
    Next i
    For i = 2 To LastPlanRow(ws)
        If CStr(ws.Cells(i, PP_SCHEMA).Value) <> "2" Then _
            Err.Raise vbObjectError + 717, , "Legacy payment plan detected. No automatic financial migration is permitted."
    Next i
    If Not ws.ProtectContents Then
        ws.Columns(PP_DUE).NumberFormat = "yyyy-mm-dd"
        ws.Columns(PP_AMOUNT).NumberFormat = "#,##0.00"
        ws.Columns(PP_PAID).NumberFormat = "#,##0.00"
        ws.Columns(PP_BALANCE).NumberFormat = "#,##0.00"
    End If
    Set EnsurePaymentPlansSheet = ws
End Function

Private Function LastPlanRow(ByVal ws As Worksheet) As Long
    LastPlanRow = ws.Cells(ws.Rows.Count, PP_PLANID).End(xlUp).Row
End Function
Private Function IsCancelled(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    IsCancelled = (NrmID(CStr(ws.Cells(r, PP_STATUS).Value)) = "CANCELLED")
End Function
Private Function WasReminded(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim text As String
    text = NrmID(CStr(ws.Cells(r, PP_REMINDED).Value))
    WasReminded = (text <> "" And text <> "FALSE" And text <> "0" And text <> "NO")
End Function

Public Function PaymentPlanBaseID(ByVal invoiceNo As String) As String
    Dim key As String
    key = NrmID(invoiceNo)
    If Left$(key, 7) = "MC-INV-" Then
        PaymentPlanBaseID = "PP-MC-" & Mid$(key, 8)
    ElseIf Left$(key, 4) = "INV-" Then
        PaymentPlanBaseID = "PP-" & Mid$(key, 5)
    Else
        Err.Raise vbObjectError + 718, , "A permanent invoice number is required."
    End If
    If Right$(PaymentPlanBaseID, 1) = "-" Then Err.Raise vbObjectError + 718, , "Invalid invoice number."
End Function

Public Function PaymentRequestRef(ByVal planID As String, ByVal installmentNo As Long) As String
    If Left$(NrmID(planID), 3) <> "PP-" Or installmentNo < 0 Then Err.Raise 5, , "Invalid payment request."
    If installmentNo = 0 Then
        PaymentRequestRef = planID & "-DEP"
    Else
        PaymentRequestRef = planID & "-" & Format$(installmentNo, "00")
    End If
End Function

Private Function NewPlanID(ByVal ws As Worksheet, ByVal invoiceNo As String) As String
    Dim base As String, candidate As String, revision As Long
    base = PaymentPlanBaseID(invoiceNo)
    candidate = base: revision = 1
    Do While FindLogRow(ws, candidate) > 0 Or PlanIDWasUsed(candidate)
        revision = revision + 1
        candidate = PaymentPlanRevisionID(invoiceNo, revision)
    Loop
    NewPlanID = candidate
End Function

Public Function PaymentPlanRevisionID(ByVal invoiceNo As String, ByVal revision As Long) As String
    If revision < 1 Then Err.Raise 5, , "Invalid plan revision."
    PaymentPlanRevisionID = PaymentPlanBaseID(invoiceNo)
    If revision > 1 Then PaymentPlanRevisionID = PaymentPlanRevisionID & "-R" & Format$(revision, "00")
End Function

Private Function PlanIDWasUsed(ByVal planID As String) As Boolean
    Dim audit As Worksheet, r As Long
    Set audit = ThisWorkbook.Worksheets("AuditLog")
    For r = 2 To audit.Cells(audit.Rows.Count, 1).End(xlUp).Row
        If NrmID(CStr(audit.Cells(r, 4).Value)) = NrmID(planID) Then PlanIDWasUsed = True: Exit Function
    Next r
End Function

Private Function PlanRows(ByVal ws As Worksheet, ByVal planID As String) As Collection
    Dim rows As New Collection, r As Long, n As Long, term As Long
    For r = 2 To LastPlanRow(ws)
        If NrmID(CStr(ws.Cells(r, PP_PLANID).Value)) = NrmID(planID) Then
            term = WholeNumber(ws.Cells(r, PP_COUNT).Value)
            Exit For
        End If
    Next r
    If term < 0 Or term > 1200 Then Err.Raise 5, , "Invalid plan term."
    For n = 0 To term
        For r = 2 To LastPlanRow(ws)
            If NrmID(CStr(ws.Cells(r, PP_PLANID).Value)) = NrmID(planID) Then
                If Num(ws.Cells(r, PP_NUMBER).Value) = n Then rows.Add r
            End If
        Next r
    Next n
    Set PlanRows = rows
End Function

Private Function InvoiceConfig(ByVal invoiceNo As String) As DocConfig
    If Left$(NrmID(invoiceNo), 7) = "MC-INV-" Then
        InvoiceConfig = GetDocConfig("MC")
    ElseIf Left$(NrmID(invoiceNo), 4) = "INV-" Then
        InvoiceConfig = GetDocConfig("INV")
    Else
        Err.Raise vbObjectError + 718, , "Invalid permanent invoice: " & invoiceNo
    End If
End Function

Private Function Money(ByVal v As Variant) As Currency
    If IsError(v) Then Err.Raise 5, , "Invalid financial value."
    If Not IsNumeric(v) Or Trim$(CStr(v)) = "" Then Err.Raise 5, , "Invalid financial value."
    If CDec(v) < 0 Or CDec(v) * 100 <> Fix(CDec(v) * 100) Then _
        Err.Raise 5, , "Financial values must be nonnegative whole cents."
    Money = CCur(v)
End Function

Private Function WholeNumber(ByVal v As Variant) As Long
    Dim value As Currency
    value = Money(v)
    If value <> Fix(value) Or value > 2147483647# Then Err.Raise 5, , "Schedule numbers must be nonnegative whole numbers."
    WholeNumber = CLng(value)
End Function

Private Sub ValidateInvoice(ByVal invoiceNo As String, ByRef cfg As DocConfig, ByRef log As Worksheet, ByRef lr As Long)
    Dim status As String, total As Currency
    cfg = InvoiceConfig(invoiceNo)
    Set log = ThisWorkbook.Worksheets(cfg.logName)
    lr = FindLogRow(log, invoiceNo)
    If lr < 2 Then Err.Raise 5, , "Permanent invoice missing: " & invoiceNo
    status = NrmID(CStr(log.Cells(lr, cfg.colStatus).Value))
    If status = "VOIDED" Or status = "CANCELLED" Then Err.Raise 5, , "Permanent invoice is cancelled or voided."
    total = Money(log.Cells(lr, cfg.colTotal).Value)
    If total <= 0 Then Err.Raise 5, , "Permanent invoice total must be positive."
    total = Money(log.Cells(lr, cfg.colPaid).Value)
    total = Money(log.Cells(lr, cfg.colBalance).Value)
End Sub

Private Sub ValidateRecipient(ByRef recipient As String, ByRef custID As String, _
        ByVal patientName As String, ByRef dept As String)
    recipient = LCase$(Trim$(recipient))
    If recipient = "private" Then recipient = "patient"
    dept = UCase$(Trim$(dept)): custID = Trim$(custID)
    If dept <> "WA" And dept <> "WD" Then Err.Raise 5, , "Department must be WA or WD."
    Select Case recipient
        Case "doctor"
            If custID = "" Or modHelpers.CustIDToDrName(custID) = "" Then Err.Raise 5, , "Select a saved customer."
        Case "patient"
            If Trim$(patientName) = "" Then Err.Raise 5, , "Enter the private patient's name."
        Case "medclaim"
            If Trim$(patientName) = "" Or custID = "" Then Err.Raise 5, , "Select a medical customer and patient."
            If FindLogRow(ThisWorkbook.Worksheets("Med Customers"), custID) = 0 Then Err.Raise 5, , "Medical customer not found."
        Case Else
            Err.Raise 5, , "Recipient must be doctor, patient or medclaim."
    End Select
End Sub

Private Sub ValidatePlan(ByVal ws As Worksheet, ByVal rows As Collection, Optional ByVal changing As Boolean = False)
    Dim item As Variant, r As Long, first As Long, term As Long, sum As Currency, n As Long
    Dim cfg As DocConfig, log As Worksheet, lr As Long, seen As Object, count As Long
    Dim invoiceNo As String, id As String, total As Currency, recipient As String, cust As String, patient As String
    Dim base As String, revisionText As String, revision As Long
    Dim source As String
    If rows.Count = 0 Then Err.Raise 5, , "Payment plan not found."
    first = CLng(rows(1)): invoiceNo = CStr(ws.Cells(first, PP_INVOICE).Value)
    id = CStr(ws.Cells(first, PP_PLANID).Value)
    base = PaymentPlanBaseID(invoiceNo)
    If NrmID(id) <> NrmID(base) Then
        If Left$(NrmID(id), Len(base) + 2) <> NrmID(base) & "-R" Then Err.Raise 5, , "Plan ID does not match the permanent invoice."
        revisionText = Mid$(NrmID(id), Len(base) + 3)
        revision = WholeNumber(revisionText)
        If revision < 2 Or NrmID(id) <> PaymentPlanRevisionID(invoiceNo, revision) Then Err.Raise 5, , "Invalid plan revision ID."
    End If
    ValidateInvoice invoiceNo, cfg, log, lr
    source = CStr(ws.Cells(first, PP_SOURCE).Value)
    If source <> "" Then
        If Left$(NrmID(source), 2) = "Q-" And cfg.docKind = "INV" Then
            If NrmID(CStr(log.Cells(lr, IL_SRCQUOTE).Value)) <> NrmID(source) Then Err.Raise 5, , "Quote / invoice source link is inconsistent."
        ElseIf NrmID(source) <> NrmID(invoiceNo) Then
            Err.Raise 5, , "Plan source is not its permanent invoice."
        End If
    End If
    total = Money(log.Cells(lr, cfg.colTotal).Value)
    term = WholeNumber(ws.Cells(first, PP_COUNT).Value)
    If term < 1 Or rows.Count <> term + 1 Then Err.Raise 5, , "Incomplete payment schedule."
    If cfg.docKind = "MC" Then
        recipient = "medclaim": cust = CStr(log.Cells(lr, ML_CUST).Value): patient = CStr(log.Cells(lr, ML_PATIENT).Value)
    Else
        recipient = LCase$(CStr(log.Cells(lr, IL_RECIP).Value))
        If recipient = "private" Then recipient = "patient"
        cust = CStr(log.Cells(lr, IL_CUST).Value): patient = CStr(log.Cells(lr, IL_PATIENT).Value)
    End If
    Set seen = CreateObject("Scripting.Dictionary")
    For Each item In rows
        r = CLng(item): n = WholeNumber(ws.Cells(r, PP_NUMBER).Value)
        If n > term Or seen.Exists(CStr(n)) Then Err.Raise 5, , "Invalid or duplicate installment number."
        seen.Add CStr(n), True
        If CStr(ws.Cells(r, PP_REQUEST).Value) <> PaymentRequestRef(id, n) Or _
            NrmID(CStr(ws.Cells(r, PP_SOURCE).Value)) <> NrmID(source) Or _
            NrmID(CStr(ws.Cells(r, PP_INVOICE).Value)) <> NrmID(invoiceNo) Or _
            Money(ws.Cells(r, PP_TOTAL).Value) <> total Or WholeNumber(ws.Cells(r, PP_COUNT).Value) <> term Then _
            Err.Raise 5, , "Plan / permanent invoice references are inconsistent."
        If NrmID(CStr(ws.Cells(r, PP_CUST).Value)) <> NrmID(cust) Or _
            NrmID(CStr(ws.Cells(r, PP_PATIENT).Value)) <> NrmID(patient) Or _
            NrmID(CStr(ws.Cells(r, PP_DEPT).Value)) <> NrmID(CStr(log.Cells(lr, cfg.colDept).Value)) Then _
            Err.Raise 5, , "Plan account does not match its permanent invoice."
        If LCase$(CStr(ws.Cells(r, PP_RECIP).Value)) <> recipient Then
            If Not (recipient = "patient" And LCase$(CStr(ws.Cells(r, PP_RECIP).Value)) = "medclaim") Then _
                Err.Raise 5, , "Plan recipient classification does not match."
        End If
        If Not IsDate(ws.Cells(r, PP_DUE).Value) Then Err.Raise 5, , "Invalid due date."
        If Money(ws.Cells(r, PP_DEPOSIT).Value) <> Money(ws.Cells(first, PP_AMOUNT).Value) Then _
            Err.Raise 5, , "Deposit metadata does not match the schedule."
        If Not IsNumeric(ws.Cells(r, PP_PERCENT).Value) Then Err.Raise 5, , "Invalid deposit percentage."
        If CDbl(ws.Cells(r, PP_PERCENT).Value) < 0 Or CDbl(ws.Cells(r, PP_PERCENT).Value) > 1 Then Err.Raise 5, , "Invalid deposit percentage."
        If IsCancelled(ws, r) Then Err.Raise 5, , "Plan is cancelled."
        If changing Then
            If WasReminded(ws, r) Or Trim$(CStr(ws.Cells(r, PP_SENDSTATE).Value)) <> "" Or _
                Trim$(CStr(ws.Cells(r, PP_PDF).Value)) <> "" Then _
                Err.Raise 5, , "An exported, sent or attempted request cannot be amended or cancelled."
            If RequestHasOutputHistory(CStr(ws.Cells(r, PP_REQUEST).Value)) Then _
                Err.Raise 5, , "Request export / send history prevents amendment or cancellation."
        End If
        sum = sum + Money(ws.Cells(r, PP_AMOUNT).Value)
    Next item
    For r = 2 To LastPlanRow(ws)
        If NrmID(CStr(ws.Cells(r, PP_PLANID).Value)) = NrmID(id) Then count = count + 1
        If Not IsCancelled(ws, r) And NrmID(CStr(ws.Cells(r, PP_INVOICE).Value)) = NrmID(invoiceNo) Then
            If NrmID(CStr(ws.Cells(r, PP_PLANID).Value)) <> NrmID(id) Then Err.Raise 5, , "Multiple active plans for one invoice."
        End If
    Next r
    If count <> rows.Count Or sum <> total Then Err.Raise 5, , "Schedule does not reconcile to the permanent invoice."
    If changing Then
        If Money(log.Cells(lr, cfg.colPaid).Value) <> 0 Then Err.Raise 5, , "A partially paid plan cannot be changed."
        If InvoiceHasPaymentHistory(invoiceNo) Then Err.Raise 5, , "Payment history prevents plan amendment / cancellation."
    End If
End Sub

Private Function InvoiceHasPaymentHistory(ByVal invoiceNo As String) As Boolean
    Dim ws As Worksheet, r As Long, audit As Worksheet
    Set ws = ThisWorkbook.Worksheets("Payments")
    For r = 2 To ws.Cells(ws.Rows.Count, PY_ID).End(xlUp).Row
        If NrmID(CStr(ws.Cells(r, PY_INV).Value)) = NrmID(invoiceNo) Then InvoiceHasPaymentHistory = True: Exit Function
    Next r
    Set audit = ThisWorkbook.Worksheets("AuditLog")
    For r = 2 To audit.Cells(audit.Rows.Count, 1).End(xlUp).Row
        If NrmID(CStr(audit.Cells(r, 4).Value)) = NrmID(invoiceNo) Then
            Select Case NrmID(CStr(audit.Cells(r, 3).Value))
                Case "PAYMENT", "PAYEDIT", "PAYDELETE", "PAYMENTPLANPAYMENT", "CREDITNOTE"
                    InvoiceHasPaymentHistory = True: Exit Function
            End Select
        End If
    Next r
End Function

Private Function RequestHasOutputHistory(ByVal request As String, Optional ByVal sendsOnly As Boolean = False) As Boolean
    Dim ws As Worksheet, r As Long, action As String
    Set ws = ThisWorkbook.Worksheets("AuditLog")
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If NrmID(CStr(ws.Cells(r, 4).Value)) = NrmID(request) Then
            action = NrmID(CStr(ws.Cells(r, 3).Value))
            Select Case action
                Case "PAYMENTPLANMAILATTEMPT", "PAYMENTPLANMAILSENT"
                    RequestHasOutputHistory = True: Exit Function
                Case "PAYMENTREQUESTEXPORTATTEMPT", "PAYMENTREQUESTEXPORT"
                    If Not sendsOnly Then RequestHasOutputHistory = True: Exit Function
            End Select
        End If
    Next r
End Function

' Kept for existing callers: requests are not financial invoices.
Public Function GeneratePlanInvoices(ByVal sourceDocNo As String, ByVal aligners As Double, _
        Optional ByVal total As Currency = 0, Optional ByVal recipientType As String = "", _
        Optional ByVal custID As String = "", Optional ByVal patientName As String = "", _
        Optional ByVal dept As String = "", Optional ByVal startDate As Date = 0, _
        Optional ByVal depositPercent As Double = -1) As String
    GeneratePlanInvoices = BuildPlan(sourceDocNo, aligners, total, recipientType, custID, _
        patientName, dept, startDate, depositPercent)
End Function

Private Function BuildPlan(ByVal sourceDocNo As String, ByVal aligners As Double, ByVal total As Currency, _
        ByVal recipient As String, ByVal cust As String, ByVal patient As String, ByVal dept As String, _
        ByVal startDate As Date, ByVal pct As Double, Optional ByVal retiring As String = "") As String
    Dim ws As Worksheet, src As Worksheet, cfg As DocConfig, lr As Long, r As Long, i As Long
    Dim invoiceNo As String, id As String, p As PaymentPlanResult, oldRows As Collection
    Dim item As Variant, snapshots As New Collection, saved As Variant, firstNew As Long
    Dim oldEvents As Boolean, owns As Boolean, errText As String, kind As String, vat As Currency
    Dim quote As Worksheet, qr As Long
    Dim payments As Worksheet, linkColumn As Long
    On Error GoTo Failed
    If mWriting Then Err.Raise 5, , "Payment plan update already running."
    ' Validate receipt storage and all existing active plans before committing any new charge.
    Set payments = ThisWorkbook.Worksheets("Payments")
    linkColumn = PaymentLinkColumn("PaymentPlanID")
    linkColumn = PaymentLinkColumn("PaymentRequestRef")
    RefreshPaymentPlans
    mWriting = True: owns = True
    oldEvents = Application.EnableEvents
    Set ws = EnsurePaymentPlansSheet()
    If ws.ProtectContents Or ThisWorkbook.Worksheets("AuditLog").ProtectContents Then Err.Raise 5, , "Plan / audit log is protected."
    sourceDocNo = Trim$(sourceDocNo)
    If retiring <> "" Then
        Set oldRows = PlanRows(ws, retiring)
        ValidatePlan ws, oldRows, True
        invoiceNo = CStr(ws.Cells(CLng(oldRows(1)), PP_INVOICE).Value)
        If total <> 0 And total <> Money(ws.Cells(CLng(oldRows(1)), PP_TOTAL).Value) Then _
            Err.Raise 5, , "Amending a schedule cannot change the permanent invoice total."
        sourceDocNo = CStr(ws.Cells(CLng(oldRows(1)), PP_SOURCE).Value)
    ElseIf Left$(NrmID(sourceDocNo), 2) = "Q-" Then
        kind = "QTE": cfg = GetDocConfig(kind)
        Set src = ThisWorkbook.Worksheets(cfg.logName)
        lr = FindLogRow(src, sourceDocNo)
        If lr = 0 Then Err.Raise 5, , "Saved quote not found."
        If Trim$(CStr(src.Cells(lr, QL_CONVINV).Value)) <> "" Or _
            NrmID(CStr(src.Cells(lr, QL_STATUS).Value)) = "CONVERTED" Then _
            Err.Raise 5, , "Quote already converted. Select its permanent invoice."
        If NrmID(CStr(src.Cells(lr, QL_STATUS).Value)) = "VOIDED" Or _
            NrmID(CStr(src.Cells(lr, QL_STATUS).Value)) = "CANCELLED" Then Err.Raise 5, , "Quote is cancelled."
        total = Money(src.Cells(lr, QL_TOTAL).Value)
        recipient = CStr(src.Cells(lr, QL_RECIP).Value): cust = CStr(src.Cells(lr, QL_CUST).Value)
        patient = CStr(src.Cells(lr, QL_PATIENT).Value): dept = CStr(src.Cells(lr, QL_DEPT).Value)
    ElseIf sourceDocNo <> "" Then
        invoiceNo = sourceDocNo
    End If
    If invoiceNo <> "" Then
        ValidateInvoice invoiceNo, cfg, src, lr
        CheckInvoicePaymentLinks invoiceNo
        invoiceNo = CStr(src.Cells(lr, cfg.colNo).Value)
        total = Money(src.Cells(lr, cfg.colTotal).Value): dept = CStr(src.Cells(lr, cfg.colDept).Value)
        If cfg.docKind = "MC" Then
            recipient = "medclaim": cust = CStr(src.Cells(lr, ML_CUST).Value): patient = CStr(src.Cells(lr, ML_PATIENT).Value)
        Else
            recipient = CStr(src.Cells(lr, IL_RECIP).Value): cust = CStr(src.Cells(lr, IL_CUST).Value)
            patient = CStr(src.Cells(lr, IL_PATIENT).Value)
            If LCase$(recipient) = "patient" Then
                If InStr(1, CStr(src.Cells(lr, IL_NOTES).Value), "; recipient medclaim", vbTextCompare) > 0 Then recipient = "medclaim"
                If Trim$(CStr(src.Cells(lr, IL_SRCQUOTE).Value)) <> "" Then
                    Set quote = ThisWorkbook.Worksheets("QuoteLog")
                    qr = FindLogRow(quote, CStr(src.Cells(lr, IL_SRCQUOTE).Value))
                    If qr > 0 Then
                        If LCase$(CStr(quote.Cells(qr, QL_RECIP).Value)) = "medclaim" Then recipient = "medclaim"
                    End If
                End If
                For r = 2 To LastPlanRow(ws)
                    If NrmID(CStr(ws.Cells(r, PP_INVOICE).Value)) = NrmID(invoiceNo) And _
                        LCase$(CStr(ws.Cells(r, PP_RECIP).Value)) = "medclaim" Then recipient = "medclaim"
                Next r
            End If
        End If
        For r = 2 To LastPlanRow(ws)
            If Not IsCancelled(ws, r) And NrmID(CStr(ws.Cells(r, PP_INVOICE).Value)) = NrmID(invoiceNo) Then
                If NrmID(CStr(ws.Cells(r, PP_PLANID).Value)) <> NrmID(retiring) Then Err.Raise 5, , "Invoice already has an active plan."
            End If
        Next r
    End If
    ValidateRecipient recipient, cust, patient, dept
    If sourceDocNo = "" And invoiceNo = "" Then
        For r = 2 To LastPlanRow(ws)
            If Trim$(CStr(ws.Cells(r, PP_SOURCE).Value)) = "" Then
                If NrmID(CStr(ws.Cells(r, PP_CUST).Value)) = NrmID(cust) And _
                    NrmID(CStr(ws.Cells(r, PP_PATIENT).Value)) = NrmID(patient) And _
                    NrmID(CStr(ws.Cells(r, PP_DEPT).Value)) = NrmID(dept) And _
                    NrmID(CStr(ws.Cells(r, PP_RECIP).Value)) = NrmID(recipient) Then _
                    Err.Raise 5, , "A manual case already exists for this account / patient. Select its permanent invoice."
            End If
        Next r
        Set src = ThisWorkbook.Worksheets("InvoiceLog")
        For r = 2 To src.Cells(src.Rows.Count, IL_NO).End(xlUp).Row
            If InStr(1, CStr(src.Cells(r, IL_NOTES).Value), "Payment plan PERMANENT; source ;", vbTextCompare) > 0 Then
                If NrmID(CStr(src.Cells(r, IL_CUST).Value)) = NrmID(cust) And _
                    NrmID(CStr(src.Cells(r, IL_PATIENT).Value)) = NrmID(patient) And _
                    NrmID(CStr(src.Cells(r, IL_DEPT).Value)) = NrmID(dept) And _
                    NrmID(CStr(src.Cells(r, IL_STATUS).Value)) <> "VOIDED" Then _
                    Err.Raise 5, , "A permanent manual invoice already exists. Select " & CStr(src.Cells(r, IL_NO).Value) & "."
            End If
        Next r
    End If
    If pct = -1 Then pct = GetDepositPercent()
    If startDate = 0 Then startDate = Date
    p = CalcPaymentPlan(aligners, total, pct)
    If Money(total) <> total Then Err.Raise 5, , "Invalid total."
    If LastPlanRow(ws) + p.InstallmentCount + 1 > ws.Rows.Count Then Err.Raise 5, , "Not enough schedule rows."
    If DateAdd("m", p.InstallmentCount, startDate) < startDate Then Err.Raise 5, , "Invalid final due date."
    Application.EnableEvents = False
    If invoiceNo = "" Then
        vat = CCur(Round(total * 0.15 / 1.15, 2))
        If kind = "QTE" Then vat = Money(src.Cells(lr, QL_VAT).Value)
        invoiceNo = modConvert.GeneratePaymentPlanInvoice("PERMANENT", sourceDocNo, recipient, cust, patient, _
            dept, Date, startDate, 0, total, vat)
        If kind = "QTE" Then
            ' The adapter durably links the saved quote only after every invoice line is written.
            If NrmID(CStr(src.Cells(lr, QL_CONVINV).Value)) <> NrmID(invoiceNo) Then Err.Raise 5, , "Quote link was not saved."
        End If
    End If
    id = NewPlanID(ws, invoiceNo)
    firstNew = LastPlanRow(ws) + 1
    LogAudit "PaymentPlanCreate", id, "", invoiceNo, "Nonfinancial schedule; total=" & CStr(total)
    For i = 0 To p.InstallmentCount
        r = firstNew + i
        ws.Cells(r, PP_PLANID).Value = id
        ws.Cells(r, PP_SOURCE).Value = sourceDocNo
        ws.Cells(r, PP_RECIP).Value = recipient: ws.Cells(r, PP_CUST).Value = cust
        ws.Cells(r, PP_DEPT).Value = dept: ws.Cells(r, PP_PATIENT).Value = patient
        ws.Cells(r, PP_TOTAL).Value = total: ws.Cells(r, PP_DEPOSIT).Value = p.Deposit
        ws.Cells(r, PP_COUNT).Value = p.InstallmentCount: ws.Cells(r, PP_NUMBER).Value = i
        ws.Cells(r, PP_DUE).Value = DateAdd("m", i, startDate)
        ws.Cells(r, PP_AMOUNT).Value = p.Schedule(i): ws.Cells(r, PP_INVOICE).Value = invoiceNo
        ws.Cells(r, PP_GENERATED).Value = False: ws.Cells(r, PP_STATUS).Value = "Pending"
        ws.Cells(r, PP_CREATED).Value = Now: ws.Cells(r, PP_MODIFIED).Value = Now
        ws.Cells(r, PP_ALIGNERS).Value = aligners: ws.Cells(r, PP_PAID).Value = 0
        ws.Cells(r, PP_BALANCE).Value = p.Schedule(i): ws.Cells(r, PP_REQUEST).Value = PaymentRequestRef(id, i)
        ws.Cells(r, PP_SCHEMA).Value = 2: ws.Cells(r, PP_PERCENT).Value = pct
    Next i
    If retiring <> "" Then
        For Each item In oldRows
            snapshots.Add Array(CLng(item), ws.Cells(CLng(item), 1).Resize(1, PP_LASTCOL).Formula)
        Next item
        RetirePlan ws, oldRows, "Revised by " & id
    End If
    ValidatePlan ws, PlanRows(ws, id)
    LogAudit "PaymentPlanCommitted", id, "", invoiceNo, "One permanent charge; requests only"
    Application.EnableEvents = oldEvents: mWriting = False
    RefreshPaymentPlans
    BuildPlan = id
    Exit Function
Failed:
    errText = Err.Description
    If Not owns Then Err.Raise vbObjectError + 721, "BuildPlan", errText
    On Error Resume Next
    For Each saved In snapshots
        ws.Cells(CLng(saved(0)), 1).Resize(1, PP_LASTCOL).Formula = saved(1)
    Next saved
    If id <> "" And firstNew > 0 Then
        For r = firstNew To LastPlanRow(ws)
            If CStr(ws.Cells(r, PP_PLANID).Value) = id Then
                ws.Cells(r, PP_STATUS).Value = "Cancelled"
                ws.Cells(r, PP_SCHEMA).Value = 2
                ws.Cells(r, PP_BALANCE).Value = 0
            End If
        Next r
    End If
    LogAudit "PaymentPlanFailed", id, "", invoiceNo, errText & "; permanent invoice retained; use it to retry"
    Application.EnableEvents = oldEvents: mWriting = False
    On Error GoTo 0
    Err.Raise vbObjectError + 721, "BuildPlan", errText & IIf(invoiceNo <> "", " Permanent invoice retained: " & invoiceNo, "")
End Function

Private Sub RetirePlan(ByVal ws As Worksheet, ByVal rows As Collection, ByVal reason As String)
    Dim item As Variant, r As Long
    For Each item In rows
        r = CLng(item)
        LogAudit "PaymentPlanCancelRow", CStr(ws.Cells(r, PP_REQUEST).Value), _
            CStr(ws.Cells(r, PP_STATUS).Value), "Cancelled", reason
        ws.Cells(r, PP_STATUS).Value = "Cancelled"
        ws.Cells(r, PP_BALANCE).Value = 0
        ws.Cells(r, PP_MODIFIED).Value = Now
    Next item
End Sub

Public Function CancelPaymentPlan(ByVal planID As String) As Boolean
    Dim ws As Worksheet, rows As Collection, item As Variant, saved As Variant, snapshots As New Collection
    Dim errText As String, owns As Boolean, oldEvents As Boolean
    On Error GoTo Failed
    If mWriting Then Err.Raise 5, , "Payment plan update already running."
    mWriting = True: owns = True
    oldEvents = Application.EnableEvents: Application.EnableEvents = False
    Set ws = EnsurePaymentPlansSheet(): Set rows = PlanRows(ws, planID)
    ValidatePlan ws, rows, True
    If ws.ProtectContents Or ThisWorkbook.Worksheets("AuditLog").ProtectContents Then Err.Raise 5, , "Plan / audit log is protected."
    For Each item In rows
        snapshots.Add Array(CLng(item), ws.Cells(CLng(item), 1).Resize(1, PP_LASTCOL).Formula)
    Next item
    RetirePlan ws, rows, "Cancel plan only; permanent invoice remains payable"
    LogAudit "PaymentPlanCancelled", planID, "", "Cancelled", "Permanent invoice unchanged"
    CancelPaymentPlan = True
    Application.EnableEvents = oldEvents: mWriting = False
    Exit Function
Failed:
    errText = Err.Description
    On Error Resume Next
    For Each saved In snapshots
        ws.Cells(CLng(saved(0)), 1).Resize(1, PP_LASTCOL).Formula = saved(1)
    Next saved
    If owns Then Application.EnableEvents = oldEvents: mWriting = False
    On Error GoTo 0
    MsgBox "Plan was not cancelled: " & errText, vbExclamation, "Payment options"
End Function

Public Function AmendPaymentPlan(ByVal planID As String, ByVal aligners As Double, _
        Optional ByVal total As Currency = 0, Optional ByVal depositPercent As Double = -1) As String
    Dim ws As Worksheet, rows As Collection, r As Long
    On Error GoTo Failed
    Set ws = EnsurePaymentPlansSheet(): Set rows = PlanRows(ws, planID)
    ValidatePlan ws, rows, True: r = CLng(rows(1))
    If depositPercent = -1 Then depositPercent = CDbl(ws.Cells(r, PP_PERCENT).Value)
    AmendPaymentPlan = BuildPlan(CStr(ws.Cells(r, PP_SOURCE).Value), aligners, total, _
        CStr(ws.Cells(r, PP_RECIP).Value), CStr(ws.Cells(r, PP_CUST).Value), _
        CStr(ws.Cells(r, PP_PATIENT).Value), CStr(ws.Cells(r, PP_DEPT).Value), _
        CDate(ws.Cells(r, PP_DUE).Value), depositPercent, planID)
    Exit Function
Failed:
    MsgBox "Plan was not amended: " & Err.Description, vbExclamation, "Payment options"
End Function

' Append by header, not by fixed index: existing workbooks may already have extra payment columns.
Public Function PaymentLinkColumn(ByVal header As String, Optional ByVal create As Boolean = False) As Long
    Dim ws As Worksheet, c As Long, found As Long, last As Long, lastUsed As Range
    Set ws = ThisWorkbook.Worksheets("Payments")
    last = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set lastUsed = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    If Not lastUsed Is Nothing Then
        If lastUsed.Column > last Then last = lastUsed.Column
    End If
    If last < PY_NOTES Then last = PY_NOTES
    For c = PY_NOTES + 1 To last
        If NrmID(CStr(ws.Cells(1, c).Value)) = NrmID(header) Then
            If found <> 0 Then Err.Raise 5, , "Duplicate Payments link header: " & header
            found = c
        End If
    Next c
    If found = 0 And create Then
        found = last + 1
        If found > ws.Columns.Count Then Err.Raise 5, , "No space for payment references."
        ws.Cells(1, found).Value = header
    End If
    PaymentLinkColumn = found
End Function

Private Function ActivePlanForInvoice(ByVal ws As Worksheet, ByVal invoiceNo As String) As String
    Dim r As Long, id As String
    For r = 2 To LastPlanRow(ws)
        If Not IsCancelled(ws, r) And NrmID(CStr(ws.Cells(r, PP_INVOICE).Value)) = NrmID(invoiceNo) Then
            id = CStr(ws.Cells(r, PP_PLANID).Value)
            If ActivePlanForInvoice <> "" And NrmID(ActivePlanForInvoice) <> NrmID(id) Then Err.Raise 5, , "Multiple active plans."
            ActivePlanForInvoice = id
        End If
    Next r
End Function

Public Sub ValidatePaymentRequest(ByVal invoiceNo As String, ByVal planID As String, ByVal requestRef As String)
    Dim ws As Worksheet, rows As Collection, item As Variant, found As Boolean
    If planID = "" And requestRef = "" Then Exit Sub
    If planID = "" Then Err.Raise 5, , "Request must include its plan."
    Set ws = EnsurePaymentPlansSheet(): Set rows = PlanRows(ws, planID)
    ValidatePlan ws, rows
    If NrmID(CStr(ws.Cells(CLng(rows(1)), PP_INVOICE).Value)) <> NrmID(invoiceNo) Then _
        Err.Raise 5, , "Payment request belongs to a different invoice / account."
    If requestRef = "" Then Exit Sub
    For Each item In rows
        If NrmID(CStr(ws.Cells(CLng(item), PP_REQUEST).Value)) = NrmID(requestRef) Then found = True
    Next item
    If Not found Then Err.Raise 5, , "Request does not belong to this active plan."
End Sub

Public Function ChoosePaymentRequest(ByVal invoiceNo As String, ByRef planID As String, _
        ByRef requestRef As String, ByRef amount As Double) As Boolean
    Dim ws As Worksheet, id As String, rows As Collection, item As Variant, choices As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PaymentPlans")
    On Error GoTo 0
    ChoosePaymentRequest = True
    If ws Is Nothing Then Exit Function
    Set ws = EnsurePaymentPlansSheet()
    id = ActivePlanForInvoice(ws, invoiceNo)
    If id = "" Then Exit Function
    RefreshPaymentPlans
    Set rows = PlanRows(ws, id): ValidatePlan ws, rows
    For Each item In rows
        choices = choices & CStr(ws.Cells(CLng(item), PP_REQUEST).Value) & "  R " & _
            Format(ws.Cells(CLng(item), PP_BALANCE).Value, "#,##0.00") & vbCrLf
    Next item
    requestRef = Trim$(InputBox("Optional request reference (blank = deposit-first / oldest installment)." & vbCrLf & _
        choices, "Payment request for " & invoiceNo))
    planID = id
    ValidatePaymentRequest invoiceNo, planID, requestRef
    If requestRef <> "" Then
        For Each item In rows
            If NrmID(CStr(ws.Cells(CLng(item), PP_REQUEST).Value)) = NrmID(requestRef) Then _
                amount = CDbl(ws.Cells(CLng(item), PP_BALANCE).Value)
        Next item
    End If
    choices = InputBox("Amount received for " & invoiceNo & " (blank cancels):", "Record plan payment", Format(amount, "0.00"))
    If Trim$(choices) = "" Then ChoosePaymentRequest = False: Exit Function
    amount = CDbl(Money(choices))
    If amount <= 0 Then Err.Raise 5, , "Payment must be positive."
End Function

Public Sub StorePaymentRequest(ByVal paymentRow As Long, ByVal invoiceNo As String, _
        ByVal planID As String, ByVal requestRef As String)
    Dim ws As Worksheet, pc As Long, rc As Long
    ValidatePaymentRequest invoiceNo, planID, requestRef
    If planID = "" And requestRef = "" Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Payments")
    If paymentRow < 2 Or paymentRow > ws.Cells(ws.Rows.Count, PY_ID).End(xlUp).Row Then Err.Raise 5, , "Payment row not found."
    If NrmID(CStr(ws.Cells(paymentRow, PY_INV).Value)) <> NrmID(invoiceNo) Then _
        Err.Raise 5, , "Payment row belongs to another invoice / account."
    pc = PaymentLinkColumn("PaymentPlanID", True): rc = PaymentLinkColumn("PaymentRequestRef", True)
    ws.Cells(paymentRow, pc).Value = planID: ws.Cells(paymentRow, rc).Value = requestRef
End Sub

Public Sub ValidatePaymentWrite(ByVal invoiceNo As String, ByVal amount As Double, ByVal paymentDate As Variant, _
        Optional ByVal planID As String = "", Optional ByVal requestRef As String = "", _
        Optional ByVal excludingRow As Long = 0)
    Dim ws As Worksheet, pay As Worksheet, cfg As DocConfig, log As Worksheet, lr As Long, r As Long
    Dim id As String, paid As Currency, value As Currency, existingAllocations As Variant
    If Not IsDate(paymentDate) Then Err.Raise 5, , "Payment date is invalid."
    value = Money(amount)
    If value <= 0 Then Err.Raise 5, , "Payment amount must be positive whole cents."
    If ThisWorkbook.Worksheets("Payments").ProtectContents Or ThisWorkbook.Worksheets("AuditLog").ProtectContents Then _
        Err.Raise 5, , "Payment / audit log is protected."
    ValidatePaymentRequest invoiceNo, planID, requestRef
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PaymentPlans")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Set ws = EnsurePaymentPlansSheet()
    id = ActivePlanForInvoice(ws, invoiceNo)
    If id = "" Then Exit Sub
    ValidatePlan ws, PlanRows(ws, id)
    existingAllocations = PlanAllocatedPayments(id, Date)
    ValidateInvoice invoiceNo, cfg, log, lr
    If log.ProtectContents Then Err.Raise 5, , "Permanent invoice log is protected."
    Set pay = ThisWorkbook.Worksheets("Payments")
    For r = 2 To pay.Cells(pay.Rows.Count, PY_ID).End(xlUp).Row
        If r <> excludingRow And NrmID(CStr(pay.Cells(r, PY_INV).Value)) = NrmID(invoiceNo) Then _
            paid = paid + Money(pay.Cells(r, PY_AMT).Value)
    Next r
    If paid + value > Money(log.Cells(lr, cfg.colTotal).Value) Then _
        Err.Raise 5, , "Plan payment exceeds the permanent invoice's remaining balance."
End Sub

Public Sub CheckInvoicePaymentLinks(ByVal invoiceNo As String)
    Dim pay As Worksheet, pc As Long, rc As Long, r As Long, id As String, ref As String
    Set pay = ThisWorkbook.Worksheets("Payments")
    pc = PaymentLinkColumn("PaymentPlanID"): rc = PaymentLinkColumn("PaymentRequestRef")
    If pc = 0 And rc = 0 Then Exit Sub
    For r = 2 To pay.Cells(pay.Rows.Count, PY_ID).End(xlUp).Row
        If NrmID(CStr(pay.Cells(r, PY_INV).Value)) = NrmID(invoiceNo) Then
            id = "": ref = ""
            If pc > 0 Then id = CStr(pay.Cells(r, pc).Value)
            If rc > 0 Then ref = CStr(pay.Cells(r, rc).Value)
            ValidatePaymentRequest invoiceNo, id, ref
        End If
    Next r
End Sub

' One application only. Excess stays unallocated on the permanent invoice, never duplicated.
Public Function AllocatePlanAmount(ByRef amounts() As Currency, ByRef paid() As Currency, _
        ByVal amount As Currency, Optional ByVal target As Long = -1) As Currency
    Dim n As Long, applied As Currency
    amount = Money(amount)
    If LBound(amounts) <> 0 Or UBound(amounts) <> UBound(paid) Or LBound(paid) <> 0 Then Err.Raise 5, , "Invalid allocation arrays."
    If target < -1 Or target > UBound(amounts) Then Err.Raise 5, , "Invalid allocation target."
    For n = 0 To UBound(amounts)
        If amounts(n) < 0 Or paid(n) < 0 Or paid(n) > amounts(n) Then Err.Raise 5, , "Invalid schedule allocation."
        If Money(amounts(n)) <> amounts(n) Or Money(paid(n)) <> paid(n) Then Err.Raise 5, , "Invalid schedule cents."
    Next n
    If target >= 0 Then
        applied = amounts(target) - paid(target)
        If applied > amount Then applied = amount
        paid(target) = paid(target) + applied: amount = amount - applied
    End If
    For n = 0 To UBound(amounts)
        applied = amounts(n) - paid(n)
        If applied > amount Then applied = amount
        paid(n) = paid(n) + applied: amount = amount - applied
    Next n
    AllocatePlanAmount = amount
End Function

Public Function PlanAllocatedPayments(ByVal planID As String, ByVal cutoff As Date) As Variant
    Dim ws As Worksheet, pay As Worksheet, rows As Collection, item As Variant, first As Long, n As Long
    Dim amounts() As Currency, paid() As Currency, pRows() As Long, dates() As Double
    Dim r As Long, count As Long, j As Long, swap As Long, dt As Double, swapDate As Double, eff As Double
    Dim pc As Long, rc As Long, pid As String, ref As String, target As Long, excess As Currency
    Dim invoiceNo As String
    Set ws = EnsurePaymentPlansSheet(): Set rows = PlanRows(ws, planID): ValidatePlan ws, rows
    first = CLng(rows(1)): invoiceNo = CStr(ws.Cells(first, PP_INVOICE).Value)
    ReDim amounts(0 To rows.Count - 1): ReDim paid(0 To rows.Count - 1)
    For Each item In rows
        n = CLng(ws.Cells(CLng(item), PP_NUMBER).Value): amounts(n) = Money(ws.Cells(CLng(item), PP_AMOUNT).Value)
    Next item
    Set pay = ThisWorkbook.Worksheets("Payments")
    pc = PaymentLinkColumn("PaymentPlanID"): rc = PaymentLinkColumn("PaymentRequestRef")
    ReDim pRows(1 To Application.Max(1, pay.Cells(pay.Rows.Count, PY_ID).End(xlUp).Row))
    ReDim dates(1 To UBound(pRows))
    For r = 2 To pay.Cells(pay.Rows.Count, PY_ID).End(xlUp).Row
        If NrmID(CStr(pay.Cells(r, PY_INV).Value)) = NrmID(invoiceNo) Then
            ' Match the normal statement's undated-payment policy.
            eff = CDbl(cutoff)
            If IsDate(pay.Cells(r, PY_DATE).Value) Then
                dt = CDbl(CDate(pay.Cells(r, PY_DATE).Value))
                If dt > 0 Then eff = dt
            End If
            If eff < CDbl(DateAdd("d", 1, DateValue(cutoff))) Then
                count = count + 1: pRows(count) = r: dates(count) = eff
            End If
        End If
    Next r
    ' Stable date / physical row order, rebuilt after every edit or deletion.
    For n = 2 To count
        swap = pRows(n): swapDate = dates(n): j = n - 1
        Do While j >= 1
            If dates(j) <= swapDate Then Exit Do
            pRows(j + 1) = pRows(j): dates(j + 1) = dates(j): j = j - 1
        Loop
        pRows(j + 1) = swap: dates(j + 1) = swapDate
    Next n
    For n = 1 To count
        r = pRows(n): pid = "": ref = "": target = -1
        If pc > 0 Then pid = CStr(pay.Cells(r, pc).Value)
        If rc > 0 Then ref = CStr(pay.Cells(r, rc).Value)
        ValidatePaymentRequest invoiceNo, pid, ref
        If pid <> "" And NrmID(pid) <> NrmID(planID) Then Err.Raise 5, , "Payment belongs to another plan revision."
        If ref <> "" Then
            For Each item In rows
                If NrmID(CStr(ws.Cells(CLng(item), PP_REQUEST).Value)) = NrmID(ref) Then target = CLng(ws.Cells(CLng(item), PP_NUMBER).Value)
            Next item
        End If
        excess = AllocatePlanAmount(amounts, paid, Money(pay.Cells(r, PY_AMT).Value), target)
    Next n
    PlanAllocatedPayments = paid
End Function

Public Sub RefreshPaymentPlans()
    Dim ws As Worksheet, r As Long, id As String, seen As Object, rows As Collection
    Dim item As Variant, allocations As Variant, paid As Currency, bal As Currency, released As Boolean, status As String
    On Error GoTo Failed
    If mWriting Or mRefreshing Then Exit Sub
    mRefreshing = True: Set ws = EnsurePaymentPlansSheet(): Set seen = CreateObject("Scripting.Dictionary")
    For r = 2 To LastPlanRow(ws)
        If Not IsCancelled(ws, r) Then
            id = CStr(ws.Cells(r, PP_PLANID).Value)
            If Not seen.Exists(NrmID(id)) Then
                seen.Add NrmID(id), True
                Set rows = PlanRows(ws, id): ValidatePlan ws, rows
                allocations = PlanAllocatedPayments(id, Date)
                released = (CCur(allocations(0)) = Money(ws.Cells(CLng(rows(1)), PP_AMOUNT).Value))
                For Each item In rows
                    paid = CCur(allocations(CLng(ws.Cells(CLng(item), PP_NUMBER).Value)))
                    bal = Money(ws.Cells(CLng(item), PP_AMOUNT).Value) - paid
                    status = "Pending"
                    If bal = 0 Then
                        status = "Paid"
                    ElseIf released Or CLng(ws.Cells(CLng(item), PP_NUMBER).Value) = 0 Then
                        status = "Active"
                        If CDate(ws.Cells(CLng(item), PP_DUE).Value) < Date Then status = "Overdue"
                    End If
                    ws.Cells(CLng(item), PP_PAID).Value = paid: ws.Cells(CLng(item), PP_BALANCE).Value = bal
                    If CStr(ws.Cells(CLng(item), PP_STATUS).Value) <> status Then
                        LogAudit "PaymentPlanStatus", CStr(ws.Cells(CLng(item), PP_REQUEST).Value), _
                            CStr(ws.Cells(CLng(item), PP_STATUS).Value), status, "Allocated receipts as of " & Format$(Date, "yyyy-mm-dd")
                    End If
                    ws.Cells(CLng(item), PP_STATUS).Value = status: ws.Cells(CLng(item), PP_MODIFIED).Value = Now
                Next item
            End If
        End If
    Next r
    mRefreshing = False
    Exit Sub
Failed:
    mRefreshing = False
    Err.Raise vbObjectError + 727, "RefreshPaymentPlans", Err.Description
End Sub

Public Function DueInstallmentRows() As Collection
    Dim rows As New Collection, ws As Worksheet, r As Long, lead As Long, status As String
    RefreshPaymentPlans: Set ws = EnsurePaymentPlansSheet(): lead = GetReminderLeadDays()
    For r = 2 To LastPlanRow(ws)
        status = CStr(ws.Cells(r, PP_STATUS).Value)
        If Not IsCancelled(ws, r) And Money(ws.Cells(r, PP_BALANCE).Value) > 0 And _
            (status = "Active" Or status = "Overdue") And Not WasReminded(ws, r) Then
            If Trim$(CStr(ws.Cells(r, PP_SENDSTATE).Value)) = "" Then
                If Not RequestHasOutputHistory(CStr(ws.Cells(r, PP_REQUEST).Value), True) Then
                    If CDate(ws.Cells(r, PP_DUE).Value) <= Date + lead Then rows.Add r
                End If
            End If
        End If
    Next r
    Set DueInstallmentRows = rows
End Function

Public Function PaymentRequestWithinMailWindow(ByVal installmentNo As Long, ByVal dueDate As Date, _
        ByVal cutoff As Date, ByVal leadDays As Long) As Boolean
    If installmentNo < 0 Or leadDays < 0 Then Err.Raise 5, , "Invalid request / reminder window."
    PaymentRequestWithinMailWindow = (installmentNo = 0 Or dueDate <= DateAdd("d", leadDays, DateValue(cutoff)))
End Function

Public Sub ShowInstallmentReminder()
    Dim rows As Collection
    On Error GoTo Failed
    Set rows = DueInstallmentRows()
    If rows.Count > 0 Then frmPaymentOptions.ShowReminder
    Exit Sub
Failed:
    MsgBox "Reminders unavailable: " & Err.Description, vbExclamation, "Payment options"
End Sub

Private Function RecipientEmail(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim contacts As Worksheet, key As String, i As Long
    ' Only a unique saved bill-to is trusted; private patients supply an address with consent.
    Select Case LCase$(CStr(ws.Cells(r, PP_RECIP).Value))
        Case "doctor"
            Set contacts = ThisWorkbook.Worksheets("Customers")
            key = NrmID(CStr(ws.Cells(r, PP_CUST).Value))
            For i = 2 To contacts.Cells(contacts.Rows.Count, 1).End(xlUp).Row
                If NrmID(CStr(contacts.Cells(i, 1).Value)) = key Then
                    RecipientEmail = Trim$(CStr(contacts.Cells(i, 7).Value))
                    Exit Function
                End If
            Next i
    End Select
End Function

Public Function ExportPaymentRequest(ByVal planRow As Long) As String
    Dim ws As Worksheet, paper As Worksheet, folder As String, path As String, base As String, suffix As Long
    Dim oldAlerts As Boolean, oldEvents As Boolean, errNo As Long, errText As String, patient As String
    Dim previousSheet As Object
    On Error GoTo Failed
    oldAlerts = Application.DisplayAlerts: oldEvents = Application.EnableEvents
    Set previousSheet = ActiveSheet
    RefreshPaymentPlans: Set ws = EnsurePaymentPlansSheet()
    If planRow < 2 Or planRow > LastPlanRow(ws) Then Err.Raise 5, , "Select a payment request."
    ValidatePlan ws, PlanRows(ws, CStr(ws.Cells(planRow, PP_PLANID).Value))
    If ws.ProtectContents Or ThisWorkbook.Worksheets("AuditLog").ProtectContents Then Err.Raise 5, , "Request / audit log is protected."
    folder = modExportPDF.PaymentPDFDirectory()
    If folder = "" Then Exit Function
    patient = Trim$(CStr(ws.Cells(planRow, PP_PATIENT).Value))
    If patient = "" Then patient = "NoName"
    base = modExportPDF.CleanName(CStr(ws.Cells(planRow, PP_REQUEST).Value) & "_" & patient)
    path = folder & Application.PathSeparator & base & ".pdf"
    Do While Dir$(path) <> ""
        suffix = suffix + 1: path = folder & Application.PathSeparator & base & "_" & CStr(suffix) & ".pdf"
    Loop
    LogAudit "PaymentRequestExportAttempt", CStr(ws.Cells(planRow, PP_REQUEST).Value), "", path, "Nonfinancial request"
    Application.EnableEvents = False
    Set paper = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    paper.Columns("A").ColumnWidth = 28: paper.Columns("B").ColumnWidth = 65
    paper.Range("A1").Value = "PAYMENT REQUEST - NOT A TAX INVOICE"
    paper.Range("A3").Value = "Request reference": paper.Range("B3").Value = ws.Cells(planRow, PP_REQUEST).Value
    paper.Range("A4").Value = "Permanent invoice": paper.Range("B4").Value = ws.Cells(planRow, PP_INVOICE).Value
    paper.Range("A5").Value = "Plan": paper.Range("B5").Value = ws.Cells(planRow, PP_PLANID).Value
    paper.Range("A6").Value = "Account": paper.Range("B6").Value = ws.Cells(planRow, PP_CUST).Value
    paper.Range("A7").Value = "Patient": paper.Range("B7").Value = ws.Cells(planRow, PP_PATIENT).Value
    paper.Range("A8").Value = "Department / recipient"
    paper.Range("B8").Value = CStr(ws.Cells(planRow, PP_DEPT).Value) & " / " & CStr(ws.Cells(planRow, PP_RECIP).Value)
    paper.Range("A10").Value = "Due date": paper.Range("B10").Value = Format(ws.Cells(planRow, PP_DUE).Value, "yyyy-mm-dd")
    paper.Range("A11").Value = "Scheduled amount": paper.Range("B11").Value = ws.Cells(planRow, PP_AMOUNT).Value
    paper.Range("A12").Value = "Allocated payments": paper.Range("B12").Value = ws.Cells(planRow, PP_PAID).Value
    paper.Range("A13").Value = "Requested balance": paper.Range("B13").Value = ws.Cells(planRow, PP_BALANCE).Value
    paper.Range("A14").Value = "Schedule status": paper.Range("B14").Value = ws.Cells(planRow, PP_STATUS).Value
    paper.Range("B11:B13").NumberFormat = """R ""#,##0.00"
    paper.Range("A15").Value = "No additional charge or VAT is raised by this request."
    paper.Range("A16").Value = "All receipts are credited to the permanent invoice above."
    paper.Range("A18").Value = "Please quote the request reference when arranging payment."
    paper.Range("A1:B18").Font.Size = 11
    paper.Range("A1").Font.Bold = True
    With paper.PageSetup
        .PrintArea = "$A$1:$B$20": .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
    End With
    paper.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ws.Cells(planRow, PP_PDF).Value = path: ws.Cells(planRow, PP_GENERATED).Value = True
    LogAudit "PaymentRequestExport", CStr(ws.Cells(planRow, PP_REQUEST).Value), "", path, "Invoice unchanged"
    Application.DisplayAlerts = False: paper.Delete: Set paper = Nothing
    previousSheet.Activate
    Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents
    ExportPaymentRequest = path
    Exit Function
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If Not paper Is Nothing Then Application.DisplayAlerts = False: paper.Delete
    If Not previousSheet Is Nothing Then previousSheet.Activate
    Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents
    On Error GoTo 0
    Err.Raise errNo, "ExportPaymentRequest", errText
End Function

Public Function ExportPlanRequests(ByVal planID As String) As Long
    Dim ws As Worksheet, rows As Collection, item As Variant, folder As String, path As String
    Dim completed As Long, errText As String
    On Error GoTo Failed
    RefreshPaymentPlans
    Set ws = EnsurePaymentPlansSheet(): Set rows = PlanRows(ws, planID)
    ValidatePlan ws, rows
    folder = modExportPDF.PaymentPDFDirectory()
    If folder = "" Then Exit Function
    For Each item In rows
        path = ExportPaymentRequest(CLng(item))
        If path = "" Then Err.Raise 5, , "Request export was cancelled."
        completed = completed + 1
    Next item
    LogAudit "PaymentPlanRequestsExport", planID, "", CStr(completed), "All nonfinancial notices; shared folder=" & folder
    ExportPlanRequests = completed
    Exit Function
Failed:
    errText = Err.Description
    On Error Resume Next
    LogAudit "PaymentPlanRequestsExportFailed", planID, "", CStr(completed), errText
    On Error GoTo 0
    Err.Raise vbObjectError + 729, "ExportPlanRequests", CStr(completed) & _
        " notices exported. Plan and permanent invoice remain saved; retry export, not plan creation. " & errText
End Function

' Compatibility entrypoint: mails the nonfinancial request, never recalls/rewrites Invoice.
Public Function MailInstallmentInvoice(ByVal planRow As Long) As Boolean
    Dim ws As Worksheet, outlook As Object, mail As Object, address As String, path As String, request As String
    Dim sent As Boolean, attempted As Boolean, errText As String
    On Error GoTo Failed
    RefreshPaymentPlans: Set ws = EnsurePaymentPlansSheet()
    If planRow < 2 Or planRow > LastPlanRow(ws) Then Err.Raise 5, , "Select a payment request."
    If ws.ProtectContents Or ThisWorkbook.Worksheets("AuditLog").ProtectContents Then Err.Raise 5, , "Tracking logs are protected."
    ValidatePlan ws, PlanRows(ws, CStr(ws.Cells(planRow, PP_PLANID).Value))
    If Money(ws.Cells(planRow, PP_BALANCE).Value) = 0 Then Err.Raise 5, , "Request is already paid."
    If CStr(ws.Cells(planRow, PP_STATUS).Value) = "Pending" Then Err.Raise 5, , "Deposit must be paid first."
    If Not PaymentRequestWithinMailWindow(WholeNumber(ws.Cells(planRow, PP_NUMBER).Value), _
        CDate(ws.Cells(planRow, PP_DUE).Value), Date, GetReminderLeadDays()) Then _
        Err.Raise 5, , "This installment is not yet within the due reminder window. It cannot be emailed early."
    If WasReminded(ws, planRow) Or Trim$(CStr(ws.Cells(planRow, PP_SENDSTATE).Value)) <> "" Then _
        Err.Raise 5, , "A previous send / attempt exists. Review Outlook and AuditLog; automatic resend is blocked."
    request = CStr(ws.Cells(planRow, PP_REQUEST).Value)
    If RequestHasOutputHistory(request, True) Then Err.Raise 5, , "Audit history shows a send / attempt. Review before any resend."
    address = RecipientEmail(ws, planRow)
    If address = "" Then address = Trim$(InputBox("Enter recipient email:", "Payment request"))
    If InStr(2, address, "@") = 0 Or InStr(address, " ") > 0 Or InStr(address, vbCr) > 0 Or _
        InStr(address, vbLf) > 0 Or InStr(address, ";") > 0 Or InStr(address, ",") > 0 Then Err.Raise 5, , "Enter one valid email address."
    If MsgBox("Send nonfinancial payment request " & request & " to " & address & " now?" & vbCrLf & _
        "Confirm the recipient has consented. No / Cancel sends nothing.", vbYesNoCancel + vbQuestion) <> vbYes Then Exit Function
    path = ExportPaymentRequest(planRow)
    If path = "" Then Exit Function
    Set outlook = CreateObject("Outlook.Application"): Set mail = outlook.CreateItem(0)
    mail.To = address: mail.Subject = "Payment request " & request
    mail.Body = "Attached is a nonfinancial payment request against invoice " & _
        CStr(ws.Cells(planRow, PP_INVOICE).Value) & ". It is not an additional invoice or charge."
    mail.Attachments.Add path
    LogAudit "PaymentPlanMailAttempt", request, "", address, "Explicit recipient / consent confirmation"
    ws.Cells(planRow, PP_SENDSTATE).Value = "Attempted"
    attempted = True
    mail.Send: sent = True
    ws.Cells(planRow, PP_REMINDED).Value = Now: ws.Cells(planRow, PP_SENDSTATE).Value = "Outlook accepted"
    ws.Cells(planRow, PP_MODIFIED).Value = Now
    LogAudit "PaymentPlanMailSent", request, "", address, "Outlook accepted; delivery not guaranteed"
    MailInstallmentInvoice = True
    Exit Function
Failed:
    errText = Err.Description
    On Error Resume Next
    If attempted And Not sent Then ws.Cells(planRow, PP_SENDSTATE).Value = "Failed or unconfirmed"
    LogAudit "PaymentPlanMailFailed", request, "", address, errText
    On Error GoTo 0
    MailInstallmentInvoice = sent
    MsgBox "Send result requires review: " & errText & vbCrLf & "Check Outlook / AuditLog before sending again.", vbExclamation
End Function

Public Sub RefreshPaymentOptionsDashboard()
    Dim ws As Worksheet, r As Long, dueMonth As Long, overdue As Long, d As Date, status As String
    On Error GoTo Failed
    RefreshPaymentPlans: Set ws = EnsurePaymentPlansSheet()
    For r = 2 To LastPlanRow(ws)
        If Not IsCancelled(ws, r) And Num(ws.Cells(r, PP_NUMBER).Value) > 0 And Num(ws.Cells(r, PP_BALANCE).Value) > 0.005 Then
            d = CDate(ws.Cells(r, PP_DUE).Value): status = CStr(ws.Cells(r, PP_STATUS).Value)
            If status = "Active" Or status = "Overdue" Then
                If Month(d) = Month(Date) And Year(d) = Year(Date) Then dueMonth = dueMonth + 1
                If status = "Overdue" Then overdue = overdue + 1
            End If
        End If
    Next r
    ThisWorkbook.Worksheets("Menu").Range("B70").Value = "Payment plans: " & CStr(dueMonth) & " due this month; " & CStr(overdue) & " overdue"
    Exit Sub
Failed:
    MsgBox "Dashboard unavailable: " & Err.Description, vbExclamation
End Sub

Public Sub RunDepositsStatement()
    RunDepositStatement
End Sub
Public Sub RunInstallmentsStatement()
    Dim ws As Worksheet, rows As New Collection, r As Long, status As String
    On Error GoTo Failed
    RefreshPaymentPlans: Set ws = EnsurePaymentPlansSheet()
    For r = 2 To LastPlanRow(ws)
        status = CStr(ws.Cells(r, PP_STATUS).Value)
        If Not IsCancelled(ws, r) And Num(ws.Cells(r, PP_NUMBER).Value) > 0 And _
            Num(ws.Cells(r, PP_BALANCE).Value) > 0.005 And (status = "Active" Or status = "Overdue") Then
            If CDate(ws.Cells(r, PP_DUE).Value) <= Date Then rows.Add r
        End If
    Next r
    If rows.Count = 0 Then
        MsgBox "No outstanding installments are due.", vbInformation
    Else
        modStatement.RunPaymentPlanStatement rows, "Installments statement - ALL recipients / ALL departments"
    End If
    Exit Sub
Failed:
    MsgBox "Installments statement unavailable: " & Err.Description, vbExclamation
End Sub

Private Sub RunDepositStatement()
    Dim ws As Worksheet, rows As New Collection, recipient As String, dept As String, customer As String
    Dim r As Long, matches As Boolean
    On Error GoTo Failed
    recipient = LCase$(Trim$(InputBox("Recipient: doctor, patient or medclaim (blank cancels)", "Deposits statement")))
    If recipient = "" Then Exit Sub
    If recipient <> "doctor" And recipient <> "patient" And recipient <> "medclaim" Then Err.Raise 5, , "Invalid recipient."
    dept = UCase$(Trim$(InputBox("Department: WA, WD or ALL", "Deposits statement", "ALL")))
    If dept = "" Then Exit Sub
    If dept <> "WA" And dept <> "WD" And dept <> "ALL" Then Err.Raise 5, , "Invalid department."
    customer = Trim$(InputBox("Customer ID / medical Bill To / patient name, or ALL", "Deposits statement", "ALL"))
    If customer = "" Then Exit Sub
    RefreshPaymentPlans: Set ws = EnsurePaymentPlansSheet()
    For r = 2 To LastPlanRow(ws)
        matches = (LCase$(CStr(ws.Cells(r, PP_RECIP).Value)) = recipient)
        If dept <> "ALL" Then matches = matches And NrmID(CStr(ws.Cells(r, PP_DEPT).Value)) = NrmID(dept)
        If NrmID(customer) <> "ALL" Then
            If recipient = "patient" Then
                matches = matches And NrmID(CStr(ws.Cells(r, PP_PATIENT).Value)) = NrmID(customer)
            Else
                matches = matches And NrmID(CStr(ws.Cells(r, PP_CUST).Value)) = NrmID(customer)
            End If
        End If
        If matches And Not IsCancelled(ws, r) And Num(ws.Cells(r, PP_BALANCE).Value) > 0.005 Then
            If Num(ws.Cells(r, PP_NUMBER).Value) = 0 Then rows.Add r
        End If
    Next r
    If rows.Count = 0 Then
        MsgBox "No matching outstanding requests.", vbInformation
    Else
        modStatement.RunPaymentPlanStatement rows, "Deposits statement - " & recipient & " / " & dept & " / " & customer
    End If
    Exit Sub
Failed:
    MsgBox "Deposits statement unavailable: " & Err.Description, vbExclamation
End Sub

Public Sub RunSinglePlanStatement(Optional ByVal planID As String = "")
    Dim ws As Worksheet, rows As Collection
    On Error GoTo Failed
    If planID = "" Then planID = Trim$(InputBox("Payment plan ID:", "Payment plan statement"))
    If planID = "" Then Exit Sub
    RefreshPaymentPlans: Set ws = EnsurePaymentPlansSheet(): Set rows = PlanRows(ws, planID)
    If rows.Count = 0 Then Err.Raise 5, , "Payment plan not found."
    modStatement.RunPaymentPlanStatement rows, "Payment plan " & planID
    Exit Sub
Failed:
    MsgBox "Plan statement unavailable: " & Err.Description, vbExclamation
End Sub
