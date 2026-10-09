Attribute VB_Name = "modPaymentOptions"
Option Explicit

Private mWriting As Boolean
Private mRefreshing As Boolean

' Supplying all settings makes this calculation independent of the workbook.
Public Function CalcPaymentPlan(ByVal aligners As Double, ByVal total As Currency, _
        Optional ByVal depositPercent As Double = -1, _
        Optional ByVal weeksPerAligner As Double = -1, _
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
    If term > 2147483646# Then Err.Raise vbObjectError + 714, , "Payment term is too large."
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
        ByVal minimum As Double, ByVal maximum As Double, _
        Optional ByVal wholeNumber As Boolean = False) As Double
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
        If SettingNumber <> Fix(SettingNumber) Then Err.Raise vbObjectError + 716, , "Setting must be a whole number: " & key
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
    Dim ws As Worksheet, headers As Variant, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PaymentPlans")
    On Error GoTo 0
    headers = Array("PlanID", "SourceDocNo", "RecipientType", "CustID", "Dept", _
        "Total", "Deposit", "InstallmentNo", "InstallmentCount", "DueDate", _
        "Amount", "InvoiceNo", "Generated", "Status", "Reminded", _
        "CreatedAt", "ModifiedAt", "PatientName", "PDFPath", "Aligners", "Paid", "Balance")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "PaymentPlans"
        For i = 0 To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
        ws.Columns(PP_DUE).NumberFormat = "yyyy-mm-dd"
        ws.Columns(PP_AMOUNT).NumberFormat = "#,##0.00"
        ws.Columns(PP_PAID).NumberFormat = "#,##0.00"
        ws.Columns(PP_BALANCE).NumberFormat = "#,##0.00"
    Else
        For i = 0 To UBound(headers)
            If CStr(ws.Cells(1, i + 1).Value) <> CStr(headers(i)) Then
                Err.Raise vbObjectError + 717, , "PaymentPlans column layout does not match modConfig."
            End If
        Next i
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

Private Function HasPayments(ByVal docNo As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Payments")
    For r = 2 To ws.Cells(ws.Rows.Count, PY_INV).End(xlUp).Row
        If NrmID(CStr(ws.Cells(r, PY_INV).Value)) = NrmID(docNo) Then
            HasPayments = True
            Exit Function
        End If
    Next r
End Function

Private Function HeaderColumn(ByVal ws As Worksheet, ParamArray names() As Variant) As Long
    Dim c As Long, name As Variant
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        For Each name In names
            If NrmID(CStr(ws.Cells(1, c).Value)) = NrmID(CStr(name)) Then
                HeaderColumn = c
                Exit Function
            End If
        Next name
    Next c
End Function

Private Function InvoiceWasSent(ByVal ws As Worksheet, ByVal r As Long, _
        Optional ByVal includeAttempts As Boolean = False) As Boolean
    Dim c As Long, v As Variant, status As String, header As String, audit As Worksheet
    status = NrmID(CStr(ws.Cells(r, IL_STATUS).Value))
    If status = "SENT" Or status = "EMAILED" Then InvoiceWasSent = True: Exit Function
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        header = NrmID(CStr(ws.Cells(1, c).Value))
        Select Case header
            Case "SENT", "SENTAT", "SENTDATE", "EMAILED", "EMAILSENT"
                v = ws.Cells(r, c).Value
                If IsError(v) Then InvoiceWasSent = True: Exit Function
                If Trim$(CStr(v)) <> "" And NrmID(CStr(v)) <> "FALSE" And _
                    CStr(v) <> "0" And NrmID(CStr(v)) <> "NO" Then
                    InvoiceWasSent = True
                    Exit Function
                End If
        End Select
    Next c
    Set audit = ThisWorkbook.Worksheets("AuditLog")
    For c = 2 To audit.Cells(audit.Rows.Count, 4).End(xlUp).Row
        If NrmID(CStr(audit.Cells(c, 4).Value)) = NrmID(CStr(ws.Cells(r, IL_NO).Value)) Then
            status = NrmID(CStr(audit.Cells(c, 3).Value))
            If status = "PAYMENTPLANMAILSENT" Or status = "PAYMENTPLANSENTOBSERVED" Or _
                (includeAttempts And (status = "PAYMENTPLANMAILATTEMPT" Or _
                    status = "PAYMENTPLANMAILDRAFT")) Then
                InvoiceWasSent = True
                Exit Function
            End If
        End If
    Next c
End Function

Private Sub AssertUnpaid(ByVal ws As Worksheet, ByVal r As Long, ByRef c As DocConfig, _
        Optional ByVal checkSent As Boolean = False)
    Dim docNo As String, status As String
    If r < 2 Then Err.Raise vbObjectError + 718, , "Saved invoice not found."
    docNo = CStr(ws.Cells(r, c.colNo).Value)
    status = NrmID(CStr(ws.Cells(r, c.colStatus).Value))
    If HasPayments(docNo) Or Num(ws.Cells(r, c.colPaid).Value) <> 0 Then _
        Err.Raise vbObjectError + 718, , "An invoice with payment history cannot be replaced or cancelled: " & docNo
    If Not IsNumeric(ws.Cells(r, c.colTotal).Value) Or Not IsNumeric(ws.Cells(r, c.colPaid).Value) Or _
        Not IsNumeric(ws.Cells(r, c.colBalance).Value) Then _
        Err.Raise vbObjectError + 718, , "Invoice financial values are invalid: " & docNo
    If Num(ws.Cells(r, c.colTotal).Value) < 0 Or Num(ws.Cells(r, c.colBalance).Value) < 0 Then _
        Err.Raise vbObjectError + 718, , "Invoice financial values are negative: " & docNo
    If status <> "UNPAID" And status <> "PENDING" And status <> "OVERDUE" Then
        If Not (status = "PAID" And Num(ws.Cells(r, c.colTotal).Value) = 0) Then _
            Err.Raise vbObjectError + 718, , "Invoice is not unpaid: " & docNo
    End If
    If checkSent Then
        If InvoiceWasSent(ws, r, True) Then Err.Raise vbObjectError + 718, , "Sent invoice, draft or unconfirmed send attempt cannot be changed: " & docNo
    End If
    If Num(ws.Cells(r, c.colBalance).Value) < Num(ws.Cells(r, c.colTotal).Value) - 0.005 Then _
        Err.Raise vbObjectError + 718, , "Invoice balance is inconsistent: " & docNo
End Sub

Private Function SnapshotText(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long) As String
    Dim c As Long, text As String
    For c = 1 To lastCol
        text = text & CStr(c) & "=" & CStr(ws.Cells(r, c).Value) & "; "
    Next c
    SnapshotText = text
End Function

' Retain lines and headers; audit captures financial values before zeroing.
Private Sub SoftVoid(ByVal ws As Worksheet, ByVal r As Long, ByRef cfg As DocConfig, ByVal reason As String)
    Dim subCol As Long, discCol As Long, vatCol As Long, notesCol As Long, modCol As Long
    If cfg.docKind = "MC" Then
        subCol = ML_SUBTOTAL: discCol = ML_DISC: vatCol = ML_VAT
        notesCol = ML_NOTES: modCol = ML_MODIFIED
    Else
        subCol = IL_SUBTOTAL: discCol = IL_DISC: vatCol = IL_VAT
        notesCol = IL_NOTES: modCol = IL_MODIFIED
    End If
    LogAudit "PaymentPlanVoid", CStr(ws.Cells(r, cfg.colNo).Value), _
        SnapshotText(ws, r, 27), "Voided; financial amounts zero", reason & "; original lines retained"
    ws.Cells(r, subCol).Value = 0
    ws.Cells(r, discCol).Value = 0
    ws.Cells(r, vatCol).Value = 0
    ws.Cells(r, cfg.colTotal).Value = 0
    ws.Cells(r, cfg.colPaid).Value = 0
    ws.Cells(r, cfg.colBalance).Value = 0
    ws.Cells(r, cfg.colStatus).Value = "Voided"
    If cfg.docKind = "MC" Then
        ws.Cells(r, ML_DISCPCT).Value = 0
        ws.Cells(r, ML_DISCFIX).Value = 0
    Else
        ws.Cells(r, IL_DISCPCT).Value = 0
        ws.Cells(r, IL_DISCFIX).Value = 0
    End If
    ws.Cells(r, notesCol).Value = CStr(ws.Cells(r, notesCol).Value) & vbLf & reason
    ws.Cells(r, modCol).Value = Now
End Sub

Private Sub ValidateRecipient(ByRef recipient As String, ByRef custID As String, _
        ByVal patientName As String, ByRef dept As String)
    recipient = LCase$(Trim$(recipient))
    If recipient = "private" Then recipient = "patient"
    dept = UCase$(Trim$(dept))
    custID = Trim$(custID)
    If dept <> "WA" And dept <> "WD" Then Err.Raise vbObjectError + 719, , "Department must be WA or WD."
    Select Case recipient
        Case "doctor"
            If custID = "" Or CustIDToDrName(custID) = "" Then Err.Raise vbObjectError + 719, , "Select a saved customer."
        Case "patient"
            If Trim$(patientName) = "" Then Err.Raise vbObjectError + 719, , "Enter the private patient's name."
        Case "medclaim"
            If Trim$(patientName) = "" Then Err.Raise vbObjectError + 719, , "Enter the medical patient's name."
            If custID = "" Then Err.Raise vbObjectError + 719, , "Select a saved medical customer."
            If FindLogRow(ThisWorkbook.Worksheets("Med Customers"), custID) = 0 Then _
                Err.Raise vbObjectError + 719, , "Medical customer was not found."
        Case Else
            Err.Raise vbObjectError + 719, , "Recipient must be doctor, patient or medclaim."
    End Select
End Sub

Public Function PaymentPlanIDBase(ByVal patientName As String, ByVal sourceDocNo As String) As String
    patientName = Trim$(patientName)
    sourceDocNo = Trim$(sourceDocNo)
    If patientName = "" Then Err.Raise 5, , "Enter the patient name for the Plan ID."
    If sourceDocNo = "" Then sourceDocNo = "Manual"
    PaymentPlanIDBase = patientName & " - " & sourceDocNo
End Function

Public Function PaymentSourceMatches(ByVal docNo As String, ByVal patient As String, _
                                     ByVal customer As String, ByVal query As String) As Boolean
    query = Trim$(query)
    PaymentSourceMatches = (query = "" Or InStr(1, docNo, query, vbTextCompare) > 0 Or _
                           InStr(1, patient, query, vbTextCompare) > 0 Or _
                           InStr(1, customer, query, vbTextCompare) > 0)
End Function

Private Function NewPlanID(ByVal ws As Worksheet, ByVal patientName As String, ByVal sourceDocNo As String) As String
    Dim n As Long, key As String, base As String
    base = PaymentPlanIDBase(patientName, sourceDocNo)
    key = base
    Do
        If FindLogRow(ws, key) = 0 Then NewPlanID = key: Exit Function
        n = n + 1
        key = base & " (" & CStr(n + 1) & ")"
    Loop
End Function

Private Function PlanRows(ByVal ws As Worksheet, ByVal planID As String) As Collection
    Dim rows As New Collection, r As Long
    For r = 2 To LastPlanRow(ws)
        If NrmID(CStr(ws.Cells(r, PP_PLANID).Value)) = NrmID(planID) Then
            If Num(ws.Cells(r, PP_NUMBER).Value) = 0 And rows.Count > 0 Then
                rows.Add r, Before:=1
            Else
                rows.Add r
            End If
        End If
    Next r
    Set PlanRows = rows
End Function

Private Sub GuardPlan(ByVal ws As Worksheet, ByVal rows As Collection)
    Dim item As Variant, il As Worksheet, cfg As DocConfig, r As Long, docNo As String
    Dim seen As Object, firstRow As Long, term As Long, n As Double, amountSum As Currency, lr As Long
    If rows.Count = 0 Then Err.Raise vbObjectError + 720, , "Payment plan not found."
    Set il = ThisWorkbook.Worksheets("InvoiceLog")
    cfg = GetDocConfig("INV")
    Set seen = CreateObject("Scripting.Dictionary")
    firstRow = CLng(rows(1))
    term = CLng(Num(ws.Cells(firstRow, PP_COUNT).Value))
    If term < 1 Or rows.Count <> term + 1 Then Err.Raise vbObjectError + 720, , "Payment plan schedule is incomplete."
    For Each item In rows
        r = CLng(item)
        n = Num(ws.Cells(r, PP_NUMBER).Value)
        If Not IsNumeric(ws.Cells(r, PP_NUMBER).Value) Or n <> Fix(n) Or n < 0 Or n > term Then _
            Err.Raise vbObjectError + 720, , "Invalid installment number."
        If seen.Exists(CStr(n)) Then Err.Raise vbObjectError + 720, , "Duplicate installment number."
        seen.Add CStr(n), True
        If NrmID(CStr(ws.Cells(r, PP_SOURCE).Value)) <> NrmID(CStr(ws.Cells(firstRow, PP_SOURCE).Value)) Or _
            Num(ws.Cells(r, PP_TOTAL).Value) <> Num(ws.Cells(firstRow, PP_TOTAL).Value) Then _
            Err.Raise vbObjectError + 720, , "Payment plan source/total is inconsistent."
        If IsCancelled(ws, r) Then Err.Raise vbObjectError + 720, , "Payment plan is already cancelled."
        If WasReminded(ws, r) Then Err.Raise vbObjectError + 720, , "A sent payment plan cannot be changed."
        docNo = CStr(ws.Cells(r, PP_INVOICE).Value)
        lr = FindLogRow(il, docNo)
        AssertUnpaid il, lr, cfg, True
        If Num(ws.Cells(r, PP_AMOUNT).Value) <> Num(il.Cells(lr, IL_TOTAL).Value) Then _
            Err.Raise vbObjectError + 720, , "Plan and invoice amounts do not agree."
        amountSum = amountSum + CCur(Num(ws.Cells(r, PP_AMOUNT).Value))
    Next item
    If amountSum <> CCur(Num(ws.Cells(firstRow, PP_TOTAL).Value)) Then _
        Err.Raise vbObjectError + 720, , "Plan amounts do not reconcile to the original total."
End Sub

Public Function GeneratePlanInvoices(ByVal sourceDocNo As String, ByVal aligners As Double, _
        Optional ByVal total As Currency = 0, Optional ByVal recipientType As String = "", _
        Optional ByVal custID As String = "", Optional ByVal patientName As String = "", _
        Optional ByVal dept As String = "", Optional ByVal startDate As Date = 0, _
        Optional ByVal depositPercent As Double = -1) As String
    GeneratePlanInvoices = BuildPlan(sourceDocNo, aligners, total, recipientType, custID, _
                                    patientName, dept, startDate, "", depositPercent)
End Function

Private Function BuildPlan(ByVal sourceDocNo As String, ByVal aligners As Double, _
        ByVal total As Currency, ByVal recipient As String, ByVal custID As String, _
        ByVal patientName As String, ByVal dept As String, ByVal startDate As Date, _
        Optional ByVal retiringPlan As String = "", Optional ByVal depositPercent As Double = -1) As String
    Dim ws As Worksheet, src As Worksheet, il As Worksheet, lines As Worksheet
    Dim cfg As DocConfig, invCfg As DocConfig, p As PaymentPlanResult, sourceRow As Long
    Dim r As Long, i As Long, pr As Long, amount As Currency, due As Date
    Dim id As String, docNo As String, kind As String, sourceTotal As Currency
    Dim oldRows As Collection, snapshots As New Collection, generated As New Collection
    Dim item As Variant, saved As Variant, errText As String, oldEvents As Boolean
    Dim oldScreen As Boolean, ownsWrite As Boolean, planStarted As Boolean
    Dim vatFraction As Double, vatAmount As Currency
    Dim targetVat As Currency, allocatedVat As Currency, remainingGross As Currency, finalDate As Date
    Dim planTotal As Currency
    On Error GoTo Failed
    If mWriting Then Err.Raise vbObjectError + 721, , "A payment plan update is already running."
    mWriting = True: ownsWrite = True
    Set ws = EnsurePaymentPlansSheet()
    Set il = ThisWorkbook.Worksheets("InvoiceLog")
    Set lines = ThisWorkbook.Worksheets("InvoiceLines")
    ' Required before any financial write: a missing/protected audit log fails closed.
    If ThisWorkbook.Worksheets("AuditLog").ProtectContents Then Err.Raise vbObjectError + 721, , "AuditLog is protected."
    If il.ProtectContents Or lines.ProtectContents Or ws.ProtectContents Then _
        Err.Raise vbObjectError + 721, , "Payment plan sheets are protected."
    Set oldRows = PlanRows(ws, retiringPlan)
    If retiringPlan <> "" Then GuardPlan ws, oldRows
    sourceDocNo = Trim$(sourceDocNo)
    If sourceDocNo <> "" Then
        If Left$(NrmID(sourceDocNo), 7) = "MC-INV-" Then
            kind = "MC"
        ElseIf Left$(NrmID(sourceDocNo), 2) = "Q-" Then
            kind = "QTE"
        ElseIf Left$(NrmID(sourceDocNo), 4) = "INV-" Then
            kind = "INV"
        Else
            Err.Raise vbObjectError + 722, , "Choose a saved quote, invoice or medical claim."
        End If
        cfg = GetDocConfig(kind)
        Set src = ThisWorkbook.Worksheets(cfg.logName)
        If src.ProtectContents Then Err.Raise vbObjectError + 722, , "Source log is protected."
        sourceRow = FindLogRow(src, sourceDocNo)
        If sourceRow = 0 Then Err.Raise vbObjectError + 722, , "Source document is not saved."
        sourceDocNo = CStr(src.Cells(sourceRow, cfg.colNo).Value)
        If kind = "QTE" Then
            If Trim$(CStr(src.Cells(sourceRow, QL_CONVINV).Value)) <> "" Or _
                NrmID(CStr(src.Cells(sourceRow, QL_STATUS).Value)) = "CONVERTED" Then
                If retiringPlan = "" Then Err.Raise vbObjectError + 722, , "This quote is converted. Select its invoice instead."
                If NrmID(CStr(src.Cells(sourceRow, QL_CONVINV).Value)) <> _
                    NrmID(CStr(ws.Cells(CLng(oldRows(1)), PP_INVOICE).Value)) Then _
                    Err.Raise vbObjectError + 722, , "Quote conversion no longer refers to this plan."
            End If
            If NrmID(CStr(src.Cells(sourceRow, QL_STATUS).Value)) = "CANCELLED" Or _
                NrmID(CStr(src.Cells(sourceRow, QL_STATUS).Value)) = "VOIDED" Then _
                Err.Raise vbObjectError + 722, , "Source quote is cancelled."
            recipient = CStr(src.Cells(sourceRow, QL_RECIP).Value)
            custID = CStr(src.Cells(sourceRow, QL_CUST).Value)
            patientName = CStr(src.Cells(sourceRow, QL_PATIENT).Value)
        Else
            If retiringPlan = "" Then AssertUnpaid src, sourceRow, cfg, (kind = "INV")
            If retiringPlan <> "" And NrmID(CStr(src.Cells(sourceRow, cfg.colStatus).Value)) <> "VOIDED" Then _
                Err.Raise vbObjectError + 722, , "Amendment source is no longer voided."
            If kind = "MC" Then
                recipient = "medclaim"
                custID = CStr(src.Cells(sourceRow, ML_CUST).Value)
                patientName = CStr(src.Cells(sourceRow, ML_PATIENT).Value)
            Else
                recipient = CStr(src.Cells(sourceRow, IL_RECIP).Value)
                custID = CStr(src.Cells(sourceRow, IL_CUST).Value)
                patientName = CStr(src.Cells(sourceRow, IL_PATIENT).Value)
            End If
        End If
        dept = CStr(src.Cells(sourceRow, cfg.colDept).Value)
        sourceTotal = CCur(Num(src.Cells(sourceRow, cfg.colTotal).Value))
        If retiringPlan = "" Or kind = "QTE" Then
            total = sourceTotal
        Else
            ' Invoice/MC sources were voided; the saved plan retains their original total.
            total = CCur(Num(ws.Cells(CLng(oldRows(1)), PP_TOTAL).Value))
        End If
        For r = 2 To LastPlanRow(ws)
            If NrmID(CStr(ws.Cells(r, PP_INVOICE).Value)) = NrmID(sourceDocNo) And Not IsCancelled(ws, r) Then _
                Err.Raise vbObjectError + 723, , "An active payment-plan invoice cannot be used as a new source."
            If NrmID(CStr(ws.Cells(r, PP_SOURCE).Value)) = NrmID(sourceDocNo) And Not IsCancelled(ws, r) Then
                If NrmID(CStr(ws.Cells(r, PP_PLANID).Value)) <> NrmID(retiringPlan) Then _
                    Err.Raise vbObjectError + 723, , "This source already has a payment plan."
            End If
        Next r
    End If
    ValidateRecipient recipient, custID, patientName, dept
    If startDate = 0 Then startDate = Date
    p = CalcPaymentPlan(aligners, total, depositPercent)
    planTotal = CCur(Fix(CDec(total) * 100 + CDec(0.5)) / 100)
    vatFraction = 0.15 / 1.15
    If Not src Is Nothing And (retiringPlan = "" Or kind = "QTE") Then
        If sourceTotal <= 0 Then Err.Raise vbObjectError + 722, , "Source total must be positive."
        Select Case kind
            Case "QTE": vatFraction = Num(src.Cells(sourceRow, QL_VAT).Value) / sourceTotal
            Case "MC": vatFraction = Num(src.Cells(sourceRow, ML_VAT).Value) / sourceTotal
            Case "INV": vatFraction = Num(src.Cells(sourceRow, IL_VAT).Value) / sourceTotal
        End Select
    ElseIf retiringPlan <> "" Then
        vatAmount = 0
        For Each item In oldRows
            r = FindLogRow(il, CStr(ws.Cells(CLng(item), PP_INVOICE).Value))
            vatAmount = vatAmount + CCur(Num(il.Cells(r, IL_VAT).Value))
        Next item
        vatFraction = vatAmount / CCur(Num(ws.Cells(CLng(oldRows(1)), PP_TOTAL).Value))
    End If
    If vatFraction < 0 Or vatFraction >= 1 Then Err.Raise vbObjectError + 722, , "Source tax amounts are inconsistent."
    targetVat = CCur(Fix(CDec(planTotal) * CDec(vatFraction) * 100 + CDec(0.5)) / 100)
    remainingGross = planTotal
    finalDate = DateAdd("m", p.InstallmentCount, startDate)
    If LastPlanRow(ws) + p.InstallmentCount + 1 > ws.Rows.Count Or _
        il.Cells(il.Rows.Count, IL_NO).End(xlUp).Row + p.InstallmentCount + 1 > il.Rows.Count Or _
        lines.Cells(lines.Rows.Count, LN_DOCNO).End(xlUp).Row + p.InstallmentCount + 1 > lines.Rows.Count Then _
        Err.Raise vbObjectError + 722, , "There are not enough worksheet rows for this plan."
    id = NewPlanID(ws, patientName, sourceDocNo)
    invCfg = GetDocConfig("INV")
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    planStarted = True
    LogAudit "PaymentPlanCreate", id, "", CStr(planTotal), "Source=" & sourceDocNo & "; aligners=" & CStr(aligners)
    For i = 0 To p.InstallmentCount
        amount = p.Schedule(i)
        due = DateAdd("m", i, startDate)
        vatAmount = CCur(Fix(CDec(amount) * CDec(vatFraction) * 100 + CDec(0.5)) / 100)
        If vatAmount > targetVat - allocatedVat Then vatAmount = targetVat - allocatedVat
        remainingGross = remainingGross - amount
        If targetVat - allocatedVat - vatAmount > remainingGross Then _
            vatAmount = targetVat - allocatedVat - remainingGross
        allocatedVat = allocatedVat + vatAmount
        docNo = modConvert.GeneratePaymentPlanInvoice(id, sourceDocNo, recipient, custID, _
            patientName, dept, startDate, due, i, amount, vatAmount)
        r = FindLogRow(il, docNo)
        If r = 0 Then Err.Raise vbObjectError + 724, , "Invoice adapter did not save the invoice."
        generated.Add Array(r, docNo)
        pr = LastPlanRow(ws) + 1
        ws.Cells(pr, PP_PLANID).Value = id
        ws.Cells(pr, PP_SOURCE).Value = sourceDocNo
        ws.Cells(pr, PP_DEPT).Value = dept
        ws.Cells(pr, PP_RECIP).Value = recipient
        ws.Cells(pr, PP_CUST).Value = custID
        ws.Cells(pr, PP_ALIGNERS).Value = aligners
        ws.Cells(pr, PP_TOTAL).Value = planTotal
        ws.Cells(pr, PP_DEPOSIT).Value = p.Deposit
        ws.Cells(pr, PP_COUNT).Value = p.InstallmentCount
        ws.Cells(pr, PP_NUMBER).Value = i
        ws.Cells(pr, PP_INVOICE).Value = docNo
        ws.Cells(pr, PP_DUE).Value = due
        ws.Cells(pr, PP_AMOUNT).Value = amount
        ws.Cells(pr, PP_PAID).Value = 0
        ws.Cells(pr, PP_BALANCE).Value = amount
        ws.Cells(pr, PP_STATUS).Value = "Pending"
        ws.Cells(pr, PP_CREATED).Value = Now
        ws.Cells(pr, PP_MODIFIED).Value = Now
        ws.Cells(pr, PP_PATIENT).Value = patientName
        ws.Cells(pr, PP_GENERATED).Value = True
    Next i
    ' Preflight again immediately before the commit; keep restoration snapshots.
    If retiringPlan <> "" Then
        GuardPlan ws, oldRows
        For Each item In oldRows
            r = FindLogRow(il, CStr(ws.Cells(CLng(item), PP_INVOICE).Value))
            snapshots.Add Array(il.Name, r, il.Cells(r, 1).Resize(1, 27).Formula)
            snapshots.Add Array(ws.Name, CLng(item), ws.Cells(CLng(item), 1).Resize(1, PP_LASTCOL).Formula)
        Next item
        RetirePlan ws, oldRows, "Replaced by " & id
    End If
    If Not src Is Nothing Then
        If retiringPlan = "" Or kind = "QTE" Then
            snapshots.Add Array(src.Name, sourceRow, src.Cells(sourceRow, 1).Resize(1, 27).Formula)
            If kind = "QTE" Then
                LogAudit "PaymentPlanQuote", sourceDocNo, SnapshotText(src, sourceRow, 27), "Converted to " & id, "Plan source"
                src.Cells(sourceRow, QL_CONVINV).Value = CStr(ws.Cells(LastPlanRow(ws) - p.InstallmentCount, PP_INVOICE).Value)
                src.Cells(sourceRow, QL_STATUS).Value = "Converted"
                src.Cells(sourceRow, QL_MODIFIED).Value = Now
            Else
                AssertUnpaid src, sourceRow, cfg, (kind = "INV")
                SoftVoid src, sourceRow, cfg, "Replaced by payment plan " & id
            End If
        End If
    End If
    LogAudit "PaymentPlanCommitted", id, "", CStr(planTotal), "All invoices saved; source=" & sourceDocNo
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    mWriting = False
    BuildPlan = id
    Exit Function
Failed:
    errText = Err.Description
    On Error Resume Next
    For Each saved In snapshots
        ThisWorkbook.Worksheets(CStr(saved(0))).Cells(CLng(saved(1)), 1).Resize(1, UBound(saved(2), 2)).Formula = saved(2)
        LogAudit "PaymentPlanRollbackRestore", id, "", CStr(saved(0)) & " row " & CStr(saved(1)), errText
    Next saved
    For Each saved In generated
        r = CLng(saved(0))
        If CStr(il.Cells(r, IL_NO).Value) = "" Then il.Cells(r, IL_NO).Value = CStr(saved(1))
        SoftVoid il, r, invCfg, "Rolled back payment plan " & id & ": " & errText
        ' Creation was audited before writing; ensure rollback still zeros amounts
        ' if the audit sheet itself becomes unavailable during the failure.
        il.Cells(r, IL_SUBTOTAL).Value = 0
        il.Cells(r, IL_DISC).Value = 0
        il.Cells(r, IL_VAT).Value = 0
        il.Cells(r, IL_TOTAL).Value = 0
        il.Cells(r, IL_PAID).Value = 0
        il.Cells(r, IL_BALANCE).Value = 0
        il.Cells(r, IL_STATUS).Value = "Voided"
    Next saved
    If Not ws Is Nothing Then
        For r = 2 To LastPlanRow(ws)
            If id <> "" And CStr(ws.Cells(r, PP_PLANID).Value) = id Then
                ws.Cells(r, PP_STATUS).Value = "Cancelled"
                ws.Cells(r, PP_BALANCE).Value = 0
                ws.Cells(r, PP_MODIFIED).Value = Now
            End If
        Next r
    End If
    If planStarted Then
        Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
        LogAudit "PaymentPlanRollback", id, "", "Cancelled", errText
    End If
    If ownsWrite Then mWriting = False
    On Error GoTo 0
    MsgBox "Payment plan was not created: " & errText, vbExclamation, "Payment options"
End Function

Private Sub RetirePlan(ByVal ws As Worksheet, ByVal rows As Collection, ByVal reason As String)
    Dim il As Worksheet, cfg As DocConfig, item As Variant, r As Long
    Set il = ThisWorkbook.Worksheets("InvoiceLog")
    If il.ProtectContents Or ws.ProtectContents Or ThisWorkbook.Worksheets("AuditLog").ProtectContents Then _
        Err.Raise vbObjectError + 725, , "Payment plan logs are protected."
    cfg = GetDocConfig("INV")
    For Each item In rows
        r = CLng(item)
        SoftVoid il, FindLogRow(il, CStr(ws.Cells(r, PP_INVOICE).Value)), cfg, reason
        LogAudit "PaymentPlanCancelRow", CStr(ws.Cells(r, PP_PLANID).Value), _
            SnapshotText(ws, r, PP_LASTCOL), "Cancelled", reason
        ws.Cells(r, PP_STATUS).Value = "Cancelled"
        ws.Cells(r, PP_BALANCE).Value = 0
        ws.Cells(r, PP_MODIFIED).Value = Now
    Next item
End Sub

Public Function CancelPaymentPlan(ByVal planID As String) As Boolean
    Dim ws As Worksheet, rows As Collection, il As Worksheet, item As Variant
    Dim snapshots As New Collection, saved As Variant, r As Long, errText As String
    Dim oldEvents As Boolean, ownsWrite As Boolean
    On Error GoTo Failed
    If mWriting Then Err.Raise vbObjectError + 725, , "A payment plan update is already running."
    mWriting = True: ownsWrite = True
    oldEvents = Application.EnableEvents
    Application.EnableEvents = False
    Set ws = EnsurePaymentPlansSheet()
    Set rows = PlanRows(ws, planID)
    GuardPlan ws, rows
    Set il = ThisWorkbook.Worksheets("InvoiceLog")
    For Each item In rows
        r = FindLogRow(il, CStr(ws.Cells(CLng(item), PP_INVOICE).Value))
        snapshots.Add Array(il.Name, r, il.Cells(r, 1).Resize(1, 27).Formula)
        snapshots.Add Array(ws.Name, CLng(item), ws.Cells(CLng(item), 1).Resize(1, PP_LASTCOL).Formula)
    Next item
    RetirePlan ws, rows, "Cancelled plan " & planID
    LogAudit "PaymentPlanCancelled", planID, "", "Cancelled", "Source remains void; financial history retained"
    CancelPaymentPlan = True
    Application.EnableEvents = oldEvents
    mWriting = False
    Exit Function
Failed:
    errText = Err.Description
    On Error Resume Next
    For Each saved In snapshots
        ThisWorkbook.Worksheets(CStr(saved(0))).Cells(CLng(saved(1)), 1).Resize(1, UBound(saved(2), 2)).Formula = saved(2)
    Next saved
    If ownsWrite Then
        LogAudit "PaymentPlanCancelRollback", planID, "", "Restored", errText
        Application.EnableEvents = oldEvents
        mWriting = False
    End If
    On Error GoTo 0
    MsgBox "Payment plan was not cancelled: " & errText, vbExclamation, "Payment options"
End Function

Public Function AmendPaymentPlan(ByVal planID As String, ByVal aligners As Double, _
        Optional ByVal total As Currency = 0, Optional ByVal depositPercent As Double = -1) As String
    Dim ws As Worksheet, rows As Collection, r As Long
    On Error GoTo Failed
    Set ws = EnsurePaymentPlansSheet()
    Set rows = PlanRows(ws, planID)
    GuardPlan ws, rows
    r = CLng(rows(1))
    If total = 0 Then total = CCur(Num(ws.Cells(r, PP_TOTAL).Value))
    If depositPercent = -1 Then
        depositPercent = Num(ws.Cells(r, PP_DEPOSIT).Value) / Num(ws.Cells(r, PP_TOTAL).Value)
    End If
    AmendPaymentPlan = BuildPlan(CStr(ws.Cells(r, PP_SOURCE).Value), aligners, total, _
        CStr(ws.Cells(r, PP_RECIP).Value), CStr(ws.Cells(r, PP_CUST).Value), _
        CStr(ws.Cells(r, PP_PATIENT).Value), CStr(ws.Cells(r, PP_DEPT).Value), _
        CDate(ws.Cells(r, PP_DUE).Value), planID, depositPercent)
    Exit Function
Failed:
    MsgBox "Payment plan was not amended: " & Err.Description, vbExclamation, "Payment options"
End Function

Public Sub RefreshPaymentPlans()
    Dim ws As Worksheet, il As Worksheet, r As Long, lr As Long, docNo As String
    Dim deposits As Object, id As String, status As String, balance As Double, oldText As String
    On Error GoTo Failed
    If mWriting Or mRefreshing Then Exit Sub
    mRefreshing = True
    Set ws = EnsurePaymentPlansSheet()
    Set il = ThisWorkbook.Worksheets("InvoiceLog")
    Set deposits = CreateObject("Scripting.Dictionary")
    For r = 2 To LastPlanRow(ws)
        If Not IsCancelled(ws, r) Then
            docNo = CStr(ws.Cells(r, PP_INVOICE).Value)
            lr = FindLogRow(il, docNo)
            If lr = 0 Then Err.Raise vbObjectError + 726, , "Plan invoice missing: " & docNo
            If NrmID(CStr(il.Cells(lr, IL_STATUS).Value)) = "VOIDED" Then _
                Err.Raise vbObjectError + 726, , "An active plan invoice was voided: " & docNo
            oldText = SnapshotText(il, lr, 27)
            status = NrmID(CStr(il.Cells(lr, IL_STATUS).Value))
            If status = "SENT" Or status = "EMAILED" Then _
                LogAudit "PaymentPlanSentObserved", docNo, status, "Sent", "Retain send evidence before reconciliation"
            LogAudit "PaymentPlanReconcile", docNo, oldText, "Reconcile from Payments", "Payment plan refresh"
            ReconcileDoc "invoice", docNo
            ws.Cells(r, PP_PAID).Value = Num(il.Cells(lr, IL_PAID).Value)
            ws.Cells(r, PP_BALANCE).Value = Num(il.Cells(lr, IL_BALANCE).Value)
            If Num(ws.Cells(r, PP_NUMBER).Value) = 0 Then
                id = NrmID(CStr(ws.Cells(r, PP_PLANID).Value))
                If deposits.Exists(id) Then Err.Raise vbObjectError + 726, , "Duplicate deposit row."
                deposits.Add id, (Num(ws.Cells(r, PP_BALANCE).Value) <= 0.005)
            End If
        End If
    Next r
    For r = 2 To LastPlanRow(ws)
        If Not IsCancelled(ws, r) Then
            id = NrmID(CStr(ws.Cells(r, PP_PLANID).Value))
            If Not deposits.Exists(id) Then Err.Raise vbObjectError + 726, , "Plan deposit row missing."
            balance = Num(ws.Cells(r, PP_BALANCE).Value)
            status = "Pending"
            If deposits(id) Then
                status = "Active"
                If balance <= 0.005 Then
                    status = "Paid"
                ElseIf CDate(ws.Cells(r, PP_DUE).Value) < Date Then
                    status = "Overdue"
                End If
            ElseIf Num(ws.Cells(r, PP_NUMBER).Value) = 0 And balance > 0.005 Then
                If CDate(ws.Cells(r, PP_DUE).Value) < Date Then status = "Overdue"
            End If
            If CStr(ws.Cells(r, PP_STATUS).Value) <> status Then
                LogAudit "PaymentPlanStatus", id, CStr(ws.Cells(r, PP_STATUS).Value), status, _
                    CStr(ws.Cells(r, PP_INVOICE).Value)
                ws.Cells(r, PP_STATUS).Value = status
            End If
            ws.Cells(r, PP_MODIFIED).Value = Now
        End If
    Next r
    mRefreshing = False
    Exit Sub
Failed:
    mRefreshing = False
    Err.Raise vbObjectError + 727, "RefreshPaymentPlans", Err.Description
End Sub

Public Function DueInstallmentRows() As Collection
    Dim rows As New Collection, ws As Worksheet, r As Long, status As String, lead As Long
    Dim il As Worksheet
    RefreshPaymentPlans
    Set ws = EnsurePaymentPlansSheet()
    Set il = ThisWorkbook.Worksheets("InvoiceLog")
    lead = GetReminderLeadDays()
    For r = 2 To LastPlanRow(ws)
        status = CStr(ws.Cells(r, PP_STATUS).Value)
        If Num(ws.Cells(r, PP_NUMBER).Value) > 0 And _
            Num(ws.Cells(r, PP_BALANCE).Value) > 0.005 And _
            (status = "Active" Or status = "Overdue") And _
            Not WasReminded(ws, r) Then
            If CDate(ws.Cells(r, PP_DUE).Value) <= Date + lead Then
                If Not InvoiceWasSent(il, FindLogRow(il, CStr(ws.Cells(r, PP_INVOICE).Value))) Then rows.Add r
            End If
        End If
    Next r
    Set DueInstallmentRows = rows
End Function

Public Sub ShowInstallmentReminder()
    Dim rows As Collection, form As Object
    On Error GoTo Failed
    Set rows = DueInstallmentRows()
    If rows.Count > 0 Then
        Set form = VBA.UserForms.Add("frmPaymentOptions")
        form.ShowReminder
        Unload form
    End If
    Exit Sub
Failed:
    MsgBox "Installment reminders unavailable: " & Err.Description, vbExclamation, "Payment options"
End Sub

Private Function RecipientEmail(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim contacts As Worksheet, name As Variant, keyCol As Long, emailCol As Long, billCol As Long
    Dim i As Long, key As String, recipient As String, billKey As String, contactSheets As Variant
    recipient = LCase$(CStr(ws.Cells(r, PP_RECIP).Value))
    If recipient = "doctor" Then
        Set contacts = ThisWorkbook.Worksheets("Customers")
        key = NrmID(CStr(ws.Cells(r, PP_CUST).Value))
        For i = 2 To contacts.Cells(contacts.Rows.Count, 1).End(xlUp).Row
            If NrmID(CStr(contacts.Cells(i, 1).Value)) = key Then
                RecipientEmail = Trim$(CStr(contacts.Cells(i, 7).Value))
                Exit Function
            End If
        Next i
    Else
        key = NrmID(CStr(ws.Cells(r, PP_PATIENT).Value))
        billKey = NrmID(CStr(ws.Cells(r, PP_CUST).Value))
        If recipient = "medclaim" Then
            contactSheets = Array("Med Customers", "Patients")
        Else
            contactSheets = Array("Patients", "Med Customers")
        End If
        For Each name In contactSheets
            Set contacts = Nothing
            On Error Resume Next
            Set contacts = ThisWorkbook.Worksheets(CStr(name))
            On Error GoTo 0
            If Not contacts Is Nothing Then
                keyCol = HeaderColumn(contacts, "PatientName", "Patient Name", "Name")
                emailCol = HeaderColumn(contacts, "Email", "E-mail", "Email Address")
                billCol = HeaderColumn(contacts, "BillTo", "Bill To", "CustID", "CustomerID", "PatientID")
                If billCol > 0 And emailCol > 0 And billKey <> "" Then
                    For i = 2 To contacts.Cells(contacts.Rows.Count, billCol).End(xlUp).Row
                        If NrmID(CStr(contacts.Cells(i, billCol).Value)) = billKey Then
                            RecipientEmail = Trim$(CStr(contacts.Cells(i, emailCol).Value))
                            If RecipientEmail <> "" Then Exit Function
                        End If
                    Next i
                End If
                If keyCol > 0 And emailCol > 0 Then
                    For i = 2 To contacts.Cells(contacts.Rows.Count, keyCol).End(xlUp).Row
                        If NrmID(CStr(contacts.Cells(i, keyCol).Value)) = key Then
                            RecipientEmail = Trim$(CStr(contacts.Cells(i, emailCol).Value))
                            If RecipientEmail <> "" Then Exit Function
                        End If
                    Next i
                End If
            End If
        Next name
    End If
End Function

Public Sub SavePlanInvoicePDFs(ByVal planID As String)
    Dim ws As Worksheet, rows As Collection, item As Variant, path As String, count As Long
    On Error GoTo Failed
    Set ws = EnsurePaymentPlansSheet()
    Set rows = PlanRows(ws, planID)
    If rows.Count = 0 Then Err.Raise 5, , "Payment plan not found."
    If MsgBox("Save the deposit and installment invoices as PDFs? Choose a filename/location for each." & _
              vbCrLf & "Cancel a Save As dialog to stop. Invoices remain saved in the workbook.", _
              vbQuestion + vbYesNo, "Save plan invoices") <> vbYes Then Exit Sub
    For Each item In rows
        If Not IsCancelled(ws, CLng(item)) Then
            path = modExportPDF.ChoosePaymentInvoicePDF(CStr(ws.Cells(CLng(item), PP_INVOICE).Value), _
                                                       CStr(ws.Cells(CLng(item), PP_PATIENT).Value))
            If path = "" Then Exit Sub
            ExportPaymentInvoicePDF CLng(item), path
            count = count + 1
        End If
    Next item
    MsgBox CStr(count) & " invoice PDFs saved.", vbInformation
    Exit Sub
Failed:
    MsgBox "PDF saving stopped after " & CStr(count) & " files: " & Err.Description & vbCrLf & _
           "The plan and invoices remain saved. Run BtnSavePaymentPlanPDFs to retry.", vbExclamation
End Sub

Private Function ExportPaymentInvoicePDF(ByVal planRow As Long, ByVal path As String) As String
    Dim ws As Worksheet, invoice As Worksheet, il As Worksheet, lr As Long, docNo As String
    Dim savedDue As Variant, savedLine As Variant, savedTotals As Variant
    Dim oldEvents As Boolean, templateEdited As Boolean, errNo As Long, errText As String
    On Error GoTo Failed
    Set ws = EnsurePaymentPlansSheet()
    If ws.ProtectContents Then Err.Raise 5, , "PaymentPlans is protected."
    If planRow < 2 Or planRow > LastPlanRow(ws) Then Err.Raise 5, , "Payment invoice not found."
    If IsCancelled(ws, planRow) Then Err.Raise 5, , "Payment invoice is cancelled."
    docNo = CStr(ws.Cells(planRow, PP_INVOICE).Value)
    Set il = ThisWorkbook.Worksheets("InvoiceLog")
    lr = FindLogRow(il, docNo)
    If lr = 0 Then Err.Raise 5, , "Saved invoice not found."
    Set invoice = ThisWorkbook.Worksheets("Invoice")
    invoice.Range("G7").ClearContents
    RecallInvoice docNo
    If NrmID(CStr(invoice.Range("G7").Value)) <> NrmID(docNo) Then _
        Err.Raise vbObjectError + 728, , "Invoice recall did not complete."
    If Not IsDate(invoice.Range("G6").Value) Then Err.Raise 5, , "Invoice date did not load."
    If CDate(invoice.Range("G6").Value) <> CDate(il.Cells(lr, IL_DATE).Value) Then _
        Err.Raise 5, , "Invoice template did not finish loading."
    savedDue = invoice.Range("G11").Formula
    savedLine = invoice.Range("E16:F16").Formula
    savedTotals = invoice.Range("H35:H37").Formula
    oldEvents = Application.EnableEvents
    templateEdited = True
    Application.EnableEvents = False
    ' Preserve the saved installment VAT allocation rather than a Pricelist lookup.
    invoice.Range("G11").Value = il.Cells(lr, IL_DUE).Value
    invoice.Range("E16").Value = il.Cells(lr, IL_SUBTOTAL).Value
    invoice.Range("F16").Value = il.Cells(lr, IL_VAT).Value
    invoice.Range("H35").Value = il.Cells(lr, IL_SUBTOTAL).Value
    invoice.Range("H36").Value = il.Cells(lr, IL_VAT).Value
    invoice.Range("H37").Value = il.Cells(lr, IL_TOTAL).Value
    path = modExportPDF.ExportInvoiceForMail(invoice, path)
    If path = "" Then Err.Raise 5, , "Invoice PDF export failed."
    invoice.Range("G11").Formula = savedDue
    invoice.Range("E16:F16").Formula = savedLine
    invoice.Range("H35:H37").Formula = savedTotals
    Application.EnableEvents = oldEvents
    templateEdited = False
    ws.Cells(planRow, PP_PDF).Value = path
    ExportPaymentInvoicePDF = path
    Exit Function
Failed:
    errNo = Err.Number: errText = Err.Description
    If templateEdited Then
        On Error Resume Next
        invoice.Range("G11").Formula = savedDue
        invoice.Range("E16:F16").Formula = savedLine
        invoice.Range("H35:H37").Formula = savedTotals
        Application.EnableEvents = oldEvents
        On Error GoTo 0
    End If
    Err.Raise errNo, "ExportPaymentInvoicePDF", errText
End Function

Public Function MailInstallmentInvoice(ByVal planRow As Long) As Boolean
    Dim ws As Worksheet, outlook As Object, mail As Object
    Dim address As String, path As String, docNo As String, status As String, choice As VbMsgBoxResult
    Dim sent As Boolean, suffix As Long
    Dim errText As String
    On Error GoTo Failed
    RefreshPaymentPlans
    Set ws = EnsurePaymentPlansSheet()
    If ws.ProtectContents Or ThisWorkbook.Worksheets("AuditLog").ProtectContents Then _
        Err.Raise vbObjectError + 728, , "Payment plan tracking logs are protected."
    If planRow < 2 Or planRow > LastPlanRow(ws) Then Err.Raise vbObjectError + 728, , "Select a payment plan invoice."
    If IsCancelled(ws, planRow) Or Num(ws.Cells(planRow, PP_BALANCE).Value) <= 0.005 Then _
        Err.Raise vbObjectError + 728, , "This invoice is cancelled or paid."
    status = CStr(ws.Cells(planRow, PP_STATUS).Value)
    If Num(ws.Cells(planRow, PP_NUMBER).Value) > 0 And status = "Pending" Then _
        Err.Raise vbObjectError + 728, , "The deposit must be paid first."
    docNo = CStr(ws.Cells(planRow, PP_INVOICE).Value)
    address = RecipientEmail(ws, planRow)
    If address = "" Then address = Trim$(InputBox("Enter recipient email:", "Installment invoice"))
    If InStr(2, address, "@") = 0 Or InStr(address, " ") > 0 Or InStr(address, vbCr) > 0 Or _
        InStr(address, vbLf) > 0 Or InStr(address, ";") > 0 Or InStr(address, ",") > 0 Then _
        Err.Raise vbObjectError + 728, , "A valid single recipient email is required."
    choice = MsgBox("Send invoice " & docNo & " to " & address & " now?" & vbCrLf & _
        "Yes = send; No or Cancel = do nothing.", vbYesNoCancel + vbQuestion, "Installment invoice")
    If choice <> vbYes Then Exit Function
    If ThisWorkbook.Path = "" Then Err.Raise vbObjectError + 728, , "Save the workbook before exporting."
    path = ThisWorkbook.Path & Application.PathSeparator & docNo & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    Do While Dir$(path) <> ""
        suffix = suffix + 1
        path = ThisWorkbook.Path & Application.PathSeparator & docNo & "_" & _
            Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(suffix) & ".pdf"
    Loop
    path = ExportPaymentInvoicePDF(planRow, path)
    Set outlook = CreateObject("Outlook.Application")
    Set mail = outlook.CreateItem(0)
    mail.To = address
    mail.Subject = "Payment plan invoice " & docNo
    mail.Body = "Please find your payment plan invoice attached. Due date: " & _
        Format$(CDate(ws.Cells(planRow, PP_DUE).Value), "dd mmm yyyy") & "."
    mail.Attachments.Add path
    LogAudit "PaymentPlanMailAttempt", docNo, "", address, "Explicit send confirmation"
    mail.Send
    sent = True
    LogAudit "PaymentPlanMailSent", docNo, "", address, "Outlook Send completed"
    ws.Cells(planRow, PP_REMINDED).Value = Now
    ws.Cells(planRow, PP_MODIFIED).Value = Now
    MailInstallmentInvoice = True
    Exit Function
Failed:
    errText = Err.Description
    If sent Then
        MailInstallmentInvoice = True
        MsgBox "Outlook accepted the email, but tracking failed: " & errText, vbExclamation, "Payment options"
    Else
        MsgBox "Email was not sent: " & errText & vbCrLf & _
            "Use the saved PDF to email manually; no reminder was marked.", vbExclamation, "Payment options"
    End If
End Function

Public Sub RefreshPaymentOptionsDashboard()
    Dim ws As Worksheet, r As Long, dueMonth As Long, overdue As Long, d As Date, status As String
    On Error GoTo Failed
    RefreshPaymentPlans
    Set ws = EnsurePaymentPlansSheet()
    For r = 2 To LastPlanRow(ws)
        If Not IsCancelled(ws, r) And Num(ws.Cells(r, PP_NUMBER).Value) > 0 And _
            Num(ws.Cells(r, PP_BALANCE).Value) > 0.005 Then
            d = CDate(ws.Cells(r, PP_DUE).Value)
            status = CStr(ws.Cells(r, PP_STATUS).Value)
            If Month(d) = Month(Date) And Year(d) = Year(Date) Then dueMonth = dueMonth + 1
            If status = "Overdue" Then overdue = overdue + 1
        End If
    Next r
    ThisWorkbook.Worksheets("Menu").Range("B70").Value = "Payment plans: " & _
        CStr(dueMonth) & " due this month; " & CStr(overdue) & " overdue"
    Exit Sub
Failed:
    MsgBox "Payment plan dashboard unavailable: " & Err.Description, vbExclamation, "Payment options"
End Sub

Public Sub RunDepositsStatement()
    RunDepositStatement
End Sub

Public Sub RunInstallmentsStatement()
    Dim ws As Worksheet, rows As New Collection, r As Long, status As String
    On Error GoTo Failed
    RefreshPaymentPlans
    Set ws = EnsurePaymentPlansSheet()
    For r = 2 To LastPlanRow(ws)
        status = CStr(ws.Cells(r, PP_STATUS).Value)
        If Not IsCancelled(ws, r) And Num(ws.Cells(r, PP_NUMBER).Value) > 0 And _
            Num(ws.Cells(r, PP_BALANCE).Value) > 0.005 And _
            (status = "Active" Or status = "Overdue") Then
            If CDate(ws.Cells(r, PP_DUE).Value) <= Date Then rows.Add r
        End If
    Next r
    If rows.Count = 0 Then
        MsgBox "No outstanding installments are due.", vbInformation, "Payment options"
    Else
        modStatement.RunPaymentPlanStatement rows, "Installments statement - ALL recipients / ALL departments"
    End If
    Exit Sub
Failed:
    MsgBox "Installments statement unavailable: " & Err.Description, vbExclamation, "Payment options"
End Sub

Private Sub RunDepositStatement()
    Dim ws As Worksheet, rows As New Collection, recipient As String, dept As String, customer As String
    Dim r As Long, title As String, matches As Boolean
    On Error GoTo Failed
    recipient = LCase$(Trim$(InputBox("Recipient: doctor, patient or medclaim (blank cancels)", "Deposits statement")))
    If recipient = "" Then Exit Sub
    If recipient <> "doctor" And recipient <> "patient" And recipient <> "medclaim" Then _
        Err.Raise vbObjectError + 729, , "Choose doctor, patient or medclaim."
    dept = UCase$(Trim$(InputBox("Department: WA, WD or ALL", "Payment plan statement", "ALL")))
    If dept = "" Then Exit Sub
    If dept <> "WA" And dept <> "WD" And dept <> "ALL" Then Err.Raise vbObjectError + 729, , "Choose WA, WD or ALL."
    customer = Trim$(InputBox("Customer ID / medical Bill To / patient name, or ALL", "Deposits statement", "ALL"))
    If customer = "" Then Exit Sub
    RefreshPaymentPlans
    Set ws = EnsurePaymentPlansSheet()
    For r = 2 To LastPlanRow(ws)
        matches = (LCase$(CStr(ws.Cells(r, PP_RECIP).Value)) = recipient)
        If dept <> "ALL" Then matches = matches And NrmID(CStr(ws.Cells(r, PP_DEPT).Value)) = NrmID(dept)
        If NrmID(customer) <> "ALL" Then
            If recipient = "doctor" Or recipient = "medclaim" Then
                matches = matches And NrmID(CStr(ws.Cells(r, PP_CUST).Value)) = NrmID(customer)
            Else
                matches = matches And NrmID(CStr(ws.Cells(r, PP_PATIENT).Value)) = NrmID(customer)
            End If
        End If
        If matches And Not IsCancelled(ws, r) And Num(ws.Cells(r, PP_BALANCE).Value) > 0.005 Then
            If Num(ws.Cells(r, PP_NUMBER).Value) = 0 Then rows.Add r
        End If
    Next r
    title = "Deposits statement - " & recipient & " / " & dept & " / " & customer
    If rows.Count = 0 Then
        MsgBox "No matching outstanding invoices.", vbInformation, "Payment options"
    Else
        modStatement.RunPaymentPlanStatement rows, title
    End If
    Exit Sub
Failed:
    MsgBox "Payment plan statement unavailable: " & Err.Description, vbExclamation, "Payment options"
End Sub

Public Sub RunSinglePlanStatement(Optional ByVal planID As String = "")
    Dim ws As Worksheet, rows As Collection
    On Error GoTo Failed
    If planID = "" Then planID = Trim$(InputBox("Payment plan ID:", "Payment plan statement"))
    If planID = "" Then Exit Sub
    RefreshPaymentPlans
    Set ws = EnsurePaymentPlansSheet()
    Set rows = PlanRows(ws, planID)
    If rows.Count = 0 Then Err.Raise vbObjectError + 730, , "Payment plan not found."
    modStatement.RunPaymentPlanStatement rows, "Payment plan " & planID
    Exit Sub
Failed:
    MsgBox "Payment plan statement unavailable: " & Err.Description, vbExclamation, "Payment options"
End Sub
