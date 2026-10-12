Attribute VB_Name = "modConvert"
Option Explicit
' ============================================================================
' ConvertQuoteToInvoice � button on the QUOTE sheet.
'   1B FIX: copy line Price INCL (G), not Excl (E) � E/F/H derive.
'   PHASE 3 FIX: copy discounts from C32/C33 (not K5/K6).
'   Copies C14/D14 ApplianceType header across.
' ============================================================================

Private Const QL_STATUS_COL   As Long = 13    ' M = Status
Private Const QL_CONVINV_COL  As Long = 14    ' N = ConvertedInvNo
Private Const IL_SRCQUOTE_COL As Long = 17    ' Q = SourceQuoteNo

Public Sub ConvertQuoteToInvoice()
    Dim wsQ As Worksheet, wsI As Worksheet
    Dim wsQLog As Worksheet, wsILog As Worksheet
    Dim quoteNo As String, qrow As Long, newInv As String
    Dim existingConv As String, irow As Long

    Set wsQ = ThisWorkbook.Sheets("Quote")
    Set wsI = ThisWorkbook.Sheets("Invoice")
    Set wsQLog = ThisWorkbook.Sheets("QuoteLog")
    Set wsILog = ThisWorkbook.Sheets("InvoiceLog")

    ' ---- 1. validate current quote ----
    quoteNo = Trim(CStr(wsQ.Range("G7").value))
    If quoteNo = "" Then
        MsgBox "No saved quote is displayed. Recall or Save a quote first.", vbExclamation
        Exit Sub
    End If

    qrow = FindLogRow(wsQLog, quoteNo)
    If qrow = 0 Then
        MsgBox "Quote " & quoteNo & " not found in QuoteLog. Save it first.", vbExclamation
        Exit Sub
    End If

    ' ---- 2. block re-conversion ----
    existingConv = Trim(CStr(wsQLog.Cells(qrow, QL_CONVINV_COL).value))
    If existingConv <> "" Then
        MsgBox "This quote was already converted to " & existingConv & ".", vbExclamation
        Exit Sub
    End If

    If MsgBox("Convert quote " & quoteNo & " to a new invoice?", _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- 3. start a clean invoice, then copy quote data across ----
    NewInvoice                                   ' resets Invoice sheet + restores formulas

    ' header fields (same layout on both sheets)
    wsI.Range("C6").value = wsQ.Range("C6").value      ' Bill-To / DrName
    wsI.Range("C14").value = wsQ.Range("C14").value    ' Appliance Type
    wsI.Range("D14").value = wsQ.Range("D14").value    ' Appliance Type (2nd cell / merged partner)
    wsI.Range("F14").value = wsQ.Range("F14").value    ' Patient Name

    ' line items � A=Qty, D=Description, G=Price Incl (authoritative; E/F/H derive)
    wsI.Range("A16:A30").value = wsQ.Range("A16:A30").value   ' Qty
    wsI.Range("D16:D30").value = wsQ.Range("D16:D30").value   ' Description
    wsI.Range("G16:G30").value = wsQ.Range("G16:G30").value   ' Price Incl (1B authoritative)

    ' options: dept, recipient, discounts (C32 %, C33 fixed)
    CopyIfExists wsQ, wsI, "K1"     ' Dept
    CopyIfExists wsQ, wsI, "K2"     ' Recipient
    CopyIfExists wsQ, wsI, "C32"    ' Discount %
    CopyIfExists wsQ, wsI, "C33"    ' Discount fixed

    Application.EnableEvents = True     ' allow sheet formulas to recalc
    Application.Calculate

    ' ---- 4. save the invoice (assigns new INV number) ----
    gSuppressClearPrompt = True
    SaveInvoice
    gSuppressClearPrompt = False

    newInv = Trim(CStr(wsI.Range("G7").value))
    If newInv = "" Then Err.Raise 513, , "Invoice number was not assigned by SaveInvoice."

    ' ---- 5. link both logs ----
    irow = FindLogRow(wsILog, newInv)
    If irow > 0 Then wsILog.Cells(irow, IL_SRCQUOTE_COL).value = quoteNo   ' Q SourceQuoteNo

    wsQLog.Cells(qrow, QL_STATUS_COL).value = "Converted"                  ' M Status
    wsQLog.Cells(qrow, QL_CONVINV_COL).value = newInv                      ' N ConvertedInvNo

    LogAudit "Convert", quoteNo, "", "-> " & newInv, "Quote converted to invoice"
    LogAudit "Convert", newInv, "", "<- " & quoteNo, "Invoice created from quote"

    Application.ScreenUpdating = True
    wsI.Activate

    MsgBox "Quote " & quoteNo & " converted to invoice " & newInv & "." & vbCrLf & _
           "Choose where to save the PDF next.", vbInformation

    ' ---- 6. launch PDF export for the new invoice ----
    ExportPDF
    Exit Sub

Fail:
    gSuppressClearPrompt = False        ' always reset the flag on error
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "ConvertQuoteToInvoice error: " & Err.Description, vbExclamation
End Sub

' Saved-source conversion; one full-total charge, never an installment charge.
Public Function GeneratePaymentPlanInvoice(ByVal planID As String, ByVal sourceDocNo As String, _
        ByVal recipientType As String, ByVal custID As String, ByVal patientName As String, _
        ByVal dept As String, ByVal invoiceDate As Date, ByVal dueDate As Date, _
        ByVal installmentNo As Long, ByVal amount As Currency, ByVal vatAmount As Currency) As String
    Dim cfg As DocConfig, il As Worksheet, lines As Worksheet, src As Worksheet, sourceLines As Worksheet
    Dim docNo As String, r As Long, lr As Long, sr As Long, c As Long, i As Long, count As Long
    Dim oldEvents As Boolean, eventsSaved As Boolean, written As Boolean, linked As Boolean
    Dim errNo As Long, errText As String, planRecipient As String, qSnapshot As Variant, header As Variant
    Dim quoteSource As Boolean, firstLine As Long, srcCfg As DocConfig
    On Error GoTo Failed
    dept = UCase$(Trim$(dept)): recipientType = LCase$(Trim$(recipientType))
    planRecipient = recipientType
    If Trim$(planID) = "" Or installmentNo <> 0 Then _
        Err.Raise vbObjectError + 730, , "Only one permanent full-total invoice may be created."
    If dept <> "WA" And dept <> "WD" Then Err.Raise vbObjectError + 730, , "Department must be WA or WD."
    If recipientType = "medclaim" Then
        If custID = "" Then Err.Raise vbObjectError + 730, , "Medical customer required."
        If FindLogRow(ThisWorkbook.Worksheets("Med Customers"), custID) = 0 Then _
            Err.Raise vbObjectError + 730, , "Medical customer not found."
        ' Ordinary Invoice recall accepts patient layout; plan classification remains medclaim.
        recipientType = "patient"
    End If
    If recipientType = "doctor" Then
        If custID = "" Or CustIDToDrName(custID) = "" Then Err.Raise vbObjectError + 730, , "Saved customer required."
    ElseIf recipientType = "patient" Then
        If Trim$(patientName) = "" Then Err.Raise vbObjectError + 730, , "Patient name required."
    Else
        Err.Raise vbObjectError + 730, , "Recipient must be doctor, patient or medclaim."
    End If
    If amount <= 0 Or vatAmount < 0 Or vatAmount > amount Then Err.Raise vbObjectError + 730, , "Invalid invoice amounts."
    If CDec(amount) * 100 <> Fix(CDec(amount) * 100) Or _
        CDec(vatAmount) * 100 <> Fix(CDec(vatAmount) * 100) Then _
        Err.Raise vbObjectError + 730, , "Invoice amounts must be whole cents."
    cfg = GetDocConfig("INV")
    Set il = ThisWorkbook.Worksheets(cfg.logName)
    Set lines = ThisWorkbook.Worksheets(cfg.linesName)
    If il.ProtectContents Or lines.ProtectContents Or ThisWorkbook.Worksheets("AuditLog").ProtectContents Then _
        Err.Raise vbObjectError + 730, , "Invoice or audit logs are protected."
    If Trim$(sourceDocNo) <> "" Then
        If Left$(NrmID(sourceDocNo), 2) <> "Q-" Then _
            Err.Raise vbObjectError + 730, , "Saved INV / MC invoices must be retained, not converted to another charge."
        quoteSource = True
        srcCfg = GetDocConfig("QTE")
        Set src = ThisWorkbook.Worksheets(srcCfg.logName)
        Set sourceLines = ThisWorkbook.Worksheets(srcCfg.linesName)
        sr = FindLogRow(src, sourceDocNo)
        If sr = 0 Then Err.Raise vbObjectError + 730, , "Source document not found."
        sourceDocNo = CStr(src.Cells(sr, srcCfg.colNo).value)
        If src.ProtectContents Then Err.Raise 5, , "Quote log is protected."
        If Trim$(CStr(src.Cells(sr, QL_CONVINV).value)) <> "" Or _
            NrmID(CStr(src.Cells(sr, QL_STATUS).value)) = "CONVERTED" Then Err.Raise 5, , "Quote already converted."
        If NrmID(CStr(src.Cells(sr, QL_STATUS).value)) = "VOIDED" Or _
            NrmID(CStr(src.Cells(sr, QL_STATUS).value)) = "CANCELLED" Then Err.Raise 5, , "Quote is cancelled."
        header = PaymentPlanQuoteHeader(src.Cells(sr, 1).Resize(1, 24).value, "PENDING", dueDate, invoiceDate)
        If CCur(header(1, IL_TOTAL)) <> amount Or CCur(header(1, IL_VAT)) <> vatAmount Or _
            NrmID(CStr(header(1, IL_DEPT))) <> NrmID(dept) Or _
            NrmID(CStr(header(1, IL_CUST))) <> NrmID(custID) Or _
            NrmID(CStr(header(1, IL_PATIENT))) <> NrmID(patientName) Then Err.Raise 5, , "Saved quote differs from conversion parameters."
        For i = 2 To sourceLines.Cells(sourceLines.Rows.Count, LN_DOCNO).End(xlUp).row
            If NrmID(CStr(sourceLines.Cells(i, LN_DOCNO).value)) = NrmID(sourceDocNo) Then
                count = count + 1
                For c = LN_QTY To LN_TOTAL
                    If IsError(sourceLines.Cells(i, c).value) Then Err.Raise 5, , "Invalid saved quote line."
                    If Not IsNumeric(sourceLines.Cells(i, c).value) Then Err.Raise 5, , "Invalid saved quote line."
                    If CDbl(sourceLines.Cells(i, c).value) < 0 Then Err.Raise 5, , "Negative saved quote line."
                Next c
            End If
        Next i
        If count = 0 Then Err.Raise 5, , "Saved quote lines are missing."
        qSnapshot = src.Cells(sr, 1).Resize(1, 24).Formula
    Else
        count = 1
    End If
    oldEvents = Application.EnableEvents: eventsSaved = True
    Application.EnableEvents = False
    docNo = NextUniqueDocNumber(il, dept, cfg.docKind)
    If Not SafeToWriteNumber(il, docNo, 0) Then Err.Raise vbObjectError + 730, , "Duplicate invoice number."
    r = il.Cells(il.Rows.Count, IL_NO).End(xlUp).row + 1
    lr = lines.Cells(lines.Rows.Count, LN_DOCNO).End(xlUp).row + 1
    If r < 2 Then r = 2
    If lr < 2 Then lr = 2
    If r > il.Rows.Count Or lr + count - 1 > lines.Rows.Count Then Err.Raise 5, , "Not enough invoice rows."
    firstLine = lr
    LogAudit "PaymentPlanInvoice", docNo, "", CStr(amount), _
        "Plan=" & planID & "; source=" & sourceDocNo & "; recipient=" & planRecipient
    written = True
    il.Cells(r, IL_NO).value = docNo
    il.Cells(r, IL_DEPT).value = dept
    il.Cells(r, IL_RECIP).value = recipientType
    il.Cells(r, IL_DATE).value = invoiceDate
    il.Cells(r, IL_DUE).value = dueDate
    il.Cells(r, IL_CUST).value = custID
    il.Cells(r, IL_PATIENT).value = patientName
    il.Cells(r, IL_APPLIANCE).value = "Aligners"
    il.Cells(r, IL_SUBTOTAL).value = amount - vatAmount
    il.Cells(r, IL_DISC).value = 0
    il.Cells(r, IL_VAT).value = vatAmount
    il.Cells(r, IL_TOTAL).value = amount
    il.Cells(r, IL_PAID).value = 0
    il.Cells(r, IL_BALANCE).value = amount
    il.Cells(r, IL_STATUS).value = "Unpaid"
    il.Cells(r, IL_CREATED).value = Now
    il.Cells(r, IL_MODIFIED).value = Now
    il.Cells(r, IL_DISCPCT).value = 0
    il.Cells(r, IL_DISCFIX).value = 0
    il.Cells(r, IL_NOTES).value = "Payment plan " & planID & "; source " & sourceDocNo & "; recipient " & planRecipient
    If quoteSource Then
        header = PaymentPlanQuoteHeader(src.Cells(sr, 1).Resize(1, 24).value, docNo, dueDate, invoiceDate)
        il.Cells(r, 1).Resize(1, 27).value = header
        For i = 2 To sourceLines.Cells(sourceLines.Rows.Count, LN_DOCNO).End(xlUp).row
            If NrmID(CStr(sourceLines.Cells(i, LN_DOCNO).value)) = NrmID(sourceDocNo) Then
                lines.Cells(lr, 1).Resize(1, LN_TOTAL).value = sourceLines.Cells(i, 1).Resize(1, LN_TOTAL).value
                lines.Cells(lr, LN_DOCNO).value = docNo
                lr = lr + 1
            End If
        Next i
        LogAudit "PaymentPlanQuote", sourceDocNo, "", docNo, "Full saved header and lines preserved"
        linked = True
        src.Cells(sr, QL_CONVINV).value = docNo
        src.Cells(sr, QL_STATUS).value = "Converted"
        src.Cells(sr, QL_MODIFIED).value = Now
    Else
        lines.Cells(lr, LN_DOCNO).value = docNo
        lines.Cells(lr, LN_LINENO).value = 1
        lines.Cells(lr, LN_DESC).value = "Aligner treatment - full treatment charge"
        lines.Cells(lr, LN_QTY).value = 1
        lines.Cells(lr, LN_EXCL).value = amount - vatAmount
        lines.Cells(lr, LN_VAT).value = vatAmount
        lines.Cells(lr, LN_INCL).value = amount
        lines.Cells(lr, LN_TOTAL).value = amount
    End If
    LogAudit "PaymentPlanPermanentInvoice", docNo, "", CStr(amount), "Full treatment charge; requests are nonfinancial"
    Application.EnableEvents = oldEvents
    GeneratePaymentPlanInvoice = docNo
    Exit Function
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If linked Then src.Cells(sr, 1).Resize(1, 24).Formula = qSnapshot
    If written Then
        ' These are only newly appended, uncommitted rows; never touch any saved source.
        il.Cells(r, 1).Resize(1, 27).ClearContents
        lines.Cells(firstLine, 1).Resize(count, LN_TOTAL).ClearContents
        LogAudit "PaymentPlanInvoiceRollback", docNo, "", "Uncommitted rows removed", errText
    End If
    If eventsSaved Then Application.EnableEvents = oldEvents
    On Error GoTo 0
    Err.Raise errNo, "GeneratePaymentPlanInvoice", errText
End Function

Public Function PaymentPlanQuoteHeader(ByVal savedHeader As Variant, ByVal invoiceNo As String, _
        ByVal dueDate As Date, Optional ByVal invoiceDate As Date = 0) As Variant
    Dim result(1 To 1, 1 To 27) As Variant, sourceCols As Variant, targetCols As Variant, i As Long
    Dim subTotal As Currency, discount As Currency, vat As Currency, total As Currency, col As Variant
    sourceCols = Array(QL_DEPT, QL_RECIP, QL_DATE, QL_CUST, QL_PATIENT, QL_APPLIANCE, _
        QL_SUBTOTAL, QL_DISC, QL_VAT, QL_TOTAL, QL_KIND, QL_MEDAID, QL_MEDNO, _
        QL_MAINMEM, QL_DOCTOR, QL_BHF, QL_DISCPCT, QL_DISCFIX, QL_NOTES)
    targetCols = Array(IL_DEPT, IL_RECIP, IL_DATE, IL_CUST, IL_PATIENT, IL_APPLIANCE, _
        IL_SUBTOTAL, IL_DISC, IL_VAT, IL_TOTAL, 13, IL_MEDAID, IL_MEDNO, _
        IL_MAINMEM, IL_DOCTOR, IL_BHF, IL_DISCPCT, IL_DISCFIX, IL_NOTES)
    For Each col In Array(QL_SUBTOTAL, QL_DISC, QL_VAT, QL_TOTAL, QL_DISCPCT, QL_DISCFIX)
        If IsError(savedHeader(1, CLng(col))) Then Err.Raise 5, , "Invalid saved quote financial header."
        If Trim$(CStr(savedHeader(1, CLng(col)))) <> "" Then
            If Not IsNumeric(savedHeader(1, CLng(col))) Then Err.Raise 5, , "Invalid saved quote financial header."
            If CDec(savedHeader(1, CLng(col))) < 0 Then Err.Raise 5, , "Negative saved quote financial header."
        ElseIf CLng(col) = QL_SUBTOTAL Or CLng(col) = QL_TOTAL Or CLng(col) = QL_VAT Then
            Err.Raise 5, , "Saved quote totals are missing."
        End If
    Next col
    For Each col In Array(QL_SUBTOTAL, QL_DISC, QL_VAT, QL_TOTAL, QL_DISCFIX)
        If Trim$(CStr(savedHeader(1, CLng(col)))) <> "" Then
            If CDec(savedHeader(1, CLng(col))) * 100 <> Fix(CDec(savedHeader(1, CLng(col))) * 100) Then _
                Err.Raise 5, , "Saved quote amounts must be whole cents."
        End If
    Next col
    subTotal = CCur(savedHeader(1, QL_SUBTOTAL))
    If Trim$(CStr(savedHeader(1, QL_DISC))) <> "" Then discount = CCur(savedHeader(1, QL_DISC))
    vat = CCur(savedHeader(1, QL_VAT)): total = CCur(savedHeader(1, QL_TOTAL))
    If total <= 0 Or vat > total Then _
        Err.Raise 5, , "Saved quote totals do not reconcile."
    ' Saved templates may log a net or pre-discount subtotal. Never apply a saved discount twice.
    If Abs(subTotal + vat - total) > 0.01 And Abs(subTotal - discount + vat - total) > 0.01 Then _
        Err.Raise 5, , "Saved quote subtotal / discount / tax do not reconcile."
    If Not IsDate(savedHeader(1, QL_DATE)) Then Err.Raise 5, , "Saved quote date is invalid."
    For i = 0 To UBound(sourceCols)
        result(1, CLng(targetCols(i))) = savedHeader(1, CLng(sourceCols(i)))
    Next i
    If LCase$(CStr(result(1, IL_RECIP))) = "medclaim" Then result(1, IL_RECIP) = "patient"
    If LCase$(CStr(result(1, IL_RECIP))) = "private" Then result(1, IL_RECIP) = "patient"
    result(1, IL_NO) = invoiceNo: result(1, IL_DUE) = dueDate
    If invoiceDate = 0 Then invoiceDate = Date
    result(1, IL_DATE) = invoiceDate
    result(1, IL_PAID) = 0: result(1, IL_BALANCE) = total: result(1, IL_STATUS) = "Unpaid"
    result(1, IL_SRCQUOTE) = savedHeader(1, QL_NO)
    result(1, IL_CREATED) = Now: result(1, IL_MODIFIED) = Now
    PaymentPlanQuoteHeader = result
End Function

' ---- copy a cell only if it exists / has a value ----
Private Sub CopyIfExists(src As Worksheet, dst As Worksheet, addr As String)
    On Error Resume Next
    dst.Range(addr).value = src.Range(addr).value
    On Error GoTo 0
End Sub
