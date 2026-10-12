Attribute VB_Name = "modConvert"
Option Explicit
' ============================================================================
' ConvertQuoteToInvoice — button on the QUOTE sheet.
'   1B FIX: copy line Price INCL (G), not Excl (E) — E/F/H derive.
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

    ' line items — A=Qty, D=Description, G=Price Incl (authoritative; E/F/H derive)
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

' Noninteractive writer shared by payment-plan callers. Never alters the source.
Public Function GeneratePaymentPlanInvoice(ByVal planID As String, ByVal sourceDocNo As String, _
        ByVal recipientType As String, ByVal custID As String, ByVal patientName As String, _
        ByVal dept As String, ByVal invoiceDate As Date, ByVal dueDate As Date, _
        ByVal installmentNo As Long, ByVal amount As Currency, ByVal vatAmount As Currency) As String
    Dim cfg As DocConfig, srcCfg As DocConfig, il As Worksheet, lines As Worksheet, src As Worksheet
    Dim docNo As String, kind As String, r As Long, lr As Long, sr As Long, c As Long
    Dim oldEvents As Boolean, eventsSaved As Boolean, written As Boolean
    Dim errNo As Long, errText As String, snapshot As String, planRecipient As String
    On Error GoTo Failed
    dept = UCase$(Trim$(dept)): recipientType = LCase$(Trim$(recipientType))
    planRecipient = recipientType
    If Trim$(planID) = "" Or installmentNo < 0 Then Err.Raise vbObjectError + 730, , "Plan ID and installment number are required."
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
    If amount < 0 Or vatAmount < 0 Or vatAmount > amount Then Err.Raise vbObjectError + 730, , "Invalid invoice amounts."
    If CDec(amount) * 100 <> Fix(CDec(amount) * 100) Or _
        CDec(vatAmount) * 100 <> Fix(CDec(vatAmount) * 100) Then _
        Err.Raise vbObjectError + 730, , "Invoice amounts must be whole cents."
    cfg = GetDocConfig("INV")
    Set il = ThisWorkbook.Worksheets(cfg.logName)
    Set lines = ThisWorkbook.Worksheets(cfg.linesName)
    If il.ProtectContents Or lines.ProtectContents Or ThisWorkbook.Worksheets("AuditLog").ProtectContents Then _
        Err.Raise vbObjectError + 730, , "Invoice or audit logs are protected."
    If Trim$(sourceDocNo) <> "" Then
        Select Case True
            Case Left$(NrmID(sourceDocNo), 7) = "MC-INV-": kind = "MC"
            Case Left$(NrmID(sourceDocNo), 4) = "INV-": kind = "INV"
            Case Left$(NrmID(sourceDocNo), 2) = "Q-": kind = "QTE"
            Case Else: Err.Raise vbObjectError + 730, , "Unsupported source document."
        End Select
        srcCfg = GetDocConfig(kind)
        Set src = ThisWorkbook.Worksheets(srcCfg.logName)
        sr = FindLogRow(src, sourceDocNo)
        If sr = 0 Then Err.Raise vbObjectError + 730, , "Source document not found."
        sourceDocNo = CStr(src.Cells(sr, srcCfg.colNo).value)
    End If
    oldEvents = Application.EnableEvents: eventsSaved = True
    Application.EnableEvents = False
    docNo = NextUniqueDocNumber(il, dept, cfg.docKind)
    If Not SafeToWriteNumber(il, docNo, 0) Then Err.Raise vbObjectError + 730, , "Duplicate invoice number."
    r = il.Cells(il.Rows.Count, IL_NO).End(xlUp).row + 1
    lr = lines.Cells(lines.Rows.Count, LN_DOCNO).End(xlUp).row + 1
    If r < 2 Then r = 2
    If lr < 2 Then lr = 2
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
    If kind = "QTE" Then il.Cells(r, IL_SRCQUOTE).value = sourceDocNo
    il.Cells(r, IL_CREATED).value = Now
    il.Cells(r, IL_MODIFIED).value = Now
    il.Cells(r, IL_DISCPCT).value = 0
    il.Cells(r, IL_DISCFIX).value = 0
    il.Cells(r, IL_NOTES).value = "Payment plan " & planID & "; source " & sourceDocNo & "; recipient " & planRecipient
    If Not src Is Nothing Then CopyPlanSourceMetadata src, sr, kind, il, r
    lines.Cells(lr, LN_DOCNO).value = docNo
    lines.Cells(lr, LN_LINENO).value = 1
    lines.Cells(lr, LN_DESC).value = "Aligner plan " & planID & _
        IIf(installmentNo = 0, " deposit", " installment " & CStr(installmentNo))
    lines.Cells(lr, LN_QTY).value = 1
    lines.Cells(lr, LN_EXCL).value = amount - vatAmount
    lines.Cells(lr, LN_VAT).value = vatAmount
    lines.Cells(lr, LN_INCL).value = amount
    lines.Cells(lr, LN_TOTAL).value = amount
    Application.EnableEvents = oldEvents
    GeneratePaymentPlanInvoice = docNo
    Exit Function
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If written Then
        For c = 1 To 27
            snapshot = snapshot & CStr(c) & "=" & CStr(il.Cells(r, c).value) & "; "
        Next c
        LogAudit "PaymentPlanInvoiceRollback", docNo, snapshot, "Voided", "Plan=" & planID & "; " & errText
        il.Cells(r, IL_NO).value = docNo
        il.Cells(r, IL_SUBTOTAL).value = 0
        il.Cells(r, IL_DISC).value = 0
        il.Cells(r, IL_VAT).value = 0
        il.Cells(r, IL_TOTAL).value = 0
        il.Cells(r, IL_PAID).value = 0
        il.Cells(r, IL_BALANCE).value = 0
        il.Cells(r, IL_STATUS).value = "Voided"
        il.Cells(r, IL_MODIFIED).value = Now
        il.Cells(r, IL_NOTES).value = "Failed plan " & planID & "; source " & sourceDocNo & "; " & errText
    End If
    If eventsSaved Then Application.EnableEvents = oldEvents
    On Error GoTo 0
    Err.Raise errNo, "GeneratePaymentPlanInvoice", errText
End Function

Private Sub CopyPlanSourceMetadata(ByVal src As Worksheet, ByVal sr As Long, ByVal kind As String, _
        ByVal dst As Worksheet, ByVal dr As Long)
    Dim applianceCol As Long, notesCol As Long, sourceNotes As String
    Select Case kind
        Case "QTE"
            applianceCol = QL_APPLIANCE: notesCol = QL_NOTES
            dst.Cells(dr, IL_MEDAID).value = src.Cells(sr, QL_MEDAID).value
            dst.Cells(dr, IL_MEDNO).value = src.Cells(sr, QL_MEDNO).value
            dst.Cells(dr, IL_MAINMEM).value = src.Cells(sr, QL_MAINMEM).value
            dst.Cells(dr, IL_DOCTOR).value = src.Cells(sr, QL_DOCTOR).value
            dst.Cells(dr, IL_BHF).value = src.Cells(sr, QL_BHF).value
        Case "MC"
            applianceCol = ML_APPLIANCE: notesCol = ML_NOTES
            dst.Cells(dr, IL_MEDAID).value = src.Cells(sr, ML_MEDAID).value
            dst.Cells(dr, IL_MEDNO).value = src.Cells(sr, ML_MEDNO).value
            dst.Cells(dr, IL_MAINMEM).value = src.Cells(sr, ML_MAINMEM).value
            dst.Cells(dr, IL_DOCTOR).value = src.Cells(sr, ML_DOCTOR).value
            dst.Cells(dr, IL_BHF).value = src.Cells(sr, ML_BHF).value
        Case "INV"
            applianceCol = IL_APPLIANCE: notesCol = IL_NOTES
            dst.Cells(dr, IL_MEDAID).Resize(1, 5).value = src.Cells(sr, IL_MEDAID).Resize(1, 5).value
    End Select
    If applianceCol > 0 Then
        If Trim$(CStr(src.Cells(sr, applianceCol).value)) <> "" Then _
            dst.Cells(dr, IL_APPLIANCE).value = src.Cells(sr, applianceCol).value
    End If
    If notesCol > 0 Then
        sourceNotes = CStr(src.Cells(sr, notesCol).value)
        If Trim$(sourceNotes) <> "" Then dst.Cells(dr, IL_NOTES).value = _
            CStr(dst.Cells(dr, IL_NOTES).value) & vbLf & "Source notes: " & sourceNotes
    End If
End Sub

' ---- copy a cell only if it exists / has a value ----
Private Sub CopyIfExists(src As Worksheet, dst As Worksheet, addr As String)
    On Error Resume Next
    dst.Range(addr).value = src.Range(addr).value
    On Error GoTo 0
End Sub



