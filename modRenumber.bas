Attribute VB_Name = "modRenumber"
' modRenumber ? dept-change renumber + manual rename + revert-converted-quote.
'   RenumberOnDeptChange : dept corrected -> next number in the NEW dept seq.
'   RenameDocNumber      : manual override (G7<>K4). Renames the SAME row+lines
'                          to a user-supplied (e.g. legacy) number. Unpaid only.
'                          Counter untouched; number accepted verbatim.
'                          Cascades old->new across Payments and CreditNotes.
'   RevertQuoteToSaved   : un-convert a quote; DELETES the (unpaid) linked inv.
' ============================================================================
Option Explicit

' Plan invoice/source identifiers are immutable; changes go through Amend Plan.
Private Function PaymentPlanLinked(ByVal docNo As String) As Boolean
    Dim ws As Worksheet, i As Long
    If NrmID(docNo) = "" Then Exit Function
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PaymentPlans")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For i = 2 To ws.Cells(ws.Rows.Count, PP_PLANID).End(xlUp).row
        If NrmID(CStr(ws.Cells(i, PP_SOURCE).value)) = NrmID(docNo) Or _
           NrmID(CStr(ws.Cells(i, PP_INVOICE).value)) = NrmID(docNo) Then
            PaymentPlanLinked = True
            Exit Function
        End If
    Next i
End Function

' ---- Is this INV/MC row unpaid? (QTE has no payment -> always True) ----
Private Function IsUnpaidRow(cfg As DocConfig, wsLog As Worksheet, lr As Long) As Boolean
    If cfg.colPaid = 0 Then IsUnpaidRow = True: Exit Function
    IsUnpaidRow = (Num(wsLog.Cells(lr, cfg.colPaid).value) <= 0.005)
End Function

' ============================================================ DEPT CHANGE =====
Public Sub RenumberOnDeptChange(ByVal docKind As String, ByVal oldNo As String, _
                                ByVal oldDept As String, ByVal newDept As String)
    Dim cfg As DocConfig, wsLog As Worksheet, wsLines As Worksheet
    Dim lr As Long, newNo As String, i As Long, lastLine As Long

    cfg = GetDocConfig(docKind)
    Set wsLog = ThisWorkbook.Sheets(cfg.logName)
    Set wsLines = ThisWorkbook.Sheets(cfg.linesName)

    lr = FindLogRow(wsLog, oldNo)
    If lr = 0 Then MsgBox oldNo & " not found for renumber.", vbExclamation: Exit Sub
    If PaymentPlanLinked(oldNo) Then
        MsgBox "This document belongs to a payment plan. Use Cancel/Amend Plan instead.", vbExclamation
        Exit Sub
    End If

    ' UNPAID-ONLY guard
    If Not IsUnpaidRow(cfg, wsLog, lr) Then
        MsgBox oldNo & " has payments recorded. Department change is only allowed " & _
               "on unpaid documents.", vbExclamation, "Renumber blocked": Exit Sub
    End If

    ' QUOTE conversion conflict block
    If docKind = "QTE" Then
        If Trim(CStr(wsLog.Cells(lr, QL_CONVINV).value)) <> "" Then
            MsgBox "This quote is linked to invoice " & _
                   CStr(wsLog.Cells(lr, QL_CONVINV).value) & "." & vbCrLf & _
                   "Void or revert that invoice before changing the department.", _
                   vbExclamation, "Renumber blocked"
            Exit Sub
        End If
    End If

    newNo = NextUniqueDocNumber(wsLog, newDept, docKind)

    Dim msg As String
    msg = "Department changed from " & oldDept & " to " & newDept & "." & vbCrLf & vbCrLf & _
          "This will reassign the number:" & vbCrLf & _
          "    " & oldNo & "  ->  " & newNo & vbCrLf & vbCrLf & _
          "The record and its line items will be updated in place. Continue?"
    If MsgBox(msg, vbQuestion + vbYesNo, "Confirm renumber") <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(cfg.sheetName)
    ws.Range("G7").value = newNo
    ws.Range("K4").value = newNo
    wsLog.Cells(lr, cfg.colNo).value = newNo
    wsLog.Cells(lr, cfg.colDept).value = newDept

    ' move line rows
    lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastLine
        If NrmID(CStr(wsLines.Cells(i, LN_DOCNO).value)) = NrmID(oldNo) Then
            wsLines.Cells(i, LN_DOCNO).value = newNo
        End If
    Next i

    ' cascade references (payments + credit notes)
    If cfg.colPaid > 0 Then RepointPayments oldNo, newNo
    RepointCreditNoteSource oldNo, newNo

    LogAudit "Renumber", oldNo, oldNo, newNo, docKind & " dept " & oldDept & "->" & newDept

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Renumbered to " & newNo & "." & vbCrLf & _
           "Re-saving the document now to sync all fields.", vbInformation

    ResaveByKind docKind
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "Renumber error: " & Err.Description, vbExclamation
End Sub

' ============================================================ MANUAL RENAME ===
' Rename oldNo -> newNo (user typed newNo into G7). Same row + lines. Unpaid
' only. Counter is NOT consumed. newNo accepted verbatim (legacy numbers OK).
Public Sub RenameDocNumber(ByVal docKind As String, ByVal oldNo As String, ByVal newNo As String)
    Dim cfg As DocConfig, wsLog As Worksheet, wsLines As Worksheet, ws As Worksheet
    Dim lr As Long, i As Long, lastLine As Long

    oldNo = Trim(oldNo): newNo = Trim(newNo)
    If newNo = "" Then MsgBox "New number is blank.", vbExclamation: Exit Sub
    If NrmID(oldNo) = NrmID(newNo) Then Exit Sub
    If PaymentPlanLinked(oldNo) Then
        cfg = GetDocConfig(docKind)
        ThisWorkbook.Sheets(cfg.sheetName).Range("G7").value = oldNo
        MsgBox "This document belongs to a payment plan. Use Cancel/Amend Plan instead.", vbExclamation
        Exit Sub
    End If

    cfg = GetDocConfig(docKind)
    Set wsLog = ThisWorkbook.Sheets(cfg.logName)
    Set wsLines = ThisWorkbook.Sheets(cfg.linesName)
    Set ws = ThisWorkbook.Sheets(cfg.sheetName)

    lr = FindLogRow(wsLog, oldNo)
    If lr = 0 Then
        ' No existing row under the anchor number: nothing to rename. Restore G7
        ' to the typed value and let a normal save proceed as a NEW doc.
        MsgBox "Original number " & oldNo & " not found in " & cfg.logName & _
               "; nothing to rename. Save it as a new document instead.", vbExclamation
        Exit Sub
    End If

    ' UNPAID-ONLY guard
    If Not IsUnpaidRow(cfg, wsLog, lr) Then
        MsgBox oldNo & " has payments recorded. The number can only be changed " & _
               "on unpaid documents.", vbExclamation, "Rename blocked"
        ws.Range("G7").value = oldNo      ' revert the edit
        Exit Sub
    End If

    ' COLLISION guard: newNo must not already belong to a different row
    Dim clash As Long
    clash = FindLogRow(wsLog, newNo)
    If clash <> 0 And clash <> lr Then
        MsgBox "Number " & newNo & " already belongs to another document (row " & _
               clash & "). Rename cancelled.", vbCritical, "Duplicate number blocked"
        ws.Range("G7").value = oldNo
        Exit Sub
    End If

    If MsgBox("Change document number:" & vbCrLf & _
              "    " & oldNo & "  ->  " & newNo & vbCrLf & vbCrLf & _
              "The record and its line items will be updated in place. Continue?", _
              vbQuestion + vbYesNo, "Confirm number change") <> vbYes Then
        ws.Range("G7").value = oldNo
        Exit Sub
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' rename header + sheet anchors
    wsLog.Cells(lr, cfg.colNo).value = newNo
    ws.Range("G7").value = newNo
    ws.Range("K4").value = newNo

    ' rename line rows
    lastLine = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastLine
        If NrmID(CStr(wsLines.Cells(i, LN_DOCNO).value)) = NrmID(oldNo) Then
            wsLines.Cells(i, LN_DOCNO).value = newNo
        End If
    Next i

    ' cascade references (unpaid => Payments unlikely, but safe; CN source links)
    If cfg.colPaid > 0 Then RepointPayments oldNo, newNo
    RepointCreditNoteSource oldNo, newNo

    LogAudit "Rename", oldNo, oldNo, newNo, docKind & " number changed (manual override)"

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Number changed to " & newNo & "." & vbCrLf & _
           "Re-saving the document now to sync all fields.", vbInformation

    ResaveByKind docKind
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "RenameDocNumber error: " & Err.Description, vbExclamation
End Sub

' ---- re-save the on-screen document under its (new) number ----
Private Sub ResaveByKind(ByVal docKind As String)
    Select Case docKind
        Case "INV": SaveInvoice
        Case "QTE": SaveQuote
        Case "MC":  SaveMedClaim
    End Select
End Sub

' ---- Repoint all Payments rows from oldNo to newNo ----
Private Sub RepointPayments(ByVal oldNo As String, ByVal newNo As String)
    Dim wsP As Worksheet, last As Long, i As Long
    On Error Resume Next
    Set wsP = ThisWorkbook.Sheets("Payments")
    On Error GoTo 0
    If wsP Is Nothing Then Exit Sub
    last = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsP.Cells(i, PY_INV).value)) = NrmID(oldNo) Then
            wsP.Cells(i, PY_INV).value = newNo
        End If
    Next i
End Sub

' ---- Repoint CreditNotes.SourceInvNo (col B = 2) from oldNo to newNo ----
Private Sub RepointCreditNoteSource(ByVal oldNo As String, ByVal newNo As String)
    Dim wsCN As Worksheet, last As Long, i As Long
    On Error Resume Next
    Set wsCN = ThisWorkbook.Sheets("CreditNotes")
    On Error GoTo 0
    If wsCN Is Nothing Then Exit Sub
    last = wsCN.Cells(wsCN.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        If NrmID(CStr(wsCN.Cells(i, 2).value)) = NrmID(oldNo) Then
            wsCN.Cells(i, 2).value = newNo
        End If
    Next i
End Sub

' ============================================================ REVERT QUOTE ====
Public Sub RevertQuoteToSaved(Optional ByVal quoteNoIn As String = "")
    Dim wsQ As Worksheet, wsI As Worksheet, wsIL As Worksheet
    Dim quoteNo As String, qr As Long, invNo As String, ir As Long
    Dim paid As Double

    Set wsQ = ThisWorkbook.Sheets("QuoteLog")
    Set wsIL = ThisWorkbook.Sheets("InvoiceLog")
    Set wsI = ThisWorkbook.Sheets("InvoiceLines")

    If quoteNoIn <> "" Then
        quoteNo = Trim(quoteNoIn)
    Else
        quoteNo = Trim(InputBox("Enter Quote number to revert (e.g. Q-WA-0012):", "Revert Quote"))
    End If
    If quoteNo = "" Then Exit Sub

    qr = FindLogRow(wsQ, quoteNo)
    If qr = 0 Then MsgBox "Quote " & quoteNo & " not found.", vbExclamation: Exit Sub

    invNo = Trim(CStr(wsQ.Cells(qr, QL_CONVINV).value))
    If PaymentPlanLinked(quoteNo) Or PaymentPlanLinked(invNo) Then
        MsgBox "This quote/invoice belongs to a payment plan. Use Cancel/Amend Plan; do not delete its financial records.", vbExclamation
        Exit Sub
    End If
    If invNo = "" Then
        MsgBox "Quote " & quoteNo & " is not converted ? nothing to revert.", vbInformation
        Exit Sub
    End If

    ir = FindLogRow(wsIL, invNo)
    If ir > 0 Then
        paid = Num(wsIL.Cells(ir, IL_PAID).value)
        If paid > 0 Then
            MsgBox "Invoice " & invNo & " has payments recorded (R " & _
                   Format(paid, "#,##0.00") & ")." & vbCrLf & _
                   "Revert cancelled ? issue a credit note or handle the payment first.", _
                   vbExclamation, "Revert blocked"
            Exit Sub
        End If
    End If

    If MsgBox("This will DELETE invoice " & invNo & " (log + line items) and revert " & _
              "quote " & quoteNo & " to 'Saved' so it can be edited / re-converted." & vbCrLf & vbCrLf & _
              "This cannot be undone. Continue?", vbCritical + vbYesNo, "Revert Quote") <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If ir > 0 Then
        Dim last As Long, i As Long
        last = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).row
        For i = last To 2 Step -1
            If NrmID(CStr(wsI.Cells(i, LN_DOCNO).value)) = NrmID(invNo) Then wsI.Rows(i).Delete
        Next i
        wsIL.Rows(ir).Delete
    End If

    wsQ.Cells(qr, QL_STATUS).value = "Saved"
    wsQ.Cells(qr, QL_CONVINV).value = ""

    LogAudit "Revert", invNo, "", "DELETED", "Invoice deleted (quote " & quoteNo & " reverted)"
    LogAudit "Revert", quoteNo, "Converted", "Saved", "Quote reverted; invoice " & invNo & " deleted"

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    On Error Resume Next
    RefreshMenuSummary
    On Error GoTo 0

    MsgBox "Quote " & quoteNo & " reverted to Saved; invoice " & invNo & " deleted.", vbInformation
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "RevertQuoteToSaved error: " & Err.Description, vbExclamation
End Sub

