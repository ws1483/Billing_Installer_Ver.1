Attribute VB_Name = "modPayment"
Option Explicit
' ============================================================================
' modPayment ? launches Record Payment + (Phase 3) Manage Payments.
'   ReconcileDoc recomputes Paid/Balance/Status from remaining Payments rows.
' ============================================================================

Sub RecordPayment()
    frmPayment.Show
End Sub

Sub ManagePayments()
    frmPaymentManage.Show     ' Part B (built next)
End Sub

' docType: "invoice" | "medclaim"
Public Sub ReconcileDoc(ByVal docType As String, ByVal docNo As String)
    Dim wsLog As Worksheet, wsP As Worksheet
    Dim lr As Long, last As Long, i As Long
    Dim paidCol As Long, balCol As Long, statusCol As Long, totalCol As Long, modCol As Long
    Dim paid As Double, total As Double, bal As Double

    If docType = "medclaim" Then
        Set wsLog = ThisWorkbook.Sheets("MedAidLog")
        paidCol = ML_PAID: balCol = ML_BALANCE: statusCol = ML_STATUS
        totalCol = ML_TOTAL: modCol = ML_MODIFIED
    Else
        Set wsLog = ThisWorkbook.Sheets("InvoiceLog")
        paidCol = IL_PAID: balCol = IL_BALANCE: statusCol = IL_STATUS
        totalCol = IL_TOTAL: modCol = IL_MODIFIED
    End If

    lr = FindLogRow(wsLog, docNo)
    If lr = 0 Then Exit Sub
    If NrmID(CStr(wsLog.Cells(lr, statusCol).value)) = "VOIDED" Then
        ZeroVoidedAmounts wsLog, lr, docType, docNo, modCol
        Exit Sub
    End If

    Set wsP = ThisWorkbook.Sheets("Payments")
    last = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).row
    paid = 0
    For i = 2 To last
        If NrmID(CStr(wsP.Cells(i, PY_INV).value)) = NrmID(docNo) Then
            paid = paid + Num(wsP.Cells(i, PY_AMT).value)
        End If
    Next i

    total = Num(wsLog.Cells(lr, totalCol).value)
    bal = total - paid: If bal < 0 Then bal = 0
    LogAudit "Reconcile", docNo, _
        "Paid=" & CStr(wsLog.Cells(lr, paidCol).value) & "; Balance=" & _
        CStr(wsLog.Cells(lr, balCol).value) & "; Status=" & CStr(wsLog.Cells(lr, statusCol).value), _
        "Paid=" & CStr(paid) & "; Balance=" & CStr(bal), "Reconciled from retained Payments rows"
    wsLog.Cells(lr, paidCol).value = paid
    wsLog.Cells(lr, balCol).value = bal
    If paid <= 0 Then
        wsLog.Cells(lr, statusCol).value = "Unpaid"
    ElseIf bal <= 0.005 Then
        wsLog.Cells(lr, statusCol).value = "Paid"
    Else
        wsLog.Cells(lr, statusCol).value = "Part-Paid"
    End If
    wsLog.Cells(lr, modCol).value = Now
    RefreshExistingPaymentPlans
End Sub

Private Sub RefreshExistingPaymentPlans()
    Dim wsPlans As Worksheet
    On Error Resume Next
    Set wsPlans = ThisWorkbook.Worksheets("PaymentPlans")
    On Error GoTo Failed
    If wsPlans Is Nothing Then Exit Sub
    modPaymentOptions.RefreshPaymentPlans
    Exit Sub
Failed:
    MsgBox "Payment was reconciled, but payment plans could not be refreshed: " & _
        Err.Description, vbExclamation, "Payment options"
End Sub

' Legacy statements sum amounts without filtering status. Preserve the full
' header snapshot before correcting legacy voided rows; never remove lines/payments.
Private Sub ZeroVoidedAmounts(ByVal ws As Worksheet, ByVal r As Long, ByVal docType As String, _
        ByVal docNo As String, ByVal modifiedCol As Long)
    Dim columns As Variant, col As Variant, c As Long, changed As Boolean, snapshot As String
    If docType = "medclaim" Then
        columns = Array(ML_SUBTOTAL, ML_DISC, ML_VAT, ML_TOTAL, ML_PAID, ML_BALANCE, ML_DISCPCT, ML_DISCFIX)
    Else
        columns = Array(IL_SUBTOTAL, IL_DISC, IL_VAT, IL_TOTAL, IL_PAID, IL_BALANCE, IL_DISCPCT, IL_DISCFIX)
    End If
    For Each col In columns
        If IsError(ws.Cells(r, CLng(col)).value) Or Num(ws.Cells(r, CLng(col)).value) <> 0 Then changed = True
    Next col
    If Not changed Then Exit Sub
    For c = 1 To 27
        snapshot = snapshot & CStr(c) & "=" & CStr(ws.Cells(r, c).value) & "; "
    Next c
    LogAudit "ReconcileVoid", docNo, snapshot, "Voided; monetary amounts zero", _
        "Original header, lines and Payments history retained"
    For Each col In columns
        ws.Cells(r, CLng(col)).value = 0
    Next col
    ws.Cells(r, modifiedCol).value = Now
End Sub
