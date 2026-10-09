VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPaymentOptions
   Caption         =   "Payment Options"
   ClientHeight    =   5796
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6780
   OleObjectBlob   =   "frmPaymentOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPaymentOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents poSource As MSForms.ComboBox
Private WithEvents poAligners As MSForms.TextBox
Private WithEvents poTotal As MSForms.TextBox
Private WithEvents poPlan As MSForms.TextBox
Private WithEvents poPreview As MSForms.CommandButton
Private WithEvents poGenerate As MSForms.CommandButton
Private WithEvents poCancelPlan As MSForms.CommandButton
Private WithEvents poAmend As MSForms.CommandButton
Private WithEvents poDeposits As MSForms.CommandButton
Private WithEvents poInstallments As MSForms.CommandButton
Private WithEvents poSingle As MSForms.CommandButton
Private WithEvents poMail As MSForms.CommandButton
Private WithEvents poClose As MSForms.CommandButton
Private poRecipient As MSForms.ComboBox
Private poDept As MSForms.ComboBox
Private poCustomer As MSForms.TextBox
Private poPatient As MSForms.TextBox
Private poStart As MSForms.TextBox
Private poSummary As MSForms.Label
Private poSchedule As MSForms.ListBox
Private mReminder As Boolean
Private mLoading As Boolean
Private mDueRows As Collection

Private Sub UserForm_Initialize()
    Dim ctl As Object, ws As Worksheet, sheetName As Variant, i As Long
    ' Reuse the existing statement form's exported container; build the new controls here.
    For Each ctl In Me.Controls
        ctl.Visible = False
    Next ctl
    Me.Width = 650: Me.Height = 550
    Set poSource = AddField("ComboBox", "Source quote/invoice (blank = manual)", 12, 12, 300)
    Set poAligners = AddField("TextBox", "Number of aligners", 330, 12, 140)
    Set poTotal = AddField("TextBox", "Total (VAT inclusive)", 480, 12, 140)
    Set poRecipient = AddField("ComboBox", "Recipient type", 12, 64, 140)
    poRecipient.AddItem "doctor": poRecipient.AddItem "patient": poRecipient.AddItem "medclaim"
    poRecipient.ListIndex = 0
    Set poDept = AddField("ComboBox", "Department", 162, 64, 90)
    poDept.AddItem "WA": poDept.AddItem "WD": poDept.ListIndex = 0
    Set poCustomer = AddField("TextBox", "Customer ID / patient bill-to", 262, 64, 180)
    Set poPatient = AddField("TextBox", "Patient name", 452, 64, 168)
    Set poStart = AddField("TextBox", "Deposit due (yyyy-mm-dd)", 12, 116, 200)
    poStart.value = Format(Date, "yyyy-mm-dd")
    Set poPlan = AddField("TextBox", "Plan ID (for cancel/amend/statement)", 222, 116, 300)
    Set poSummary = Me.Controls.Add("Forms.Label.1", "poSummaryLabel", True)
    poSummary.Left = 12: poSummary.Top = 174: poSummary.Width = 608: poSummary.Height = 36
    Set poSchedule = Me.Controls.Add("Forms.ListBox.1", "poScheduleList", True)
    poSchedule.Left = 12: poSchedule.Top = 214: poSchedule.Width = 608: poSchedule.Height = 190
    poSchedule.ColumnCount = 4: poSchedule.ColumnWidths = "90 pt;110 pt;130 pt;250 pt"
    Set poPreview = AddButton("poPreviewButton", "Preview", 12, 414, 90)
    Set poGenerate = AddButton("poGenerateButton", "Generate Invoices", 112, 414, 140)
    Set poCancelPlan = AddButton("poCancelPlanButton", "Cancel Plan", 262, 414, 110)
    Set poAmend = AddButton("poAmendButton", "Amend Plan", 382, 414, 110)
    Set poMail = AddButton("poMailButton", "Mail now", 502, 414, 118)
    poMail.Visible = False
    Set poDeposits = AddButton("poDepositsButton", "Deposits Statement", 12, 458, 150)
    Set poInstallments = AddButton("poInstallmentsButton", "Installments Statement", 172, 458, 170)
    Set poSingle = AddButton("poSingleButton", "Plan Statement", 352, 458, 150)
    Set poClose = AddButton("poCloseButton", "Close", 512, 458, 108)
    mLoading = True
    For Each sheetName In Array("QuoteLog", "InvoiceLog", "MedAidLog")
        Set ws = ThisWorkbook.Sheets(CStr(sheetName))
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
            If Trim(CStr(ws.Cells(i, 1).value)) <> "" Then poSource.AddItem ws.Cells(i, 1).value
        Next i
    Next sheetName
    mLoading = False
    poSummary.Caption = "Choose a saved source, or enter total, recipient and department for a manual plan."
End Sub

Private Function AddField(kind As String, caption As String, x As Single, y As Single, w As Single) As Object
    Dim lbl As Object, ctl As Object, n As String
    n = "poField" & CStr(Me.Controls.Count)
    Set lbl = Me.Controls.Add("Forms.Label.1", n & "Label", True)
    lbl.Caption = caption: lbl.Left = x: lbl.Top = y: lbl.Width = w: lbl.Height = 18
    Set ctl = Me.Controls.Add("Forms." & kind & ".1", n, True)
    ctl.Left = x: ctl.Top = y + 20: ctl.Width = w: ctl.Height = 22
    Set AddField = ctl
End Function

Private Function AddButton(n As String, caption As String, x As Single, y As Single, w As Single) As MSForms.CommandButton
    Dim ctl As MSForms.CommandButton
    Set ctl = Me.Controls.Add("Forms.CommandButton.1", n, True)
    ctl.Caption = caption: ctl.Left = x: ctl.Top = y: ctl.Width = w: ctl.Height = 28
    Set AddButton = ctl
End Function

Private Sub poSource_Change()
    Dim ws As Worksheet, lr As Long, source As String
    If mLoading Or mReminder Then Exit Sub
    source = modHelpers.NrmID(CStr(poSource.value))
    If source = "" Then
        SetSourceLocks False
        Exit Sub
    End If
    On Error GoTo Fail
    mLoading = True
    If Left(source, 2) = "Q-" Then
        Set ws = ThisWorkbook.Sheets("QuoteLog")
        lr = modHelpers.FindLogRow(ws, source)
        If lr = 0 Then GoTo Done
        poTotal.value = ws.Cells(lr, QL_TOTAL).value
        poRecipient.value = ws.Cells(lr, QL_RECIP).value
        poDept.value = ws.Cells(lr, QL_DEPT).value
        poCustomer.value = ws.Cells(lr, QL_CUST).value
        poPatient.value = ws.Cells(lr, QL_PATIENT).value
    ElseIf Left(source, 7) = "MC-INV-" Then
        Set ws = ThisWorkbook.Sheets("MedAidLog")
        lr = modHelpers.FindLogRow(ws, source)
        If lr = 0 Then GoTo Done
        poTotal.value = ws.Cells(lr, ML_TOTAL).value
        poRecipient.value = "medclaim"
        poDept.value = ws.Cells(lr, ML_DEPT).value
        poCustomer.value = ws.Cells(lr, ML_CUST).value
        poPatient.value = ws.Cells(lr, ML_PATIENT).value
    Else
        Set ws = ThisWorkbook.Sheets("InvoiceLog")
        lr = modHelpers.FindLogRow(ws, source)
        If lr = 0 Then GoTo Done
        poTotal.value = ws.Cells(lr, IL_TOTAL).value
        poRecipient.value = ws.Cells(lr, IL_RECIP).value
        poDept.value = ws.Cells(lr, IL_DEPT).value
        poCustomer.value = ws.Cells(lr, IL_CUST).value
        poPatient.value = ws.Cells(lr, IL_PATIENT).value
    End If
Done:
    SetSourceLocks (lr > 0)
    mLoading = False
    UpdatePreview
    Exit Sub
Fail:
    mLoading = False
    MsgBox "Source lookup error: " & Err.Description, vbExclamation
End Sub

Private Sub SetSourceLocks(locked As Boolean)
    poTotal.Locked = locked
    poRecipient.Enabled = Not locked
    poDept.Enabled = Not locked
    poCustomer.Locked = locked
    poPatient.Locked = locked
End Sub

Private Sub poAligners_Change()
    If Not mLoading And Not mReminder Then UpdatePreview
End Sub

Private Sub poPlan_Change()
    Dim ws As Worksheet, i As Long
    If mLoading Or mReminder Then Exit Sub
    poSource.Enabled = True
    poGenerate.Enabled = True
    If Trim(CStr(poPlan.value)) = "" Then
        SetSourceLocks (Trim(CStr(poSource.value)) <> "")
        Exit Sub
    End If
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PaymentPlans")
    On Error GoTo Fail
    If ws Is Nothing Then Exit Sub
    For i = 2 To ws.Cells(ws.Rows.Count, PP_PLANID).End(xlUp).row
        If modHelpers.NrmID(CStr(ws.Cells(i, PP_PLANID).value)) = modHelpers.NrmID(CStr(poPlan.value)) Then
            mLoading = True
            poSource.value = ws.Cells(i, PP_SOURCE).value
            poAligners.value = ws.Cells(i, PP_ALIGNERS).value
            poTotal.value = ws.Cells(i, PP_TOTAL).value
            poRecipient.value = ws.Cells(i, PP_RECIP).value
            poCustomer.value = ws.Cells(i, PP_CUST).value
            poPatient.value = ws.Cells(i, PP_PATIENT).value
            poDept.value = ws.Cells(i, PP_DEPT).value
            poStart.value = Format(ws.Cells(i, PP_DUE).value, "yyyy-mm-dd")
            SetSourceLocks True
            poTotal.Locked = False
            poSource.Enabled = False
            poGenerate.Enabled = False
            mLoading = False
            UpdatePreview
            Exit Sub
        End If
    Next i
    Exit Sub
Fail:
    mLoading = False
    MsgBox "Plan lookup error: " & Err.Description, vbExclamation
End Sub

Private Sub poTotal_Change()
    If Not mLoading And Not mReminder Then UpdatePreview
End Sub

Private Sub poPreview_Click()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    Dim p As PaymentPlanResult, i As Long, startD As Date
    On Error GoTo Invalid
    poSchedule.Clear
    If Not IsDate(poStart.value) Then Err.Raise 5, , "Enter a valid deposit due date."
    startD = CDate(poStart.value)
    If modHelpers.Num(poAligners.value) <> Fix(modHelpers.Num(poAligners.value)) Then _
        Err.Raise 5, , "Number of aligners must be a whole number."
    p = CalcPaymentPlan(modHelpers.Num(poAligners.value), modHelpers.Num(poTotal.value))
    poSummary.Caption = Format(GetDepositPercent(), "0.##%") & " deposit; " & _
        Format(p.TreatmentWeeks, "0.##") & " treatment weeks; " & p.InstallmentCount & _
        " months; monthly R" & Format(p.MonthlyAmount, "#,##0.00") & " (final month adjusted)."
    For i = 0 To p.InstallmentCount
        poSchedule.AddItem IIf(i = 0, "Deposit", "Month " & i)
        poSchedule.List(poSchedule.ListCount - 1, 1) = Format(DateAdd("m", i, startD), "yyyy-mm-dd")
        poSchedule.List(poSchedule.ListCount - 1, 2) = "R " & Format(p.Schedule(i), "#,##0.00")
    Next i
    Exit Sub
Invalid:
    poSummary.Caption = "Preview: " & Err.Description
End Sub

Private Sub poGenerate_Click()
    Dim id As String
    On Error GoTo Fail
    If Not IsDate(poStart.value) Then Err.Raise 5, , "Enter a valid deposit due date."
    If modHelpers.Num(poAligners.value) <> Fix(modHelpers.Num(poAligners.value)) Then _
        Err.Raise 5, , "Number of aligners must be a whole number."
    If MsgBox("Generate the deposit and all installment invoices? An unpaid source invoice will be replaced, not charged twice.", _
        vbQuestion + vbYesNo, "Payment Options") <> vbYes Then Exit Sub
    id = GeneratePlanInvoices(CStr(poSource.value), modHelpers.Num(poAligners.value), _
        modHelpers.Num(poTotal.value), CStr(poRecipient.value), CStr(poCustomer.value), _
        CStr(poPatient.value), CStr(poDept.value), CDate(poStart.value))
    If id <> "" Then poPlan.value = id: MsgBox "Plan created: " & id, vbInformation
    Exit Sub
Fail:
    MsgBox "Generate plan error: " & Err.Description, vbExclamation
End Sub

Private Sub poCancelPlan_Click()
    If Trim(CStr(poPlan.value)) = "" Then MsgBox "Enter the Plan ID.", vbExclamation: Exit Sub
    If MsgBox("Cancel this plan and void its unpaid, unsent invoices?", vbQuestion + vbYesNo) = vbYes Then
        If CancelPaymentPlan(CStr(poPlan.value)) Then MsgBox "Plan cancelled; audit records retained.", vbInformation
    End If
End Sub

Private Sub poAmend_Click()
    Dim id As String
    If Trim(CStr(poPlan.value)) = "" Then MsgBox "Enter the Plan ID.", vbExclamation: Exit Sub
    If MsgBox("Replace this unpaid, unsent plan using the current source total and aligner count?", _
        vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    id = AmendPaymentPlan(CStr(poPlan.value), modHelpers.Num(poAligners.value), modHelpers.Num(poTotal.value))
    If id <> "" Then poPlan.value = id: MsgBox "Replacement plan: " & id, vbInformation
End Sub

Private Sub poDeposits_Click()
    Me.Hide
    RunDepositsStatement
    Me.Show
End Sub

Private Sub poInstallments_Click()
    Me.Hide
    RunInstallmentsStatement
    Me.Show
End Sub

Private Sub poSingle_Click()
    Me.Hide
    RunSinglePlanStatement CStr(poPlan.value)
    Me.Show
End Sub

Public Sub ShowReminder()
    Dim pr As Variant, ws As Worksheet, ctl As Object
    mReminder = True
    Me.Caption = "Installment invoices - Mail now"
    Set mDueRows = DueInstallmentRows()
    If mDueRows.Count = 0 Then Exit Sub
    Set ws = ThisWorkbook.Sheets("PaymentPlans")
    For Each ctl In Me.Controls
        ctl.Visible = False
    Next ctl
    poSummary.Visible = True: poSchedule.Visible = True: poMail.Visible = True: poClose.Visible = True
    poSummary.Top = 12: poSummary.Caption = "All unpaid, unsent installments due soon or overdue. Select an invoice and click Mail now."
    poSchedule.Top = 55: poSchedule.Height = 349: poSchedule.Clear
    For Each pr In mDueRows
        poSchedule.AddItem ws.Cells(CLng(pr), PP_INVOICE).value
        poSchedule.List(poSchedule.ListCount - 1, 1) = Format(ws.Cells(CLng(pr), PP_DUE).value, "yyyy-mm-dd")
        poSchedule.List(poSchedule.ListCount - 1, 2) = "R " & Format(ws.Cells(CLng(pr), PP_AMOUNT).value, "#,##0.00")
        poSchedule.List(poSchedule.ListCount - 1, 3) = ws.Cells(CLng(pr), PP_PATIENT).value
    Next pr
    Me.Show
End Sub

Private Sub poMail_Click()
    Dim i As Long
    i = poSchedule.ListIndex
    If i < 0 Then MsgBox "Select an invoice to mail.", vbInformation: Exit Sub
    If MailInstallmentInvoice(CLng(mDueRows(i + 1))) Then
        poSchedule.RemoveItem i
        mDueRows.Remove i + 1
        If mDueRows.Count = 0 Then Unload Me
    End If
End Sub

Private Sub poClose_Click()
    Unload Me
End Sub
