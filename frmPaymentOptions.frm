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

Private WithEvents m_poSource As MSForms.ComboBox
Private WithEvents m_poAligners As MSForms.TextBox
Private WithEvents m_poTotal As MSForms.TextBox
Private WithEvents m_poPlan As MSForms.TextBox
Private WithEvents m_poPreview As MSForms.CommandButton
Private WithEvents m_poGenerate As MSForms.CommandButton
Private WithEvents m_poCancelPlan As MSForms.CommandButton
Private WithEvents m_poAmend As MSForms.CommandButton
Private WithEvents m_poDeposits As MSForms.CommandButton
Private WithEvents m_poInstallments As MSForms.CommandButton
Private WithEvents m_poSingle As MSForms.CommandButton
Private WithEvents m_poMail As MSForms.CommandButton
Private WithEvents m_poClose As MSForms.CommandButton
Private m_poRecipient As MSForms.ComboBox
Private m_poDept As MSForms.ComboBox
Private m_poCustomer As MSForms.TextBox
Private m_poPatient As MSForms.TextBox
Private m_poStart As MSForms.TextBox
Private m_poSummary As MSForms.Label
Private m_poSchedule As MSForms.ListBox
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
    Set m_poSource = AddField("ComboBox", "Source quote/invoice (blank = manual)", 12, 12, 300)
    Set m_poAligners = AddField("TextBox", "Number of aligners", 330, 12, 140)
    Set m_poTotal = AddField("TextBox", "Total (VAT inclusive)", 480, 12, 140)
    Set m_poRecipient = AddField("ComboBox", "Recipient type", 12, 64, 140)
    m_poRecipient.AddItem "doctor": m_poRecipient.AddItem "patient": m_poRecipient.AddItem "medclaim"
    m_poRecipient.ListIndex = 0
    Set m_poDept = AddField("ComboBox", "Department", 162, 64, 90)
    m_poDept.AddItem "WA": m_poDept.AddItem "WD": m_poDept.ListIndex = 0
    Set m_poCustomer = AddField("TextBox", "Customer ID / patient bill-to", 262, 64, 180)
    Set m_poPatient = AddField("TextBox", "Patient name", 452, 64, 168)
    Set m_poStart = AddField("TextBox", "Deposit due (yyyy-mm-dd)", 12, 116, 200)
    m_poStart.value = Format(Date, "yyyy-mm-dd")
    Set m_poPlan = AddField("TextBox", "Plan ID (for cancel/amend/statement)", 222, 116, 300)
    Set m_poSummary = Me.Controls.Add("Forms.Label.1", "poSummaryLabel", True)
    m_poSummary.Left = 12: m_poSummary.Top = 174: m_poSummary.Width = 608: m_poSummary.Height = 36
    Set m_poSchedule = Me.Controls.Add("Forms.ListBox.1", "poScheduleList", True)
    m_poSchedule.Left = 12: m_poSchedule.Top = 214: m_poSchedule.Width = 608: m_poSchedule.Height = 190
    m_poSchedule.ColumnCount = 4: m_poSchedule.ColumnWidths = "90 pt;110 pt;130 pt;250 pt"
    Set m_poPreview = AddButton("poPreviewButton", "Preview", 12, 414, 90)
    Set m_poGenerate = AddButton("poGenerateButton", "Generate Invoices", 112, 414, 140)
    Set m_poCancelPlan = AddButton("poCancelPlanButton", "Cancel Plan", 262, 414, 110)
    Set m_poAmend = AddButton("poAmendButton", "Amend Plan", 382, 414, 110)
    Set m_poMail = AddButton("poMailButton", "Mail now", 502, 414, 118)
    m_poMail.Visible = False
    Set m_poDeposits = AddButton("poDepositsButton", "Deposits Statement", 12, 458, 150)
    Set m_poInstallments = AddButton("poInstallmentsButton", "Installments Statement", 172, 458, 170)
    Set m_poSingle = AddButton("poSingleButton", "Plan Statement", 352, 458, 150)
    Set m_poClose = AddButton("poCloseButton", "Close", 512, 458, 108)
    mLoading = True
    For Each sheetName In Array("QuoteLog", "InvoiceLog", "MedAidLog")
        Set ws = ThisWorkbook.Sheets(CStr(sheetName))
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
            If Trim(CStr(ws.Cells(i, 1).value)) <> "" Then m_poSource.AddItem ws.Cells(i, 1).value
        Next i
    Next sheetName
    mLoading = False
    m_poSummary.Caption = "Choose a saved source, or enter total, recipient and department for a manual plan."
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

Private Sub m_poSource_Change()
    Dim ws As Worksheet, lr As Long, source As String
    If mLoading Or mReminder Then Exit Sub
    source = modHelpers.NrmID(CStr(m_poSource.value))
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
        m_poTotal.value = ws.Cells(lr, QL_TOTAL).value
        m_poRecipient.value = ws.Cells(lr, QL_RECIP).value
        m_poDept.value = ws.Cells(lr, QL_DEPT).value
        m_poCustomer.value = ws.Cells(lr, QL_CUST).value
        m_poPatient.value = ws.Cells(lr, QL_PATIENT).value
    ElseIf Left(source, 7) = "MC-INV-" Then
        Set ws = ThisWorkbook.Sheets("MedAidLog")
        lr = modHelpers.FindLogRow(ws, source)
        If lr = 0 Then GoTo Done
        m_poTotal.value = ws.Cells(lr, ML_TOTAL).value
        m_poRecipient.value = "medclaim"
        m_poDept.value = ws.Cells(lr, ML_DEPT).value
        m_poCustomer.value = ws.Cells(lr, ML_CUST).value
        m_poPatient.value = ws.Cells(lr, ML_PATIENT).value
    Else
        Set ws = ThisWorkbook.Sheets("InvoiceLog")
        lr = modHelpers.FindLogRow(ws, source)
        If lr = 0 Then GoTo Done
        m_poTotal.value = ws.Cells(lr, IL_TOTAL).value
        m_poRecipient.value = ws.Cells(lr, IL_RECIP).value
        m_poDept.value = ws.Cells(lr, IL_DEPT).value
        m_poCustomer.value = ws.Cells(lr, IL_CUST).value
        m_poPatient.value = ws.Cells(lr, IL_PATIENT).value
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
    m_poTotal.Locked = locked
    m_poRecipient.Enabled = Not locked
    m_poDept.Enabled = Not locked
    m_poCustomer.Locked = locked
    m_poPatient.Locked = locked
End Sub

Private Sub m_poAligners_Change()
    If Not mLoading And Not mReminder Then UpdatePreview
End Sub

Private Sub m_poPlan_Change()
    Dim ws As Worksheet, i As Long
    If mLoading Or mReminder Then Exit Sub
    m_poSource.Enabled = True
    m_poGenerate.Enabled = True
    If Trim(CStr(m_poPlan.value)) = "" Then
        SetSourceLocks (Trim(CStr(m_poSource.value)) <> "")
        Exit Sub
    End If
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PaymentPlans")
    On Error GoTo Fail
    If ws Is Nothing Then Exit Sub
    For i = 2 To ws.Cells(ws.Rows.Count, PP_PLANID).End(xlUp).row
        If modHelpers.NrmID(CStr(ws.Cells(i, PP_PLANID).value)) = modHelpers.NrmID(CStr(m_poPlan.value)) Then
            mLoading = True
            m_poSource.value = ws.Cells(i, PP_SOURCE).value
            m_poAligners.value = ws.Cells(i, PP_ALIGNERS).value
            m_poTotal.value = ws.Cells(i, PP_TOTAL).value
            m_poRecipient.value = ws.Cells(i, PP_RECIP).value
            m_poCustomer.value = ws.Cells(i, PP_CUST).value
            m_poPatient.value = ws.Cells(i, PP_PATIENT).value
            m_poDept.value = ws.Cells(i, PP_DEPT).value
            m_poStart.value = Format(ws.Cells(i, PP_DUE).value, "yyyy-mm-dd")
            SetSourceLocks True
            m_poTotal.Locked = False
            m_poSource.Enabled = False
            m_poGenerate.Enabled = False
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

Private Sub m_poTotal_Change()
    If Not mLoading And Not mReminder Then UpdatePreview
End Sub

Private Sub m_poPreview_Click()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    Dim p As PaymentPlanResult, i As Long, startD As Date
    On Error GoTo Invalid
    m_poSchedule.Clear
    If Not IsDate(m_poStart.value) Then Err.Raise 5, , "Enter a valid deposit due date."
    startD = CDate(m_poStart.value)
    If modHelpers.Num(m_poAligners.value) <> Fix(modHelpers.Num(m_poAligners.value)) Then _
        Err.Raise 5, , "Number of aligners must be a whole number."
    p = CalcPaymentPlan(modHelpers.Num(m_poAligners.value), modHelpers.Num(m_poTotal.value))
    m_poSummary.Caption = Format(GetDepositPercent(), "0.##%") & " deposit; " & _
        Format(p.TreatmentWeeks, "0.##") & " treatment weeks; " & p.InstallmentCount & _
        " months; monthly R" & Format(p.MonthlyAmount, "#,##0.00") & " (final month adjusted)."
    For i = 0 To p.InstallmentCount
        m_poSchedule.AddItem IIf(i = 0, "Deposit", "Month " & i)
        m_poSchedule.List(m_poSchedule.ListCount - 1, 1) = Format(DateAdd("m", i, startD), "yyyy-mm-dd")
        m_poSchedule.List(m_poSchedule.ListCount - 1, 2) = "R " & Format(p.Schedule(i), "#,##0.00")
    Next i
    Exit Sub
Invalid:
    m_poSummary.Caption = "Preview: " & Err.Description
End Sub

Private Sub m_poGenerate_Click()
    Dim id As String
    On Error GoTo Fail
    If Not IsDate(m_poStart.value) Then Err.Raise 5, , "Enter a valid deposit due date."
    If modHelpers.Num(m_poAligners.value) <> Fix(modHelpers.Num(m_poAligners.value)) Then _
        Err.Raise 5, , "Number of aligners must be a whole number."
    If MsgBox("Generate the deposit and all installment invoices? An unpaid source invoice will be replaced, not charged twice.", _
        vbQuestion + vbYesNo, "Payment Options") <> vbYes Then Exit Sub
    id = GeneratePlanInvoices(CStr(m_poSource.value), modHelpers.Num(m_poAligners.value), _
        modHelpers.Num(m_poTotal.value), CStr(m_poRecipient.value), CStr(m_poCustomer.value), _
        CStr(m_poPatient.value), CStr(m_poDept.value), CDate(m_poStart.value))
    If id <> "" Then m_poPlan.value = id: MsgBox "Plan created: " & id, vbInformation
    Exit Sub
Fail:
    MsgBox "Generate plan error: " & Err.Description, vbExclamation
End Sub

Private Sub m_poCancelPlan_Click()
    If Trim(CStr(m_poPlan.value)) = "" Then MsgBox "Enter the Plan ID.", vbExclamation: Exit Sub
    If MsgBox("Cancel this plan and void its unpaid, unsent invoices?", vbQuestion + vbYesNo) = vbYes Then
        If CancelPaymentPlan(CStr(m_poPlan.value)) Then MsgBox "Plan cancelled; audit records retained.", vbInformation
    End If
End Sub

Private Sub m_poAmend_Click()
    Dim id As String
    If Trim(CStr(m_poPlan.value)) = "" Then MsgBox "Enter the Plan ID.", vbExclamation: Exit Sub
    If MsgBox("Replace this unpaid, unsent plan using the current source total and aligner count?", _
        vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    id = AmendPaymentPlan(CStr(m_poPlan.value), modHelpers.Num(m_poAligners.value), modHelpers.Num(m_poTotal.value))
    If id <> "" Then m_poPlan.value = id: MsgBox "Replacement plan: " & id, vbInformation
End Sub

Private Sub m_poDeposits_Click()
    Me.Hide
    RunDepositsStatement
    Me.Show
End Sub

Private Sub m_poInstallments_Click()
    Me.Hide
    RunInstallmentsStatement
    Me.Show
End Sub

Private Sub m_poSingle_Click()
    Me.Hide
    RunSinglePlanStatement CStr(m_poPlan.value)
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
    m_poSummary.Visible = True: m_poSchedule.Visible = True: m_poMail.Visible = True: m_poClose.Visible = True
    m_poSummary.Top = 12: m_poSummary.Caption = "All unpaid, unsent installments due soon or overdue. Select an invoice and click Mail now."
    m_poSchedule.Top = 55: m_poSchedule.Height = 349: m_poSchedule.Clear
    For Each pr In mDueRows
        m_poSchedule.AddItem ws.Cells(CLng(pr), PP_INVOICE).value
        m_poSchedule.List(m_poSchedule.ListCount - 1, 1) = Format(ws.Cells(CLng(pr), PP_DUE).value, "yyyy-mm-dd")
        m_poSchedule.List(m_poSchedule.ListCount - 1, 2) = "R " & Format(ws.Cells(CLng(pr), PP_AMOUNT).value, "#,##0.00")
        m_poSchedule.List(m_poSchedule.ListCount - 1, 3) = ws.Cells(CLng(pr), PP_PATIENT).value
    Next pr
    Me.Show
End Sub

Private Sub m_poMail_Click()
    Dim i As Long
    i = m_poSchedule.ListIndex
    If i < 0 Then MsgBox "Select an invoice to mail.", vbInformation: Exit Sub
    If MailInstallmentInvoice(CLng(mDueRows(i + 1))) Then
        m_poSchedule.RemoveItem i
        mDueRows.Remove i + 1
        If mDueRows.Count = 0 Then Unload Me
    End If
End Sub

Private Sub m_poClose_Click()
    Unload Me
End Sub
