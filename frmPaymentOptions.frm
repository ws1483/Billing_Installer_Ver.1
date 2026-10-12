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
Private WithEvents m_poDepositPercent As MSForms.TextBox
Private WithEvents m_poFolder As MSForms.CommandButton
Private WithEvents m_poExport As MSForms.CommandButton
Private WithEvents m_poExportAll As MSForms.CommandButton
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
Private mSourceNumbers As Collection

Private Sub UserForm_Initialize()
    Dim ctl As Object, ws As Worksheet, sheetName As Variant, i As Long
    ' Reuse the existing statement form's exported container; build the new controls here.
    For Each ctl In Me.Controls
        ctl.Visible = False
    Next ctl
    mLoading = True
    Me.Width = 650: Me.Height = 590
    Set m_poSource = AddField("ComboBox", "Source quote/invoice (blank = manual)", 12, 12, 300)
    m_poSource.MatchEntry = fmMatchEntryNone
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
    Set m_poDepositPercent = AddField("TextBox", "Deposit %", 532, 116, 88)
    m_poDepositPercent.value = Format(GetDepositPercent() * 100, "0.##")
    Set m_poSummary = Me.Controls.Add("Forms.Label.1", "poSummaryLabel", True)
    m_poSummary.Left = 12: m_poSummary.Top = 174: m_poSummary.Width = 608: m_poSummary.Height = 36
    Set m_poSchedule = Me.Controls.Add("Forms.ListBox.1", "poScheduleList", True)
    m_poSchedule.Left = 12: m_poSchedule.Top = 214: m_poSchedule.Width = 608: m_poSchedule.Height = 190
    m_poSchedule.ColumnCount = 4: m_poSchedule.ColumnWidths = "90 pt;110 pt;130 pt;250 pt"
    Set m_poPreview = AddButton("poPreviewButton", "Preview", 12, 414, 90)
    Set m_poGenerate = AddButton("poGenerateButton", "Create Plan", 112, 414, 140)
    Set m_poCancelPlan = AddButton("poCancelPlanButton", "Cancel Plan", 262, 414, 110)
    Set m_poAmend = AddButton("poAmendButton", "Amend Plan", 382, 414, 110)
    Set m_poMail = AddButton("poMailButton", "Mail now", 502, 414, 118)
    m_poMail.Visible = False
    Set m_poDeposits = AddButton("poDepositsButton", "Deposits Statement", 12, 458, 150)
    Set m_poInstallments = AddButton("poInstallmentsButton", "Installments Statement", 172, 458, 170)
    Set m_poSingle = AddButton("poSingleButton", "Plan Statement", 352, 458, 150)
    Set m_poClose = AddButton("poCloseButton", "Close", 512, 458, 108)
    Set m_poFolder = AddButton("poFolderButton", "PDF folder", 12, 500, 150)
    Set m_poExport = AddButton("poExportButton", "Export request", 172, 500, 170)
    m_poExport.Visible = False
    Set m_poExportAll = AddButton("poExportAllButton", "Export all requests", 352, 500, 170)
    m_poExportAll.Visible = False
    Set mSourceNumbers = New Collection
    For Each sheetName In Array("QuoteLog", "InvoiceLog", "MedAidLog")
        Set ws = ThisWorkbook.Sheets(CStr(sheetName))
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
            If Trim(CStr(ws.Cells(i, 1).value)) <> "" Then
                mSourceNumbers.Add CStr(ws.Cells(i, 1).value)
                m_poSource.AddItem ws.Cells(i, 1).value
            End If
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
    Dim ws As Worksheet, lr As Long, source As String, typed As String, item As Variant, cursor As Long
    If mLoading Or mReminder Then Exit Sub
    On Error GoTo Fail
    typed = CStr(m_poSource.value): cursor = m_poSource.SelStart
    source = modHelpers.NrmID(CStr(m_poSource.value))
    If m_poSource.ListIndex < 0 Then
        mLoading = True
        m_poSource.Clear
        For Each item In mSourceNumbers
            If InStr(1, CStr(item), typed, vbTextCompare) > 0 Then m_poSource.AddItem CStr(item)
        Next item
        m_poSource.value = typed
        m_poSource.SelStart = cursor
        mLoading = False
    End If
    If source = "" Then
        SetSourceLocks False
        Exit Sub
    End If
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
            m_poDepositPercent.value = Format(CDbl(ws.Cells(i, PP_PERCENT).value) * 100, "0.##")
            m_poSource.Enabled = False
            m_poGenerate.Enabled = False
            mLoading = False
            UpdatePreview
            LoadSavedSchedule ws, CStr(m_poPlan.value)
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
    m_poMail.Visible = False: m_poExport.Visible = False: m_poExportAll.Visible = False
    If Not IsDate(m_poStart.value) Then Err.Raise 5, , "Enter a valid deposit due date."
    startD = CDate(m_poStart.value)
    If modHelpers.Num(m_poAligners.value) <> Fix(modHelpers.Num(m_poAligners.value)) Then _
        Err.Raise 5, , "Number of aligners must be a whole number."
    p = CalcPaymentPlan(modHelpers.Num(m_poAligners.value), modHelpers.Num(m_poTotal.value), CaseDepositPercent())
    m_poSummary.Caption = Format(CaseDepositPercent(), "0.##%") & " deposit; " & _
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
    If MsgBox("Create nonfinancial payment requests against ONE permanent full-total invoice?" & vbCrLf & _
        "A saved invoice / medical claim is retained unchanged; a quote or manual case creates one invoice.", _
        vbQuestion + vbYesNo, "Payment Options") <> vbYes Then Exit Sub
    id = GeneratePlanInvoices(CStr(m_poSource.value), modHelpers.Num(m_poAligners.value), _
        modHelpers.Num(m_poTotal.value), CStr(m_poRecipient.value), CStr(m_poCustomer.value), _
        CStr(m_poPatient.value), CStr(m_poDept.value), CDate(m_poStart.value), CaseDepositPercent())
    If id <> "" Then
        m_poPlan.value = id
        MsgBox "Plan created: " & id, vbInformation
        If MsgBox("Save the deposit AND every installment notice to one shared PDF folder now?" & vbCrLf & _
            "These are nonfinancial notices; pending / paid notices are not emailed.", _
            vbYesNo + vbQuestion, "Export all payment requests") = vbYes Then m_poExportAll_Click
    End If
    Exit Sub
Fail:
    MsgBox "Generate plan error: " & Err.Description, vbExclamation
End Sub

Private Sub m_poCancelPlan_Click()
    If Trim(CStr(m_poPlan.value)) = "" Then MsgBox "Enter the Plan ID.", vbExclamation: Exit Sub
    If MsgBox("Cancel only this unpaid, unexported, unsent schedule? The permanent invoice stays payable.", vbQuestion + vbYesNo) = vbYes Then
        If CancelPaymentPlan(CStr(m_poPlan.value)) Then MsgBox "Plan cancelled; audit records retained.", vbInformation
    End If
End Sub

Private Sub m_poAmend_Click()
    Dim id As String
    If Trim(CStr(m_poPlan.value)) = "" Then MsgBox "Enter the Plan ID.", vbExclamation: Exit Sub
    If MsgBox("Revise this unpaid, unexported, unsent schedule? The permanent invoice / total is unchanged.", _
        vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    On Error GoTo Fail
    id = AmendPaymentPlan(CStr(m_poPlan.value), modHelpers.Num(m_poAligners.value), _
        modHelpers.Num(m_poTotal.value), CaseDepositPercent())
    If id <> "" Then m_poPlan.value = id: MsgBox "Replacement plan: " & id, vbInformation
    Exit Sub
Fail:
    MsgBox "Amend error: " & Err.Description, vbExclamation
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
    Me.Caption = "Payment requests - Mail now"
    Set mDueRows = DueInstallmentRows()
    If mDueRows.Count = 0 Then Exit Sub
    Set ws = ThisWorkbook.Sheets("PaymentPlans")
    For Each ctl In Me.Controls
        ctl.Visible = False
    Next ctl
    m_poSummary.Visible = True: m_poSchedule.Visible = True: m_poMail.Visible = True: m_poClose.Visible = True
    m_poSummary.Top = 12: m_poSummary.Caption = "Released unpaid requests due soon / overdue. Select a request and click Mail now."
    m_poSchedule.Top = 55: m_poSchedule.Height = 349: m_poSchedule.Clear
    For Each pr In mDueRows
        m_poSchedule.AddItem ws.Cells(CLng(pr), PP_REQUEST).value
        m_poSchedule.List(m_poSchedule.ListCount - 1, 1) = Format(ws.Cells(CLng(pr), PP_DUE).value, "yyyy-mm-dd")
        m_poSchedule.List(m_poSchedule.ListCount - 1, 2) = "R " & Format(ws.Cells(CLng(pr), PP_AMOUNT).value, "#,##0.00")
        m_poSchedule.List(m_poSchedule.ListCount - 1, 3) = ws.Cells(CLng(pr), PP_PATIENT).value
    Next pr
    Me.Show
End Sub

Private Sub m_poMail_Click()
    Dim i As Long
    On Error GoTo Fail
    i = m_poSchedule.ListIndex
    If i < 0 Or mDueRows Is Nothing Then MsgBox "Select a saved request to mail.", vbInformation: Exit Sub
    If MailInstallmentInvoice(CLng(mDueRows(i + 1))) Then
        If mReminder Then
            m_poSchedule.RemoveItem i
            mDueRows.Remove i + 1
            If mDueRows.Count = 0 Then Unload Me
        Else
            LoadSavedSchedule EnsurePaymentPlansSheet(), CStr(m_poPlan.value)
        End If
    End If
    Exit Sub
Fail:
    MsgBox "Mail / refresh error: " & Err.Description & ". Check Outlook / AuditLog before retrying.", vbExclamation
End Sub

Private Function CaseDepositPercent() As Double
    If Not IsNumeric(m_poDepositPercent.value) Then Err.Raise 5, , "Enter deposit percentage from 0 to 100."
    CaseDepositPercent = CDbl(m_poDepositPercent.value) / 100
    If CaseDepositPercent < 0 Or CaseDepositPercent > 1 Then Err.Raise 5, , "Deposit percentage must be 0 to 100."
End Function

Private Sub m_poDepositPercent_Change()
    If Not mLoading And Not mReminder Then UpdatePreview
End Sub

Private Sub m_poFolder_Click()
    Dim folder As String
    folder = modExportPDF.PaymentPDFDirectory(True)
    If folder <> "" Then m_poSummary.Caption = "PDF folder: " & folder
End Sub

Private Sub m_poExport_Click()
    Dim path As String, i As Long
    On Error GoTo Fail
    i = m_poSchedule.ListIndex
    If i < 0 Or mDueRows Is Nothing Then MsgBox "Select a saved payment request.", vbInformation: Exit Sub
    path = ExportPaymentRequest(CLng(mDueRows(i + 1)))
    If path <> "" Then MsgBox "Payment request exported: " & path, vbInformation
    Exit Sub
Fail:
    MsgBox "Request export error: " & Err.Description, vbExclamation
End Sub

Private Sub m_poExportAll_Click()
    Dim count As Long
    On Error GoTo Fail
    If Trim$(CStr(m_poPlan.value)) = "" Then MsgBox "Enter a saved plan ID.", vbInformation: Exit Sub
    count = ExportPlanRequests(CStr(m_poPlan.value))
    If count > 0 Then MsgBox CStr(count) & " nonfinancial notices exported. No additional invoice was created.", vbInformation
    Exit Sub
Fail:
    MsgBox "PDF export error: " & Err.Description, vbExclamation
End Sub

Private Sub LoadSavedSchedule(ByVal ws As Worksheet, ByVal planID As String)
    Dim r As Long
    RefreshPaymentPlans
    Set mDueRows = New Collection
    m_poSchedule.Clear
    For r = 2 To ws.Cells(ws.Rows.Count, PP_PLANID).End(xlUp).row
        If modHelpers.NrmID(CStr(ws.Cells(r, PP_PLANID).value)) = modHelpers.NrmID(planID) Then
            If CStr(ws.Cells(r, PP_STATUS).value) <> "Cancelled" Then
                mDueRows.Add r
                m_poSchedule.AddItem ws.Cells(r, PP_REQUEST).value
                m_poSchedule.List(m_poSchedule.ListCount - 1, 1) = Format(ws.Cells(r, PP_DUE).value, "yyyy-mm-dd")
                m_poSchedule.List(m_poSchedule.ListCount - 1, 2) = "R " & Format(ws.Cells(r, PP_BALANCE).value, "#,##0.00")
                m_poSchedule.List(m_poSchedule.ListCount - 1, 3) = ws.Cells(r, PP_STATUS).value
            End If
        End If
    Next r
    m_poMail.Visible = (mDueRows.Count > 0): m_poExport.Visible = (mDueRows.Count > 0)
    m_poExportAll.Visible = (mDueRows.Count > 0)
    m_poSummary.Caption = "Nonfinancial requests; balances allocated to permanent invoice. Deposit releases later requests."
End Sub

Private Sub m_poClose_Click()
    Unload Me
End Sub
