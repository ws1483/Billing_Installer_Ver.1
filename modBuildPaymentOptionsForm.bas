Attribute VB_Name = "modBuildPaymentOptionsForm"
Option Explicit

Private Const vbext_ct_MSForm As Long = 3
Private Const FORM_NAME As String = "frmPaymentOptions"
Private Const TEAL As Long = &H00B0872E&
Private Const WHITE As Long = &H00FFFFFF&

Public Sub BuildPaymentOptionsForm()
    Dim project As Object, component As Object, existing As Object, form As Object
    Dim item As Object, detail As String, componentCount As Long, backupName As String

    On Error GoTo AccessDenied
    Set project = Application.VBE.ActiveVBProject
    componentCount = project.VBComponents.Count
    On Error GoTo Failed
    If Not project Is ThisWorkbook.VBProject Then
        MsgBox "Select this workbook's VBA project in the VBE, then re-run BuildPaymentOptionsForm.", _
               vbExclamation, "Payment Options"
        Exit Sub
    End If
    If project.Protection <> 0 Then
        MsgBox "Unlock this workbook's VBA project before building Payment Options.", vbExclamation
        Exit Sub
    End If
    For Each item In project.VBComponents
        If StrComp(item.Name, FORM_NAME, vbTextCompare) = 0 Then
            Set existing = item
            Exit For
        End If
    Next item
    If Not existing Is Nothing Then
        If existing.Type <> vbext_ct_MSForm Then
            MsgBox FORM_NAME & " is not a UserForm. Rename that component before building.", vbExclamation
            Exit Sub
        End If
        If MsgBox("Replace the existing Payment Options form? Any custom changes to it will be lost.", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Payment Options") <> vbYes Then Exit Sub
    End If
    For Each item In VBA.UserForms
        If StrComp(TypeName(item), FORM_NAME, vbTextCompare) = 0 Then
            MsgBox "Close Payment Options before rebuilding it.", vbExclamation
            Exit Sub
        End If
    Next item

    Set component = project.VBComponents.Add(vbext_ct_MSForm)
    component.Properties("Width") = 680
    component.Properties("Height") = 450
    Set form = component.Designer
    With form
        .Caption = "Payment Options"
        .BackColor = TEAL
        .ForeColor = &H80000012&
        .BorderStyle = 0
        .Font.Name = "Tahoma"
        .Font.Size = 9
        .Cycle = 0
        .KeepScrollBarsVisible = 3
        .SpecialEffect = 0
    End With
    component.Properties("StartUpPosition") = 1
    component.Properties("ShowModal") = True

    AddLabel form, "Recipient", 12, 10, 300, True
    AddField form, "ComboBox", "poRecipient", "Recipient Type", 12, 32, 170
    AddField form, "ComboBox", "poDept", "Department", 192, 32, 120
    AddField form, "ComboBox", "poSource", "Source Quote/Invoice (blank = manual)", 332, 32, 330
    AddField form, "TextBox", "poCustomer", "Customer ID / patient bill-to", 332, 78, 330
    AddField form, "TextBox", "poPatient", "Patient Name", 12, 78, 300
    AddField form, "TextBox", "poAligners", "No. of Aligners", 12, 124, 145
    AddField form, "TextBox", "poTotal", "Quote Total (VAT inclusive)", 167, 124, 145
    AddField form, "TextBox", "poStart", "Deposit Due (yyyy-mm-dd)", 332, 124, 160
    AddField form, "TextBox", "poPlanID", "Plan ID (lookup / amend / cancel)", 502, 124, 160
    AddLabel form, "Plan Summary", 12, 174, 650, True
    Set item = AddControl(form, "TextBox", "poPlan", 12, 194, 650, 34)
    item.MultiLine = True
    item.Locked = True
    item.TabStop = False
    AddLabel form, "Schedule Preview", 12, 234, 650, True
    Set item = AddControl(form, "TextBox", "poPreview", 12, 254, 650, 100)
    item.MultiLine = True
    item.ScrollBars = 3
    item.WordWrap = False
    item.Locked = True
    Set item = AddControl(form, "ListBox", "poSchedule", 12, 254, 650, 100)
    item.ColumnCount = 4
    item.ColumnWidths = "120 pt;100 pt;100 pt;300 pt"
    item.Visible = False

    AddButton form, "poGenerate", "Generate Invoices", 12, 364, 170
    AddButton form, "poAmend", "Amend Plan", 192, 364, 150
    AddButton form, "poCancelPlan", "Cancel Plan", 352, 364, 150
    AddButton form, "poMail", "Mail Installment", 512, 364, 150
    AddButton form, "poDeposits", "Deposits Statement", 12, 402, 170
    AddButton form, "poInstallments", "Installments Statement", 192, 402, 150
    AddButton form, "poSingle", "Single Plan Statement", 352, 402, 150
    AddButton form, "poClose", "Close", 512, 402, 150
    form.Controls("poClose").Cancel = True

    InjectPaymentOptionsCode component.CodeModule
    ' Retain the old form until construction and the final rename have succeeded.
    If Not existing Is Nothing Then
        backupName = AvailableComponentName(project)
        existing.Name = backupName
    End If
    component.Name = FORM_NAME
    If Not existing Is Nothing Then project.VBComponents.Remove existing
    On Error GoTo 0
    MsgBox "Payment Options was built. Run BtnPaymentOptions to open it.", vbInformation
    Exit Sub
AccessDenied:
    MsgBox "Enable File > Options > Trust Center > Trust Center Settings > Macro Settings > " & _
           "Trust access to the VBA project object model, then re-run BuildPaymentOptionsForm." & _
           vbCrLf & Err.Description, vbExclamation, "Payment Options"
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not component Is Nothing Then project.VBComponents.Remove component
    If backupName <> "" Then existing.Name = FORM_NAME
    On Error GoTo 0
    MsgBox "Payment Options could not be built: " & detail, vbExclamation, "Payment Options"
End Sub

Private Function AvailableComponentName(ByVal project As Object) As String
    Dim i As Long, item As Object, candidate As String, found As Boolean
    Do
        i = i + 1
        candidate = "poPreviousForm" & CStr(i)
        found = False
        For Each item In project.VBComponents
            If StrComp(item.Name, candidate, vbTextCompare) = 0 Then found = True
        Next item
        If Not found Then Exit Do
    Loop
    AvailableComponentName = candidate
End Function

Private Function AddControl(ByVal form As Object, ByVal kind As String, ByVal name As String, _
                            ByVal x As Single, ByVal y As Single, ByVal w As Single, _
                            ByVal h As Single) As Object
    Dim ctl As Object
    Set ctl = form.Controls.Add("Forms." & kind & ".1", name, True)
    With ctl
        .Left = x: .Top = y: .Width = w: .Height = h
        .Font.Name = "Tahoma"
        .Font.Size = 9
        .BackColor = WHITE
        .ForeColor = &H80000012&
    End With
    Set AddControl = ctl
End Function

Private Sub AddLabel(ByVal form As Object, ByVal caption As String, ByVal x As Single, _
                     ByVal y As Single, ByVal w As Single, Optional ByVal header As Boolean = False)
    Dim ctl As Object
    Set ctl = AddControl(form, "Label", "poLabel" & CStr(form.Controls.Count), x, y, w, 18)
    ctl.Caption = caption
    ctl.BackColor = TEAL
    ctl.ForeColor = WHITE
    ctl.Font.Bold = header
End Sub

Private Sub AddField(ByVal form As Object, ByVal kind As String, ByVal name As String, _
                     ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal w As Single)
    Dim ctl As Object
    AddLabel form, caption, x, y, w
    Set ctl = AddControl(form, kind, name, x, y + 18, w, 22)
    If kind = "ComboBox" And name <> "poSource" Then ctl.Style = 2
End Sub

Private Sub AddButton(ByVal form As Object, ByVal name As String, ByVal caption As String, _
                      ByVal x As Single, ByVal y As Single, ByVal w As Single)
    Dim ctl As Object
    Set ctl = AddControl(form, "CommandButton", name, x, y, w, 28)
    ctl.Caption = caption
    ctl.BackColor = &H8000000F&
End Sub

Private Sub CodeLine(ByVal code As Object, ByVal text As String)
    code.InsertLines code.CountOfLines + 1, text
End Sub

Private Sub InjectPaymentOptionsCode(ByVal code As Object)
    If code.CountOfLines > 0 Then code.DeleteLines 1, code.CountOfLines
    CodeLine code, "Option Explicit"
    CodeLine code, ""
    CodeLine code, "Private mLoading As Boolean"
    CodeLine code, "Private mReminder As Boolean"
    CodeLine code, "Private mDueRows As Collection"
    CodeLine code, ""
    CodeLine code, "Private Sub UserForm_Initialize()"
    CodeLine code, "    Dim ws As Worksheet, sheetName As Variant, i As Long"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    mLoading = True"
    CodeLine code, "    poRecipient.AddItem ""Doctor"""
    CodeLine code, "    poRecipient.AddItem ""Private Patient"""
    CodeLine code, "    poRecipient.AddItem ""Med Claim"""
    CodeLine code, "    poRecipient.ListIndex = 0"
    CodeLine code, "    poDept.AddItem ""All"": poDept.AddItem ""WA"": poDept.AddItem ""WD"": poDept.AddItem ""MC"""
    CodeLine code, "    poDept.Value = ""WA"""
    CodeLine code, "    poStart.Value = Format(Date, ""yyyy-mm-dd"")"
    CodeLine code, "    For Each sheetName In Array(""QuoteLog"", ""InvoiceLog"", ""MedAidLog"")"
    CodeLine code, "        Set ws = Nothing"
    CodeLine code, "        On Error Resume Next"
    CodeLine code, "        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))"
    CodeLine code, "        On Error GoTo Failed"
    CodeLine code, "        If Not ws Is Nothing Then"
    CodeLine code, "            For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row"
    CodeLine code, "                If Trim(CStr(ws.Cells(i, 1).Value)) <> """" Then poSource.AddItem ws.Cells(i, 1).Value"
    CodeLine code, "            Next i"
    CodeLine code, "        End If"
    CodeLine code, "    Next sheetName"
    CodeLine code, "    mLoading = False"
    CodeLine code, "    poPlan.Value = ""Choose a source, or enter a total, aligner count, recipient and department."""
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    mLoading = False"
    CodeLine code, "    MsgBox ""Payment Options initialization: "" & Err.Description, vbExclamation"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Function RecipientCode() As String"
    CodeLine code, "    Select Case poRecipient.ListIndex"
    CodeLine code, "        Case 0: RecipientCode = ""doctor"""
    CodeLine code, "        Case 1: RecipientCode = ""patient"""
    CodeLine code, "        Case 2: RecipientCode = ""medclaim"""
    CodeLine code, "    End Select"
    CodeLine code, "End Function"
    CodeLine code, ""
    CodeLine code, "Private Sub SetRecipient(ByVal value As String)"
    CodeLine code, "    Select Case LCase$(Trim$(value))"
    CodeLine code, "        Case ""doctor"": poRecipient.ListIndex = 0"
    CodeLine code, "        Case ""patient"": poRecipient.ListIndex = 1"
    CodeLine code, "        Case ""medclaim"": poRecipient.ListIndex = 2"
    CodeLine code, "        Case Else: Err.Raise 5, , ""Unknown recipient type: "" & value"
    CodeLine code, "    End Select"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poSource_Change()"
    CodeLine code, "    Dim ws As Worksheet, lr As Long, source As String"
    CodeLine code, "    If mLoading Or mReminder Then Exit Sub"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    source = modHelpers.NrmID(CStr(poSource.Value))"
    CodeLine code, "    If source = """" Then"
    CodeLine code, "        SetSourceLocks False"
    CodeLine code, "        UpdatePreview"
    CodeLine code, "        Exit Sub"
    CodeLine code, "    End If"
    CodeLine code, "    mLoading = True"
    CodeLine code, "    If Left$(source, 2) = ""Q-"" Then"
    CodeLine code, "        Set ws = ThisWorkbook.Worksheets(""QuoteLog"")"
    CodeLine code, "        lr = modHelpers.FindLogRow(ws, source)"
    CodeLine code, "        If lr = 0 Then Err.Raise 5, , ""Source quote not found."""
    CodeLine code, "        poTotal.Value = ws.Cells(lr, QL_TOTAL).Value"
    CodeLine code, "        SetRecipient CStr(ws.Cells(lr, QL_RECIP).Value)"
    CodeLine code, "        poDept.Value = UCase$(Trim$(CStr(ws.Cells(lr, QL_DEPT).Value)))"
    CodeLine code, "        poCustomer.Value = ws.Cells(lr, QL_CUST).Value"
    CodeLine code, "        poPatient.Value = ws.Cells(lr, QL_PATIENT).Value"
    CodeLine code, "    ElseIf Left$(source, 7) = ""MC-INV-"" Then"
    CodeLine code, "        Set ws = ThisWorkbook.Worksheets(""MedAidLog"")"
    CodeLine code, "        lr = modHelpers.FindLogRow(ws, source)"
    CodeLine code, "        If lr = 0 Then Err.Raise 5, , ""Source med claim not found."""
    CodeLine code, "        SetRecipient ""medclaim"""
    CodeLine code, "        poTotal.Value = ws.Cells(lr, ML_TOTAL).Value"
    CodeLine code, "        poDept.Value = UCase$(Trim$(CStr(ws.Cells(lr, ML_DEPT).Value)))"
    CodeLine code, "        poCustomer.Value = ws.Cells(lr, ML_CUST).Value"
    CodeLine code, "        poPatient.Value = ws.Cells(lr, ML_PATIENT).Value"
    CodeLine code, "    Else"
    CodeLine code, "        Set ws = ThisWorkbook.Worksheets(""InvoiceLog"")"
    CodeLine code, "        lr = modHelpers.FindLogRow(ws, source)"
    CodeLine code, "        If lr = 0 Then Err.Raise 5, , ""Source invoice not found."""
    CodeLine code, "        poTotal.Value = ws.Cells(lr, IL_TOTAL).Value"
    CodeLine code, "        SetRecipient CStr(ws.Cells(lr, IL_RECIP).Value)"
    CodeLine code, "        poDept.Value = UCase$(Trim$(CStr(ws.Cells(lr, IL_DEPT).Value)))"
    CodeLine code, "        poCustomer.Value = ws.Cells(lr, IL_CUST).Value"
    CodeLine code, "        poPatient.Value = ws.Cells(lr, IL_PATIENT).Value"
    CodeLine code, "    End If"
    CodeLine code, "    SetSourceLocks True"
    CodeLine code, "    mLoading = False"
    CodeLine code, "    UpdatePreview"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    mLoading = False"
    CodeLine code, "    MsgBox ""Source lookup error: "" & Err.Description, vbExclamation"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub SetSourceLocks(ByVal locked As Boolean)"
    CodeLine code, "    poTotal.Locked = locked"
    CodeLine code, "    poRecipient.Enabled = Not locked"
    CodeLine code, "    poDept.Enabled = Not locked"
    CodeLine code, "    poCustomer.Locked = locked"
    CodeLine code, "    poPatient.Locked = locked"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poAligners_Change()"
    CodeLine code, "    If Not mLoading And Not mReminder Then UpdatePreview"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poTotal_Change()"
    CodeLine code, "    If Not mLoading And Not mReminder Then UpdatePreview"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poStart_Change()"
    CodeLine code, "    If Not mLoading And Not mReminder Then UpdatePreview"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poPlanID_Change()"
    CodeLine code, "    Dim ws As Worksheet, i As Long"
    CodeLine code, "    If mLoading Or mReminder Then Exit Sub"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    poSource.Enabled = True"
    CodeLine code, "    poGenerate.Enabled = True"
    CodeLine code, "    If Trim(CStr(poPlanID.Value)) = """" Then"
    CodeLine code, "        SetSourceLocks (Trim(CStr(poSource.Value)) <> """")"
    CodeLine code, "        Exit Sub"
    CodeLine code, "    End If"
    CodeLine code, "    Set ws = modPaymentOptions.EnsurePaymentPlansSheet()"
    CodeLine code, "    For i = 2 To ws.Cells(ws.Rows.Count, PP_PLANID).End(xlUp).Row"
    CodeLine code, "        If modHelpers.NrmID(CStr(ws.Cells(i, PP_PLANID).Value)) = modHelpers.NrmID(CStr(poPlanID.Value)) Then"
    CodeLine code, "            mLoading = True"
    CodeLine code, "            poSource.Value = ws.Cells(i, PP_SOURCE).Value"
    CodeLine code, "            poAligners.Value = ws.Cells(i, PP_ALIGNERS).Value"
    CodeLine code, "            poTotal.Value = ws.Cells(i, PP_TOTAL).Value"
    CodeLine code, "            SetRecipient CStr(ws.Cells(i, PP_RECIP).Value)"
    CodeLine code, "            poCustomer.Value = ws.Cells(i, PP_CUST).Value"
    CodeLine code, "            poPatient.Value = ws.Cells(i, PP_PATIENT).Value"
    CodeLine code, "            poDept.Value = UCase$(Trim$(CStr(ws.Cells(i, PP_DEPT).Value)))"
    CodeLine code, "            poStart.Value = Format(ws.Cells(i, PP_DUE).Value, ""yyyy-mm-dd"")"
    CodeLine code, "            SetSourceLocks True"
    CodeLine code, "            poTotal.Locked = (Trim(CStr(poSource.Value)) <> """")"
    CodeLine code, "            poSource.Enabled = False"
    CodeLine code, "            poGenerate.Enabled = False"
    CodeLine code, "            mLoading = False"
    CodeLine code, "            UpdatePreview"
    CodeLine code, "            Exit Sub"
    CodeLine code, "        End If"
    CodeLine code, "    Next i"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    mLoading = False"
    CodeLine code, "    MsgBox ""Plan lookup error: "" & Err.Description, vbExclamation"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub UpdatePreview()"
    CodeLine code, "    Dim p As PaymentPlanResult, i As Long, startD As Date, schedule As String"
    CodeLine code, "    On Error GoTo Invalid"
    CodeLine code, "    poPreview.Value = """""
    CodeLine code, "    If Not IsDate(poStart.Value) Then Err.Raise 5, , ""Enter a valid deposit due date."""
    CodeLine code, "    startD = CDate(poStart.Value)"
    CodeLine code, "    p = modPaymentOptions.CalcPaymentPlan(modHelpers.Num(poAligners.Value), modHelpers.Num(poTotal.Value))"
    CodeLine code, "    poPlan.Value = Format(modPaymentOptions.GetDepositPercent(), ""0.##%"") & "" deposit; "" & _"
    CodeLine code, "        Format(p.TreatmentWeeks, ""0.##"") & "" treatment weeks; "" & p.InstallmentCount & _"
    CodeLine code, "        "" months; monthly R"" & Format(p.MonthlyAmount, ""#,##0.00"") & "" (final month adjusted)."""
    CodeLine code, "    For i = 0 To p.InstallmentCount"
    CodeLine code, "        schedule = schedule & IIf(i = 0, ""Deposit"", ""Month "" & i) & vbTab & _"
    CodeLine code, "            Format(DateAdd(""m"", i, startD), ""yyyy-mm-dd"") & vbTab & _"
    CodeLine code, "            ""R "" & Format(p.Schedule(i), ""#,##0.00"") & vbCrLf"
    CodeLine code, "    Next i"
    CodeLine code, "    poPreview.Value = schedule"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Invalid:"
    CodeLine code, "    poPlan.Value = ""Preview: "" & Err.Description"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poGenerate_Click()"
    CodeLine code, "    Dim id As String"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    If Not IsDate(poStart.Value) Then Err.Raise 5, , ""Enter a valid deposit due date."""
    CodeLine code, "    If MsgBox(""Generate the deposit and all installment invoices? An unpaid source invoice will be replaced, not charged twice."", _"
    CodeLine code, "        vbQuestion + vbYesNo, ""Payment Options"") <> vbYes Then Exit Sub"
    CodeLine code, "    id = modPaymentOptions.GeneratePlanInvoices(CStr(poSource.Value), modHelpers.Num(poAligners.Value), _"
    CodeLine code, "        modHelpers.Num(poTotal.Value), RecipientCode(), CStr(poCustomer.Value), _"
    CodeLine code, "        CStr(poPatient.Value), CStr(poDept.Value), CDate(poStart.Value))"
    CodeLine code, "    If id <> """" Then poPlanID.Value = id: MsgBox ""Plan created: "" & id, vbInformation"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    MsgBox ""Generate plan error: "" & Err.Description, vbExclamation"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poCancelPlan_Click()"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    If Trim(CStr(poPlanID.Value)) = """" Then MsgBox ""Enter the Plan ID."", vbExclamation: Exit Sub"
    CodeLine code, "    If MsgBox(""Cancel this plan and void its unpaid, unsent invoices?"", vbQuestion + vbYesNo) = vbYes Then"
    CodeLine code, "        If modPaymentOptions.CancelPaymentPlan(CStr(poPlanID.Value)) Then MsgBox ""Plan cancelled; audit records retained."", vbInformation"
    CodeLine code, "    End If"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    MsgBox ""Cancel plan error: "" & Err.Description, vbExclamation"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poAmend_Click()"
    CodeLine code, "    Dim id As String"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    If Trim(CStr(poPlanID.Value)) = """" Then MsgBox ""Enter the Plan ID."", vbExclamation: Exit Sub"
    CodeLine code, "    If MsgBox(""Replace this unpaid, unsent plan using the current total and aligner count?"", _"
    CodeLine code, "        vbQuestion + vbYesNo) <> vbYes Then Exit Sub"
    CodeLine code, "    id = modPaymentOptions.AmendPaymentPlan(CStr(poPlanID.Value), modHelpers.Num(poAligners.Value), modHelpers.Num(poTotal.Value))"
    CodeLine code, "    If id <> """" Then poPlanID.Value = id: MsgBox ""Replacement plan: "" & id, vbInformation"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    MsgBox ""Amend plan error: "" & Err.Description, vbExclamation"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poDeposits_Click()"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    Me.Hide"
    CodeLine code, "    modPaymentOptions.RunDepositsStatement"
    CodeLine code, "    Me.Show"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    MsgBox ""Deposits statement error: "" & Err.Description, vbExclamation"
    CodeLine code, "    Me.Show"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poInstallments_Click()"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    Me.Hide"
    CodeLine code, "    modPaymentOptions.RunInstallmentsStatement"
    CodeLine code, "    Me.Show"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    MsgBox ""Installments statement error: "" & Err.Description, vbExclamation"
    CodeLine code, "    Me.Show"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poSingle_Click()"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    Me.Hide"
    CodeLine code, "    modPaymentOptions.RunSinglePlanStatement CStr(poPlanID.Value)"
    CodeLine code, "    Me.Show"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    MsgBox ""Plan statement error: "" & Err.Description, vbExclamation"
    CodeLine code, "    Me.Show"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Public Sub ShowReminder()"
    CodeLine code, "    Dim pr As Variant, ws As Worksheet, ctl As Object"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    Set mDueRows = modPaymentOptions.DueInstallmentRows()"
    CodeLine code, "    If mDueRows.Count = 0 Then Exit Sub"
    CodeLine code, "    mReminder = True"
    CodeLine code, "    Me.Caption = ""Installment invoices - Mail now"""
    CodeLine code, "    Set ws = ThisWorkbook.Worksheets(""PaymentPlans"")"
    CodeLine code, "    For Each ctl In Me.Controls"
    CodeLine code, "        ctl.Visible = False"
    CodeLine code, "    Next ctl"
    CodeLine code, "    poPlan.Visible = True: poSchedule.Visible = True: poMail.Visible = True: poClose.Visible = True"
    CodeLine code, "    poPlan.Top = 12"
    CodeLine code, "    poPlan.Value = ""Unpaid, unsent installments due soon or overdue. Select an invoice and click Mail Installment."""
    CodeLine code, "    poSchedule.Top = 55: poSchedule.Height = 299: poSchedule.Clear"
    CodeLine code, "    For Each pr In mDueRows"
    CodeLine code, "        poSchedule.AddItem ws.Cells(CLng(pr), PP_INVOICE).Value"
    CodeLine code, "        poSchedule.List(poSchedule.ListCount - 1, 1) = Format(ws.Cells(CLng(pr), PP_DUE).Value, ""yyyy-mm-dd"")"
    CodeLine code, "        poSchedule.List(poSchedule.ListCount - 1, 2) = ""R "" & Format(ws.Cells(CLng(pr), PP_AMOUNT).Value, ""#,##0.00"")"
    CodeLine code, "        poSchedule.List(poSchedule.ListCount - 1, 3) = ws.Cells(CLng(pr), PP_PATIENT).Value"
    CodeLine code, "    Next pr"
    CodeLine code, "    Me.Show"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    MsgBox ""Installment reminder error: "" & Err.Description, vbExclamation"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poMail_Click()"
    CodeLine code, "    Dim i As Long"
    CodeLine code, "    On Error GoTo Failed"
    CodeLine code, "    If Not mReminder Then"
    CodeLine code, "        Me.Hide"
    CodeLine code, "        modPaymentOptions.ShowInstallmentReminder"
    CodeLine code, "        Me.Show"
    CodeLine code, "        Exit Sub"
    CodeLine code, "    End If"
    CodeLine code, "    i = poSchedule.ListIndex"
    CodeLine code, "    If i < 0 Then MsgBox ""Select an invoice to mail."", vbInformation: Exit Sub"
    CodeLine code, "    If modPaymentOptions.MailInstallmentInvoice(CLng(mDueRows(i + 1))) Then"
    CodeLine code, "        poSchedule.RemoveItem i"
    CodeLine code, "        mDueRows.Remove i + 1"
    CodeLine code, "        If mDueRows.Count = 0 Then Unload Me"
    CodeLine code, "    End If"
    CodeLine code, "    Exit Sub"
    CodeLine code, "Failed:"
    CodeLine code, "    MsgBox ""Mail installment error: "" & Err.Description, vbExclamation"
    CodeLine code, "    If Not mReminder Then Me.Show"
    CodeLine code, "End Sub"
    CodeLine code, ""
    CodeLine code, "Private Sub poClose_Click()"
    CodeLine code, "    Unload Me"
    CodeLine code, "End Sub"
End Sub
