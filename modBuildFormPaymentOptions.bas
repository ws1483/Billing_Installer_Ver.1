Attribute VB_Name = "modBuildFormPaymentOptions"
Option Explicit

' Import optionally; run on a workbook copy using the trusted source export.
Public Sub BuildFormPaymentOptions()
    Dim sourcePath As Variant, sourceCode As String
    Dim project As Object, component As Object, newForm As Object
    Dim failure As String, stage As String
    On Error GoTo Fail

    stage = "selecting the form export"
    sourcePath = Application.GetOpenFilename("VBA UserForms (*.frm), *.frm", , _
        "Select the updated frmPaymentOptions.frm from this repository")
    If VarType(sourcePath) = vbBoolean Then Exit Sub
    stage = "reading the form export"
    sourceCode = ReadPaymentOptionsCode(CStr(sourcePath))

    stage = "accessing the VBA project"
    Set project = ThisWorkbook.VBProject
    For Each component In project.VBComponents
        If StrComp(component.Name, "frmPaymentOptions", vbTextCompare) = 0 Then
            MsgBox "frmPaymentOptions already exists. Export it as a backup, then remove it " & _
                "in the VBE before running this builder. No changes were made.", vbInformation
            Exit Sub
        End If
    Next component

    stage = "creating the UserForm"
    Set newForm = project.VBComponents.Add(3) ' vbext_ct_MSForm; no VBIDE reference required.
    newForm.Name = "frmPaymentOptions"
    stage = "setting the designer layout"
    With newForm.Designer
        .Caption = "Payment Options"
        .Width = 650
        .Height = 550
    End With
    stage = "setting the form startup position"
    newForm.Properties("StartUpPosition").Value = 1
    stage = "loading the form code"
    With newForm.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString sourceCode
    End With
    MsgBox "frmPaymentOptions created with an empty designer and the selected form's code. " & _
        "Run Debug > Compile VBAProject, then test Payment Options on this workbook copy.", vbInformation
    Exit Sub

Fail:
    failure = "Error " & CStr(Err.Number) & " while " & stage & ": " & Err.Description
    On Error Resume Next
    If Not newForm Is Nothing Then project.VBComponents.Remove newForm
    On Error GoTo 0
    MsgBox "Could not build Payment Options: " & failure & vbCrLf & _
        "Use an unlocked workbook copy and enable Trust access to the VBA project object model " & _
        "only while building from trusted source. Existing components are not replaced.", vbExclamation
End Sub

Private Function ReadPaymentOptionsCode(sourcePath As String) As String
    Dim fileNo As Integer, fileOpen As Boolean, lineText As String
    Dim sourceText As String, lines As Variant, i As Long
    Dim formNameFound As Boolean, codeStarted As Boolean, sourceCode As String
    Dim errorNumber As Long, errorText As String
    On Error GoTo Fail

    fileNo = FreeFile
    Open sourcePath For Binary Access Read As #fileNo
    fileOpen = True
    sourceText = Space$(LOF(fileNo))
    If Len(sourceText) > 0 Then Get #fileNo, , sourceText
    Close #fileNo
    fileOpen = False
    sourceText = Replace(sourceText, vbCrLf, vbLf)
    sourceText = Replace(sourceText, vbCr, vbLf)
    lines = Split(sourceText, vbLf)
    For i = 0 To UBound(lines)
        lineText = CStr(lines(i))
        If i = UBound(lines) And lineText = "" Then Exit For
        If Trim$(lineText) = "Attribute VB_Name = ""frmPaymentOptions""" Then formNameFound = True
        If Trim$(lineText) = "Option Explicit" Then codeStarted = True
        If codeStarted And Left$(LTrim$(lineText), 10) <> "Attribute " Then
            sourceCode = sourceCode & lineText & vbCrLf
        End If
    Next i
    If Not formNameFound Or Not codeStarted Then
        Err.Raise vbObjectError + 2100, "BuildFormPaymentOptions", _
            "Select the repository's exported frmPaymentOptions.frm, not a different form or .frx."
    End If
    ReadPaymentOptionsCode = sourceCode
    Exit Function

Fail:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If fileOpen Then Close #fileNo
    On Error GoTo 0
    Err.Raise errorNumber, "BuildFormPaymentOptions", errorText
End Function
