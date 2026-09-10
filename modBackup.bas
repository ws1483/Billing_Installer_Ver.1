Attribute VB_Name = "modBackup"
Option Explicit
' ============================================================================
' modBackup ? SaveCopyAs a timestamped copy (OneDrive-safe), keep newest 3.
'   PHASE 0: backup folder is read from Settings!B21 if present; otherwise it
'   falls back to the original hard-coded path. Navigation helpers unchanged.
' ============================================================================
Private Const BACKUP_DIR_FALLBACK As String = _
    "F:\One Drive\OneDrive\Documents\GreydataDental\Accounting\WeDental Billing\Backup"
Private Const KEEP_COUNT As Long = 3

' Resolve the backup folder: Settings!B21 override, else fallback constant.
Private Function BackupDir() As String
    Dim s As String
    On Error Resume Next
    s = Trim$(CStr(ThisWorkbook.Sheets("Settings").Range("B21").value))
    On Error GoTo 0
    If s = "" Then s = BACKUP_DIR_FALLBACK
    BackupDir = s
End Function

Public Sub RunBackup()
    Dim stamp As String, destPath As String, dir As String
    Dim fso As Object

    On Error GoTo Fail
    dir = BackupDir()
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(dir) Then
        fso.CreateFolder dir
        If Not fso.FolderExists(dir) Then
            MsgBox "Backup folder not found/creatable:" & vbCrLf & dir, _
                   vbExclamation, "Backup"
            Exit Sub
        End If
    End If

    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

    stamp = Format(Now, "yyyy-mm-dd_hhnn")
    destPath = dir & "\" & "WeDental_Backup_" & stamp & ".xlsm"

    ThisWorkbook.SaveCopyAs destPath
    PruneBackups fso, dir

    MsgBox "Backup created:" & vbCrLf & destPath & vbCrLf & vbCrLf & _
           "Keeping the latest " & KEEP_COUNT & " backups.", vbInformation, "Backup"
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    MsgBox "Backup error: " & Err.Description & vbCrLf & vbCrLf & _
           "Target: " & destPath, vbExclamation, "Backup"
End Sub

Private Sub PruneBackups(fso As Object, ByVal dir As String)
    Dim fld As Object, f As Object
    Dim names() As String, dates() As Double, n As Long, i As Long, j As Long
    Dim tS As String, tD As Double

    Set fld = fso.GetFolder(dir)

    n = 0
    ReDim names(1 To 1000)
    ReDim dates(1 To 1000)
    For Each f In fld.Files
        If LCase$(f.Name) Like "wedental_backup_*.xlsm" Then
            n = n + 1
            If n > UBound(names) Then
                ReDim Preserve names(1 To UBound(names) + 1000)
                ReDim Preserve dates(1 To UBound(dates) + 1000)
            End If
            names(n) = f.Path
            dates(n) = CDbl(f.DateLastModified)
        End If
    Next f
    If n <= KEEP_COUNT Then Exit Sub

    For i = 1 To n - 1
        For j = 1 To n - i
            If dates(j) < dates(j + 1) Then
                tD = dates(j): dates(j) = dates(j + 1): dates(j + 1) = tD
                tS = names(j): names(j) = names(j + 1): names(j + 1) = tS
            End If
        Next j
    Next i

    For i = KEEP_COUNT + 1 To n
        On Error Resume Next
        fso.DeleteFile names(i), True
        On Error GoTo 0
    Next i
End Sub

' ---- One-click "reset my view" safety button ----
Sub ShowRibbon()
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    On Error GoTo 0
    MsgBox "Full view restored.", vbInformation
End Sub

' ===== Navigation =====
Public Sub GoToMenu()
    NavTo "Menu"
End Sub
Public Sub GoToQuote()
    NavTo "Quote"
End Sub
Public Sub GoToInvoice()
    NavTo "Invoice"
End Sub

Private Sub NavTo(sheetName As String)
    Dim ws As Worksheet
    On Error GoTo Fail
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Activate
    ws.Range("A1").Select
    Exit Sub
Fail:
    MsgBox "Cannot navigate to '" & sheetName & "'." & vbCrLf & Err.Description, vbExclamation
End Sub
Public Sub ProbeMenu()
    Dim wsM As Worksheet
    On Error Resume Next
    Set wsM = ThisWorkbook.Sheets("Menu")
    On Error GoTo 0
    If wsM Is Nothing Then
        MsgBox "No sheet named exactly 'Menu' found. Check the tab name / trailing spaces.", vbCritical
        Exit Sub
    End If
    MsgBox "Menu sheet found." & vbCrLf & _
           "G17 (showType) = [" & wsM.Range("G17").value & "]" & vbCrLf & _
           "C5 (dept)      = [" & wsM.Range("C5").value & "]" & vbCrLf & _
           "B17 (from)     = [" & wsM.Range("B17").value & "]" & vbCrLf & _
           "C17 (to)       = [" & wsM.Range("C17").value & "]" & vbCrLf & _
           "QuoteLog rows  = " & ThisWorkbook.Sheets("QuoteLog").Cells(Rows.Count, "A").End(xlUp).row & vbCrLf & _
           "InvoiceLog rows= " & ThisWorkbook.Sheets("InvoiceLog").Cells(Rows.Count, "A").End(xlUp).row
End Sub
Public Sub DedupeInvoiceLines()
    Dim ws As Worksheet, last As Long, i As Long
    Dim seen As Object, key As String, removed As Long
    Set seen = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("InvoiceLines")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    Application.ScreenUpdating = False
    For i = last To 2 Step -1
        key = UCase$(Replace(Trim$(CStr(ws.Cells(i, 1).value)), " ", "")) & "|" & _
              CStr(ws.Cells(i, 2).value) & "|" & _
              UCase$(Trim$(CStr(ws.Cells(i, 5).value)))     ' DocNo|LineNo|Desc
        If seen.Exists(key) Then
            ws.Rows(i).Delete
            removed = removed + 1
        Else
            seen.Add key, 1
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox "Removed " & removed & " duplicate line row(s) from InvoiceLines.", vbInformation
End Sub

