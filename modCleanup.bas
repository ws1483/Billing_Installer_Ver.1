Attribute VB_Name = "modCleanup"
' modCleanup ? integrity sweep for line sheets vs their logs.
'   CleanOrphanLines: for each (log, lines) pair, delete every LINE row whose
'     doc number does NOT exist in the LOG (e.g. INV-...-313 in InvoiceLines but
'     absent from InvoiceLog). Dry-run count + confirm, then delete + audit.
'   Also REPORTS (does not delete) any LOG doc that has zero line rows.
'   Covers: InvoiceLog/InvoiceLines, QuoteLog/QuoteLines,
'           MedAidLog/MedAidLines, CreditNotes/CreditNoteLines.
' ============================================================================
Option Explicit

Private Type Pair
    logName As String
    linesName As String
    label As String
End Type

Public Sub CleanOrphanLines()
    Dim pairs(1 To 4) As Pair, p As Long
    pairs(1).logName = "InvoiceLog":  pairs(1).linesName = "InvoiceLines":    pairs(1).label = "Invoice"
    pairs(2).logName = "QuoteLog":    pairs(2).linesName = "QuoteLines":      pairs(2).label = "Quote"
    pairs(3).logName = "MedAidLog":   pairs(3).linesName = "MedAidLines":     pairs(3).label = "Med Claim"
    pairs(4).logName = "CreditNotes": pairs(4).linesName = "CreditNoteLines": pairs(4).label = "Credit Note"

    Dim totalOrphans As Long, detail As String, emptyDocs As String
    Dim orphanRows As Collection

    ' ---- PASS 1: count (dry run) ----
    For p = 1 To 4
        Dim c As Long
        c = CountOrphans(pairs(p).logName, pairs(p).linesName)
        If c > 0 Then
            totalOrphans = totalOrphans + c
            detail = detail & "  - " & pairs(p).label & ": " & c & " orphan line row(s)" & vbCrLf
        End If
        Dim ed As String
        ed = LogsWithoutLines(pairs(p).logName, pairs(p).linesName)
        If ed <> "" Then emptyDocs = emptyDocs & "  [" & pairs(p).label & "] " & ed & vbCrLf
    Next p

    If totalOrphans = 0 Then
        Dim m As String
        m = "No orphan line rows found. Line sheets are consistent with their logs."
        If emptyDocs <> "" Then _
            m = m & vbCrLf & vbCrLf & "Note ? logs with NO line rows (not deleted):" & vbCrLf & emptyDocs
        MsgBox m, vbInformation, "Cleanup ? nothing to do"
        Exit Sub
    End If

    Dim ask As String
    ask = "Found " & totalOrphans & " orphan line row(s) whose document is NOT in the log:" & _
          vbCrLf & vbCrLf & detail & vbCrLf & "Delete them now?"
    If emptyDocs <> "" Then _
        ask = ask & vbCrLf & vbCrLf & "(FYI ? logs with NO line rows, NOT deleted:" & vbCrLf & emptyDocs & ")"
    If MsgBox(ask, vbExclamation + vbYesNo, "Confirm orphan cleanup") <> vbYes Then Exit Sub

    ' ---- PASS 2: delete ----
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim deleted As Long
    For p = 1 To 4
        deleted = deleted + DeleteOrphans(pairs(p).logName, pairs(p).linesName, pairs(p).label)
    Next p

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    LogAudit "Cleanup", "(lines)", totalOrphans & " found", deleted & " deleted", "Orphan line cleanup"
    MsgBox deleted & " orphan line row(s) deleted.", vbInformation, "Cleanup complete"
    Exit Sub
Fail:
    Application.EnableEvents = True: Application.ScreenUpdating = True
    MsgBox "CleanOrphanLines error: " & Err.Description, vbExclamation
End Sub

' ---- count orphan line rows (doc number absent from the log) ----
Private Function CountOrphans(logName As String, linesName As String) As Long
    Dim wsLog As Worksheet, wsLines As Worksheet
    Dim keys As Object, last As Long, i As Long, docNo As String, cnt As Long
    If Not SheetExists(logName) Or Not SheetExists(linesName) Then Exit Function
    Set wsLog = ThisWorkbook.Sheets(logName)
    Set wsLines = ThisWorkbook.Sheets(linesName)
    Set keys = LogKeySet(wsLog)

    last = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        docNo = NrmID(CStr(wsLines.Cells(i, LN_DOCNO).value))
        If docNo <> "" Then
            If Not keys.Exists(docNo) Then cnt = cnt + 1
        End If
    Next i
    CountOrphans = cnt
End Function

' ---- delete orphan line rows; returns count deleted ----
Private Function DeleteOrphans(logName As String, linesName As String, label As String) As Long
    Dim wsLog As Worksheet, wsLines As Worksheet
    Dim keys As Object, last As Long, i As Long, docNo As String, del As Long
    If Not SheetExists(logName) Or Not SheetExists(linesName) Then Exit Function
    Set wsLog = ThisWorkbook.Sheets(logName)
    Set wsLines = ThisWorkbook.Sheets(linesName)
    Set keys = LogKeySet(wsLog)

    last = wsLines.Cells(wsLines.Rows.Count, "A").End(xlUp).row
    For i = last To 2 Step -1
        docNo = NrmID(CStr(wsLines.Cells(i, LN_DOCNO).value))
        If docNo <> "" Then
            If Not keys.Exists(docNo) Then
                wsLines.Rows(i).Delete
                del = del + 1
            End If
        End If
    Next i
    DeleteOrphans = del
End Function

' ---- report logs whose doc number has NO line rows (comma list) ----
Private Function LogsWithoutLines(logName As String, linesName As String) As String
    Dim wsLog As Worksheet, wsLines As Worksheet
    Dim lineKeys As Object, last As Long, i As Long, docNo As String, out As String
    If Not SheetExists(logName) Or Not SheetExists(linesName) Then Exit Function
    Set wsLog = ThisWorkbook.Sheets(logName)
    Set wsLines = ThisWorkbook.Sheets(linesName)
    Set lineKeys = LogKeySet(wsLines)   ' reuse: keys from col A of lines

    last = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        docNo = NrmID(CStr(wsLog.Cells(i, "A").value))
        If docNo <> "" Then
            If Not lineKeys.Exists(docNo) Then _
                out = out & CStr(wsLog.Cells(i, "A").value) & ", "
        End If
    Next i
    If Len(out) > 2 Then out = Left$(out, Len(out) - 2)
    LogsWithoutLines = out
End Function

' ---- set of normalized doc numbers from column A of a sheet ----
Private Function LogKeySet(ws As Worksheet) As Object
    Dim d As Object, last As Long, i As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    last = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To last
        k = NrmID(CStr(ws.Cells(i, "A").value))
        If k <> "" Then If Not d.Exists(k) Then d.Add k, 1
    Next i
    Set LogKeySet = d
End Function

Private Function SheetExists(nm As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not (s Is Nothing)
End Function

