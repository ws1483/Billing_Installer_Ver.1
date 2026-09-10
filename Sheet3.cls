VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim isect As Range, c As Range

    ' recipient toggle
    If Not Application.Intersect(Target, Me.Range("K2")) Is Nothing Then
        SetRecipientType Me, CStr(Me.Range("K2").value)
        Exit Sub
    End If

    ' patient-mode: Doctor picked (G12) -> auto-fill BHF (G13)
    If Not Application.Intersect(Target, Me.Range("G12")) Is Nothing Then
        If LCase(Trim(CStr(Me.Range("K2").value))) = "patient" Then
            Application.EnableEvents = False
            Me.Range("G13").value = LookupBHFByDrName(CStr(Me.Range("G12").value))
            Application.EnableEvents = True
        End If
        Exit Sub
    End If

    ' price auto-fill
    Set isect = Application.Intersect(Target, Me.Range("D16:D30"))
    If isect Is Nothing Then Exit Sub
    Application.EnableEvents = False
    On Error GoTo Clean
    For Each c In isect.Cells
        FillLinePrice Me, c.row
    Next c
Clean:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' date picker on document date G6
    If Not Application.Intersect(Target, Me.Range("G6")) Is Nothing Then
        Cancel = True
        Dim d As Variant
        d = PickDate(Target.value)
        If IsDate(d) Then
            Application.EnableEvents = False
            Target.value = d
            Application.EnableEvents = True
        End If
    End If
End Sub

