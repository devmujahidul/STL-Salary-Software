from __future__ import annotations

from pathlib import Path

from .constants import ALL_LABEL, SUBSECTION_PREFIX

VBA_THISWORKBOOK = f"""
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    ApplyFooterSignatures
    ' Ctrl+Shift+T opens the duty-transfer form (Alt+F8 also works).
    Application.OnKey "^+t", "ShowDutyTransferForm"
End Sub

' Excel prints hidden columns - so before any print (or print preview) we
' hide the Designation column on the attendance sheets, and a queued macro
' un-hides it as soon as the print job finishes. On screen the column stays
' visible; it only disappears on paper.
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    On Error Resume Next
    HideDesignationForPrint
    Application.OnTime Now, _
        "'" & ThisWorkbook.Name & "'!RestoreDesignationColumns"
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo SafeExit
    If Sh.Name = "Lists" Or Sh.Name = "_Lists" Then Exit Sub
    If Intersect(Target, Sh.Range("A4")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    ApplySubSectionFilter Sh
SafeExit:
    Application.EnableEvents = True
End Sub
"""

VBA_MODULE = f"""
Option Explicit

' Module-level state for the Designation print-hide feature. VBA only accepts
' module-level declarations here, at the top of the module - declaring this
' between procedures fails to compile ("Variable not defined").
Private DesignationState As Collection

Public Sub ApplySubSectionFilter(ByVal ws As Worksheet)
    Dim raw As String
    Dim chosen As String
    Dim lastCol As Long
    Dim lastRow As Long
    Dim filterCol As Long
    Dim parts As Variant
    Dim i As Long
    Dim formulaText As String
    Dim rng As Range

    raw = CStr(ws.Range("A4").Value)
    chosen = raw
    If InStr(1, chosen, "{SUBSECTION_PREFIX}", vbTextCompare) = 1 Then
        chosen = Trim$(Mid$(chosen, Len("{SUBSECTION_PREFIX}") + 1))
    End If

    lastCol = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 6 Then Exit Sub
    filterCol = lastCol
    Set rng = ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, lastCol))

    If chosen = "" Or chosen = "{ALL_LABEL}" Then
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        rng.AutoFilter
    ElseIf InStr(chosen, "+") > 0 Then
        parts = Split(chosen, "+")
        For i = LBound(parts) To UBound(parts)
            parts(i) = Trim$(CStr(parts(i)))
        Next i
        rng.AutoFilter Field:=filterCol, Criteria1:=parts, Operator:=xlFilterValues
    Else
        rng.AutoFilter Field:=filterCol, Criteria1:=chosen
    End If

    ' Put the live "Sub Section: ..." formula back into A4. The normal
    ' (non-array) formula recalculates whenever the filter changes.
    formulaText = LabelFormula(ws)
    On Error Resume Next
    If Len(formulaText) > 0 Then
        ws.Range("A4").Formula = formulaText
        ws.Range("A4").Calculate
    Else
        ws.Range("A4").Value = "{SUBSECTION_PREFIX}" & chosen
    End If
    On Error GoTo 0
End Sub

Private Function LabelFormula(ByVal ws As Worksheet) As String
    Dim nm As Name
    Dim text As String

    On Error Resume Next
    Set nm = ws.Parent.Names(ws.Name & "_LabelFormula")
    If Not nm Is Nothing Then
        text = CStr(nm.RefersToRange.Value)
    End If
    On Error GoTo 0
    LabelFormula = text
End Function

' ---- print support -------------------------------------------------------

' The file hides the Designation column by default; this stores each sheet's
' pre-print state so RestoreDesignationColumns puts it back exactly as it was
' (hidden by default, or visible if the user un-hid it on purpose).
' (The DesignationState declaration itself lives at the top of the module.)

' Hide the Designation column (D) on the attendance sheets. Excel does not
' print hidden columns, so the printout drops Designation.
Public Sub HideDesignationForPrint()
    Set DesignationState = New Collection
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Lists" Then
            DesignationState.Add ws.Columns("D").Hidden, ws.Name
            ws.Columns("D").Hidden = True
        End If
    Next ws
End Sub

' Restore the pre-print hidden/visible state for every sheet.
Public Sub RestoreDesignationColumns()
    If DesignationState Is Nothing Then
        ' No remembered state (e.g. run manually): keep the file default,
        ' which is hidden.
        Dim wsDefault As Worksheet
        For Each wsDefault In ThisWorkbook.Worksheets
            If wsDefault.Name <> "Lists" Then wsDefault.Columns("D").Hidden = True
        Next wsDefault
        Exit Sub
    End If
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Lists" Then
            Dim wasHidden As Variant
            On Error Resume Next
            wasHidden = DesignationState(ws.Name)
            On Error GoTo 0
            If Not IsEmpty(wasHidden) Then ws.Columns("D").Hidden = CBool(wasHidden)
        End If
    Next ws
    Set DesignationState = Nothing
End Sub

' ---------------------------------------------------------------------------

' The footer signature names live on the Lists sheet (column A, from row 2).
' Editing the names there and rerunning this macro refreshes the signature
' footer on every attendance sheet without touching the code.
Public Sub ApplyFooterSignatures()
    Dim src As Worksheet, ws As Worksheet
    Dim lastR As Long, n As Long, i As Long

    On Error Resume Next
    Set src = ThisWorkbook.Worksheets("Lists")
    On Error GoTo 0
    If src Is Nothing Then Exit Sub

    lastR = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    n = lastR - 1                 ' names start on row 2
    If n < 1 Then Exit Sub

    Dim zoneL As String, zoneC As String, zoneR As String
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Lists" Then
            ws.PageSetup.LeftFooter = ZoneText(src, n, 1)
            ws.PageSetup.CenterFooter = ZoneText(src, n, 2)
            ws.PageSetup.RightFooter = ZoneText(src, n, 3)
        End If
    Next ws
End Sub

' Zone 1/2/3 = left/center/right footer. Names 4+ are appended to the right
' zone. Each name prints as 'Name: ____' in 9pt like the Excel-side default.
Private Function ZoneText(ByVal src As Worksheet, ByVal n As Long, ByVal zoneIdx As Long) As String
    Dim slot As Long
    Dim txt As String
    If zoneIdx < n Then
        ' 4+ names: the right zone also carries every extra name.
        If zoneIdx = 3 Then
            slot = zoneIdx
            txt = FooterNameText(src, slot)
            slot = slot + 1
            Do While slot <= n
                txt = txt & "        " & FooterNameText(src, slot)
                slot = slot + 1
            Loop
            ZoneText = txt
        Else
            ZoneText = FooterNameText(src, zoneIdx)
        End If
    ElseIf zoneIdx <= n Then
        ZoneText = FooterNameText(src, zoneIdx)
    Else
        ZoneText = ""
    End If
End Function

Private Function FooterNameText(ByVal src As Worksheet, ByVal rowOffset As Long) As String
    Dim nm As String
    nm = Trim$(CStr(src.Cells(rowOffset + 1, 1).Value))
    If Len(nm) > 0 Then
        FooterNameText = "&9" & nm & ": " & String$(22, "_")
    Else
        FooterNameText = ""
    End If
End Function
"""


VBA_TRANSFER_MODULE = r'''
Option Explicit

' ===================== Duty transfer (per sub section) =====================
' One command moves every qualifying worker's daily duty from a source date
' to a target date, per the payroll rules:
'
'   * Sub section scope: chosen on the form (All = whole workbook).
'   * Source date: only PRESENT workers move - code P, or a worked weekend
'     (code W with hours on the totals row). Anyone absent (A) or on leave
'     is never touched.
'   * Target date: the cells must read A for that worker, otherwise the
'     worker is skipped and nothing is overwritten.
'   * The duty moves with its OT and EOT hours.
'   * Weekday source: after the move the source day becomes A with 0/0.
'   * Weekend (W) source with full-day hours H: the target weekday keeps
'     8 regular hours, OT = H - 8 (max 2) and any hours beyond that go to
'     EOT. If H <= 8 no hours are lost: OT takes up to 2, the rest to EOT.
'     The source weekend keeps its W stamp but drops to 0/0.
' Every run is logged on a hidden sheet so "Undo last transfer" restores
' the previous state of the whole batch.

Private Const HISTORY_SHEET_NAME As String = "_TransferLog"

Public Sub ShowDutyTransferForm()
    frmDutyTransfer.Show 0
End Sub

Public Function ExecuteDutyTransfer( _
    ByVal sectionName As String, _
    ByVal fromDay As String, _
    ByVal toDay As String) As String

    Dim ws As Worksheet
    Dim r As Long, lastRow As Long, fCol As Long
    Dim fromCol As Long, toCol As Long
    Dim sec As String, srcCode As String, tgt As String
    Dim fullHours As Double, newOT As Double, newEOT As Double
    Dim moved As Long, srcWeekday As Long, srcWeekend As Long
    Dim skippedOccupied As Long, savedCells As Long
    Dim batchRow As Long

    sectionName = Trim$(sectionName)
    fromDay = Trim$(fromDay)
    toDay = Trim$(toDay)
    If fromDay = "" Or toDay = "" Then
        ExecuteDutyTransfer = "Pick both the source and the target date first."
        Exit Function
    End If
    If fromDay = toDay Then
        ExecuteDutyTransfer = "The source and the target dates are the same."
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Failed

    batchRow = 0
    For Each ws In ThisWorkbook.Worksheets
        If Not IsAttendanceSheet(ws) Then GoTo NextSheet
        If LastDataRow(ws) < 6 Then GoTo NextSheet
        fCol = FilterColumn(ws)
        lastRow = LastDataRow(ws)
        fromCol = FindDateColumn(ws, fromDay)
        toCol = FindDateColumn(ws, toDay)
        If fromCol = 0 Or toCol = 0 Then GoTo NextSheet

        r = 6
        Do While r <= lastRow
            sec = Trim$(CStr(ws.Cells(r, fCol).Value))
            If sectionName <> "All" And StrComp(sec, sectionName, vbTextCompare) <> 0 Then GoTo NextWorker

            srcCode = UCase$(Trim$(CStr(ws.Cells(r, fromCol).Value)))
            If srcCode = "W" Then
                If Val(ws.Cells(r + 2, fromCol).Value) <= 0 Then GoTo NextWorker
            ElseIf srcCode <> "P" Then
                GoTo NextWorker
            End If

            tgt = UCase$(Trim$(CStr(ws.Cells(r, toCol).Value)))
            If tgt <> "A" Then
                skippedOccupied = skippedOccupied + 1
                GoTo NextWorker
            End If

            If batchRow = 0 Then batchRow = StartHistoryBatch()
            savedCells = savedCells + SaveCellState(ws, r, fromCol)
            savedCells = savedCells + SaveCellState(ws, r + 1, fromCol)
            savedCells = savedCells + SaveCellState(ws, r + 2, fromCol)
            savedCells = savedCells + SaveCellState(ws, r, toCol)
            savedCells = savedCells + SaveCellState(ws, r + 1, toCol)
            savedCells = savedCells + SaveCellState(ws, r + 2, toCol)

            If srcCode = "W" Then
                fullHours = Val(ws.Cells(r + 2, fromCol).Value)
                If fullHours > 8 Then
                    newOT = fullHours - 8
                    If newOT > 2 Then
                        newEOT = (fullHours - 8) - 2
                        newOT = 2
                    Else
                        newEOT = 0
                    End If
                Else
                    newOT = fullHours
                    If newOT > 2 Then
                        newEOT = fullHours - 2
                        newOT = 2
                    Else
                        newEOT = 0
                    End If
                End If
                ws.Cells(r, toCol).Value = "P"
                ws.Cells(r + 1, toCol).Value = newOT
                ws.Cells(r + 2, toCol).Value = newEOT
                ws.Cells(r + 1, fromCol).Value = 0
                ws.Cells(r + 2, fromCol).Value = 0
                ' the weekend keeps its W stamp, hours drop to zero
                srcWeekend = srcWeekend + 1
            Else
                ws.Cells(r, toCol).Value = "P"
                ws.Cells(r + 1, toCol).Value = ws.Cells(r + 1, fromCol).Value
                ws.Cells(r + 2, toCol).Value = ws.Cells(r + 2, fromCol).Value
                ws.Cells(r, fromCol).Value = "A"
                ws.Cells(r + 1, fromCol).Value = 0
                ws.Cells(r + 2, fromCol).Value = 0
                srcWeekday = srcWeekday + 1
            End If
            moved = moved + 1
NextWorker:
            r = r + 3
        Loop
NextSheet:
    Next ws

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If batchRow > 0 And savedCells = 0 Then
        ' Nothing changed: drop the empty batch marker.
        On Error Resume Next
        HistorySheet().Rows(batchRow).Delete
        On Error GoTo 0
        batchRow = 0
    End If

    ExecuteDutyTransfer = CStr(moved) & " worker(s) moved from day " & fromDay & " to day " & toDay & _
        " (" & srcWeekday & " weekday, " & srcWeekend & " weekend)." & vbCrLf & _
        CStr(skippedOccupied) & " worker(s) skipped: target date is not A."
    Exit Function

Failed:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ExecuteDutyTransfer = "Transfer stopped: " & Err.Description
End Function

Public Function UndoLastTransfer() As String
    Dim hs As Worksheet, ws As Worksheet
    Dim lastRow As Long, batchRow As Long, r As Long, restored As Long

    Set hs = HistorySheet()
    lastRow = hs.Cells(hs.Rows.Count, 1).End(xlUp).Row
    batchRow = 0
    For r = lastRow To 1 Step -1
        If CStr(hs.Cells(r, 1).Value) = "BATCH" Then
            batchRow = r
            Exit For
        End If
    Next r
    If batchRow = 0 Then
        UndoLastTransfer = "Nothing to undo - no transfer has run yet."
        Exit Function
    End If

    Application.ScreenUpdating = False
    On Error GoTo Failed
    For r = batchRow + 1 To lastRow
        If CStr(hs.Cells(r, 1).Value) = "CELL" Then
            Set ws = ThisWorkbook.Worksheets(CStr(hs.Cells(r, 2).Value))
            ws.Range(CStr(hs.Cells(r, 3).Value)).Value = hs.Cells(r, 4).Value
            restored = restored + 1
        End If
    Next r
    hs.Rows(batchRow & ":" & lastRow).Delete
    Application.ScreenUpdating = True
    UndoLastTransfer = "Last transfer undone: " & restored & " cell(s) restored."
    Exit Function
Failed:
    Application.ScreenUpdating = True
    UndoLastTransfer = "Undo failed: " & Err.Description
End Function

Public Function HasUndo() As Boolean
    Dim hs As Worksheet
    Dim lastRow As Long, r As Long
    On Error GoTo NoUndo
    Set hs = HistorySheet()
    lastRow = hs.Cells(hs.Rows.Count, 1).End(xlUp).Row
    If lastRow > 0 Then
        For r = lastRow To 1 Step -1
            If CStr(hs.Cells(r, 1).Value) = "BATCH" Then
                If r < lastRow And CStr(hs.Cells(lastRow, 1).Value) = "CELL" Then HasUndo = True
                Exit For
            End If
        Next r
    End If
    Exit Function
NoUndo:
End Function

' ---------------- helpers ----------------

Private Function StartHistoryBatch() As Long
    Dim hs As Worksheet
    Set hs = HistorySheet()
    StartHistoryBatch = hs.Cells(hs.Rows.Count, 1).End(xlUp).Row + 1
    hs.Cells(StartHistoryBatch, 1).Value = "BATCH"
    hs.Cells(StartHistoryBatch, 2).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Private Function SaveCellState(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Long
    Dim hs As Worksheet
    Dim nxt As Long
    Set hs = HistorySheet()
    nxt = hs.Cells(hs.Rows.Count, 1).End(xlUp).Row + 1
    hs.Cells(nxt, 1).Value = "CELL"
    hs.Cells(nxt, 2).Value = ws.Name
    hs.Cells(nxt, 3).Value = ws.Cells(r, c).Address(False, False)
    hs.Cells(nxt, 4).Value = ws.Cells(r, c).Value
    SaveCellState = 1
End Function

Private Function HistorySheet() As Worksheet
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name = HISTORY_SHEET_NAME Then
            Set HistorySheet = sh
            Exit Function
        End If
    Next sh
    Set HistorySheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    HistorySheet.Name = HISTORY_SHEET_NAME
    HistorySheet.Visible = xlSheetVeryHidden
End Function

Private Function IsAttendanceSheet(ByVal ws As Worksheet) As Boolean
    Select Case ws.Name
        Case "Lists", "_Lists", HISTORY_SHEET_NAME
            IsAttendanceSheet = False
        Case Else
            IsAttendanceSheet = (CStr(ws.Range("A5").Value) = "SL")
    End Select
End Function

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    LastDataRow = ws.Cells(ws.Rows.Count, FilterColumn(ws)).End(xlUp).Row
End Function

' Row 5 headers run contiguously through the SubSection Filter column, so the
' first blank header cell sits one column past it.
Private Function FilterColumn(ByVal ws As Worksheet) As Long
    Dim c As Long
    For c = 5 To ws.Columns.Count
        If CStr(ws.Cells(5, c).Value) = "" Then Exit For
    Next c
    FilterColumn = c - 1
End Function

Private Function FindDateColumn(ByVal ws As Worksheet, ByVal dayText As String) As Long
    Dim c As Long
    Dim v As String
    For c = 5 To ws.Columns.Count
        v = CStr(ws.Cells(5, c).Value)
        If v = "" Then Exit For
        If StrComp(v, dayText, vbTextCompare) = 0 Then
            FindDateColumn = c
            Exit Function
        End If
    Next c
End Function

' ---------------- form fill helpers ----------------

Public Sub FillSectionCombo(ByVal frm As Object)
    Dim cbo As Object
    Dim ws As Worksheet
    Dim fCol As Long, listCol As Long, r As Long, lastRow As Long
    Dim name As String, picked As String, k As Long

    Set cbo = frm.Controls("cboSection")
    cbo.Clear
    cbo.AddItem "All"
    For Each ws In ThisWorkbook.Worksheets
        If IsAttendanceSheet(ws) Then
            fCol = FilterColumn(ws)
            listCol = fCol + 3
            lastRow = ws.Cells(ws.Rows.Count, listCol).End(xlUp).Row
            For r = 3 To lastRow            ' row 2 holds the "All" entry
                name = Trim$(CStr(ws.Cells(r, listCol).Value))
                If name <> "" And name <> "All" Then
                    If ComboIndex(cbo, name) < 0 Then cbo.AddItem name
                End If
            Next r
        End If
    Next ws
    ' Default: the sub section currently shown on the active sheet, else All.
    picked = ActiveSubSection()
    k = -1
    If picked <> "" Then k = ComboIndex(cbo, picked)
    If k < 0 Then k = 0
    On Error Resume Next
    cbo.ListIndex = k
    On Error GoTo 0
End Sub

Public Sub FillDateCombo(ByVal frm As Object)
    Dim ws As Worksheet
    Dim c As Long
    Dim v As String
    Set ws = ReferenceAttendanceSheet()
    For c = 5 To ws.Columns.Count
        v = CStr(ws.Cells(5, c).Value)
        If v = "" Then Exit For
        If IsNumeric(v) Then
            frm.Controls("cboFrom").AddItem v
            frm.Controls("cboTo").AddItem v
        End If
    Next c
End Sub

Private Function ReferenceAttendanceSheet() As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set ReferenceAttendanceSheet = ActiveSheet
    On Error GoTo 0
    If Not ReferenceAttendanceSheet Is Nothing Then
        If IsAttendanceSheet(ReferenceAttendanceSheet) Then Exit Function
    End If
    Set ReferenceAttendanceSheet = Nothing
    For Each sh In ThisWorkbook.Worksheets
        If IsAttendanceSheet(sh) Then
            Set ReferenceAttendanceSheet = sh
            Exit Function
        End If
    Next sh
End Function

Private Function ActiveSubSection() As String
    Dim ws As Worksheet
    Dim txt As String
    Set ws = ReferenceAttendanceSheet()
    If ws Is Nothing Then Exit Function
    txt = CStr(ws.Range("A4").Value)
    If InStr(1, txt, "Sub Section: ", vbTextCompare) = 1 Then
        txt = Trim$(Mid$(txt, Len("Sub Section: ") + 1))
    End If
    If txt = "All" Then txt = ""
    ActiveSubSection = txt
End Function

Private Function ComboIndex(ByVal cbo As Object, ByVal text As String) As Long
    Dim i As Long
    ComboIndex = -1
    For i = 0 To cbo.ListCount - 1
        If StrComp(CStr(cbo.List(i)), text, vbTextCompare) = 0 Then
            ComboIndex = i
            Exit Function
        End If
    Next i
End Function
'''

VBA_FORM_CODE = r'''
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo Fail
    lblHint.Caption = _
        "Move each qualifying worker's duty from the source date to the " & vbCrLf & _
        "target date. Rules:" & vbCrLf & vbCrLf & _
        "- Only workers PRESENT on the source date move (code P, or a " & vbCrLf & _
        "  worked weekend W). Absent/leave workers are never touched." & vbCrLf & _
        "- The target date must show A for the worker, else they are skipped." & vbCrLf & _
        "- Weekend sources: the target weekday keeps 8 regular hours; OT = " & vbCrLf & _
        "  full day minus 8 (max 2), the remaining hours go to EOT." & vbCrLf & _
        "- Monthly P/A/OT/EOT totals recalculate automatically." & vbCrLf & _
        "- Undo restores the LAST batch only."
    FillSectionCombo Me
    FillDateCombo Me
    btnUndo.Enabled = HasUndo()
    Exit Sub
Fail:
    MsgBox "Could not initialise the transfer form: " & Err.Description, vbExclamation, "Transfer Duty"
End Sub

Private Sub btnGo_Click()
    On Error GoTo Fail
    Dim s As String, fd As String, td As String, res As String
    s = ComboText(cboSection)
    fd = ComboText(cboFrom)
    td = ComboText(cboTo)
    If fd = "" Or td = "" Then
        MsgBox "Pick both a source date and a target date.", vbExclamation, "Transfer Duty"
        Exit Sub
    End If
    If fd = td Then
        MsgBox "The source and target dates must be different.", vbExclamation, "Transfer Duty"
        Exit Sub
    End If
    If MsgBox("Move every qualifying duty from date " & fd & " to date " & td & "?" & vbCrLf & vbCrLf & _
              "Sub section: " & s & vbCrLf & _
              "Present workers on " & fd & " (P or worked W) move " & vbCrLf & _
              "only if date " & td & " shows A for them.", _
              vbYesNo + vbQuestion, "Transfer duty") <> vbYes Then Exit Sub

    res = ExecuteDutyTransfer(s, fd, td)
    MsgBox res, vbInformation, "Transfer duty"
    btnUndo.Enabled = HasUndo()
    Exit Sub
Fail:
    MsgBox "Transfer stopped: " & Err.Description, vbExclamation, "Transfer Duty"
End Sub

Private Sub btnUndo_Click()
    On Error GoTo Fail
    MsgBox UndoLastTransfer(), vbInformation, "Transfer duty"
    btnUndo.Enabled = HasUndo()
    Exit Sub
Fail:
    MsgBox "Undo failed: " & Err.Description, vbExclamation, "Transfer Duty"
End Sub

Private Function ComboText(ByVal cbo As Object) As String
    On Error Resume Next
    If IsNull(cbo.Value) Then
        ComboText = ""
    Else
        ComboText = Trim$(CStr(cbo.Value))
    End If
End Function
'''


def _add_ctl(designer, kind: str, name: str, **attrs):
    ctl = designer.Controls.Add(kind, name)
    for attr, value in attrs.items():
        try:
            setattr(ctl, attr, value)
            if attr == "Font":
                ctl.Font.Size = value
        except Exception:
            pass  # secondary property; default is fine
    return ctl


def _build_transfer_form(project) -> bool:
    """Design a modeless UserForm (frmDutyTransfer) inside the workbook."""
    form_name = "frmDutyTransfer"
    for comp in list(project.VBComponents):
        if comp.Name == form_name:
            project.VBComponents.Remove(comp)
    form = project.VBComponents.Add(3)  # vbext_ct_MSForm
    form.Name = form_name
    for key, value in (
        ("Caption", "Transfer Duty"),
        ("Width", 370),
        ("Height", 320),
    ):
        try:
            form.Properties(key).Value = value
        except Exception:
            pass

    des = form.Designer
    _add_ctl(des, "Forms.Label.1", "lblSection", Caption="Sub section (All = whole file)", Left=12, Top=14, Width=280, Height=14)
    _add_ctl(des, "Forms.ComboBox.1", "cboSection", Left=12, Top=32, Width=220, Height=18)
    _add_ctl(des, "Forms.Label.1", "lblFrom", Caption="From date (duty must be present)", Left=12, Top=58, Width=280, Height=14)
    _add_ctl(des, "Forms.ComboBox.1", "cboFrom", Left=12, Top=76, Width=220, Height=18)
    _add_ctl(des, "Forms.Label.1", "lblTo", Caption="To date (must be absent / A)", Left=12, Top=102, Width=280, Height=14)
    _add_ctl(des, "Forms.ComboBox.1", "cboTo", Left=12, Top=120, Width=220, Height=18)
    _add_ctl(des, "Forms.Label.1", "lblHint", Caption="Rules overview", Left=12, Top=148, Width=330, Height=110)
    _add_ctl(des, "Forms.CommandButton.1", "btnGo", Caption="Transfer", Left=12, Top=262, Width=155, Height=26)
    _add_ctl(des, "Forms.CommandButton.1", "btnUndo", Caption="Undo last transfer", Left=175, Top=262, Width=167, Height=26)

    form.CodeModule.AddFromString(VBA_FORM_CODE)
    return True


def attach_filter_macros(xlsx_path: str | Path) -> tuple[bool, str]:
    """Add dropdown-filter macros if Excel allows VBA project access.

    Returns (success, reason). On success the reason is empty; otherwise it
    names the exact blocker (missing pywin32, no Excel, VBA project access
    disabled, file locked, ...) so the console can show a concrete fix.
    """
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError:
        return False, (
            "Python package 'pywin32' is not installed in this Python.\n"
            "        Fix: run  python -m pip install pywin32"
        )

    xlsx_path = Path(xlsx_path).resolve()
    pythoncom.CoInitialize()
    excel = None
    try:
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            return False, (
                "Microsoft Excel could not be started (is it installed?):\n"
                f"        {exc}".strip()
            )
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        try:
            wb = excel.Workbooks.Open(str(xlsx_path))
        except Exception as exc:
            return False, (
                f"Excel could not open the generated workbook:\n"
                f"        {exc}".strip()
            )
        try:
            project = wb.VBProject
        except Exception:
            wb.Close(SaveChanges=False)
            return False, (
                "Excel is blocking programmatic access to macros.\n"
                "        Fix in Excel: File > Options > Trust Center >\n"
                "        Trust Center Settings... > Macro Settings >\n"
                "        tick 'Trust access to the VBA project object model',\n"
                "        then run this program again."
            )

        _replace_thisworkbook(project, VBA_THISWORKBOOK)
        _replace_std_module(project, "FilterHelpers", VBA_MODULE)
        _replace_std_module(project, "DutyTransferHelpers", VBA_TRANSFER_MODULE)
        _build_transfer_form(project)

        xlsm_path = xlsx_path.with_suffix(".xlsm")
        try:
            wb.SaveAs(str(xlsm_path), FileFormat=52)
        except Exception as exc:
            wb.Close(SaveChanges=False)
            return False, (
                f"Excel could not save {xlsm_path.name}:\n"
                f"        {exc}".strip() + "\n"
                "        (close the file if it is open in Excel, or check\n"
                "        write permission on the output folder)"
            )
        wb.Close(SaveChanges=False)
        return True, ""
    except Exception as exc:
        return False, f"Unexpected Excel automation error: {exc}".strip()
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def _replace_thisworkbook(project, code: str) -> None:
    target = project.VBComponents("ThisWorkbook")
    count = target.CodeModule.CountOfLines
    if count:
        target.CodeModule.DeleteLines(1, count)
    target.CodeModule.AddFromString(code)


def _replace_std_module(project, name: str, code: str) -> None:
    existing = None
    for component in project.VBComponents:
        if component.Name == name:
            existing = component
            break
    if existing is not None:
        project.VBComponents.Remove(existing)
    created = project.VBComponents.Add(1)
    created.Name = name
    created.CodeModule.AddFromString(code)
