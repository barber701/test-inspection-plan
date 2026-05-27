#!/usr/bin/env python3
"""
setup_vba.py
Inserts Rev-7 VBA code and action buttons into Test_Inspection_Plan_Rev_7.xlsm
via Excel COM automation (requires Excel to be installed).
Run once; safe to re-run (buttons are removed then re-added).
"""

import sys, os, winreg, time
import win32com.client

FILE = r"E:\Coding Projects\Test and Inspection Plan\Test_Inspection_Plan_Rev_7.xlsm"

# ── Colour helper ──────────────────────────────────────────────────────────
def rgb(r, g, b):
    """Convert R,G,B to Excel's BGR integer (Windows COLORREF)."""
    return r + (g << 8) + (b << 16)

C_NAVY  = rgb(31,  56,  100)
C_GREEN = rgb(30,  70,   32)
C_GOLD  = rgb(166, 124,   0)
C_WHITE = rgb(255, 255, 255)
C_AMBER = rgb(255, 192,   0)

# ══════════════════════════════════════════════════════════════════════════
# VBA SOURCE CODE
# ══════════════════════════════════════════════════════════════════════════

VBA_THISWORKBOOK = r"""
Private Sub Workbook_Open()

    Dim wsCtrl      As Worksheet
    Dim currentMode As String
    Dim choice      As Integer

    Set wsCtrl = ThisWorkbook.Sheets("Control")
    currentMode = LCase(Trim(wsCtrl.Range("C4").Value))

    If currentMode <> "automated" And currentMode <> "manual" Then

        choice = MsgBox( _
            "Welcome to the Test & Inspection Plan." & vbCrLf & vbCrLf & _
            "How would you like to complete this plan?" & vbCrLf & vbCrLf & _
            "  YES  =  Automated  (import data from NetInspect)" & vbCrLf & _
            "  NO   =  Manual      (fill in the cards by hand)", _
            vbQuestion + vbYesNo, "Test & Inspection Plan -- Select Method")

        If choice = vbYes Then
            wsCtrl.Range("C4").Value = "Automated"
        Else
            wsCtrl.Range("C4").Value = "Manual"
        End If

    End If

    Call RefreshControlDisplay
    wsCtrl.Activate
    wsCtrl.Range("C11").Select

End Sub
"""

VBA_CTRL_SHEET = r"""
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$C$4" Then
        Call RefreshControlDisplay
    End If
End Sub
"""

VBA_MAIN = r"""
Option Explicit

Private Const CARDS_START_ROW      As Long = 6
Private Const ROWS_PER_CARD        As Long = 9
Private Const MANUAL_ROWS_PER_CARD As Long = 13
Private Const MANUAL_CARDS         As Long = 50

Private Const ROW_AUTO_FIRST  As Long = 6
Private Const ROW_AUTO_LAST   As Long = 9
Private Const ROW_MAN_FIRST   As Long = 19
Private Const ROW_MAN_LAST    As Long = 21

Private Const NI_OP     As Integer = 4
Private Const NI_SAMP   As Integer = 7
Private Const NI_DESC   As Integer = 8
Private Const NI_TYPE   As Integer = 9
Private Const NI_SPEC   As Integer = 10
Private Const NI_UNITS  As Integer = 11
Private Const NI_BUBBLE As Integer = 12
Private Const NI_REFLOC As Integer = 18
Private Const NI_PROC   As Integer = 15

Private Function IsManualMode() As Boolean
    Dim mode As String
    mode = Trim(ThisWorkbook.Sheets("Control").Range("C4").Value)
    IsManualMode = (LCase(mode) = "manual")
End Function

'==========================================================================
Public Sub RefreshControlDisplay()

    Dim wsCtrl As Worksheet
    Dim mode   As String

    Set wsCtrl = ThisWorkbook.Sheets("Control")
    mode = LCase(Trim(wsCtrl.Range("C4").Value))

    Application.ScreenUpdating = False

    Select Case mode
        Case "manual"
            wsCtrl.Rows(ROW_AUTO_FIRST & ":" & ROW_AUTO_LAST).Hidden = True
            wsCtrl.Rows(ROW_MAN_FIRST  & ":" & ROW_MAN_LAST).Hidden  = False
            wsCtrl.Range("B" & ROW_MAN_FIRST).Value = _
                Chr(9654) & "  STEP 1  --  OPEN MANUAL ENTRY SHEET"
            ThisWorkbook.Sheets("Manual Entry").Visible = xlSheetVisible

        Case "automated"
            wsCtrl.Rows(ROW_AUTO_FIRST & ":" & ROW_AUTO_LAST).Hidden = False
            wsCtrl.Rows(ROW_MAN_FIRST  & ":" & ROW_MAN_LAST).Hidden  = True
            wsCtrl.Range("B6").Value = _
                Chr(9654) & "  STEP 1  --  IMPORT NETINSPECT FILE"

        Case Else
            wsCtrl.Rows(ROW_AUTO_FIRST & ":" & ROW_AUTO_LAST).Hidden = False
            wsCtrl.Rows(ROW_MAN_FIRST  & ":" & ROW_MAN_LAST).Hidden  = False
    End Select

    Application.ScreenUpdating = True
End Sub

'==========================================================================
Private Sub Bdr(rng As Range, edge As Integer, wt As Integer, clr As Long)
    With rng.Borders(edge)
        .LineStyle = xlContinuous
        .Weight = wt
        .Color = clr
    End With
End Sub

Private Sub OutlineMed(rng As Range, clr As Long)
    Call Bdr(rng, xlEdgeLeft,   xlMedium, clr)
    Call Bdr(rng, xlEdgeRight,  xlMedium, clr)
    Call Bdr(rng, xlEdgeTop,    xlMedium, clr)
    Call Bdr(rng, xlEdgeBottom, xlMedium, clr)
End Sub

Private Sub SidesMedTB(rng As Range, botMed As Boolean, sideClr As Long)
    Dim cT As Long: cT = RGB(191, 191, 191)
    Call Bdr(rng, xlEdgeLeft,  xlMedium, sideClr)
    Call Bdr(rng, xlEdgeRight, xlMedium, sideClr)
    Call Bdr(rng, xlEdgeTop,   xlThin,   cT)
    If botMed Then
        Call Bdr(rng, xlEdgeBottom, xlMedium, sideClr)
    Else
        Call Bdr(rng, xlEdgeBottom, xlThin, cT)
    End If
End Sub

Private Sub LeftCell(rng As Range, botMed As Boolean, sideClr As Long)
    Dim cT As Long: cT = RGB(191, 191, 191)
    Call Bdr(rng, xlEdgeLeft,  xlMedium, sideClr)
    Call Bdr(rng, xlEdgeRight, xlThin,   cT)
    Call Bdr(rng, xlEdgeTop,   xlThin,   cT)
    If botMed Then
        Call Bdr(rng, xlEdgeBottom, xlMedium, sideClr)
    Else
        Call Bdr(rng, xlEdgeBottom, xlThin, cT)
    End If
End Sub

Private Sub RightCell(rng As Range, botMed As Boolean, sideClr As Long)
    Dim cT As Long: cT = RGB(191, 191, 191)
    Call Bdr(rng, xlEdgeLeft,  xlThin,   cT)
    Call Bdr(rng, xlEdgeRight, xlMedium, sideClr)
    Call Bdr(rng, xlEdgeTop,   xlThin,   cT)
    If botMed Then
        Call Bdr(rng, xlEdgeBottom, xlMedium, sideClr)
    Else
        Call Bdr(rng, xlEdgeBottom, xlThin, cT)
    End If
End Sub

Private Sub MidCell(rng As Range, botMed As Boolean, sideClr As Long)
    Dim cT As Long: cT = RGB(191, 191, 191)
    Call Bdr(rng, xlEdgeLeft,  xlThin, cT)
    Call Bdr(rng, xlEdgeRight, xlThin, cT)
    Call Bdr(rng, xlEdgeTop,   xlThin, cT)
    If botMed Then
        Call Bdr(rng, xlEdgeBottom, xlMedium, sideClr)
    Else
        Call Bdr(rng, xlEdgeBottom, xlThin, cT)
    End If
End Sub

'==========================================================================
Sub ImportNetInspect()

    If IsManualMode() Then
        MsgBox "The Input Method is set to Manual." & vbCrLf & vbCrLf & _
               "Switch to 'Automated' on the Control sheet to use the import function, " & _
               "or use the Manual Entry sheet to fill in characteristics directly.", _
               vbInformation, "Manual Mode Active"
        Exit Sub
    End If

    Dim fd          As FileDialog
    Dim sPath       As String
    Dim wbSrc       As Workbook
    Dim wsSrc       As Worksheet
    Dim wsRpt       As Worksheet
    Dim wsCtrl      As Worksheet
    Dim lastSrcRow  As Long
    Dim i           As Long
    Dim importCount As Long
    Dim cardRow     As Long

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select NetInspect Export File"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
    fd.AllowMultiSelect = False
    If fd.Show = False Then Exit Sub
    sPath = fd.SelectedItems(1)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Reading NetInspect file..."

    On Error GoTo ImportErr
    Set wbSrc = Workbooks.Open(sPath, ReadOnly:=True, UpdateLinks:=False)
    Set wsSrc = wbSrc.Sheets(1)

    If InStr(1, CStr(wsSrc.Cells(2, NI_OP).Value), "Operation", vbTextCompare) = 0 Then
        MsgBox "This does not look like a standard NetInspect export.", vbExclamation
        wbSrc.Close False
        GoTo CleanUp
    End If

    Set wsRpt  = ThisWorkbook.Sheets("Inspection Report")
    Set wsCtrl = ThisWorkbook.Sheets("Control")

    Application.StatusBar = "Clearing previous data..."
    Dim lastClear As Long
    lastClear = wsRpt.Cells(wsRpt.Rows.Count, "A").End(xlUp).Row
    If lastClear >= CARDS_START_ROW Then
        wsRpt.Rows(CARDS_START_ROW & ":" & lastClear + 10).Delete
    End If

    lastSrcRow  = wsSrc.Cells(wsSrc.Rows.Count, NI_OP).End(xlUp).Row
    importCount = 0
    cardRow     = CARDS_START_ROW

    For i = 3 To lastSrcRow
        If Trim(CStr(wsSrc.Cells(i, NI_OP).Value)) <> "" Then
            importCount = importCount + 1
            Application.StatusBar = "Building card " & importCount & "..."
            Call BuildCard(wsRpt, cardRow, importCount, _
                CStr(wsSrc.Cells(i, NI_OP).Value), _
                CStr(wsSrc.Cells(i, NI_SAMP).Value), _
                CStr(wsSrc.Cells(i, NI_DESC).Value), _
                CStr(wsSrc.Cells(i, NI_TYPE).Value), _
                CStr(wsSrc.Cells(i, NI_SPEC).Value), _
                CStr(wsSrc.Cells(i, NI_UNITS).Value), _
                CStr(wsSrc.Cells(i, NI_BUBBLE).Value), _
                CStr(wsSrc.Cells(i, NI_REFLOC).Value), _
                CStr(wsSrc.Cells(i, NI_PROC).Value))
            cardRow = cardRow + ROWS_PER_CARD
        End If
    Next i

    wbSrc.Close False
    wsCtrl.Range("C8").Value = "IMPORTED: " & Dir(sPath) & " - " & importCount & " items"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    If importCount > 0 Then
        MsgBox importCount & " items imported successfully.", vbInformation, "Import Complete"
        ThisWorkbook.Sheets("Inspection Report").Activate
        ThisWorkbook.Sheets("Inspection Report").Range("A1").Select
    End If
    Exit Sub

ImportErr:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    If Not wbSrc Is Nothing Then wbSrc.Close False
    MsgBox "Import failed: " & Err.Description, vbCritical, "Error"
End Sub

'==========================================================================
Sub GoToManualEntry()

    If Not IsManualMode() Then
        Dim ans As Integer
        ans = MsgBox("The Input Method is currently set to 'Automated'." & vbCrLf & vbCrLf & _
                     "Switch to 'Manual' mode on the Control sheet first, or click Yes to " & _
                     "open the Manual Entry sheet anyway.", _
                     vbQuestion + vbYesNo, "Mode Check")
        If ans = vbNo Then Exit Sub
    End If

    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets("Manual Entry")
    wsM.Visible = xlSheetVisible
    wsM.Activate
    wsM.Range("A1").Select

    MsgBox "Manual Entry sheet is ready." & vbCrLf & vbCrLf & _
           "Fill in each card:" & vbCrLf & _
           "  - Dark blue row      : Card number header" & vbCrLf & _
           "  - Amber row          : Field labels (Op No., Sampling, Type, Spec, Units, Bubble #)" & vbCrLf & _
           "  - White row          : Enter values in the individual cells" & vbCrLf & _
           "  - Green row          : Equipment Name and Equipment Number" & vbCrLf & _
           "  - Dark gold row      : Inspection Requirement / Description label" & vbCrLf & _
           "  - Yellow row         : Type the inspection requirement here" & vbCrLf & _
           "  - Blue row           : Reference Location and Process Reference" & vbCrLf & _
           "  - Dark green rows    : TAR Holder fields (dropdowns where applicable)" & vbCrLf & vbCrLf & _
           "50 blank cards are pre-built. Leave unused cards empty.", _
           vbInformation, "Manual Entry"
End Sub

'==========================================================================
Public Sub RebuildManualSheet()

    Dim confirm As Integer
    confirm = MsgBox("This will rebuild all " & MANUAL_CARDS & " blank Manual Entry cards." & vbCrLf & _
                     "Any existing data on this sheet will be cleared." & vbCrLf & vbCrLf & _
                     "Continue?", vbQuestion + vbYesNo, "Rebuild Manual Sheet")
    If confirm = vbNo Then Exit Sub

    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets("Manual Entry")
    wsM.Visible = xlSheetVisible

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim clearRow As Long
    clearRow = CARDS_START_ROW + MANUAL_CARDS * MANUAL_ROWS_PER_CARD + 5
    With wsM.Range("A" & CARDS_START_ROW & ":F" & clearRow)
        .ClearContents
        .ClearFormats
    End With

    Dim i        As Long
    Dim startRow As Long
    For i = 1 To MANUAL_CARDS
        startRow = CARDS_START_ROW + (i - 1) * MANUAL_ROWS_PER_CARD
        Application.StatusBar = "Building manual card " & i & " of " & MANUAL_CARDS & "..."
        Call BuildManualCard(wsM, startRow, i)
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False

    wsM.Activate
    wsM.Range("A1").Select
    MsgBox MANUAL_CARDS & " manual cards rebuilt successfully.", vbInformation, "Rebuild Complete"
End Sub

'==========================================================================
Private Sub BuildManualCard(ws As Worksheet, startRow As Long, cardNum As Long)

    Dim r0  As Long: r0 = startRow
    Dim r1  As Long: r1 = r0 + 1
    Dim r2  As Long: r2 = r0 + 2
    Dim r3  As Long: r3 = r0 + 3
    Dim r4  As Long: r4 = r0 + 4
    Dim r5  As Long: r5 = r0 + 5
    Dim r6  As Long: r6 = r0 + 6
    Dim r7  As Long: r7 = r0 + 7
    Dim r12 As Long: r12 = r0 + 12

    Dim cM  As Long: cM = RGB(47, 84, 150)
    Dim cT  As Long: cT = RGB(191, 191, 191)

    ws.Rows(r0).RowHeight   = 20
    ws.Rows(r1).RowHeight   = 14
    ws.Rows(r2).RowHeight   = 18
    ws.Rows(r3).RowHeight   = 18
    ws.Rows(r4).RowHeight   = 14
    ws.Rows(r5).RowHeight   = 42
    ws.Rows(r6).RowHeight   = 16
    ws.Rows(r7).RowHeight   = 16
    ws.Rows(r0 + 8).RowHeight  = 22
    ws.Rows(r0 + 9).RowHeight  = 22
    ws.Rows(r0 + 10).RowHeight = 22
    ws.Rows(r0 + 11).RowHeight = 28
    ws.Rows(r12).RowHeight  = 7

    With ws.Range("A" & r0 & ":F" & r0)
        .Merge
        .Value = "  ITEM " & cardNum
        .Interior.Color = RGB(31, 56, 100)
        .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 9
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .WrapText = False
    End With
    Call OutlineMed(ws.Range("A" & r0 & ":F" & r0), cM)
    Call Bdr(ws.Range("A" & r0 & ":F" & r0), xlEdgeBottom, xlThin, cT)

    Dim lblClr  As Long: lblClr  = RGB(255, 217, 102)
    Dim lblFont As Long: lblFont = RGB(31, 56, 100)
    Dim labels(1 To 6) As String
    labels(1) = "Op No.":  labels(2) = "Sampling"
    labels(3) = "Type":    labels(4) = "Spec"
    labels(5) = "Units":   labels(6) = "Bubble #"

    Dim c As Integer
    For c = 1 To 6
        With ws.Cells(r1, c)
            .Value = labels(c)
            .Interior.Color = lblClr
            .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 8
            .Font.Color = lblFont
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
    Next c
    Call Bdr(ws.Range("A" & r1 & ":F" & r1), xlEdgeLeft,  xlMedium, cM)
    Call Bdr(ws.Range("A" & r1 & ":F" & r1), xlEdgeRight, xlMedium, cM)
    Call Bdr(ws.Range("A" & r1 & ":F" & r1), xlEdgeTop,   xlThin,   cT)
    Call Bdr(ws.Range("A" & r1 & ":F" & r1), xlEdgeBottom,xlThin,   cT)
    With ws.Range("A" & r1 & ":F" & r1).Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = cT
    End With

    For c = 1 To 6
        With ws.Cells(r2, c)
            .Interior.Color = RGB(255, 255, 255)
            .Font.Name = "Arial": .Font.Size = 9: .Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
    Next c
    Call Bdr(ws.Range("A" & r2 & ":F" & r2), xlEdgeLeft,  xlMedium, cM)
    Call Bdr(ws.Range("A" & r2 & ":F" & r2), xlEdgeRight, xlMedium, cM)
    Call Bdr(ws.Range("A" & r2 & ":F" & r2), xlEdgeTop,   xlThin,   cT)
    Call Bdr(ws.Range("A" & r2 & ":F" & r2), xlEdgeBottom,xlThin,   cT)
    With ws.Range("A" & r2 & ":F" & r2).Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = cT
    End With

    With ws.Range("A" & r3 & ":B" & r3)
        .Merge: .Value = "  Equipment Name:"
        .Interior.Color = RGB(237, 247, 238)
        .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 8
        .Font.Color = RGB(30, 70, 32)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call LeftCell(ws.Range("A" & r3 & ":B" & r3), False, cM)

    With ws.Cells(r3, 3)
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Arial": .Font.Size = 9: .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call MidCell(ws.Cells(r3, 3), False, cM)

    With ws.Cells(r3, 4)
        .Value = "  Equipment No.:"
        .Interior.Color = RGB(237, 247, 238)
        .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 8
        .Font.Color = RGB(30, 70, 32)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call MidCell(ws.Cells(r3, 4), False, cM)

    With ws.Cells(r3, 5)
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Arial": .Font.Size = 9: .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call MidCell(ws.Cells(r3, 5), False, cM)

    ws.Cells(r3, 6).Interior.ColorIndex = xlNone
    Call Bdr(ws.Cells(r3, 6), xlEdgeRight,  xlMedium, cM)
    Call Bdr(ws.Cells(r3, 6), xlEdgeTop,    xlThin,   cT)
    Call Bdr(ws.Cells(r3, 6), xlEdgeBottom, xlThin,   cT)

    With ws.Range("A" & r4 & ":F" & r4)
        .Merge: .Value = "  INSPECTION REQUIREMENT / DESCRIPTION:"
        .Interior.Color = RGB(166, 124, 0)
        .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 8
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call SidesMedTB(ws.Range("A" & r4 & ":F" & r4), False, cM)

    With ws.Range("A" & r5 & ":F" & r5)
        .Merge: .Value = ""
        .Interior.Color = RGB(255, 253, 231)
        .Font.Name = "Arial": .Font.Size = 9: .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlTop: .WrapText = True
    End With
    Call SidesMedTB(ws.Range("A" & r5 & ":F" & r5), False, cM)

    With ws.Range("A" & r6 & ":C" & r6)
        .Merge: .Value = "  Ref: "
        .Interior.Color = RGB(222, 234, 246)
        .Font.Name = "Arial": .Font.Size = 8: .Font.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call LeftCell(ws.Range("A" & r6 & ":C" & r6), False, cM)

    With ws.Range("D" & r6 & ":E" & r6)
        .Merge: .Value = "  Process Ref: "
        .Interior.Color = RGB(222, 234, 246)
        .Font.Name = "Arial": .Font.Size = 8: .Font.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call MidCell(ws.Range("D" & r6 & ":E" & r6), False, cM)
    ws.Cells(r6, 6).Interior.ColorIndex = xlNone
    Call Bdr(ws.Cells(r6, 6), xlEdgeRight,  xlMedium, cM)
    Call Bdr(ws.Cells(r6, 6), xlEdgeTop,    xlThin,   cT)
    Call Bdr(ws.Cells(r6, 6), xlEdgeBottom, xlThin,   cT)

    With ws.Range("A" & r7 & ":F" & r7)
        .Merge: .Value = "  TAR HOLDER - PLEASE COMPLETE:"
        .Interior.Color = RGB(30, 70, 32)
        .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 8
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call SidesMedTB(ws.Range("A" & r7 & ":F" & r7), False, cM)

    Dim tarL(1 To 4) As String, tarR(1 To 4) As String
    tarL(1) = "Verification Method":     tarR(1) = "Equipment Type"
    tarL(2) = "MSA Verification Method": tarR(2) = "MSA Ref / Justification"
    tarL(3) = "Responsible Person":      tarR(3) = "Control Method"
    tarL(4) = "Reaction Plan":           tarR(4) = "Comments"

    Dim idx As Integer, tr As Long, isLast As Boolean
    For idx = 1 To 4
        tr = r7 + idx: isLast = (idx = 4)
        With ws.Range("A" & tr & ":B" & tr)
            .Merge: .Value = "  " & tarL(idx) & ":"
            .Interior.Color = RGB(237, 247, 238)
            .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 8
            .Font.Color = RGB(30, 70, 32)
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
        Call LeftCell(ws.Range("A" & tr & ":B" & tr), isLast, cM)
        With ws.Cells(tr, 3)
            .Interior.Color = RGB(255, 255, 255)
            .Font.Name = "Arial": .Font.Size = 9
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
        Call MidCell(ws.Cells(tr, 3), isLast, cM)
        With ws.Cells(tr, 4)
            .Value = "  " & tarR(idx) & ":"
            .Interior.Color = RGB(237, 247, 238)
            .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 8
            .Font.Color = RGB(30, 70, 32)
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
        Call MidCell(ws.Cells(tr, 4), isLast, cM)
        With ws.Cells(tr, 5)
            .Interior.Color = RGB(255, 255, 255)
            .Font.Name = "Arial": .Font.Size = 9
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
        Call MidCell(ws.Cells(tr, 5), isLast, cM)
        ws.Cells(tr, 6).Interior.ColorIndex = xlNone
        Call Bdr(ws.Cells(tr, 6), xlEdgeRight, xlMedium, cM)
        Call Bdr(ws.Cells(tr, 6), xlEdgeTop,   xlThin,   cT)
        If isLast Then
            Call Bdr(ws.Cells(tr, 6), xlEdgeBottom, xlMedium, cM)
        Else
            Call Bdr(ws.Cells(tr, 6), xlEdgeBottom, xlThin, cT)
        End If
    Next idx

    With ws.Cells(r7 + 1, 3).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="=DropdownLists!$A$2:$A$8"
        .IgnoreBlank = True:  .InCellDropdown = True
        .ShowInput   = False: .ShowError      = False
    End With
    With ws.Cells(r7 + 1, 5).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="=DropdownLists!$C$2:$C$99"
        .IgnoreBlank = True:  .InCellDropdown = True
        .ShowInput   = False: .ShowError      = False
    End With
    With ws.Cells(r7 + 2, 3).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="=DropdownLists!$B$2:$B$12"
        .IgnoreBlank = True:  .InCellDropdown = True
        .ShowInput   = False: .ShowError      = False
    End With

    ws.Rows(r12).RowHeight = 7
    ws.Range("A" & r12 & ":F" & r12).ClearFormats
End Sub

'==========================================================================
Private Sub BuildCard(ws As Worksheet, startRow As Long, itemNum As Long, _
                      opNo As String, sampling As String, description As String, _
                      measType As String, spec As String, units As String, _
                      bubble As String, refLoc As String, process As String)

    Dim r0 As Long: r0 = startRow
    Dim r1 As Long: r1 = r0 + 1
    Dim r2 As Long: r2 = r0 + 2
    Dim r3 As Long: r3 = r0 + 3
    Dim r8 As Long: r8 = r0 + 8

    Dim cM As Long: cM = RGB(47, 84, 150)
    Dim cT As Long: cT = RGB(191, 191, 191)

    ws.Rows(r0).RowHeight = 22
    ws.Rows(r1).RowHeight = 42
    ws.Rows(r2).RowHeight = 18
    ws.Rows(r3).RowHeight = 16
    ws.Rows(r0 + 4).RowHeight = 22
    ws.Rows(r0 + 5).RowHeight = 22
    ws.Rows(r0 + 6).RowHeight = 22
    ws.Rows(r0 + 7).RowHeight = 28
    ws.Rows(r8).RowHeight = 7

    With ws.Range("A" & r0 & ":F" & r0)
        .Merge
        .Value = "  ITEM " & itemNum & "     Op: " & opNo & _
                 "     Sampling: " & sampling & _
                 "     Type: " & measType & _
                 "     Spec: " & spec & IIf(units <> "", " " & units, "") & _
                 "     Bubble #: " & bubble
        .Interior.Color = RGB(31, 56, 100)
        .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 9
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .WrapText = False
    End With
    Call OutlineMed(ws.Range("A" & r0 & ":F" & r0), cM)
    Call Bdr(ws.Range("A" & r0 & ":F" & r0), xlEdgeBottom, xlThin, cT)

    With ws.Range("A" & r1 & ":F" & r1)
        .Merge
        .Value = description
        .Interior.Color = RGB(255, 253, 231)
        .Font.Name = "Arial": .Font.Size = 9: .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    Call SidesMedTB(ws.Range("A" & r1 & ":F" & r1), False, cM)

    With ws.Range("A" & r2 & ":C" & r2)
        .Merge: .Value = "  Ref: " & refLoc
        .Interior.Color = RGB(222, 234, 246)
        .Font.Name = "Arial": .Font.Size = 8: .Font.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call LeftCell(ws.Range("A" & r2 & ":C" & r2), False, cM)

    With ws.Range("D" & r2 & ":E" & r2)
        .Merge: .Value = "  Process Ref: " & process
        .Interior.Color = RGB(222, 234, 246)
        .Font.Name = "Arial": .Font.Size = 8: .Font.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call MidCell(ws.Range("D" & r2 & ":E" & r2), False, cM)
    Call Bdr(ws.Cells(r2, 6), xlEdgeRight, xlMedium, cM)
    Call Bdr(ws.Cells(r2, 6), xlEdgeTop,   xlThin,   cT)
    Call Bdr(ws.Cells(r2, 6), xlEdgeBottom,xlThin,   cT)
    ws.Cells(r2, 6).Interior.ColorIndex = xlNone

    With ws.Range("A" & r3 & ":F" & r3)
        .Merge: .Value = "  TAR HOLDER - PLEASE COMPLETE:"
        .Interior.Color = RGB(30, 70, 32)
        .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 8
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    Call SidesMedTB(ws.Range("A" & r3 & ":F" & r3), False, cM)

    Dim tarL(1 To 4) As String, tarR(1 To 4) As String
    tarL(1) = "Verification Method":     tarR(1) = "Equipment Type"
    tarL(2) = "MSA Verification Method": tarR(2) = "MSA Ref / Justification"
    tarL(3) = "Responsible Person":      tarR(3) = "Control Method"
    tarL(4) = "Reaction Plan":           tarR(4) = "Comments"

    Dim idx As Integer, tr As Long, isLast As Boolean
    For idx = 1 To 4
        tr = r3 + idx: isLast = (idx = 4)
        With ws.Range("A" & tr & ":B" & tr)
            .Merge: .Value = "  " & tarL(idx) & ":"
            .Interior.Color = RGB(237, 247, 238)
            .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 8
            .Font.Color = RGB(30, 70, 32)
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
        Call LeftCell(ws.Range("A" & tr & ":B" & tr), isLast, cM)
        With ws.Cells(tr, 3)
            .Interior.Color = RGB(255, 255, 255)
            .Font.Name = "Arial": .Font.Size = 9
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
        Call MidCell(ws.Cells(tr, 3), isLast, cM)
        With ws.Cells(tr, 4)
            .Value = "  " & tarR(idx) & ":"
            .Interior.Color = RGB(237, 247, 238)
            .Font.Name = "Arial": .Font.Bold = True: .Font.Size = 8
            .Font.Color = RGB(30, 70, 32)
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
        Call MidCell(ws.Cells(tr, 4), isLast, cM)
        With ws.Cells(tr, 5)
            .Interior.Color = RGB(255, 255, 255)
            .Font.Name = "Arial": .Font.Size = 9
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        End With
        Call MidCell(ws.Cells(tr, 5), isLast, cM)
        ws.Cells(tr, 6).Interior.ColorIndex = xlNone
        Call Bdr(ws.Cells(tr, 6), xlEdgeRight, xlMedium, cM)
        Call Bdr(ws.Cells(tr, 6), xlEdgeTop,   xlThin,   cT)
        If isLast Then
            Call Bdr(ws.Cells(tr, 6), xlEdgeBottom, xlMedium, cM)
        Else
            Call Bdr(ws.Cells(tr, 6), xlEdgeBottom, xlThin, cT)
        End If
    Next idx

    With ws.Cells(r3 + 1, 3).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="=DropdownLists!$A$2:$A$8"
        .IgnoreBlank = True: .InCellDropdown = True
        .ShowInput = False: .ShowError = False
    End With
    With ws.Cells(r3 + 1, 5).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="=DropdownLists!$C$2:$C$99"
        .IgnoreBlank = True: .InCellDropdown = True
        .ShowInput = False: .ShowError = False
    End With
    With ws.Cells(r3 + 2, 3).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="=DropdownLists!$B$2:$B$12"
        .IgnoreBlank = True: .InCellDropdown = True
        .ShowInput = False: .ShowError = False
    End With

    ws.Rows(r8).RowHeight = 7
    ws.Range("A" & r8 & ":F" & r8).ClearFormats
End Sub

'==========================================================================
Sub SendToTARHolder()

    Dim wsCtrl   As Worksheet
    Dim tarEmail As String
    Dim mvaNo    As String
    Dim prepBy   As String
    Dim planDate As String

    Set wsCtrl = ThisWorkbook.Sheets("Control")
    mvaNo    = Trim(wsCtrl.Range("C11").Value)
    planDate = Trim(wsCtrl.Range("C13").Value)
    prepBy   = Trim(wsCtrl.Range("C12").Value)
    tarEmail = Trim(wsCtrl.Range("C16").Value)

    If tarEmail = "" Then
        MsgBox "Please enter the TAR Holder email in Step 3.", vbExclamation: Exit Sub
    End If
    If mvaNo = "" Or mvaNo = "DAOFXXXX" Then
        MsgBox "Please enter the MVA Number in Step 2.", vbExclamation: Exit Sub
    End If

    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook to a folder first.", vbExclamation
        Application.Dialogs(xlDialogSaveAs).Show
        If ThisWorkbook.Path = "" Then Exit Sub
    End If
    ThisWorkbook.Save

    Dim modeNote As String
    If IsManualMode() Then
        modeNote = "This plan has been prepared using Manual Entry mode." & vbCrLf & _
                   "Please complete the green fields on the Manual Entry sheet."
    Else
        modeNote = "All inspection characteristics have been automatically imported from NetInspect." & vbCrLf & _
                   "Please complete the green fields on the Inspection Report sheet."
    End If

    On Error GoTo EmailErr
    Dim olApp  As Object: Set olApp = CreateObject("Outlook.Application")
    Dim olMail As Object: Set olMail = olApp.CreateItem(0)

    Dim body As String
    body = "Dear TAR Holder," & vbCrLf & vbCrLf & _
           "Please find attached the Test & Inspection Plan for review." & vbCrLf & vbCrLf & _
           "  MVA Number : " & mvaNo & vbCrLf & _
           IIf(planDate <> "", "  Date        : " & planDate & vbCrLf, "") & _
           IIf(prepBy <> "", "  Prepared By : " & prepBy & vbCrLf, "") & _
           vbCrLf & modeNote & vbCrLf & vbCrLf & _
           "ACTION REQUIRED:" & vbCrLf & _
           "  1. Open the workbook (enable macros)" & vbCrLf & _
           "  2. Complete the green fields on the " & _
              IIf(IsManualMode(), "Manual Entry", "Inspection Report") & " sheet" & vbCrLf & _
           "  3. Sign the Approval sheet" & vbCrLf & _
           "  4. Click Export as PDF on the Approval sheet" & vbCrLf & vbCrLf & _
           "Kind regards," & vbCrLf & IIf(prepBy <> "", prepBy, "[Operator]")

    With olMail
        .To = tarEmail
        .Subject = "ACTION REQUIRED: Test & Inspection Plan - " & mvaNo
        .body = body
        .Attachments.Add ThisWorkbook.FullName
        .Display
    End With

    MsgBox "Email draft opened in Outlook. Please review and send.", vbInformation, "Email Ready"
    Exit Sub
EmailErr:
    MsgBox "Email error: " & Err.Description & vbCrLf & "Check Outlook is open.", vbCritical, "Error"
End Sub

'==========================================================================
Sub ExportAsPDF()

    Dim wsCtrl   As Worksheet
    Dim mvaNo    As String
    Dim savePath As String
    Dim fd       As FileDialog
    Dim rptSheet As String

    Set wsCtrl = ThisWorkbook.Sheets("Control")
    mvaNo = Trim(wsCtrl.Range("C11").Value)

    If mvaNo = "" Or mvaNo = "DAOFXXXX" Then
        MsgBox "Please enter the MVA Number on the Control sheet.", vbExclamation: Exit Sub
    End If

    If IsManualMode() Then
        rptSheet = "Manual Entry"
        ThisWorkbook.Sheets("Manual Entry").Visible = xlSheetVisible
    Else
        rptSheet = "Inspection Report"
    End If

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Choose folder to save PDF"
    fd.InitialFileName = IIf(ThisWorkbook.Path <> "", ThisWorkbook.Path, "C:\")
    If fd.Show = False Then Exit Sub

    savePath = fd.SelectedItems(1) & "\Test_Inspection_Plan_" & mvaNo & "_Approved.pdf"

    On Error GoTo PDFErr
    Application.ScreenUpdating = False

    Dim wsRpt As Worksheet
    Set wsRpt = ThisWorkbook.Sheets(rptSheet)
    Dim lastRow As Long
    lastRow = wsRpt.Cells(wsRpt.Rows.Count, "A").End(xlUp).Row
    If lastRow < CARDS_START_ROW Then lastRow = CARDS_START_ROW + 10
    wsRpt.PageSetup.PrintArea = "A1:F" & lastRow
    wsRpt.PageSetup.Orientation = xlPortrait
    wsRpt.PageSetup.FitToPagesWide = 1
    wsRpt.PageSetup.FitToPagesTall = False

    ThisWorkbook.Sheets(Array(rptSheet, "Approval")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True

    ThisWorkbook.Sheets("Approval").Select
    Application.ScreenUpdating = True
    MsgBox "PDF saved to:" & vbCrLf & savePath, vbInformation, "Export Complete"
    Exit Sub

PDFErr:
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets("Approval").Select
    MsgBox "PDF export error: " & Err.Description, vbCritical, "Error"
End Sub
"""

# ══════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════

def enable_vba_access():
    """Enable programmatic VBA access via registry (required for AddFromString)."""
    tried = False
    for ver in ["16.0", "15.0", "14.0", "12.0"]:
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                rf"Software\Microsoft\Office\{ver}\Excel\Security",
                0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            winreg.CloseKey(key)
            tried = True
            break
        except OSError:
            pass
    return tried


def set_vba_code(vb_project, module_name, code, mod_type=1):
    """Insert or replace a standard VBA module."""
    # mod_type: 1=stdModule, 100=Document (sheet/workbook)
    existing = None
    for comp in vb_project.VBComponents:
        if comp.Name == module_name:
            existing = comp
            break
    if existing is None:
        existing = vb_project.VBComponents.Add(mod_type)
        existing.Name = module_name
    cm = existing.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(code)


def add_button(ws, name, label, row_num, col_start_addr, col_end_addr,
               on_action, bg_color, h_pad=2):
    """Remove any existing button with this name then add a styled rectangle shape."""
    for shp in list(ws.Shapes):
        if shp.Name == name:
            shp.Delete()

    row    = ws.Rows(row_num)
    rng    = ws.Range(f"{col_start_addr}{row_num}:{col_end_addr}{row_num}")
    top    = row.Top + h_pad
    height = row.Height - h_pad * 2
    left   = rng.Left
    width  = rng.Width

    shp = ws.Shapes.AddShape(1, left, top, width, height)  # msoShapeRectangle=1
    shp.Name     = name
    shp.OnAction = on_action

    shp.Fill.ForeColor.RGB = bg_color
    shp.Line.Visible = False  # msoFalse = 0 but False works via COM

    tf = shp.TextFrame
    tf.Characters().Text = label
    tf.HorizontalAlignment = -4108  # xlHAlignCenter
    tf.VerticalAlignment   = -4108  # xlVAlignCenter
    tf.AutoSize            = False

    fnt = tf.Characters().Font
    fnt.Name  = "Calibri"
    fnt.Size  = 10
    fnt.Bold  = True
    fnt.Color = C_WHITE


def fix_control_sheet(ws):
    """Tidy rows 17-21 of the Control sheet and add row heights."""
    # Clear the stray text labels in 18-19 (will be replaced by shapes + VBA labels)
    ws.Cells(18, 2).ClearContents()
    ws.Cells(19, 2).ClearContents()

    # Ensure adequate row heights for button rows
    for r, h in [(8, 26), (17, 26), (20, 26), (21, 26)]:
        ws.Rows(r).RowHeight = h

    # Input Method dropdown in C4
    with_dv = ws.Cells(4, 3).Validation
    with_dv.Delete()
    ws.Cells(4, 3).Validation.Add(
        Type=3,            # xlValidateList
        AlertStyle=1,      # xlValidAlertInformation
        Formula1='"Automated,Manual"'
    )
    ws.Cells(4, 3).Validation.InCellDropdown = True
    ws.Cells(4, 3).Validation.IgnoreBlank   = True

    # Add dropdown via COM properly
    with_dv2 = ws.Range("C4").Validation
    with_dv2.Delete()
    ws.Range("C4").Validation.Add(Type=3, AlertStyle=1, Formula1='"Automated,Manual"')


# ══════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════

def main():
    if not os.path.exists(FILE):
        print(f"ERROR: File not found: {FILE}")
        sys.exit(1)

    print("Enabling VBA object model access ...")
    enable_vba_access()

    print("Opening Excel ...")
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible       = False
    xl.DisplayAlerts = False

    try:
        wb = xl.Workbooks.Open(
            os.path.abspath(FILE),
            UpdateLinks=False,
            ReadOnly=False
        )

        vbp = wb.VBProject
        print("Inserting VBA: ThisWorkbook ...")
        tw_comp = vbp.VBComponents("ThisWorkbook")
        cm = tw_comp.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(VBA_THISWORKBOOK)

        print("Inserting VBA: Control sheet module ...")
        ws_ctrl = wb.Sheets("Control")
        ctrl_comp = vbp.VBComponents(ws_ctrl.CodeName)
        cm2 = ctrl_comp.CodeModule
        if cm2.CountOfLines > 0:
            cm2.DeleteLines(1, cm2.CountOfLines)
        cm2.AddFromString(VBA_CTRL_SHEET)

        print("Inserting VBA: Main module ...")
        set_vba_code(vbp, "TIPMain", VBA_MAIN, mod_type=1)

        print("Fixing Control sheet layout ...")

        # Add C4 dropdown via Validation
        ws_ctrl.Range("C4").Validation.Delete()
        ws_ctrl.Range("C4").Validation.Add(
            Type=3, AlertStyle=1, Formula1='"Automated,Manual"')
        ws_ctrl.Range("C4").Validation.InCellDropdown = True

        # Row heights
        for r, h in [(8, 26), (17, 26), (20, 26), (21, 26)]:
            ws_ctrl.Rows(r).RowHeight = h

        # Clear stray text in 18-19
        ws_ctrl.Cells(18, 2).ClearContents()
        ws_ctrl.Cells(19, 2).ClearContents()

        print("Adding Control sheet buttons ...")
        add_button(ws_ctrl,
                   name="btnImportNetInspect",
                   label=chr(8659) + "  Import NetInspect File",
                   row_num=8, col_start_addr="B", col_end_addr="C",
                   on_action="ImportNetInspect",
                   bg_color=C_NAVY)

        add_button(ws_ctrl,
                   name="btnSendToTAR",
                   label=chr(9993) + "  Send to TAR Holder",
                   row_num=17, col_start_addr="B", col_end_addr="C",
                   on_action="SendToTARHolder",
                   bg_color=C_NAVY)

        add_button(ws_ctrl,
                   name="btnOpenManual",
                   label=chr(9654) + "  Open Manual Entry Sheet",
                   row_num=20, col_start_addr="B", col_end_addr="C",
                   on_action="GoToManualEntry",
                   bg_color=C_GREEN)

        add_button(ws_ctrl,
                   name="btnRebuildManual",
                   label=chr(9167) + "  Rebuild Manual Sheet",
                   row_num=21, col_start_addr="B", col_end_addr="C",
                   on_action="RebuildManualSheet",
                   bg_color=C_GREEN)

        print("Adding Approval sheet Export PDF button ...")
        ws_appr = wb.Sheets("Approval")
        # Ensure row 9 exists with adequate height
        ws_appr.Rows(9).RowHeight = 26
        add_button(ws_appr,
                   name="btnExportPDF",
                   label=chr(8594) + "  Export as PDF",
                   row_num=9, col_start_addr="B", col_end_addr="D",
                   on_action="ExportAsPDF",
                   bg_color=C_NAVY)

        print("Saving workbook ...")
        wb.Save()
        print("Done. Closing Excel ...")
        wb.Close(SaveChanges=False)

    except Exception as ex:
        import traceback
        print("ERROR:", ex)
        traceback.print_exc()
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
    finally:
        xl.Quit()

    print()
    print("setup_vba.py complete.")
    print("Open Test_Inspection_Plan_Rev_7.xlsm -- macros must be enabled.")
    print("On open, answer YES (Automated) or NO (Manual).")


if __name__ == "__main__":
    main()
