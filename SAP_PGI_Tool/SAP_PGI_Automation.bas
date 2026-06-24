Attribute VB_Name = "SAP_PGI_Automation"
Option Explicit

Private Const SHEET_NAME        As String = "Deliveries"
Private Const BACKUP_SHEET      As String = "Backup"
Private Const COL_DELIVERY      As Integer = 1   ' Column A
Private Const COL_TRACKING      As Integer = 2   ' Column B
Private Const COL_CATEGORY      As Integer = 3   ' Column C
Private Const COL_STATUS        As Integer = 4   ' Column D
Private Const COL_QUANTITY      As Integer = 5   ' Column E (machines only)
Private Const DATA_START_ROW    As Integer = 2
Private Const BACKUP_COUNT_COL  As Integer = 7   ' Column G in backup - stores exact row count

Private Const WAIT_SHORT        As Long = 500
Private Const WAIT_MEDIUM       As Long = 800

' ============================================================
'  RESET BUTTON
' ============================================================
Public Sub ResetDeliveries()
    Dim ws      As Worksheet
    Dim answer  As Integer

    answer = MsgBox("This will delete ALL delivery and tracking numbers." & vbNewLine & _
                    "Are you sure you want to reset?", _
                    vbYesNo + vbQuestion, "Reset Confirmation")
    If answer = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)

    ' Backup BEFORE clearing
    BackupData ws

    ' Clear everything from row 2 down - fixed range, no End(xlUp)
    ws.Range("A" & DATA_START_ROW & ":E1000").ClearContents
    ws.Range("A" & DATA_START_ROW & ":E1000").Interior.ColorIndex = xlNone
    ws.Range("A" & DATA_START_ROW & ":E1000").Font.Bold = False

    ws.Cells(DATA_START_ROW, COL_DELIVERY).Select
    MsgBox "All deliveries cleared. Ready for new entries.", vbInformation, "Reset Complete"
End Sub

' ============================================================
'  UNDO RESET BUTTON
' ============================================================
Public Sub UndoReset()
    Dim ws      As Worksheet
    Dim wsBak   As Worksheet
    Dim i       As Long
    Dim bakRows As Long

    ' Check backup sheet exists
    On Error Resume Next
    Set wsBak = ThisWorkbook.Sheets(BACKUP_SHEET)
    On Error GoTo 0

    If wsBak Is Nothing Then
        MsgBox "No backup found. Please use Reset first before trying to Undo.", _
               vbInformation, "No Backup"
        Exit Sub
    End If

    ' Read exact row count stored during backup
    bakRows = 0
    If wsBak.Cells(1, BACKUP_COUNT_COL).Value <> "" Then
        bakRows = CLng(wsBak.Cells(1, BACKUP_COUNT_COL).Value)
    End If

    If bakRows <= 0 Then
        MsgBox "Backup appears to be empty. Nothing to restore.", _
               vbInformation, "Nothing to Restore"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)

    ' Clear current data first
    ws.Range("A" & DATA_START_ROW & ":E1000").ClearContents
    ws.Range("A" & DATA_START_ROW & ":E1000").Interior.ColorIndex = xlNone
    ws.Range("A" & DATA_START_ROW & ":E1000").Font.Bold = False

    ' Restore EXACTLY bakRows rows - no guessing, no End(xlUp)
    Dim destRow As Long
    destRow = DATA_START_ROW

    For i = DATA_START_ROW To DATA_START_ROW + bakRows - 1
        ' Restore each column individually to preserve formatting
        ws.Cells(destRow, COL_DELIVERY).Value  = wsBak.Cells(i, COL_DELIVERY).Value
        ws.Cells(destRow, COL_TRACKING).Value  = wsBak.Cells(i, COL_TRACKING).Value
        ws.Cells(destRow, COL_CATEGORY).Value  = wsBak.Cells(i, COL_CATEGORY).Value
        ws.Cells(destRow, COL_STATUS).Value    = wsBak.Cells(i, COL_STATUS).Value
        ws.Cells(destRow, COL_QUANTITY).Value  = wsBak.Cells(i, COL_QUANTITY).Value

        ' Restore status cell color
        Dim bakColor As Long
        bakColor = wsBak.Cells(i, COL_STATUS).Interior.Color
        If wsBak.Cells(i, COL_STATUS).Interior.ColorIndex <> xlNone Then
            ws.Cells(destRow, COL_STATUS).Interior.Color = bakColor
        End If
        ws.Cells(destRow, COL_STATUS).Font.Bold = wsBak.Cells(i, COL_STATUS).Font.Bold

        destRow = destRow + 1
    Next i

    MsgBox "Undo complete. " & bakRows & " rows restored.", vbInformation, "Undo Complete"
End Sub

' ============================================================
'  BACKUP - called by ResetDeliveries
' ============================================================
Private Sub BackupData(ws As Worksheet)
    Dim wsBak   As Worksheet
    Dim i       As Long
    Dim bakRow  As Long
    Dim cellVal As String
    Dim savedRows As Long

    ' Delete old backup sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(BACKUP_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create new hidden backup sheet
    Set wsBak = ThisWorkbook.Sheets.Add
    wsBak.Name = BACKUP_SHEET
    wsBak.Visible = xlSheetVeryHidden

    ' Copy header row
    ws.Rows(1).Copy wsBak.Rows(1)

    ' Scan the sheet - only save rows that currently have a valid delivery number
    ' Valid = cell is not empty AND is a number AND has 5+ digits
    bakRow = DATA_START_ROW
    savedRows = 0

    For i = DATA_START_ROW To 1000
        cellVal = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))

        ' Only backup rows with a real delivery number visible on screen RIGHT NOW
        If cellVal <> "" And IsNumeric(cellVal) And Len(cellVal) >= 5 Then
            wsBak.Cells(bakRow, COL_DELIVERY).Value  = ws.Cells(i, COL_DELIVERY).Value
            wsBak.Cells(bakRow, COL_TRACKING).Value  = ws.Cells(i, COL_TRACKING).Value
            wsBak.Cells(bakRow, COL_CATEGORY).Value  = ws.Cells(i, COL_CATEGORY).Value
            wsBak.Cells(bakRow, COL_STATUS).Value    = ws.Cells(i, COL_STATUS).Value
            wsBak.Cells(bakRow, COL_QUANTITY).Value  = ws.Cells(i, COL_QUANTITY).Value

            ' Copy status cell color
            If ws.Cells(i, COL_STATUS).Interior.ColorIndex <> xlNone Then
                wsBak.Cells(bakRow, COL_STATUS).Interior.Color = ws.Cells(i, COL_STATUS).Interior.Color
            End If
            wsBak.Cells(bakRow, COL_STATUS).Font.Bold = ws.Cells(i, COL_STATUS).Font.Bold

            bakRow = bakRow + 1
            savedRows = savedRows + 1
        End If
    Next i

    ' Store EXACT count in column G row 1 of backup sheet
    ' UndoReset reads this - no guessing required
    wsBak.Cells(1, BACKUP_COUNT_COL).Value = savedRows
End Sub

' ============================================================
'  RUN PGI - Parts & Conversion only
' ============================================================
Public Sub RunPGI()
    Dim ws         As Worksheet
    Dim sapSession As Object
    Dim lastRow    As Long
    Dim i          As Long
    Dim delivery   As String
    Dim tracking   As String
    Dim category   As String
    Dim doneCount  As Integer
    Dim errorCount As Integer
    Dim skipCount  As Integer

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).Row

    If lastRow < DATA_START_ROW Then
        MsgBox "No deliveries found.", vbInformation, "No Data"
        Exit Sub
    End If

    Set sapSession = GetSAPSession()
    If sapSession Is Nothing Then Exit Sub

    If MsgBox("Start PGI for Parts & Conversion now?", _
              vbYesNo + vbQuestion, "Confirm Start") = vbNo Then Exit Sub

    doneCount = 0 : errorCount = 0 : skipCount = 0

    For i = DATA_START_ROW To lastRow
        delivery = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))
        tracking = Trim(CStr(ws.Cells(i, COL_TRACKING).Value))
        category = UCase(Trim(CStr(ws.Cells(i, COL_CATEGORY).Value)))

        If delivery = "" Then GoTo NextRowPGI
        If InStr(ws.Cells(i, COL_STATUS).Value, "Done") > 0 Then GoTo NextRowPGI

        If category = "MACHINE" Or category = "MACHINES" Then
            If ws.Cells(i, COL_STATUS).Value = "" Then
                SetStatus ws, i, "Pending - Machine", "YELLOW"
            End If
            skipCount = skipCount + 1
            GoTo NextRowPGI
        End If

        SetStatus ws, i, "Processing...", "BLUE"
        DoEvents

        If ProcessPartsConversion(sapSession, delivery, tracking) Then
            SetStatus ws, i, "Done - PGI Posted", "GREEN"
            doneCount = doneCount + 1
        Else
            SetStatus ws, i, "ERROR - Check Manually", "RED"
            errorCount = errorCount + 1
        End If

NextRowPGI:
    Next i

    MsgBox "Parts & Conversion PGI Complete!" & vbNewLine & vbNewLine & _
           "Done:    " & doneCount & vbNewLine & _
           "Errors:  " & errorCount & vbNewLine & _
           "Skipped (Machines): " & skipCount, _
           vbInformation, "Process Complete"
End Sub

' ============================================================
'  RUN MACHINE PGI - Machines only
' ============================================================
Public Sub RunMachinePGI()
    Dim ws         As Worksheet
    Dim sapSession As Object
    Dim lastRow    As Long
    Dim i          As Long
    Dim delivery   As String
    Dim tracking   As String
    Dim category   As String
    Dim quantity   As Integer
    Dim doneCount  As Integer
    Dim errorCount As Integer
    Dim skipCount  As Integer

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).Row

    If lastRow < DATA_START_ROW Then
        MsgBox "No deliveries found.", vbInformation, "No Data"
        Exit Sub
    End If

    Set sapSession = GetSAPSession()
    If sapSession Is Nothing Then Exit Sub

    If MsgBox("Start PGI for Machines now?", _
              vbYesNo + vbQuestion, "Confirm Start") = vbNo Then Exit Sub

    doneCount = 0 : errorCount = 0 : skipCount = 0

    For i = DATA_START_ROW To lastRow
        delivery = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))
        tracking = Trim(CStr(ws.Cells(i, COL_TRACKING).Value))
        category = UCase(Trim(CStr(ws.Cells(i, COL_CATEGORY).Value)))

        If delivery = "" Then GoTo NextRowMachine
        If InStr(ws.Cells(i, COL_STATUS).Value, "Done") > 0 Then GoTo NextRowMachine

        If category <> "MACHINE" And category <> "MACHINES" Then
            skipCount = skipCount + 1
            GoTo NextRowMachine
        End If

        quantity = 0
        If ws.Cells(i, COL_QUANTITY).Value <> "" Then
            quantity = CInt(ws.Cells(i, COL_QUANTITY).Value)
        End If

        If quantity <= 0 Then
            SetStatus ws, i, "ERROR - Enter Quantity in Column E", "RED"
            errorCount = errorCount + 1
            GoTo NextRowMachine
        End If

        SetStatus ws, i, "Processing...", "BLUE"
        DoEvents

        If ProcessMachine(sapSession, delivery, tracking, quantity) Then
            SetStatus ws, i, "Done - PGI Posted (" & quantity & " machines)", "GREEN"
            doneCount = doneCount + 1
        Else
            SetStatus ws, i, "ERROR - Check Manually", "RED"
            errorCount = errorCount + 1
        End If

NextRowMachine:
    Next i

    MsgBox "Machine PGI Complete!" & vbNewLine & vbNewLine & _
           "Done:    " & doneCount & vbNewLine & _
           "Errors:  " & errorCount & vbNewLine & _
           "Skipped: " & skipCount, _
           vbInformation, "Process Complete"
End Sub

' ============================================================
'  CORE: Parts & Conversion
' ============================================================
Private Function ProcessPartsConversion(sapSession As Object, _
                                         delivery As String, _
                                         tracking As String) As Boolean
    On Error GoTo HandleError

    sapSession.StartTransaction "VL02N"
    SAPWait WAIT_MEDIUM

    sapSession.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = delivery
    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_MEDIUM

    Dim statusBar As String
    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "does not exist") > 0 Or _
       InStr(LCase(statusBar), "not found") > 0 Then GoTo HandleError

    sapSession.findById("wnd[0]/tbar[1]/btn[8]").press
    SAPWait WAIT_SHORT

    sapSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04").Select
    SAPWait WAIT_SHORT

    sapSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/" & _
                        "ssubSUBSCREEN_BODY:SAPMV50A:2108/txtLIKP-BOLNR").Text = tracking
    SAPWait WAIT_SHORT

    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_SHORT

    sapSession.findById("wnd[0]/tbar[1]/btn[20]").press
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "error") > 0 Or _
       InStr(LCase(statusBar), "not possible") > 0 Then GoTo HandleError

    ProcessPartsConversion = True
    Exit Function

HandleError:
    ProcessPartsConversion = False
End Function

' ============================================================
'  CORE: Machine (loops PGI per quantity)
' ============================================================
Private Function ProcessMachine(sapSession As Object, _
                                 delivery As String, _
                                 tracking As String, _
                                 quantity As Integer) As Boolean
    On Error GoTo HandleError

    Dim statusBar As String
    Dim j         As Integer

    ' 1. Open VL02N (works from any screen)
    sapSession.StartTransaction "VL02N"
    SAPWait WAIT_MEDIUM

    ' 2. Enter delivery number
    sapSession.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = delivery
    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "does not exist") > 0 Or _
       InStr(LCase(statusBar), "not found") > 0 Then GoTo HandleError

    ' 3. Open Header Details
    sapSession.findById("wnd[0]/tbar[1]/btn[8]").press
    SAPWait WAIT_SHORT

    ' 4. Enter tracking number in BilOfLad (Shipment tab)
    sapSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/" & _
                        "ssubSUBSCREEN_BODY:SAPMV50A:2108/txtLIKP-BOLNR").Text = tracking
    SAPWait WAIT_SHORT

    ' 5. Save and wait 5s for SAP to finish processing
    sapSession.findById("wnd[0]/tbar[0]/btn[11]").press
    SAPWait 5000

    ' 6. Loop PGI once per machine
    For j = 1 To quantity

        ' For machines 2, 3, etc.: re-open VL02N fresh and go to delivery overview
        If j > 1 Then
            sapSession.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00002"
            SAPWait WAIT_MEDIUM
            sapSession.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = delivery
            sapSession.findById("wnd[0]").sendVKey 0
            SAPWait WAIT_MEDIUM
        End If

        ' Click Post Goods Issue
        sapSession.findById("wnd[0]/tbar[1]/btn[20]").press
        SAPWait WAIT_MEDIUM

        ' Confirm Maintain Serial Numbers popup - click green tick
        sapSession.findById("wnd[1]/tbar[0]/btn[0]").press
        SAPWait WAIT_SHORT

        ' Save
        sapSession.findById("wnd[0]/tbar[0]/btn[11]").press
        SAPWait WAIT_MEDIUM

    Next j

    ProcessMachine = True
    Exit Function

HandleError:
    ProcessMachine = False
End Function

' ============================================================
'  HELPER: Connect to SAP
' ============================================================
Private Function GetSAPSession() As Object
    Dim sapGui  As Object
    Dim sapApp  As Object
    Dim sapConn As Object
    Dim sapSess As Object

    On Error GoTo NoSAP

    Set sapGui  = GetObject("SAPGUI")
    Set sapApp  = sapGui.GetScriptingEngine
    Set sapConn = sapApp.Children(0)
    Set sapSess = sapConn.Children(0)

    Set GetSAPSession = sapSess
    Exit Function

NoSAP:
    MsgBox "Cannot connect to SAP GUI." & vbNewLine & vbNewLine & _
           "Please make sure:" & vbNewLine & _
           "1. SAP GUI is open on your PC" & vbNewLine & _
           "2. You are logged into SAP" & vbNewLine & _
           "3. SAP GUI Scripting is enabled", _
           vbCritical, "SAP Connection Error"
    Set GetSAPSession = Nothing
End Function

' ============================================================
'  HELPER: Wait
' ============================================================
Private Sub SAPWait(milliseconds As Long)
    Application.Wait Now + (milliseconds / 86400000#)
End Sub

' ============================================================
'  HELPER: Set status
' ============================================================
Private Sub SetStatus(ws As Worksheet, row As Long, _
                      statusText As String, colorName As String)
    With ws.Cells(row, COL_STATUS)
        .Value = statusText
        Select Case colorName
            Case "GREEN"  : .Interior.Color = RGB(144, 238, 144)
            Case "RED"    : .Interior.Color = RGB(255, 99, 71)
            Case "YELLOW" : .Interior.Color = RGB(255, 255, 153)
            Case "ORANGE" : .Interior.Color = RGB(255, 200, 100)
            Case "BLUE"   : .Interior.Color = RGB(173, 216, 230)
            Case Else     : .Interior.ColorIndex = xlNone
        End Select
        .Font.Bold = (colorName = "RED")
    End With
    DoEvents
End Sub
