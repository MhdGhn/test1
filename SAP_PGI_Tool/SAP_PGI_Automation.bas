Attribute VB_Name = "SAP_PGI_Automation"
Option Explicit

Private Const SHEET_NAME        As String = "Deliveries"
Private Const BACKUP_SHEET      As String = "Backup"
Private Const COL_DELIVERY      As Integer = 1
Private Const COL_TRACKING      As Integer = 2
Private Const COL_CATEGORY      As Integer = 3
Private Const COL_STATUS        As Integer = 4
Private Const COL_QUANTITY      As Integer = 5
Private Const DATA_START_ROW    As Integer = 2
Private Const BACKUP_COUNT_COL  As Integer = 7
Private Const SALES_ORDER_CELL  As String = "J18"

Private Const WAIT_SHORT        As Long = 500
Private Const WAIT_MEDIUM       As Long = 800

' ============================================================
'  CREATE DELIVERY FROM SALES ORDER
' ============================================================
Public Sub CreateDelivery()
    Dim ws          As Worksheet
    Dim sapSession  As Object
    Dim salesOrder  As String
    Dim newDelivery As String
    Dim nextRow     As Long
    Dim lastRow     As Long

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)

    salesOrder = Trim(CStr(ws.Range(SALES_ORDER_CELL).Value))

    If salesOrder = "" Then
        MsgBox "Please enter a Sales Order number in cell J18 first.", _
               vbInformation, "No Sales Order"
        Exit Sub
    End If

    Set sapSession = GetSAPSession()
    If sapSession Is Nothing Then Exit Sub

    If MsgBox("Create delivery for Sales Order: " & salesOrder & "?", _
              vbYesNo + vbQuestion, "Confirm") = vbNo Then Exit Sub

    newDelivery = ProcessCreateDelivery(sapSession, salesOrder)

    If newDelivery = "" Then
        MsgBox "Failed to create delivery for Sales Order " & salesOrder & "." & vbNewLine & _
               "Please check SAP manually.", vbCritical, "Error"
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).row
    nextRow = lastRow + 1
    If lastRow < DATA_START_ROW Then nextRow = DATA_START_ROW

    ws.Cells(nextRow, COL_DELIVERY).Value = CLng(newDelivery)
    ws.Range(SALES_ORDER_CELL).Value = ""

    MsgBox "Delivery created successfully!" & vbNewLine & vbNewLine & _
           "Sales Order:     " & salesOrder & vbNewLine & _
           "New Delivery:    " & newDelivery & vbNewLine & vbNewLine & _
           "Delivery number added to row " & nextRow & " in Column A.", _
           vbInformation, "Delivery Created"

    ws.Cells(nextRow, COL_TRACKING).Select
End Sub

' ============================================================
'  CORE: Create Delivery in VL01N
' ============================================================
Private Function ProcessCreateDelivery(sapSession As Object, _
                                        salesOrder As String) As String
    On Error GoTo HandleError

    Dim statusBar As String
    Dim dateStr   As String
    Dim rowIndex  As Integer
    Dim basePath  As String

    basePath = "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/" & _
               "ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/"

    dateStr = Format(DateAdd("yyyy", 1, Date), "DD.MM.YYYY")

    sapSession.StartTransaction "VL01N"
    SAPWait WAIT_MEDIUM

    sapSession.findById("wnd[0]/usr/ctxtLIKP-VSTEL").Text = "A230"
    SAPWait 100

    sapSession.findById("wnd[0]/usr/ctxtLV50C-DATBI").Text = dateStr
    SAPWait 100

    sapSession.findById("wnd[0]/usr/ctxtLV50C-VBELN").Text = salesOrder
    SAPWait 100

    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "does not exist") > 0 Or _
       InStr(LCase(statusBar), "not found") > 0 Or _
       InStr(LCase(statusBar), "error") > 0 Then GoTo HandleError

    sapSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
    SAPWait WAIT_SHORT

    rowIndex = 0
    Do
        On Error Resume Next
        sapSession.findById(basePath & "txtLIPSD-PIKMG[6," & rowIndex & "]").Text = "1"
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo HandleError
            Exit Do
        End If
        On Error GoTo HandleError
        SAPWait 100
        rowIndex = rowIndex + 1
    Loop

    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_SHORT

    sapSession.findById("wnd[0]/tbar[0]/btn[11]").press
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text

    Dim words()     As String
    Dim k           As Integer
    Dim newDelivery As String

    newDelivery = ""

    Dim cleanMsg As String
    cleanMsg = Replace(statusBar, ".", " ")
    cleanMsg = Replace(cleanMsg, ",", " ")
    words = Split(cleanMsg, " ")

    For k = 0 To UBound(words)
        Dim w As String
        w = Trim(words(k))
        If Len(w) = 8 And IsNumeric(w) Then
            newDelivery = w
            Exit For
        End If
    Next k

    If newDelivery = "" Then GoTo HandleError

    ProcessCreateDelivery = newDelivery
    Exit Function

HandleError:
    On Error Resume Next
    ProcessCreateDelivery = ""
End Function

' ============================================================
'  IMPORT DELIVERY NUMBERS FROM PDF (clipboard)
' ============================================================
Public Sub ImportFromClipboard()
    Dim ws          As Worksheet
    Dim clipText    As String
    Dim words()     As String
    Dim j           As Long
    Dim word        As String
    Dim nextRow     As Long
    Dim addedCount  As Integer
    Dim skipCount   As Integer
    Dim existing    As String
    Dim lastRow     As Long

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    clipText = GetClipboardText()

    If clipText = "" Then
        MsgBox "Clipboard is empty." & vbNewLine & vbNewLine & _
               "Please:" & vbNewLine & _
               "1. Open your PDF in any viewer" & vbNewLine & _
               "2. Press Ctrl+A to select all text" & vbNewLine & _
               "3. Press Ctrl+C to copy" & vbNewLine & _
               "4. Click this button again", _
               vbInformation, "Nothing to Import"
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).row
    existing = ""
    If lastRow >= DATA_START_ROW Then
        Dim i As Long
        For i = DATA_START_ROW To lastRow
            Dim dv As String
            dv = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))
            If dv <> "" Then existing = existing & "|" & dv & "|"
        Next i
    End If

    nextRow = lastRow + 1
    If lastRow < DATA_START_ROW Then nextRow = DATA_START_ROW

    addedCount = 0
    skipCount = 0

    clipText = Replace(clipText, vbCrLf, " ")
    clipText = Replace(clipText, vbLf, " ")
    clipText = Replace(clipText, vbCr, " ")
    clipText = Replace(clipText, vbTab, " ")
    clipText = Replace(clipText, ",", " ")
    clipText = Replace(clipText, ";", " ")

    words = Split(clipText, " ")

    For j = 0 To UBound(words)
        word = Trim(words(j))
        If Len(word) = 8 And IsNumeric(word) Then
            If InStr(existing, "|" & word & "|") > 0 Then
                skipCount = skipCount + 1
            Else
                ws.Cells(nextRow, COL_DELIVERY).Value = CLng(word)
                existing = existing & "|" & word & "|"
                nextRow = nextRow + 1
                addedCount = addedCount + 1
            End If
        End If
    Next j

    If addedCount = 0 And skipCount = 0 Then
        MsgBox "No 8-digit delivery numbers found in the copied text.", _
               vbInformation, "Nothing Found"
    Else
        MsgBox "Import complete!" & vbNewLine & vbNewLine & _
               "Added:    " & addedCount & " new delivery numbers" & vbNewLine & _
               "Skipped:  " & skipCount & " (already in sheet)", _
               vbInformation, "Import Complete"
        ws.Cells(DATA_START_ROW, COL_DELIVERY).Select
    End If
End Sub

' ============================================================
'  IMPORT TRACKING NUMBER FROM COURIER WEBSITE (clipboard)
' ============================================================
Public Sub ImportTrackingFromClipboard()
    Dim ws        As Worksheet
    Dim clipText  As String
    Dim words()   As String
    Dim j         As Long
    Dim word      As String
    Dim tracking  As String
    Dim targetRow As Long
    Dim lastRow   As Long
    Dim i         As Long

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    clipText = GetClipboardText()

    If clipText = "" Then
        MsgBox "Clipboard is empty." & vbNewLine & vbNewLine & _
               "Please copy the tracking number from the courier website first.", _
               vbInformation, "Nothing to Import"
        Exit Sub
    End If

    clipText = Replace(clipText, vbCrLf, " ")
    clipText = Replace(clipText, vbLf, " ")
    clipText = Replace(clipText, vbCr, " ")
    clipText = Replace(clipText, vbTab, " ")
    clipText = Replace(clipText, ",", " ")
    clipText = Replace(clipText, ";", " ")

    tracking = ""
    words = Split(clipText, " ")

    For j = 0 To UBound(words)
        word = Trim(words(j))
        If Len(word) >= 8 Then
            If IsTrackingNumber(word) Then
                tracking = word
                Exit For
            End If
        End If
    Next j

    If tracking = "" Then
        MsgBox "No tracking number found in the copied text." & vbNewLine & vbNewLine & _
               "Expected format: IGT... or ECN... followed by digits.", _
               vbInformation, "Not Found"
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).row
    targetRow = 0

    For i = DATA_START_ROW To lastRow
        Dim delVal As String
        delVal = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))
        Dim trkVal As String
        trkVal = Trim(CStr(ws.Cells(i, COL_TRACKING).Value))
        If delVal <> "" And trkVal = "" Then
            targetRow = i
            Exit For
        End If
    Next i

    If targetRow = 0 Then
        MsgBox "No empty tracking number slot found.", _
               vbInformation, "No Slot Available"
        Exit Sub
    End If

    ws.Cells(targetRow, COL_TRACKING).Value = tracking
    ws.Cells(targetRow, COL_TRACKING).Select

    MsgBox "Tracking number added to row " & targetRow & ":" & vbNewLine & vbNewLine & _
           "Delivery:  " & ws.Cells(targetRow, COL_DELIVERY).Value & vbNewLine & _
           "Tracking:  " & tracking, _
           vbInformation, "Tracking Added"
End Sub

' ============================================================
'  HELPER: Check if word is a tracking number
' ============================================================
Private Function IsTrackingNumber(word As String) As Boolean
    Dim k           As Integer
    Dim letterCount As Integer
    Dim digitCount  As Integer
    Dim c           As String

    letterCount = 0
    digitCount = 0

    For k = 1 To Len(word)
        c = Mid(word, k, 1)
        If c >= "A" And c <= "Z" Then
            If digitCount > 0 Then
                IsTrackingNumber = False
                Exit Function
            End If
            letterCount = letterCount + 1
        ElseIf c >= "0" And c <= "9" Then
            digitCount = digitCount + 1
        Else
            IsTrackingNumber = False
            Exit Function
        End If
    Next k

    IsTrackingNumber = (letterCount >= 2 And letterCount <= 4 And digitCount >= 6)
End Function

' ============================================================
'  HELPER: Read clipboard text
' ============================================================
Private Function GetClipboardText() As String
    Dim obj As Object
    On Error GoTo ClipError
    Set obj = CreateObject("HTMLfile")
    GetClipboardText = obj.ParentWindow.ClipboardData.GetData("text")
    Set obj = Nothing
    Exit Function
ClipError:
    GetClipboardText = ""
End Function

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
    BackupData ws

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

    On Error Resume Next
    Set wsBak = ThisWorkbook.Sheets(BACKUP_SHEET)
    On Error GoTo 0

    If wsBak Is Nothing Then
        MsgBox "No backup found. Please use Reset first before trying to Undo.", _
               vbInformation, "No Backup"
        Exit Sub
    End If

    wsBak.Visible = xlSheetVisible
    bakRows = 0
    If wsBak.Cells(1, BACKUP_COUNT_COL).Value <> "" Then
        bakRows = CLng(wsBak.Cells(1, BACKUP_COUNT_COL).Value)
    End If
    wsBak.Visible = xlSheetVeryHidden

    If bakRows <= 0 Then
        MsgBox "Backup appears to be empty. Nothing to restore.", _
               vbInformation, "Nothing to Restore"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)

    ws.Range("A" & DATA_START_ROW & ":E1000").ClearContents
    ws.Range("A" & DATA_START_ROW & ":E1000").Interior.ColorIndex = xlNone
    ws.Range("A" & DATA_START_ROW & ":E1000").Font.Bold = False

    wsBak.Visible = xlSheetVisible

    Dim destRow As Long
    destRow = DATA_START_ROW

    For i = DATA_START_ROW To DATA_START_ROW + bakRows - 1
        ws.Cells(destRow, COL_DELIVERY).Value = wsBak.Cells(i, COL_DELIVERY).Value
        ws.Cells(destRow, COL_TRACKING).Value = wsBak.Cells(i, COL_TRACKING).Value
        ws.Cells(destRow, COL_CATEGORY).Value = wsBak.Cells(i, COL_CATEGORY).Value
        ws.Cells(destRow, COL_STATUS).Value = wsBak.Cells(i, COL_STATUS).Value
        ws.Cells(destRow, COL_QUANTITY).Value = wsBak.Cells(i, COL_QUANTITY).Value

        If wsBak.Cells(i, COL_STATUS).Interior.ColorIndex <> xlNone Then
            ws.Cells(destRow, COL_STATUS).Interior.Color = wsBak.Cells(i, COL_STATUS).Interior.Color
        End If
        ws.Cells(destRow, COL_STATUS).Font.Bold = wsBak.Cells(i, COL_STATUS).Font.Bold

        destRow = destRow + 1
    Next i

    wsBak.Visible = xlSheetVeryHidden
    MsgBox "Undo complete. " & bakRows & " rows restored.", vbInformation, "Undo Complete"
End Sub

' ============================================================
'  BACKUP
' ============================================================
Private Sub BackupData(ws As Worksheet)
    Dim wsBak     As Worksheet
    Dim i         As Long
    Dim bakRow    As Long
    Dim cellVal   As String
    Dim savedRows As Long

    On Error Resume Next
    Set wsBak = ThisWorkbook.Sheets(BACKUP_SHEET)
    On Error GoTo 0

    If wsBak Is Nothing Then
        Set wsBak = ThisWorkbook.Sheets.Add
        wsBak.Name = BACKUP_SHEET
    End If

    wsBak.Visible = xlSheetVisible
    wsBak.Cells.Clear
    ws.Rows(1).Copy wsBak.Rows(1)

    bakRow = DATA_START_ROW
    savedRows = 0

    For i = DATA_START_ROW To 1000
        cellVal = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))
        If cellVal <> "" And IsNumeric(cellVal) And Len(cellVal) >= 5 Then
            wsBak.Cells(bakRow, COL_DELIVERY).Value = ws.Cells(i, COL_DELIVERY).Value
            wsBak.Cells(bakRow, COL_TRACKING).Value = ws.Cells(i, COL_TRACKING).Value
            wsBak.Cells(bakRow, COL_CATEGORY).Value = ws.Cells(i, COL_CATEGORY).Value
            wsBak.Cells(bakRow, COL_STATUS).Value = ws.Cells(i, COL_STATUS).Value
            wsBak.Cells(bakRow, COL_QUANTITY).Value = ws.Cells(i, COL_QUANTITY).Value

            If ws.Cells(i, COL_STATUS).Interior.ColorIndex <> xlNone Then
                wsBak.Cells(bakRow, COL_STATUS).Interior.Color = ws.Cells(i, COL_STATUS).Interior.Color
            End If
            wsBak.Cells(bakRow, COL_STATUS).Font.Bold = ws.Cells(i, COL_STATUS).Font.Bold

            bakRow = bakRow + 1
            savedRows = savedRows + 1
        End If
    Next i

    wsBak.Cells(1, BACKUP_COUNT_COL).Value = savedRows
    wsBak.Visible = xlSheetVeryHidden
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
    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).row

    If lastRow < DATA_START_ROW Then
        MsgBox "No deliveries found.", vbInformation, "No Data"
        Exit Sub
    End If

    Set sapSession = GetSAPSession()
    If sapSession Is Nothing Then Exit Sub

    If MsgBox("Start PGI for Parts & Conversion now?", _
              vbYesNo + vbQuestion, "Confirm Start") = vbNo Then Exit Sub

    doneCount = 0: errorCount = 0: skipCount = 0

    For i = DATA_START_ROW To lastRow
        delivery = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))
        tracking = Trim(CStr(ws.Cells(i, COL_TRACKING).Value))
        category = UCase(Trim(CStr(ws.Cells(i, COL_CATEGORY).Value)))

        If delivery = "" Then GoTo NextRowPGI
        If InStr(ws.Cells(i, COL_STATUS).Value, "Done") > 0 Then GoTo NextRowPGI

        If category = "MACHINE" Or category = "MACHINES" Then
            If ws.Cells(i, COL_STATUS).Value = "" Then
                SetStatus ws, i, "Pending - Machine", "NONE"
            End If
            skipCount = skipCount + 1
            GoTo NextRowPGI
        End If

        SetStatus ws, i, "Processing...", "NONE"
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
    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).row

    If lastRow < DATA_START_ROW Then
        MsgBox "No deliveries found.", vbInformation, "No Data"
        Exit Sub
    End If

    Set sapSession = GetSAPSession()
    If sapSession Is Nothing Then Exit Sub

    If MsgBox("Start PGI for Machines now?", _
              vbYesNo + vbQuestion, "Confirm Start") = vbNo Then Exit Sub

    doneCount = 0: errorCount = 0: skipCount = 0

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

        SetStatus ws, i, "Processing...", "NONE"
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
'  RUN PICKING - Conversion only
' ============================================================
Public Sub RunConversionPicking()
    Dim ws         As Worksheet
    Dim sapSession As Object
    Dim lastRow    As Long
    Dim i          As Long
    Dim delivery   As String
    Dim category   As String
    Dim quantity   As Integer
    Dim doneCount  As Integer
    Dim errorCount As Integer
    Dim skipCount  As Integer

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).row

    If lastRow < DATA_START_ROW Then
        MsgBox "No deliveries found.", vbInformation, "No Data"
        Exit Sub
    End If

    Set sapSession = GetSAPSession()
    If sapSession Is Nothing Then Exit Sub

    If MsgBox("Start Picking for Conversion & Parts deliveries now?", _
              vbYesNo + vbQuestion, "Confirm Start") = vbNo Then Exit Sub

    doneCount = 0: errorCount = 0: skipCount = 0

    For i = DATA_START_ROW To lastRow
        delivery = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))
        category = UCase(Trim(CStr(ws.Cells(i, COL_CATEGORY).Value)))

        If delivery = "" Then GoTo NextRowConvPick
        If InStr(ws.Cells(i, COL_STATUS).Value, "Picked") > 0 Then GoTo NextRowConvPick

        If category = "CONVERSION" Then

            SetStatus ws, i, "Picking...", "NONE"
            DoEvents

            If ProcessConversionPicking(sapSession, delivery) Then
                SetStatus ws, i, "Picked - OK", "ORANGE"
                doneCount = doneCount + 1
            Else
                SetStatus ws, i, "ERROR - Check Manually", "RED"
                errorCount = errorCount + 1
            End If

        ElseIf category = "PART" Or category = "PARTS" Then

            quantity = 0
            If ws.Cells(i, COL_QUANTITY).Value <> "" Then
                quantity = CInt(ws.Cells(i, COL_QUANTITY).Value)
            End If

            If quantity <= 0 Then
                SetStatus ws, i, "ERROR - Enter Quantity in Column E", "RED"
                errorCount = errorCount + 1
                GoTo NextRowConvPick
            End If

            SetStatus ws, i, "Picking...", "NONE"
            DoEvents

            If ProcessPartsPicking(sapSession, delivery, quantity) Then
                SetStatus ws, i, "Picked - OK (" & quantity & " parts)", "ORANGE"
                doneCount = doneCount + 1
            Else
                SetStatus ws, i, "ERROR - Check Manually", "RED"
                errorCount = errorCount + 1
            End If

        Else
            skipCount = skipCount + 1
        End If

NextRowConvPick:
    Next i

    MsgBox "Picking Complete!" & vbNewLine & vbNewLine & _
           "Done:    " & doneCount & vbNewLine & _
           "Errors:  " & errorCount & vbNewLine & _
           "Skipped: " & skipCount, _
           vbInformation, "Process Complete"
End Sub

' ============================================================
'  RUN PICKING - Machines only
' ============================================================
Public Sub RunMachinePicking()
    Dim ws         As Worksheet
    Dim sapSession As Object
    Dim lastRow    As Long
    Dim i          As Long
    Dim delivery   As String
    Dim category   As String
    Dim quantity   As Integer
    Dim doneCount  As Integer
    Dim errorCount As Integer
    Dim skipCount  As Integer

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).row

    If lastRow < DATA_START_ROW Then
        MsgBox "No deliveries found.", vbInformation, "No Data"
        Exit Sub
    End If

    Set sapSession = GetSAPSession()
    If sapSession Is Nothing Then Exit Sub

    If MsgBox("Start Picking for Machine deliveries now?", _
              vbYesNo + vbQuestion, "Confirm Start") = vbNo Then Exit Sub

    doneCount = 0: errorCount = 0: skipCount = 0

    For i = DATA_START_ROW To lastRow
        delivery = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))
        category = UCase(Trim(CStr(ws.Cells(i, COL_CATEGORY).Value)))

        If delivery = "" Then GoTo NextRowMachPick
        If InStr(ws.Cells(i, COL_STATUS).Value, "Picked") > 0 Then GoTo NextRowMachPick

        If category <> "MACHINE" And category <> "MACHINES" Then
            skipCount = skipCount + 1
            GoTo NextRowMachPick
        End If

        quantity = 0
        If ws.Cells(i, COL_QUANTITY).Value <> "" Then
            quantity = CInt(ws.Cells(i, COL_QUANTITY).Value)
        End If

        If quantity <= 0 Then
            SetStatus ws, i, "ERROR - Enter Quantity in Column E", "RED"
            errorCount = errorCount + 1
            GoTo NextRowMachPick
        End If

        SetStatus ws, i, "Picking...", "NONE"
        DoEvents

        If ProcessMachinePicking(sapSession, delivery, quantity) Then
            SetStatus ws, i, "Picked - OK (" & quantity & " machines)", "ORANGE"
            doneCount = doneCount + 1
        Else
            SetStatus ws, i, "ERROR - Check Manually", "RED"
            errorCount = errorCount + 1
        End If

NextRowMachPick:
    Next i

    MsgBox "Machine Picking Complete!" & vbNewLine & vbNewLine & _
           "Done:    " & doneCount & vbNewLine & _
           "Errors:  " & errorCount & vbNewLine & _
           "Skipped: " & skipCount, _
           vbInformation, "Process Complete"
End Sub

' ============================================================
'  CORE: Conversion Picking
' ============================================================
Private Function ProcessConversionPicking(sapSession As Object, _
                                           delivery As String) As Boolean
    On Error GoTo HandleError

    Dim statusBar As String
    Dim rowIndex  As Integer

    sapSession.StartTransaction "VL02N"
    SAPWait WAIT_MEDIUM

    sapSession.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = delivery
    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "does not exist") > 0 Or _
       InStr(LCase(statusBar), "not found") > 0 Then GoTo HandleError

    sapSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
    SAPWait WAIT_SHORT

    rowIndex = 0
    Do
        On Error Resume Next
        sapSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/" & _
                            "ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/" & _
                            "txtLIPSD-PIKMG[6," & rowIndex & "]").Text = "1"
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo HandleError
            Exit Do
        End If
        On Error GoTo HandleError
        SAPWait 100
        rowIndex = rowIndex + 1
    Loop

    If rowIndex = 0 Then GoTo HandleError

    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_SHORT

    sapSession.findById("wnd[0]/tbar[0]/btn[11]").press
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "error") > 0 Then GoTo HandleError

    ProcessConversionPicking = True
    Exit Function

HandleError:
    On Error Resume Next
    ProcessConversionPicking = False
End Function

' ============================================================
'  CORE: Parts Picking (storage A200, quantity from Column E)
' ============================================================
Private Function ProcessPartsPicking(sapSession As Object, _
                                      delivery As String, _
                                      quantity As Integer) As Boolean
    On Error GoTo HandleError

    Dim statusBar As String
    Dim basePath  As String

    basePath = "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/" & _
               "ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/"

    sapSession.StartTransaction "VL02N"
    SAPWait WAIT_MEDIUM

    sapSession.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = delivery
    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "does not exist") > 0 Or _
       InStr(LCase(statusBar), "not found") > 0 Then GoTo HandleError

    sapSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
    SAPWait WAIT_SHORT

    sapSession.findById(basePath & "ctxtLIPS-LGORT[3,0]").Text = "A200"
    sapSession.findById(basePath & "txtLIPSD-PIKMG[6,0]").Text = CStr(quantity)
    SAPWait 100

    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_SHORT

    sapSession.findById("wnd[0]/tbar[0]/btn[11]").press
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "error") > 0 Then GoTo HandleError

    ProcessPartsPicking = True
    Exit Function

HandleError:
    On Error Resume Next
    ProcessPartsPicking = False
End Function

' ============================================================
'  CORE: Machine Picking
'
'  Detects whether batch splits are already configured by
'  pressing [9,0] and checking if a popup window (wnd[1])
'  appears.
'
'  PATH 1 - Popup appeared (batch splits NOT yet set up):
'    Close popup, go to ZVSER F00004, assign serial numbers
'    for all machines, come back to VL02N, then pick.
'
'  PATH 2 - No popup (batch splits ALREADY configured):
'    Collapse what we just opened, then pick directly.
'
'  Picking loop (both paths):
'    For each item: expand [9,j] -> fill [6,1]="1" ->
'    sendVKey 0 -> collapse [9,0]
' ============================================================
Private Function ProcessMachinePicking(sapSession As Object, _
                                        delivery As String, _
                                        quantity As Integer) As Boolean
    On Error GoTo HandleError

    Dim statusBar     As String
    Dim j             As Integer
    Dim basePath      As String
    Dim hasBatchSplit As Boolean
    Dim wndTitle      As String
    Dim shellPath     As String

    basePath = "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/" & _
               "ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/"
    shellPath = "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell"

    ' ── Open VL02N ──────────────────────────────────────────
    sapSession.StartTransaction "VL02N"
    SAPWait WAIT_MEDIUM
    sapSession.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = delivery
    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "does not exist") > 0 Or _
       InStr(LCase(statusBar), "not found") > 0 Then GoTo HandleError

    sapSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
    SAPWait WAIT_SHORT

    ' ── Press expand button on first item to test mode ──────
    sapSession.findById(basePath & "btnRV50A-CHMULT[9,0]").SetFocus
    SAPWait 100
    sapSession.findById(basePath & "btnRV50A-CHMULT[9,0]").press
    SAPWait WAIT_SHORT

    ' Detect popup: if wnd[1] exists a batch-split config
    ' popup appeared, meaning serials are not assigned yet.
    hasBatchSplit = False
    On Error Resume Next
    wndTitle = sapSession.findById("wnd[1]").Text
    If Err.Number = 0 Then hasBatchSplit = True
    Err.Clear
    On Error GoTo HandleError

    ' ── PATH 1: Popup appeared - set up batch splits first ──
    If hasBatchSplit Then
        ' Close the popup cleanly
        sapSession.findById("wnd[1]/tbar[0]/btn[12]").press
        SAPWait WAIT_SHORT
        ' Back to SAP Easy Access
        sapSession.findById("wnd[0]/tbar[0]/btn[3]").press
        SAPWait WAIT_MEDIUM

        ' Navigate to ZVSER (F00004) and assign serial numbers
        sapSession.findById(shellPath).selectedNode = "F00004"
        SAPWait WAIT_SHORT
        sapSession.findById(shellPath).doubleClickNode "F00004"
        SAPWait WAIT_MEDIUM

        sapSession.findById("wnd[0]/usr/ctxtIT_VBELN-LOW").Text = delivery
        sapSession.findById("wnd[0]/usr/ctxtIT_VBELN-LOW").SetFocus
        SAPWait 100

        sapSession.findById("wnd[0]/tbar[1]/btn[8]").press
        SAPWait WAIT_MEDIUM

        ' Select all delivery rows
        sapSession.findById("wnd[0]/usr/btnTC_LIPS_MARK").press
        SAPWait WAIT_SHORT

        ' Click Batch Split button
        sapSession.findById("wnd[0]/tbar[1]/btn[5]").press
        SAPWait WAIT_MEDIUM

        ' Assign one serial number per machine
        For j = 1 To quantity
            sapSession.findById("wnd[1]/usr/tblZSDE_BATCH_SPLIT_FOR_DELIVERYTC_OBJKA") _
                .getAbsoluteRow(0).Selected = True
            SAPWait 100
            sapSession.findById("wnd[1]/usr/tblZSDE_BATCH_SPLIT_FOR_DELIVERYTC_OBJKA/" & _
                                "txtI_OBJKA-SERNR[0,0]").SetFocus
            SAPWait 100
            sapSession.findById("wnd[1]/tbar[0]/btn[5]").press
            SAPWait WAIT_MEDIUM
        Next j

        ' Go back to SAP Easy Access (3x back)
        sapSession.findById("wnd[0]/tbar[0]/btn[12]").press
        SAPWait WAIT_SHORT
        sapSession.findById("wnd[0]/tbar[0]/btn[12]").press
        SAPWait WAIT_SHORT
        sapSession.findById("wnd[0]/tbar[0]/btn[12]").press
        SAPWait WAIT_MEDIUM

        ' Re-open VL02N ready for picking
        sapSession.StartTransaction "VL02N"
        SAPWait WAIT_MEDIUM
        sapSession.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = delivery
        sapSession.findById("wnd[0]").sendVKey 0
        SAPWait WAIT_MEDIUM
        sapSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
        SAPWait WAIT_SHORT

    ' ── PATH 2: No popup - batch splits already configured ──
    Else
        ' Collapse the row we just expanded before the loop
        sapSession.findById(basePath & "btnRV50A-CHMULT[9,0]").SetFocus
        SAPWait 100
        sapSession.findById(basePath & "btnRV50A-CHMULT[9,0]").press
        SAPWait WAIT_SHORT
    End If

    ' ── Picking loop (same for both paths) ──────────────────
    ' Each item: expand toggle, fill picked qty = 1, sendVKey 0, collapse toggle.
    ' After expansion SAP scrolls that item to row 0, so collapse is always [9,0].
    For j = 0 To quantity - 1
        sapSession.findById(basePath & "btnRV50A-CHMULT[9," & j & "]").SetFocus
        SAPWait 100
        sapSession.findById(basePath & "btnRV50A-CHMULT[9," & j & "]").press
        SAPWait WAIT_SHORT

        sapSession.findById(basePath & "txtLIPSD-PIKMG[6,1]").Text = "1"
        SAPWait 100
        sapSession.findById("wnd[0]").sendVKey 0
        SAPWait 100

        sapSession.findById(basePath & "btnRV50A-CHMULT[9,0]").SetFocus
        SAPWait 100
        sapSession.findById(basePath & "btnRV50A-CHMULT[9,0]").press
        SAPWait WAIT_MEDIUM
    Next j

    ' ── Save ────────────────────────────────────────────────
    sapSession.findById("wnd[0]/tbar[0]/btn[11]").press
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "error") > 0 Or _
       InStr(LCase(statusBar), "not possible") > 0 Then GoTo HandleError

    ProcessMachinePicking = True
    Exit Function

HandleError:
    On Error Resume Next
    ProcessMachinePicking = False
End Function

' ============================================================
'  CORE: Parts & Conversion PGI
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
'  CORE: Machine PGI
' ============================================================
Private Function ProcessMachine(sapSession As Object, _
                                 delivery As String, _
                                 tracking As String, _
                                 quantity As Integer) As Boolean
    On Error GoTo HandleError

    Dim statusBar As String
    Dim j         As Integer

    sapSession.StartTransaction "VL02N"
    SAPWait WAIT_MEDIUM

    sapSession.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = delivery
    sapSession.findById("wnd[0]").sendVKey 0
    SAPWait WAIT_MEDIUM

    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "does not exist") > 0 Or _
       InStr(LCase(statusBar), "not found") > 0 Then GoTo HandleError

    sapSession.findById("wnd[0]/tbar[1]/btn[8]").press
    SAPWait WAIT_SHORT

    sapSession.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\04/" & _
                        "ssubSUBSCREEN_BODY:SAPMV50A:2108/txtLIKP-BOLNR").Text = tracking
    SAPWait WAIT_SHORT

    sapSession.findById("wnd[0]/tbar[0]/btn[11]").press
    SAPWait 5000

    For j = 1 To quantity
        sapSession.findById("wnd[0]/tbar[1]/btn[20]").press
        SAPWait WAIT_MEDIUM

        sapSession.findById("wnd[1]/tbar[0]/btn[0]").press
        SAPWait WAIT_SHORT

        sapSession.findById("wnd[0]/tbar[0]/btn[11]").press
        SAPWait WAIT_MEDIUM

        If j < quantity Then
            sapSession.findById("wnd[0]/tbar[0]/btn[3]").press
            SAPWait WAIT_SHORT
        End If
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

    Set sapGui = GetObject("SAPGUI")
    Set sapApp = sapGui.GetScriptingEngine
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
'  HELPER: Set status with color rules
' ============================================================
Private Sub SetStatus(ws As Worksheet, row As Long, _
                      statusText As String, colorName As String)
    With ws.Cells(row, COL_STATUS)
        .Value = statusText
        Select Case colorName
            Case "GREEN"
                .Interior.Color = RGB(144, 238, 144)
                .Font.Bold = False
            Case "RED"
                .Interior.Color = RGB(255, 99, 71)
                .Font.Bold = True
            Case "ORANGE"
                .Interior.Color = RGB(255, 165, 0)
                .Font.Bold = False
            Case Else
                .Interior.ColorIndex = xlNone
                .Font.Bold = False
        End Select
    End With
    DoEvents
End Sub

' ============================================================
'  FIX COLORS FOR EXISTING ROWS
' ============================================================
Public Sub FixStatusColors()
    Dim ws      As Worksheet
    Dim i       As Long
    Dim lastRow As Long
    Dim status  As String

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).row

    For i = DATA_START_ROW To lastRow
        status = LCase(Trim(CStr(ws.Cells(i, COL_STATUS).Value)))

        With ws.Cells(i, COL_STATUS)
            If InStr(status, "done") > 0 Then
                .Interior.Color = RGB(144, 238, 144)
                .Font.Bold = False
            ElseIf InStr(status, "error") > 0 Then
                .Interior.Color = RGB(255, 99, 71)
                .Font.Bold = True
            ElseIf InStr(status, "picked") > 0 Then
                .Interior.Color = RGB(255, 165, 0)
                .Font.Bold = False
            ElseIf status = "" Then
                .Interior.ColorIndex = xlNone
                .Font.Bold = False
            End If
        End With
    Next i

    MsgBox "Status colors updated!", vbInformation, "Done"
End Sub

' ============================================================
'  ALIGN BUTTONS
' ============================================================
Sub AlignButtons()
    Dim ws  As Worksheet
    Dim btn As Button
    Set ws = ThisWorkbook.Sheets("Deliveries")

    Dim btnLeft   As Double
    Dim btnWidth  As Double
    Dim btnHeight As Double
    Dim startTop  As Double
    Dim gap       As Double
    Dim placed    As Integer

    Dim btnOrder(8) As String
    btnOrder(0) = "IMPORT DELIVERY (from PDF)"
    btnOrder(1) = "RUN PICKING  (Conversion)"
    btnOrder(2) = "RUN PICKING  (Machines)"
    btnOrder(3) = "IMPORT TRACKING (from web)"
    btnOrder(4) = "RUN PGI  (Conversion & Parts)"
    btnOrder(5) = "RUN PGI  (Machines)"
    btnOrder(6) = "RESET TEMPLATE"
    btnOrder(7) = "UNDO RESET"
    btnOrder(8) = "Create Delivery number"

    btnLeft = ws.Columns("H").Left + 5
    btnWidth = 180
    btnHeight = 28
    startTop = ws.Rows(2).Top
    gap = 10
    placed = 0

    Dim j As Integer
    For j = 0 To 8
        For Each btn In ws.Buttons
            If Trim(btn.Caption) = Trim(btnOrder(j)) Then
                btn.Left = btnLeft
                btn.Top = startTop + placed * (btnHeight + gap)
                btn.Width = btnWidth
                btn.Height = btnHeight
                placed = placed + 1
                Exit For
            End If
        Next btn
    Next j

    MsgBox "All buttons aligned!", vbInformation, "Done"
End Sub
