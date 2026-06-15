Attribute VB_Name = "SAP_PGI_Automation"
Option Explicit

' ============================================================
'  CONSTANTS - Adjust these if your SAP screen layout differs
' ============================================================
Private Const SHEET_NAME        As String = "Deliveries"
Private Const COL_DELIVERY      As Integer = 1   ' Column A
Private Const COL_TRACKING      As Integer = 2   ' Column B
Private Const COL_CATEGORY      As Integer = 3   ' Column C
Private Const COL_STATUS        As Integer = 4   ' Column D
Private Const DATA_START_ROW    As Integer = 2   ' Row 2 (row 1 = headers)

' SAP screen wait times (milliseconds)
Private Const WAIT_SHORT        As Long = 1500
Private Const WAIT_MEDIUM       As Long = 2500

' ============================================================
'  RESET BUTTON - Clears all delivery data, keeps headers
' ============================================================
Public Sub ResetDeliveries()
    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim answer  As Integer

    answer = MsgBox("This will delete ALL delivery and tracking numbers." & vbNewLine & _
                    "Are you sure you want to reset?", _
                    vbYesNo + vbQuestion, "Reset Confirmation")

    If answer = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).Row

    If lastRow >= DATA_START_ROW Then
        ws.Range("A" & DATA_START_ROW & ":D" & lastRow).ClearContents
        ws.Range("A" & DATA_START_ROW & ":D" & lastRow).Interior.ColorIndex = xlNone
    End If

    ws.Cells(DATA_START_ROW, COL_DELIVERY).Select
    MsgBox "All deliveries cleared. Ready for new entries.", vbInformation, "Reset Complete"
End Sub

' ============================================================
'  MAIN PGI BUTTON - Processes all Parts & Conversion rows
' ============================================================
Public Sub RunPGI()
    Dim ws          As Worksheet
    Dim sapSession  As Object
    Dim lastRow     As Long
    Dim i           As Long
    Dim delivery    As String
    Dim tracking    As String
    Dim category    As String
    Dim doneCount   As Integer
    Dim errorCount  As Integer
    Dim skipCount   As Integer

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).Row

    If lastRow < DATA_START_ROW Then
        MsgBox "No deliveries found. Please enter delivery numbers first.", _
               vbInformation, "No Data"
        Exit Sub
    End If

    ' --- Connect to SAP ---
    Set sapSession = GetSAPSession()
    If sapSession Is Nothing Then Exit Sub

    ' --- Confirm before starting ---
    Dim total As Long
    total = lastRow - DATA_START_ROW + 1
    If MsgBox("Ready to process " & total & " rows." & vbNewLine & _
              "SAP must be open and logged in." & vbNewLine & vbNewLine & _
              "Start PGI process now?", _
              vbYesNo + vbQuestion, "Confirm Start") = vbNo Then Exit Sub

    doneCount = 0
    errorCount = 0
    skipCount = 0

    ' --- Loop through each row ---
    For i = DATA_START_ROW To lastRow
        delivery = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))
        tracking = Trim(CStr(ws.Cells(i, COL_TRACKING).Value))
        category = UCase(Trim(CStr(ws.Cells(i, COL_CATEGORY).Value)))

        ' Skip empty rows
        If delivery = "" Then GoTo NextRow

        ' Skip already processed rows
        If ws.Cells(i, COL_STATUS).Value <> "" And _
           InStr(ws.Cells(i, COL_STATUS).Value, "Done") > 0 Then
            GoTo NextRow
        End If

        ' Skip Machines (different workflow - handled separately)
        If category = "MACHINE" Or category = "MACHINES" Then
            SetStatus ws, i, "Pending - Machine", "YELLOW"
            skipCount = skipCount + 1
            GoTo NextRow
        End If

        ' Skip rows with no category that are not Parts/Conversion
        If category <> "PARTS" And category <> "PART" And _
           category <> "CONVERSION" And category <> "CONV" Then
            If category <> "" Then
                SetStatus ws, i, "Skipped - Unknown Category", "ORANGE"
                skipCount = skipCount + 1
                GoTo NextRow
            End If
        End If

        ' --- Process this delivery ---
        SetStatus ws, i, "Processing...", "BLUE"
        DoEvents

        If ProcessSingleDelivery(sapSession, delivery, tracking) Then
            SetStatus ws, i, "Done - PGI Posted", "GREEN"
            doneCount = doneCount + 1
        Else
            SetStatus ws, i, "ERROR - Check Manually", "RED"
            errorCount = errorCount + 1
        End If

NextRow:
    Next i

    ' --- Summary ---
    MsgBox "PGI Process Complete!" & vbNewLine & vbNewLine & _
           "Done:    " & doneCount & vbNewLine & _
           "Errors:  " & errorCount & vbNewLine & _
           "Skipped: " & skipCount & vbNewLine & vbNewLine & _
           "Check the Status column for details.", _
           vbInformation, "Process Complete"
End Sub

' ============================================================
'  CORE: Open VL02N, fill delivery + tracking, post PGI
' ============================================================
Private Function ProcessSingleDelivery(sapSession As Object, _
                                        delivery As String, _
                                        tracking As String) As Boolean
    On Error GoTo HandleError

    ' 1. Open VL02N transaction
    sapSession.StartTransaction "VL02N"
    SAPWait WAIT_MEDIUM

    ' 2. Enter delivery number and press Enter
    sapSession.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = delivery
    sapSession.findById("wnd[0]").sendVKey 0   ' 0 = Enter
    SAPWait WAIT_MEDIUM

    ' 3. Check if delivery opened correctly (look for error in status bar)
    Dim statusBar As String
    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "does not exist") > 0 Or _
       InStr(LCase(statusBar), "not found") > 0 Or _
       InStr(LCase(statusBar), "no authorization") > 0 Then
        GoTo HandleError
    End If

    ' 4. Click "Header" menu to open Header Details
    '    In VL02N: Menu path is Edit > Header... OR toolbar button
    '    NOTE: If this line fails, record a SAP script to find the correct ID
    sapSession.findById("wnd[0]/mbar/menu[1]/menu[0]").Select   ' Edit > Header
    SAPWait WAIT_SHORT

    ' 5. Click on Shipment tab
    '    NOTE: Tab ID may differ in your system - record to verify
    sapSession.findById("wnd[0]/usr/tabsTABSHEAD/tabpTABSH").Select
    SAPWait WAIT_SHORT

    ' 6. Enter tracking number in Bill of Lading field
    sapSession.findById("wnd[0]/usr/tabsTABSHEAD/tabpTABSH/" & _
                        "ssubSUBTABSHEAD:SAPMV50A:1112/ctxtLIKP-BOLNR").Text = tracking
    SAPWait WAIT_SHORT

    ' 7. Click Post Goods Issue button
    '    Try toolbar button first (btn[8] is common for PGI in VL02N)
    '    NOTE: Record a SAP script to confirm the correct button ID
    sapSession.findById("wnd[0]/tbar[1]/btn[8]").press
    SAPWait WAIT_MEDIUM

    ' 8. Check status bar for success
    statusBar = sapSession.findById("wnd[0]/sbar").Text
    If InStr(LCase(statusBar), "error") > 0 Or _
       InStr(LCase(statusBar), "not possible") > 0 Then
        GoTo HandleError
    End If

    ProcessSingleDelivery = True
    Exit Function

HandleError:
    ProcessSingleDelivery = False
End Function

' ============================================================
'  HELPER: Connect to open SAP GUI session
' ============================================================
Private Function GetSAPSession() As Object
    Dim sapGui    As Object
    Dim sapApp    As Object
    Dim sapConn   As Object
    Dim sapSess   As Object

    On Error GoTo NoSAP

    Set sapGui  = GetObject("SAPGUI")
    Set sapApp  = sapGui.GetScriptingEngine
    Set sapConn = sapApp.Children(0)    ' First open connection
    Set sapSess = sapConn.Children(0)   ' First session

    Set GetSAPSession = sapSess
    Exit Function

NoSAP:
    MsgBox "Cannot connect to SAP GUI." & vbNewLine & vbNewLine & _
           "Please make sure:" & vbNewLine & _
           "1. SAP GUI is open on your PC" & vbNewLine & _
           "2. You are logged into SAP" & vbNewLine & _
           "3. SAP GUI Scripting is enabled" & vbNewLine & vbNewLine & _
           "To enable scripting: SAP GUI > Help > Settings > Scripting tab", _
           vbCritical, "SAP Connection Error"
    Set GetSAPSession = Nothing
End Function

' ============================================================
'  HELPER: Wait for SAP screen to load
' ============================================================
Private Sub SAPWait(milliseconds As Long)
    Application.Wait Now + (milliseconds / 86400000#)
End Sub

' ============================================================
'  HELPER: Set status cell color and text
' ============================================================
Private Sub SetStatus(ws As Worksheet, row As Integer, _
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
