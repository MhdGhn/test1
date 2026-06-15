'================================================================
' SAP PGI TOOL - AUTO BUILDER
' Double-click this file to create the Excel tool on your Desktop
'================================================================

Dim xlApp, xlWb, xlWs, xlMod
Dim btnReset, btnPGI
Dim savePath

savePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\SAP_PGI_Tool.xlsm"

MsgBox "This will create SAP_PGI_Tool.xlsm on your Desktop." & vbNewLine & _
       "Click OK to continue.", vbInformation, "SAP PGI Tool Builder"

' --- Create Excel ---
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set xlWb = xlApp.Workbooks.Add

' --- Set up sheet ---
Set xlWs = xlWb.Sheets(1)
xlWs.Name = "Deliveries"

' --- Column widths ---
xlWs.Columns("A").ColumnWidth = 22
xlWs.Columns("B").ColumnWidth = 28
xlWs.Columns("C").ColumnWidth = 16
xlWs.Columns("D").ColumnWidth = 28

' --- Header row ---
With xlWs.Range("A1:D1")
    .Value = Array("Delivery Number", "Tracking Number", "Category", "Status")
    .Font.Bold = True
    .Font.Color = RGB(255, 255, 255)
    .Interior.Color = RGB(31, 73, 125)
    .HorizontalAlignment = -4108  ' xlCenter
    .RowHeight = 24
End With

' --- Freeze top row ---
xlWs.Rows("2:2").Select
xlApp.ActiveWindow.FreezePanes = True
xlWs.Range("A2").Select

' --- Category dropdown (A2:C101 for up to 100 rows) ---
Dim valRange
Set valRange = xlWs.Range("C2:C101")
With xlWb.ActiveSheet.Range("C2:C101").Validation
    .Delete
    .Add 3, 1, 1, "PARTS,CONVERSION,MACHINE"
    .ShowError = True
    .ErrorTitle = "Invalid Category"
    .ErrorMessage = "Please select: PARTS, CONVERSION or MACHINE"
End With

' --- Alternating row colors for readability ---
Dim r As Integer
For r = 2 To 101
    If r Mod 2 = 0 Then
        xlWs.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 242, 242)
    End If
Next r

' --- Button area - Row above data (merge some cells for buttons) ---
xlWs.Rows("1:1").RowHeight = 24

' Add a row at top for buttons by inserting a row
xlWs.Rows("1:1").Insert

' Button row formatting
With xlWs.Range("A1:D1")
    .Interior.Color = RGB(240, 240, 240)
    .RowHeight = 36
End With

' Re-do headers in row 2
With xlWs.Range("A2:D2")
    .Value = Array("Delivery Number", "Tracking Number", "Category", "Status")
    .Font.Bold = True
    .Font.Color = RGB(255, 255, 255)
    .Interior.Color = RGB(31, 73, 125)
    .HorizontalAlignment = -4108
    .RowHeight = 24
End With

' --- Add RESET Button ---
Set btnReset = xlWs.Buttons.Add(5, 5, 130, 26)
With btnReset
    .Caption = "RESET DELIVERIES"
    .OnAction = "ResetDeliveries"
    .Font.Bold = True
    .Font.Size = 9
End With

' --- Add RUN PGI Button ---
Set btnPGI = xlWs.Buttons.Add(145, 5, 200, 26)
With btnPGI
    .Caption = "RUN PGI  (Parts & Conversion)"
    .OnAction = "RunPGI"
    .Font.Bold = True
    .Font.Size = 9
End With

' --- Freeze row 2 (header) ---
xlWs.Range("A3").Select
xlApp.ActiveWindow.FreezePanes = False
xlApp.ActiveWindow.FreezePanes = True
xlWs.Range("A3").Select

' --- Add VBA Code ---
Set xlMod = xlWb.VBProject.VBComponents.Add(1)  ' 1 = vbext_ct_StdModule
xlMod.Name = "SAP_PGI_Automation"

Dim code As String
code = "Option Explicit" & vbNewLine & _
"" & vbNewLine & _
"Private Const SHEET_NAME     As String  = ""Deliveries""" & vbNewLine & _
"Private Const COL_DELIVERY   As Integer = 1" & vbNewLine & _
"Private Const COL_TRACKING   As Integer = 2" & vbNewLine & _
"Private Const COL_CATEGORY   As Integer = 3" & vbNewLine & _
"Private Const COL_STATUS     As Integer = 4" & vbNewLine & _
"Private Const DATA_START_ROW As Integer = 3" & vbNewLine & _
"Private Const WAIT_SHORT     As Long    = 1500" & vbNewLine & _
"Private Const WAIT_MEDIUM    As Long    = 2500" & vbNewLine & _
"" & vbNewLine & _
"Public Sub ResetDeliveries()" & vbNewLine & _
"    Dim ws      As Worksheet" & vbNewLine & _
"    Dim lastRow As Long" & vbNewLine & _
"    Dim answer  As Integer" & vbNewLine & _
"    answer = MsgBox(""This will delete ALL delivery and tracking numbers."" & vbNewLine & ""Are you sure?"", vbYesNo + vbQuestion, ""Reset Confirmation"")" & vbNewLine & _
"    If answer = vbNo Then Exit Sub" & vbNewLine & _
"    Set ws = ThisWorkbook.Sheets(SHEET_NAME)" & vbNewLine & _
"    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).Row" & vbNewLine & _
"    If lastRow >= DATA_START_ROW Then" & vbNewLine & _
"        ws.Range(""A"" & DATA_START_ROW & "":D"" & lastRow).ClearContents" & vbNewLine & _
"        ws.Range(""A"" & DATA_START_ROW & "":D"" & lastRow).Interior.ColorIndex = xlNone" & vbNewLine & _
"        Dim r As Long" & vbNewLine & _
"        For r = DATA_START_ROW To DATA_START_ROW + 98" & vbNewLine & _
"            If r Mod 2 <> 0 Then ws.Range(""A"" & r & "":D"" & r).Interior.Color = RGB(242, 242, 242)" & vbNewLine & _
"        Next r" & vbNewLine & _
"    End If" & vbNewLine & _
"    ws.Cells(DATA_START_ROW, COL_DELIVERY).Select" & vbNewLine & _
"    MsgBox ""All deliveries cleared. Ready for new entries."", vbInformation, ""Reset Complete""" & vbNewLine & _
"End Sub" & vbNewLine & _
"" & vbNewLine & _
"Public Sub RunPGI()" & vbNewLine & _
"    Dim ws         As Worksheet" & vbNewLine & _
"    Dim sapSession As Object" & vbNewLine & _
"    Dim lastRow    As Long" & vbNewLine & _
"    Dim i          As Long" & vbNewLine & _
"    Dim delivery   As String" & vbNewLine & _
"    Dim tracking   As String" & vbNewLine & _
"    Dim category   As String" & vbNewLine & _
"    Dim doneCount  As Integer" & vbNewLine & _
"    Dim errCount   As Integer" & vbNewLine & _
"    Dim skipCount  As Integer" & vbNewLine & _
"    Set ws = ThisWorkbook.Sheets(SHEET_NAME)" & vbNewLine & _
"    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).Row" & vbNewLine & _
"    If lastRow < DATA_START_ROW Then" & vbNewLine & _
"        MsgBox ""No deliveries found. Please enter delivery numbers first."", vbInformation, ""No Data""" & vbNewLine & _
"        Exit Sub" & vbNewLine & _
"    End If" & vbNewLine & _
"    Set sapSession = GetSAPSession()" & vbNewLine & _
"    If sapSession Is Nothing Then Exit Sub" & vbNewLine & _
"    Dim total As Long" & vbNewLine & _
"    total = lastRow - DATA_START_ROW + 1" & vbNewLine & _
"    If MsgBox(""Ready to process "" & total & "" rows."" & vbNewLine & ""SAP must be open and logged in."" & vbNewLine & vbNewLine & ""Start PGI process now?"", vbYesNo + vbQuestion, ""Confirm Start"") = vbNo Then Exit Sub" & vbNewLine & _
"    doneCount = 0 : errCount = 0 : skipCount = 0" & vbNewLine & _
"    For i = DATA_START_ROW To lastRow" & vbNewLine & _
"        delivery = Trim(CStr(ws.Cells(i, COL_DELIVERY).Value))" & vbNewLine & _
"        tracking = Trim(CStr(ws.Cells(i, COL_TRACKING).Value))" & vbNewLine & _
"        category = UCase(Trim(CStr(ws.Cells(i, COL_CATEGORY).Value)))" & vbNewLine & _
"        If delivery = """" Then GoTo NextRow" & vbNewLine & _
"        If InStr(ws.Cells(i, COL_STATUS).Value, ""Done"") > 0 Then GoTo NextRow" & vbNewLine & _
"        If category = ""MACHINE"" Or category = ""MACHINES"" Then" & vbNewLine & _
"            SetStatus ws, i, ""Pending - Machine"", ""YELLOW""" & vbNewLine & _
"            skipCount = skipCount + 1" & vbNewLine & _
"            GoTo NextRow" & vbNewLine & _
"        End If" & vbNewLine & _
"        SetStatus ws, i, ""Processing..."", ""BLUE""" & vbNewLine & _
"        DoEvents" & vbNewLine & _
"        If ProcessSingleDelivery(sapSession, delivery, tracking) Then" & vbNewLine & _
"            SetStatus ws, i, ""Done - PGI Posted"", ""GREEN""" & vbNewLine & _
"            doneCount = doneCount + 1" & vbNewLine & _
"        Else" & vbNewLine & _
"            SetStatus ws, i, ""ERROR - Check Manually"", ""RED""" & vbNewLine & _
"            errCount = errCount + 1" & vbNewLine & _
"        End If" & vbNewLine & _
"NextRow:" & vbNewLine & _
"    Next i" & vbNewLine & _
"    MsgBox ""PGI Process Complete!"" & vbNewLine & vbNewLine & ""Done:    "" & doneCount & vbNewLine & ""Errors:  "" & errCount & vbNewLine & ""Skipped: "" & skipCount & vbNewLine & vbNewLine & ""Check Status column for details."", vbInformation, ""Process Complete""" & vbNewLine & _
"End Sub" & vbNewLine & _
"" & vbNewLine & _
"Private Function ProcessSingleDelivery(sapSession As Object, delivery As String, tracking As String) As Boolean" & vbNewLine & _
"    On Error GoTo HandleError" & vbNewLine & _
"    sapSession.StartTransaction ""VL02N""" & vbNewLine & _
"    SAPWait WAIT_MEDIUM" & vbNewLine & _
"    sapSession.findById(""wnd[0]/usr/ctxtLIKP-VBELN"").Text = delivery" & vbNewLine & _
"    sapSession.findById(""wnd[0]"").sendVKey 0" & vbNewLine & _
"    SAPWait WAIT_MEDIUM" & vbNewLine & _
"    Dim statusBar As String" & vbNewLine & _
"    statusBar = sapSession.findById(""wnd[0]/sbar"").Text" & vbNewLine & _
"    If InStr(LCase(statusBar), ""does not exist"") > 0 Or InStr(LCase(statusBar), ""not found"") > 0 Then GoTo HandleError" & vbNewLine & _
"    sapSession.findById(""wnd[0]/mbar/menu[1]/menu[0]"").Select" & vbNewLine & _
"    SAPWait WAIT_SHORT" & vbNewLine & _
"    sapSession.findById(""wnd[0]/usr/tabsTABSHEAD/tabpTABSH"").Select" & vbNewLine & _
"    SAPWait WAIT_SHORT" & vbNewLine & _
"    sapSession.findById(""wnd[0]/usr/tabsTABSHEAD/tabpTABSH/ssubSUBTABSHEAD:SAPMV50A:1112/ctxtLIKP-BOLNR"").Text = tracking" & vbNewLine & _
"    SAPWait WAIT_SHORT" & vbNewLine & _
"    sapSession.findById(""wnd[0]/tbar[1]/btn[8]"").press" & vbNewLine & _
"    SAPWait WAIT_MEDIUM" & vbNewLine & _
"    statusBar = sapSession.findById(""wnd[0]/sbar"").Text" & vbNewLine & _
"    If InStr(LCase(statusBar), ""error"") > 0 Or InStr(LCase(statusBar), ""not possible"") > 0 Then GoTo HandleError" & vbNewLine & _
"    ProcessSingleDelivery = True" & vbNewLine & _
"    Exit Function" & vbNewLine & _
"HandleError:" & vbNewLine & _
"    ProcessSingleDelivery = False" & vbNewLine & _
"End Function" & vbNewLine & _
"" & vbNewLine & _
"Private Function GetSAPSession() As Object" & vbNewLine & _
"    Dim sapGui As Object, sapApp As Object, sapConn As Object, sapSess As Object" & vbNewLine & _
"    On Error GoTo NoSAP" & vbNewLine & _
"    Set sapGui  = GetObject(""SAPGUI"")" & vbNewLine & _
"    Set sapApp  = sapGui.GetScriptingEngine" & vbNewLine & _
"    Set sapConn = sapApp.Children(0)" & vbNewLine & _
"    Set sapSess = sapConn.Children(0)" & vbNewLine & _
"    Set GetSAPSession = sapSess" & vbNewLine & _
"    Exit Function" & vbNewLine & _
"NoSAP:" & vbNewLine & _
"    MsgBox ""Cannot connect to SAP GUI."" & vbNewLine & vbNewLine & ""Please make sure:"" & vbNewLine & ""1. SAP GUI is open on your PC"" & vbNewLine & ""2. You are logged into SAP"" & vbNewLine & ""3. SAP GUI Scripting is enabled"", vbCritical, ""SAP Connection Error""" & vbNewLine & _
"    Set GetSAPSession = Nothing" & vbNewLine & _
"End Function" & vbNewLine & _
"" & vbNewLine & _
"Private Sub SAPWait(milliseconds As Long)" & vbNewLine & _
"    Application.Wait Now + (milliseconds / 86400000#)" & vbNewLine & _
"End Sub" & vbNewLine & _
"" & vbNewLine & _
"Private Sub SetStatus(ws As Worksheet, row As Long, statusText As String, colorName As String)" & vbNewLine & _
"    With ws.Cells(row, COL_STATUS)" & vbNewLine & _
"        .Value = statusText" & vbNewLine & _
"        Select Case colorName" & vbNewLine & _
"            Case ""GREEN""  : .Interior.Color = RGB(144, 238, 144)" & vbNewLine & _
"            Case ""RED""    : .Interior.Color = RGB(255, 99, 71)" & vbNewLine & _
"            Case ""YELLOW"" : .Interior.Color = RGB(255, 255, 153)" & vbNewLine & _
"            Case ""ORANGE"" : .Interior.Color = RGB(255, 200, 100)" & vbNewLine & _
"            Case ""BLUE""   : .Interior.Color = RGB(173, 216, 230)" & vbNewLine & _
"            Case Else     : .Interior.ColorIndex = xlNone" & vbNewLine & _
"        End Select" & vbNewLine & _
"        .Font.Bold = (colorName = ""RED"")" & vbNewLine & _
"    End With" & vbNewLine & _
"    DoEvents" & vbNewLine & _
"End Sub"

xlMod.CodeModule.AddFromString code

' --- Save as .xlsm ---
xlWb.SaveAs savePath, 52   ' 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
xlWb.Close False
xlApp.Quit

Set xlMod    = Nothing
Set btnPGI   = Nothing
Set btnReset = Nothing
Set xlWs     = Nothing
Set xlWb     = Nothing
Set xlApp    = Nothing

MsgBox "SUCCESS!" & vbNewLine & vbNewLine & _
       "SAP_PGI_Tool.xlsm has been created on your Desktop." & vbNewLine & vbNewLine & _
       "When you open it:" & vbNewLine & _
       "1. Click 'Enable Macros' if prompted" & vbNewLine & _
       "2. Make sure SAP GUI is open and you are logged in" & vbNewLine & _
       "3. Enter your deliveries and click RUN PGI", _
       vbInformation, "SAP PGI Tool Ready"
