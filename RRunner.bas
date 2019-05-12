Attribute VB_Name = "RRunner"
Option Explicit
' ###################################################################
' Configuration Parameters
' ###################################################################
' Path to the R Scripts and where the temporary files will be created
Private Const R_SCRIPTS_PATH = ".\r"
Private Const WORKING_PATH = ".\tmp"
' Time to wait for the R Script answer in milliseconds
Private Const TIMEOUT_MILLISECONDS = 10000
Private Const R_IN_FILE_NAME = "_RInput_"
Private Const R_OUT_FILE_NAME = "_ROutput_"

' ###################################################################
' License information
' ###################################################################
' Copyright 2019 fizban99
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' ###################################################################


Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal HWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal HWnd As Long, ByVal wCmd As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare Function KillTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function SetTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
   
Private Const WM_CHAR As Long = &H102
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5


Private WorkingPath As String
Private RScriptsPath As String
Private mInRange As Boolean
Private mCalculatedCells2R As New Collection
Private mCalculatedCells2P As New Collection
Private mWindowsTimerID As Long
Private mApplicationTimerTime As Date
Private mInProcessP As Boolean
Private mInProcessR As Boolean


Private Sub SetStatus(s As String)
    Application.StatusBar = s
End Sub
' Get the window handle of a windows given a part of its caption
Private Function GetHandleFromPartialCaption(ByRef lWnd As Long, ByVal sCaption As String, Optional ChildCaption As String) As Boolean

    Dim lhWndP As Long, lhWndC As Long
    Dim sStr As String, found As Boolean, r As Long
    GetHandleFromPartialCaption = False
    lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
    Do While lhWndP <> 0
        sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
                      
        GetWindowText lhWndP, sStr, Len(sStr)
        sStr = Left$(sStr, Len(sStr) - 1)
        If InStr(1, sStr, sCaption) > 0 Then
            found = True
            lWnd = lhWndP
            Exit Do
        End If
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop
    If found And ChildCaption <> "" Then
        lhWndC = GetWindow(lhWndP, GW_CHILD)
        found = False
        Do Until lhWndC = 0
                sStr = Space$(255)
                r = GetWindowText(lhWndC, sStr, 255)
                sStr = Left$(sStr, r)
                If InStr(1, sStr, ChildCaption) <> 0 Then
                    lWnd = lhWndC
                    found = True
                    Exit Do
                End If
                lhWndC = GetWindow(lhWndC, GW_CHILD)
        Loop
    End If
    GetHandleFromPartialCaption = found
End Function



' Converts a relative path into an absolute path
Private Function AbsolutePath(ByVal path As String) As String
    If Right(path, 1) = "\" Or Right(path, 1) = "/" Then
        path = Left(path, Len(path) - 1)
    End If
    If Left(path, 1) = "." Then
        AbsolutePath = ThisWorkbook.path & Right(path, Len(path) - 1)
    Else
        AbsolutePath = path
    End If
End Function

' Tries to log an error to the error.log and displays it on the Console
Private Sub logError(errorText As String)
    Dim f As Long
    f = FreeFile()
    On Error Resume Next
    Open WorkingPath + "\error.log" For Append As f
    Print #f, errorText
    Close f
    On Error GoTo 0
    Post2Console "### RRunner: " + errorText + vbCr
End Sub


' Sends a given text to the R Console using PostMessage
Private Function Post2Console(Script As String) As Boolean
    Dim h As Long, x As String, i As Long, WindowFound As Boolean
         
 
 ' Find MDI Child first
    WindowFound = GetHandleFromPartialCaption(h, "RGui", "R Console")
    
    ' If not found, look for the console in SDI mode
    If Not WindowFound Then
        WindowFound = GetHandleFromPartialCaption(h, "R Console")
    End If
    
    If WindowFound Then
        For i = 1 To Len(Script)
            PostMessage h, WM_CHAR, Asc(Mid(Script, i)), 0
        Next i
        Post2Console = True
    Else
        Post2Console = False
        If Not mInRange Then
            MsgBox "R Console not found. Please start the R Console first."
        End If
    End If

End Function

' Displays the contents of the error.log on the R Console
Private Sub ShowErrorOnConsole()
    If FileExists(WorkingPath & "\error.log") Then
        Post2Console ("writeLines(readLines('" + Replace(WorkingPath, "\", "/") + "/error.log" + "'))") + vbCr
    End If
End Sub

' Function that accepts multiple outputs and multiple inputs as dictionaries
Public Function RunRScript(RangesToExport As Dictionary, RangesToImport As Dictionary, PicturesToImport As Dictionary, Script As Variant) As Boolean
    RunRScript = RunRScriptMain(RangesToExport, RangesToImport, PicturesToImport, Script, False, (0))
End Function


Private Function generateScriptFile(scriptContent As String, filename As String, ChartName As String)
    Dim fso As New Scripting.FileSystemObject
    Dim outputTextFile As TextStream, helper As String
    Dim wpath As String, saveChart As String
    
    If ChartName <> "" Then
        saveChart = "saveChart ('" + ChartName + "')"
    End If
    wpath = Replace(WorkingPath, "\", "/")
    helper = "library(readxl) \nlibrary(writexl)\ngetTable <- function(tableName) {read_excel('" & wpath & "/_RInput_.xlsx', sheet = tableName)} \nwriteResult <- function(tablenames, col_names = TRUE) {  write_xlsx(tablenames, path = '" & wpath & "/_ROutput_.xlsx', col_names = col_names, format_headers = FALSE)}\nsaveChart <- function(name,  pxwidth = 1024, pxheight = 768, dpi=150) {\nggsave(filename = paste('" & wpath & "/',name,'.png',sep = ''),dpi=dpi, units='in', width=pxwidth/dpi, height=pxheight/dpi)}\ndone <- function() {  file.create('" & wpath & "/done') \ncloseAllConnections()}"
    helper = Replace(helper, "\n", vbNewLine)
    On Error GoTo generateScriptFileErr
    Set outputTextFile = fso.CreateTextFile(filename, True)
    scriptContent = "this.dir <- dirname(parent.frame(2)$ofile)" + vbNewLine + _
                      "setwd (this.dir)" + vbNewLine + _
                      helper + vbNewLine + _
                      "result = data.frame()" + vbNewLine + _
                      scriptContent + vbNewLine + _
                      "writeResult(tablenames = list('result'=result))" + vbNewLine + _
                      saveChart + vbNewLine + _
                      "done()" + vbNewLine + _
                      "rm(list=ls())" + vbNewLine
                     
    outputTextFile.Write scriptContent
    outputTextFile.Close
    On Error GoTo 0
    generateScriptFile = True
    Exit Function
generateScriptFileErr:
    On Error GoTo 0
    logError (err.Description)
    generateScriptFile = False
    Exit Function
End Function


' Function that accepts multiple outputs and multiple inputs as dictionaries
Public Function RunRScriptMain(RangesToExport As Dictionary, RangesToImport As Dictionary, PicturesToImport As Dictionary, RScript As Variant, inRange As Boolean, ByRef result As Variant) As Boolean
    Dim x As String, KeyVal As String, errorFile As String, scriptContent As String, i As Long, filepath As String
    Dim rng As Range, Key As Variant, OutputKey As String, doneFile As String, Script As String, ChartName As String
        
    Application.ScreenUpdating = False
    WorkingPath = AbsolutePath(WORKING_PATH)
    mInRange = inRange
    If Not FolderExists(WorkingPath) Then
        On Error GoTo errorMkdir:
        MkDir WorkingPath
        On Error GoTo 0
    End If
    
    If Not FolderExists(RScriptsPath) Then
        On Error GoTo errorMkdir:
        MkDir RScriptsPath
        On Error GoTo 0
    End If
    
    RScriptsPath = AbsolutePath(R_SCRIPTS_PATH)
    doneFile = WorkingPath + "\done"
    ' Remove output file before running the script
    If FileExists(doneFile) Then
        Kill doneFile
    End If
    If FileExists(doneFile) Then
        Kill doneFile
    End If
    'remove all potential conflicting pictures
    If Not PicturesToImport Is Nothing Then
        For Each Key In PicturesToImport.Keys
            filepath = WorkingPath + "\" + Key + ".png"
            If FileExists(filepath) Then
                Kill filepath
            End If
        Next Key
    End If
    errorFile = WorkingPath + "\error.log"
    If FileExists(errorFile) Then
        Kill errorFile
    End If

    'Generate script if necessary
    If TypeName(RScript) = "Range" Then
           If RScript.rows.Count = 1 And RScript.Columns.Count = 1 Then
             RScript = RScript.Value
           Else
             RScript = RScript.Value
             For i = LBound(RScript) To UBound(RScript)
                scriptContent = scriptContent + RScript(i, 1) + vbNewLine
             Next i
             
           End If
    End If
    ' Try to guess if the first parameter is a script name, a reference to a script or r code
    If TypeName(RScript) <> "Variant()" Then
        If Right(Trim(UCase(RScript)), 2) <> ".R" Then
            scriptContent = RScript
        End If
    End If
    
    
    If scriptContent <> "" Then
        If Not PicturesToImport Is Nothing Then
            ChartName = PicturesToImport.Keys(0)
        End If
        RScript = "_temp_.r"
        If Not generateScriptFile(scriptContent, RScriptsPath + "\" + RScript, ChartName) Then
            RunRScriptMain = False
            GoTo end_function
        End If
    End If

    ' Generate input files for the R script
    If Not ExportFiles(RangesToExport) Then
        RunRScriptMain = False
        ShowErrorOnConsole
        GoTo end_function
    End If
    
    Script = "source('" + Replace(RScriptsPath, "\", "/") + "/" + RScript + "')" + vbCr
    If Post2Console(Script) Then
        If WaitForFile(doneFile, inRange) Then
            For Each Key In RangesToImport.Keys
                OutputKey = Key
                LoadOutput OutputKey, RangesToImport(OutputKey), result
            Next Key
            If Not PicturesToImport Is Nothing Then
                For Each Key In PicturesToImport.Keys
                    OutputKey = Key
                    LoadOutputPicture OutputKey, PicturesToImport(OutputKey)
                Next Key
            End If
            RunRScriptMain = True
        Else
            RunRScriptMain = False
            ShowErrorOnConsole
        End If
    Else
        RunRScriptMain = False
    End If
    
end_function:
    SetStatus ("")
    Application.ScreenUpdating = True
    Exit Function

errorMkdir:
    On Error GoTo 0
    logError "Error: Could not create " + WorkingPath + " or " + RScriptsPath
    RunRScriptMain = False
    Application.ScreenUpdating = True
End Function

' Check every 100 milliseconds if the expected file is available
Private Function WaitForFile(filename As String, inRange As Boolean) As Boolean
    Dim StartTickCount As Long
    Dim TickCountNow As Long
    Dim EndTickCount As Long, done As Boolean

    SetStatus "Waiting for R Script to finish..."
    StartTickCount = GetTickCount()
    EndTickCount = StartTickCount + TIMEOUT_MILLISECONDS
    done = False
    Do While Not done
        DoEvents
        Sleep (100)
        If FileExists(filename) Then
            done = True
            WaitForFile = True
        End If
        If GetTickCount() > EndTickCount Then
            done = True
            WaitForFile = False
        End If
    Loop
    SetStatus ""
End Function

' Helper function to check if a file exists
Private Function FileExists(filename As String) As Boolean
    If Dir(filename) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If

End Function


' Helper function to check if a folder exists
Private Function FolderExists(folderName As String) As Boolean
    If Dir(folderName, vbDirectory) <> "" Then
        FolderExists = True
    Else
        FolderExists = False
    End If

End Function


' Load a xlsx file into a range
Private Sub LoadOutput(TableName As String, rng As Range, ByRef result As Variant)
    Dim wb As Workbook, tblArr As Variant, destRng As Variant, firstRow As Variant, usedRng As Range
    Dim rows As Long, cols As Long
    
    SetStatus "Retrieving output..."
    Set wb = Application.Workbooks.Open(WorkingPath + "\" & R_OUT_FILE_NAME & ".xlsx")
    Set usedRng = wb.Worksheets(TableName).UsedRange
    rows = usedRng.rows.Count
    cols = usedRng.Columns.Count

    tblArr = usedRng.Value
    wb.Close

    Set destRng = rng.Resize(rows, cols)
    destRng.Value = tblArr
    SetStatus ""
End Sub



' Simplified call to run a script with a single outRange called result
Public Function RunR2Range(Script As Variant, RangeToImport As Range, ParamArray RangesToExport() As Variant) As Boolean
    Dim i As Integer
    Dim inp As New Dictionary, out As New Dictionary, outP As Dictionary
    
    For i = LBound(RangesToExport) To UBound(RangesToExport) Step 2
        Set inp(RangesToExport(i)) = RangesToExport(i + 1)
    Next i
    
    Set out("result") = RangeToImport
    RunR2Range = RunRScript(inp, out, outP, Script)
End Function

' Simplified call to run a script with a single outRange called Result
Public Function RunR2Plot(Script As Variant, RangeToExport As Range, ChartToLoad As ChartObject, PlotName As String) As Boolean
    Dim i As Integer
    Dim inp As New Dictionary, out As New Dictionary, outP As New Dictionary
    
    Set inp(PlotName) = RangeToExport
    
    
    Set outP(PlotName) = ChartToLoad
    RunR2Plot = RunRScript(inp, out, outP, Script)
End Function



'Export a range into a sheet of a given workbook
Private Function ExportTable(TableName As String, rng As Range, tempWB As Workbook) As Boolean

    Dim tblArr As Variant
    Dim cols As Long, lrow As Long
    Dim sht As Worksheet, destRng As Range
    
    ExportTable = True
    ' Find the last not empty row. This allows to send a range as columns
    ' https://www.excelcampus.com/vba/find-last-row-column-cell/
    If rng.rows.Count > 2 Then
        On Error GoTo empty_range:
        lrow = rng.Find(What:="*", _
                        After:=rng(1, 1), _
                        lookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
        On Error GoTo 0
        tblArr = rng.Resize(lrow).Value
        cols = UBound(tblArr, 2)
    Else
        lrow = rng.rows.Count
        cols = rng.Columns.Count
        tblArr = rng.Value
    End If
    
    SetStatus "Exporting table " + TableName + "..."
    Set sht = tempWB.Worksheets.Add
    sht.Name = TableName
    Set destRng = sht.Range(sht.Cells(1, 1), sht.Cells(lrow, cols))
    destRng.Value = tblArr
    Exit Function
empty_range:
    On Error GoTo 0
    logError "Empty Range selected to export."
    ExportTable = False
End Function

'Export a dictionary of ranges to the interface workbook
Private Function ExportFiles(InputRange As Dictionary) As Boolean
    Dim Key As Variant, KeyVal As String
    Dim tempWB As Workbook, filepath As String
    
    ExportFiles = True
    filepath = WorkingPath + "\" + R_IN_FILE_NAME + ".xlsx"
    SetStatus "Exporting ranges to R..."
    Set tempWB = Application.Workbooks.Add()
    For Each Key In InputRange.Keys
        KeyVal = Key
        If Not ExportTable(KeyVal, InputRange(Key), tempWB) Then
            ExportFiles = False
            tempWB.Close False
            Exit Function
        End If
    Next Key
    Application.DisplayAlerts = False
    tempWB.SaveAs filepath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    tempWB.Close False
    Application.DisplayAlerts = True
    SetStatus ""
End Function

'Insert a file picture inside a given ChartObject
Private Sub LoadOutputPicture(PictureName As String, ChartArea As ChartObject)
    Dim filepath As String, sh As Shape, pic As Shape, meas As Variant, wia As Object
    Dim orWidth As Long, orHeight As Long
    filepath = WorkingPath + "\" + PictureName + ".png"
    For Each sh In ChartArea.Chart.Shapes
        sh.Delete
    Next sh

    On Error Resume Next
    Set pic = ChartArea.Chart.Shapes.AddPicture(filepath, msoFalse, msoCTrue, 0, 0, -1, -1)
    If err.Number <> 0 Then
        logError err.Description & ": " & filepath
        Exit Sub
    End If
    On Error GoTo 0
    Set wia = CreateObject("WIA.ImageFile")
    'Load the ImageFile object with the specified File.
    wia.LoadFile filepath
    'Get the necessary properties.
    orWidth = wia.Width
    orHeight = wia.Height
    'Release the ImageFile object.
    Set wia = Nothing
    With pic
        .LockAspectRatio = msoFalse
        .Placement = xlFreeFloating
        If orWidth / orHeight < ChartArea.Width / ChartArea.Height Then
            .Height = ChartArea.Height
            .Width = .Height * orWidth / orHeight
        Else
            .Width = ChartArea.Width
            .Height = .Width * orHeight / orWidth
        End If
        .LockAspectRatio = msoTrue
    End With

End Sub
   

' This is the second of two timer routines. Because this timer routine is
' triggered by Application.OnTime it is safe, i.e., Excel will not allow the
' timer to fire unless the environment is safe (no open model dialogs or cell
' being edited).
Private Function RunRInCell2RangeUDF()
    Dim i As Integer, result As Variant, tmpApplicationCalculation As Long, Script As Variant, Cell As Range
    Dim inp As New Dictionary, out As New Dictionary, outP As Dictionary, RangesToExport As Variant
    Dim k As Variant, DestCell As Range, inputR As Range
    
    mInProcessR = True
    Application.ScreenUpdating = False
    tmpApplicationCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Do While mCalculatedCells2R.Count > 0
        ' Retrieve parameters from original UDF
        Set Cell = mCalculatedCells2R(1)(0)
        Set DestCell = mCalculatedCells2R(1)(1)
        If TypeName(mCalculatedCells2R(1)(2)) = "Range" Then
            Set inputR = mCalculatedCells2R(1)(2)
            Set Script = inputR
        Else
            Script = inputR
        End If
        
        RangesToExport = mCalculatedCells2R(1)(3)
        mCalculatedCells2R.Remove 1
        inp.RemoveAll
        For i = LBound(RangesToExport) To UBound(RangesToExport)
            If Not inp.Exists("range" & (i + 1)) Then
                If Not IsMissing(RangesToExport(i)) Then
                    If TypeName(RangesToExport(i)) = "Range" Then
                        inp.Add ("range" & (i + 1)), RangesToExport(i)
                    Else
                        logError "Incorrect destination range in RunRInCell2Range"
                        GoTo exit_function
                    End If
                Else
                    logError ("Missing parameter in call to RunRInCell2Range")
                End If
            Else
                logError "Duplicate call with " & "range" & (i + 1)
                GoTo exit_function
            End If
        Next i
        
        Set out("result") = DestCell
        logError "Current calculation: " & Cell.Address(external:=True)
        RunRScriptMain inp, out, outP, Script, True, result
    Loop

exit_function:
    Application.Calculation = tmpApplicationCalculation
    Application.ScreenUpdating = True
    DoEvents
    mInProcessR = False
End Function


' This is the second of two timer routines. Because this timer routine is
' triggered by Application.OnTime it is safe, i.e., Excel will not allow the
' timer to fire unless the environment is safe (no open model dialogs or cell
' being edited).
Private Function RunRInCell2PlotUDF()
    Dim i As Integer, result As Variant, tmpApplicationCalculation As Long, Script As Variant, Cell As Range
    Dim inp As New Dictionary, outP As New Dictionary, outR As New Dictionary, RangesToExport As Variant
    Dim k As Variant, DestChart As ChartObject, ChartName As String, inputR As Range
    
    mInProcessP = True
    Application.ScreenUpdating = False
    tmpApplicationCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Do While mCalculatedCells2P.Count > 0
        ' Retrieve parameters from original UDF
        Set Cell = mCalculatedCells2P(1)(0)
        ChartName = mCalculatedCells2P(1)(1)
        On Error Resume Next
        Set DestChart = Cell.Worksheet.ChartObjects(ChartName)
        If err <> 0 Then
            On Error GoTo 0
            logError ChartName & " does not exist"
            mCalculatedCells2P.Remove 1
            GoTo exit_function
        End If
        On Error GoTo 0
        If TypeName(mCalculatedCells2P(1)(2)) = "Range" Then
            Set inputR = mCalculatedCells2P(1)(2)
            Set Script = inputR
        Else
            Script = inputR
        End If
        RangesToExport = mCalculatedCells2P(1)(3)
        mCalculatedCells2P.Remove 1
        inp.RemoveAll
        For i = LBound(RangesToExport) To UBound(RangesToExport)
             inp.Add ("range" & (i + 1)), RangesToExport(i)
        Next i
        
        Set outP(ChartName) = DestChart
        logError "Current plot: " & Cell.Address(external:=True)
        RunRScriptMain inp, outR, outP, Script, True, result
    Loop

exit_function:
    Application.Calculation = tmpApplicationCalculation
    Application.ScreenUpdating = True
    mInProcessP = False
    SetStatus ""
End Function




' Proxy UDF function: https://stackoverflow.com/questions/8520732/i-dont-want-my-excel-add-in-to-return-an-array-instead-i-need-a-udf-to-change
' This is a UDF that returns starts a windows timer
' that starts a second Appliction.OnTime timer that performs activities not
' allowed in a UDF.
Public Function RunRInCell2Range(Script As Variant, Trigger As Variant, DestCell As Range, ParamArray RangesToExport()) As Variant

    Dim RunMacro As String, Content As Variant
    RunMacro = "OFF"
    If TypeName(Trigger) = "Range" Then
        If Trigger.Value Then RunMacro = "ON"
    Else
        If Trigger Then RunMacro = "ON"
    End If
    If RunMacro = "OFF" Or mInProcessR Then
        RunRInCell2Range = "(R Code " & RunMacro & ")"
        Exit Function
    End If
   ' Cache the caller's reference and parameters so it can be dealt with in a non-UDF routine
   Content = Array(Application.Caller, DestCell, Script, RangesToExport)
   On Error Resume Next
   mCalculatedCells2R.Add Content, Application.Caller.Address(external:=True)
   On Error GoTo 0

   ' Setting/resetting the timer should be the last action taken in the UDF
   If mWindowsTimerID <> 0 Then KillTimer 0&, mWindowsTimerID
   mWindowsTimerID = SetTimer(0&, 0&, 1, AddressOf RunRInCell2RangeProxy)
   RunRInCell2Range = "(R Code ON)"
End Function



' This is the first of two timer routines. This one is called by the Windows
' timer. Since a Windows timer cannot run code if a cell is being edited or a
' dialog is open this routine schedules a second safe timer using
' Application.OnTime which is ignored in a UDF.
Private Sub UDFCaller(funct As String)
   ' Stop the Windows timer
   On Error Resume Next
   KillTimer 0&, mWindowsTimerID
   On Error GoTo 0
   mWindowsTimerID = 0

   ' Cancel any previous OnTime timers
   If mApplicationTimerTime <> 0 Then
      On Error Resume Next
      Application.OnTime mApplicationTimerTime, funct, , False
      On Error GoTo 0
   End If

   ' Schedule timer
   mApplicationTimerTime = Now
   Application.OnTime mApplicationTimerTime, funct

End Sub

'Proxy of RunInCell2RangeUDF through UDFCaller
Private Sub RunRInCell2RangeProxy()
    UDFCaller "RunRInCell2RangeUDF"
End Sub

'Proxy of RunInCell2PlotUDF through UDFCaller
Private Sub RunRInCell2PlotProxy()
    UDFCaller "RunRInCell2PlotUDF"
End Sub
' Proxy UDF function: https://stackoverflow.com/questions/8520732/i-dont-want-my-excel-add-in-to-return-an-array-instead-i-need-a-udf-to-change
' This is a UDF that returns starts a windows timer
' that starts a second Appliction.OnTime timer that performs activities not
' allowed in a UDF.
Public Function RunRInCell2Plot(Script As Variant, Trigger As Range, ChartToLoad As String, ParamArray RangesToExport()) As Variant

    Dim RunMacro As String, Content As Variant
    RunMacro = "OFF"
    If TypeName(Trigger) = "Range" Then
        If Trigger.Value Then RunMacro = "ON"
    Else
        If Trigger Then RunMacro = True
    End If
    If RunMacro = "OFF" Or mInProcessP Then
        RunRInCell2Plot = "(R Code " & RunMacro & ")"
        Exit Function
    End If
   ' Cache the caller's reference and parameters so it can be dealt with in a non-UDF routine
   Content = Array(Application.Caller, ChartToLoad, Script, RangesToExport)
   'On Error Resume Next
   mCalculatedCells2P.Add Content, Application.Caller.Address(external:=True)
   'On Error GoTo 0

   ' Setting/resetting the timer should be the last action taken in the UDF
   If mWindowsTimerID <> 0 Then KillTimer 0&, mWindowsTimerID
   mWindowsTimerID = SetTimer(0&, 0&, 1, AddressOf RunRInCell2PlotProxy)
   RunRInCell2Plot = "(R Code ON)"
End Function
