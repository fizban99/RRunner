Attribute VB_Name = "RRunner"
Option Explicit
' ###################################################################
' Configuration Parameters
' ###################################################################
' Path to the R Scripts and where the temporary files will be created
Private Const WORKING_PATH = "."
' Time to wait for the R Script answer in milliseconds
Private Const TimeOutMilliseconds = 10000
Private Const INTERFACE_IN_FILE_NAME = "_Input_"
Private Const INTERFACE_OUT_FILE_NAME = "_Output_"
' ###################################################################

Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    
Private Const WM_CHAR As Long = &H102
Private Const GW_HWNDNEXT = 2

Dim WorkingPath As String


Private Sub SetStatus(s As String)
    'Application.ScreenUpdating = True
    Application.StatusBar = s
    'Application.ScreenUpdating = False
End Sub
' Get the window handle of a windows given a part of its caption
Private Function GetHandleFromPartialCaption(ByRef lWnd As Long, ByVal sCaption As String) As Boolean

    Dim lhWndP As Long
    Dim sStr As String
    GetHandleFromPartialCaption = False
    lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
    Do While lhWndP <> 0
        sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
                      
        GetWindowText lhWndP, sStr, Len(sStr)
        sStr = Left$(sStr, Len(sStr) - 1)
        If InStr(1, sStr, sCaption) > 0 Then
            GetHandleFromPartialCaption = True
            lWnd = lhWndP
            Exit Do
        End If
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop

End Function

' Function that accepts multiple outputs and multiple inputs as dictionaries
Public Function RunRScript(InputRange As Dictionary, OutputRange As Dictionary, OutputPictures As Dictionary, script As String) As Boolean
    Dim h As Long, x As String, i As Long, KeyVal As String, WindowFound As Boolean
    Dim rng As Range, Key As Variant, OutputKey As String, doneFile As String
        
    Application.ScreenUpdating = False
    If WORKING_PATH = "." Then
        WorkingPath = ThisWorkbook.Path
    Else
        WorkingPath = WORKING_PATH
    End If
    doneFile = WorkingPath + "\done"
    ' Remove output file before running the script
    If FileExists(doneFile) Then
        Kill doneFile
    End If
    
    ' Generate input files for the R script
    ExportFiles InputRange
    
    
    script = "source('" + Replace(WorkingPath, "\", "/") + "/" + script + "')" + vbNewLine
    WindowFound = GetHandleFromPartialCaption(h, "R Console")
    If WindowFound Then
        For i = 1 To Len(script)
            PostMessage h, WM_CHAR, Asc(Mid(script, i)), 0
        Next i
        If WaitForFile(doneFile) Then
            For Each Key In OutputRange.Keys
                OutputKey = Key
                LoadOutput OutputKey, OutputRange(OutputKey)
            Next Key
            If Not OutputPictures Is Nothing Then
                For Each Key In OutputPictures.Keys
                    OutputKey = Key
                    LoadOutputPicture OutputKey, OutputPictures(OutputKey)
                Next Key
            End If
            RunRScript = True
        Else
            RunRScript = False
        End If
    Else
        RunRScript = False
        MsgBox "R Console not found. Please start the R Console first."
    End If
    Application.ScreenUpdating = True
End Function

' Check every 100 milliseconds if the expected file is available
Private Function WaitForFile(fileName As String) As Boolean
    Dim StartTickCount As Long
    Dim TickCountNow As Long
    Dim EndTickCount As Long, done As Boolean

    SetStatus "Waiting for R Script to finish..."
    StartTickCount = GetTickCount()
    EndTickCount = StartTickCount + TimeOutMilliseconds
    done = False
    Do While Not done
        DoEvents
        Sleep (100)
        If FileExists(fileName) Then
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
Private Function FileExists(fileName As String) As Boolean
    If Dir(fileName) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If

End Function





' Load a xlsx file into a range
Private Sub LoadOutput(TableName As String, rng As Range)
    Dim wb As Workbook, tblArr As Variant, destRng As Variant
    Dim rows As Long, cols As Long
    
    SetStatus "Retrieving output..."
    Set wb = Application.Workbooks.Open(WorkingPath + "\" & INTERFACE_OUT_FILE_NAME & ".xlsx")
    'wb.Windows(1).Visible = False
    tblArr = wb.Worksheets(TableName).UsedRange.Value
    wb.Close
    rows = UBound(tblArr, 1)
    cols = UBound(tblArr, 2)
    Set destRng = rng.Resize(rows, cols)
    destRng.Value = tblArr
    SetStatus ""
End Sub



' Simplified call to run a script with a single outRange called Resultado
Public Function RunR2Range(script As String, outRange As Range, ParamArray Ranges() As Variant) As Boolean
    Dim i As Integer
    Dim inp As New Dictionary, out As New Dictionary, outP As Dictionary
    
    For i = LBound(Ranges) To UBound(Ranges) Step 2
        Set inp(Ranges(i)) = Ranges(i + 1)
    Next i
    
    Set out("Resultado") = outRange
    RunR2Range = RunRScript(inp, out, outP, script)
End Function

' Simplified call to run a script with a single outRange called Resultado
Public Function RunR2Plot(script As String, inpRange As Range, outChart As ChartObject, PlotName As String) As Boolean
    Dim i As Integer
    Dim inp As New Dictionary, out As New Dictionary, outP As New Dictionary
    
    Set inp(PlotName) = inpRange
    
    
    Set outP(PlotName) = outChart
    RunR2Plot = RunRScript(inp, out, outP, script)
End Function



'Export a range into a sheet of a given workbook
Private Sub ExportTable(TableName As String, rng As Range, tempWB As Workbook)

    Dim tblArr As Variant
    Dim cols As Long, lrow As Long
    Dim sht As Worksheet, destRng As Range
    
    ' Find the last not empty row. This allows to send a range as columns
    ' https://www.excelcampus.com/vba/find-last-row-column-cell/
    lrow = rng.Find(What:="*", _
                    After:=rng(1, 1), _
                    lookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).row
    tblArr = rng.Resize(lrow).Value
    cols = UBound(tblArr, 2)
    SetStatus "Exporting table " + TableName + "..."
    Set sht = tempWB.Worksheets.Add
    sht.Name = TableName
    Set destRng = sht.Range(sht.Cells(1, 1), sht.Cells(lrow, cols))
    destRng.Value = tblArr
    
End Sub

'Export a dictionary of ranges to the interface workbook
Private Sub ExportFiles(InputRange As Dictionary)
    Dim Key As Variant, KeyVal As String
    Dim tempWB As Workbook, filePath As String
    
    filePath = WorkingPath + "\" + INTERFACE_IN_FILE_NAME + ".xlsx"

    Set tempWB = Application.Workbooks.Add()
    tempWB.Windows(1).Visible = False
    For Each Key In InputRange.Keys
        KeyVal = Key
        ExportTable KeyVal, InputRange(Key), tempWB
    Next Key
    Application.DisplayAlerts = False
    tempWB.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    tempWB.Close False
    Application.DisplayAlerts = True
    SetStatus ""
End Sub

'Insert a file picture inside a given ChartObject
Private Sub LoadOutputPicture(PictureName As String, ChartArea As ChartObject)
    Dim filePath As String, sh As Shape
    filePath = WorkingPath + "\" + PictureName + ".png"
    For Each sh In ChartArea.Chart.Shapes
        sh.Delete
    Next sh
    ChartArea.Chart.Shapes.AddPicture filePath, msoFalse, msoCTrue, 0, 0, -1, -1
End Sub
    

