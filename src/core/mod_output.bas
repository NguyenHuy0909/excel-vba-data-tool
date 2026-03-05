Attribute VB_Name = "mod_output"
Option Explicit

'=========================================
' MODULE: mod_output
' PURPOSE: Output orchestration to Data sheet.
' PROJECT: GID Excel Tool
'=========================================

Public Sub WriteDataToSheet()
    On Error GoTo ERR_HANDLER

    Dim wsTool As Worksheet, wsData As Worksheet
    Dim rowIndex As Long, lastFileRow As Long
    Dim filePath As String
    Dim rowHeader As Long, startColumn As Long
    Dim firstOutputColumn As Long
    Dim headerRow As Long, firstRow As Long

    DebugLog "Start WriteDataToSheet"

    Set wsTool = GetWorksheetByConfig("TOOL_SHEET")
    Set wsData = GetWorksheetByConfig("DATA_SHEET")

    headerRow = GetConfigLong("HEADER_ROW")
    firstRow = GetConfigLong("TOOL_FIRST_ROW")

    Application.ScreenUpdating = False

    firstOutputColumn = GetFirstOutputColumnFromDataSheet(wsData, headerRow)
    lastFileRow = wsTool.Cells(wsTool.Rows.Count, CStr(GetConfig("TOOL_GID_PATH_COL"))).End(xlUp).Row

    For rowIndex = firstRow To lastFileRow
        filePath = CStr(wsTool.Range(CStr(GetConfig("TOOL_GID_PATH_COL")) & rowIndex).Value)
        SetCurrentFileContext filePath, ExtractRpmFromPath(filePath), CurrentNode, ExtractComponentFromPath(filePath)
        DebugLog "Processing GID file: " & filePath

        If Not GetFileExistsFromPath(filePath) Then
            MsgBox "Not found *.GID file path. Please check Load Folder Path", vbCritical
            Application.ScreenUpdating = True
            Exit Sub
        End If

        rowHeader = headerRow
        startColumn = GetNextHeaderColumnFromDataSheet(wsData, headerRow)

        ReadGIDHeader filePath, wsData, startColumn, rowHeader
        ReadGIDData filePath, wsData, startColumn, rowHeader

        DebugLog "Applying result filters"
        FilterResultColumns wsData, wsTool
        FilterResultColumns wsData, wsTool
    Next rowIndex

    RemoveDuplicateFirstOutputColumn wsData, wsTool
    AddResultTitles wsData, wsTool, firstOutputColumn

    Application.ScreenUpdating = True
    wsData.Select
    DebugLog "End WriteDataToSheet"
    Exit Sub

ERR_HANDLER:
    Application.ScreenUpdating = True
    ErrorHandler "WriteDataToSheet"
End Sub

Public Sub ConvertUnitsToSI()
    On Error GoTo ERR_HANDLER

    Dim wsData As Worksheet
    Dim lastColumn As Long, lastRow As Long
    Dim dataRange As Range, unitRange As Range
    Dim factor As Double
    Dim startRow As Long

    DebugLog "Start ConvertUnitsToSI"

    Set wsData = GetWorksheetByConfig("DATA_SHEET")
    startRow = GetConfigLong("DATA_START_ROW")

    lastColumn = wsData.Cells(startRow, wsData.Columns.Count).End(xlToLeft).Column
    lastRow = wsData.Cells(wsData.Rows.Count, lastColumn).End(xlUp).Row

    Set dataRange = wsData.Range(wsData.Cells(startRow, 2), wsData.Cells(lastRow, lastColumn))
    Set unitRange = wsData.Range(wsData.Cells(startRow - 1, 2), wsData.Cells(startRow - 1, lastColumn))

    factor = CDbl(GetConfig("ACC_CONVERT"))
    dataRange.Value = wsData.Evaluate(dataRange.Address & "*" & CStr(factor))
    unitRange.Replace What:="mm/s^2", Replacement:="[m/s^2]", LookAt:=xlPart

    wsData.Select
    DebugLog "End ConvertUnitsToSI"
    Exit Sub

ERR_HANDLER:
    ErrorHandler "ConvertUnitsToSI"
End Sub

Public Sub ClearDataSheet()
    On Error GoTo ERR_HANDLER

    Dim wsData As Worksheet
    DebugLog "Start ClearDataSheet"

    Set wsData = GetWorksheetByConfig("DATA_SHEET")

    Application.ScreenUpdating = False
    wsData.Cells.ClearContents
    Application.ScreenUpdating = True

    DebugLog "End ClearDataSheet"
    Exit Sub

ERR_HANDLER:
    Application.ScreenUpdating = True
    ErrorHandler "ClearDataSheet"
End Sub

Private Function GetFirstOutputColumnFromDataSheet(ByVal wsData As Worksheet, ByVal headerRow As Long) As Long
    Dim lastColumn As Long

    If Application.WorksheetFunction.CountA(wsData.Cells) = 0 Then
        GetFirstOutputColumnFromDataSheet = 2
        Exit Function
    End If

    lastColumn = wsData.Cells(headerRow, wsData.Columns.Count).End(xlToLeft).Column
    GetFirstOutputColumnFromDataSheet = lastColumn + 1
End Function

Private Function GetNextHeaderColumnFromDataSheet(ByVal wsData As Worksheet, ByVal headerRow As Long) As Long
    Dim lastColumn As Long

    lastColumn = wsData.Cells(headerRow, wsData.Columns.Count).End(xlToLeft).Column
    If lastColumn = 1 Then
        GetNextHeaderColumnFromDataSheet = 1
    Else
        GetNextHeaderColumnFromDataSheet = lastColumn + 1
    End If
End Function

Private Function ExtractRpmFromPath(ByVal filePath As String) As String
    Dim parts As Variant
    Dim i As Long
    parts = Split(filePath, "\")
    For i = LBound(parts) To UBound(parts)
        If InStr(1, CStr(parts(i)), CStr(GetConfig("RPM_FOLDER_PATTERN")), vbTextCompare) > 0 Then
            ExtractRpmFromPath = CStr(parts(i))
            Exit Function
        End If
    Next i
End Function

Private Function ExtractComponentFromPath(ByVal filePath As String) As String
    Dim fileName As String
    Dim parts As Variant

    fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
    parts = Split(fileName, "-")
    If UBound(parts) >= LBound(parts) Then
        ExtractComponentFromPath = CStr(parts(LBound(parts)))
    End If
End Function
