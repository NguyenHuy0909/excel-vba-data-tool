Attribute VB_Name = "mod_config"
Option Explicit

'=========================================
' MODULE: mod_config
' PURPOSE: Centralized configuration management via CONFIG worksheet.
' PROJECT: GID Excel Tool
'=========================================

Private configCache As Object ' Scripting.Dictionary

Public Sub LoadConfig(Optional ByVal forceReload As Boolean = False)
    On Error GoTo ERR_HANDLER

    Dim wsConfig As Worksheet
    Dim lastRow As Long, r As Long
    Dim keyName As String

    DebugLog "Start LoadConfig"

    If Not forceReload Then
        If Not configCache Is Nothing Then Exit Sub
    End If

    Set configCache = CreateObject("Scripting.Dictionary")
    configCache.CompareMode = vbTextCompare

    Set wsConfig = EnsureConfigSheet()
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        keyName = Trim$(CStr(wsConfig.Cells(r, 1).Value))
        If keyName <> vbNullString Then
            configCache(keyName) = wsConfig.Cells(r, 2).Value
        End If
    Next r

    DebugLog "End LoadConfig, keys=" & CStr(configCache.Count)
    Exit Sub

ERR_HANDLER:
    ErrorHandler "LoadConfig"
End Sub

Public Function GetConfig(ByVal keyName As String) As Variant
    On Error GoTo ERR_HANDLER

    LoadConfig

    If configCache.Exists(keyName) Then
        GetConfig = configCache(keyName)
    Else
        Err.Raise vbObjectError + 513, "GetConfig", "Missing CONFIG key: " & keyName
    End If
    Exit Function

ERR_HANDLER:
    ErrorHandler "GetConfig"
End Function

Private Function EnsureConfigSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CONFIG")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "CONFIG"
    End If

    If Trim$(CStr(ws.Cells(1, 1).Value)) = vbNullString Then
        SeedDefaultConfig ws
    End If

    Set EnsureConfigSheet = ws
End Function

Private Sub SeedDefaultConfig(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "KEY"
    ws.Cells(1, 2).Value = "VALUE"
    ws.Cells(1, 3).Value = "DESCRIPTION"

    WriteConfigRow ws, 2, "RESULT_FOLDER", "results", "Subfolder containing GID output files"
    WriteConfigRow ws, 3, "RPM_FOLDER_PATTERN", "rpm", "Keyword to identify case set folders"
    WriteConfigRow ws, 4, "DATA_SHEET", "Data", "Output worksheet name"
    WriteConfigRow ws, 5, "TOOL_SHEET", "Tool", "Control worksheet name"
    WriteConfigRow ws, 6, "HEADER_ROW", 6, "Header row index on data sheet"
    WriteConfigRow ws, 7, "DATA_START_ROW", 8, "First numeric data row"
    WriteConfigRow ws, 8, "CHANNEL_TIME", "Time", "Time channel label"
    WriteConfigRow ws, 9, "CHANNEL_ANGLE", "Angle", "Angle channel label"
    WriteConfigRow ws, 10, "ACC_CONVERT", 0.001, "Acceleration conversion factor"
    WriteConfigRow ws, 11, "VELO_CONVERT", 0.001, "Velocity conversion factor"
    WriteConfigRow ws, 12, "DISP_CONVERT", 0.001, "Displacement conversion factor"
    WriteConfigRow ws, 13, "GID_EXTENSION", "GID", "GID extension without dot"
    WriteConfigRow ws, 14, "GID_FILE_MARKER", "abs_GID", "Keyword in valid GID filename"
    WriteConfigRow ws, 15, "TOOL_CASESET_INPUT", "X1", "Case set input cell"
    WriteConfigRow ws, 16, "TOOL_NODE_INPUT", "X2", "Node ID input cell"
    WriteConfigRow ws, 17, "TOOL_DOF_INPUT", "X3", "DOF input cell"
    WriteConfigRow ws, 18, "TOOL_FOLDER_CELL", "C1", "Folder path input cell"
    WriteConfigRow ws, 19, "TOOL_CASESET_RANGE_COL", "S", "Case set result folder column"
    WriteConfigRow ws, 20, "TOOL_GID_PATH_COL", "Z", "Detected GID full path column"
    WriteConfigRow ws, 21, "TOOL_OUTPUT_RANGE", "AE1:AI1", "Selected output names"
    WriteConfigRow ws, 22, "TOOL_FIRST_ROW", 5, "First list row in Tool sheet"
    WriteConfigRow ws, 23, "TOOL_EX_INDEX_COL", 1, "EX index column"
    WriteConfigRow ws, 24, "TOOL_EX_NAME_COL", 2, "EX name column"
    WriteConfigRow ws, 25, "TOOL_EX_PATH_COL", 3, "EX path column"
    WriteConfigRow ws, 26, "TOOL_EX_DATE_COL", 9, "EX modified date column"
    WriteConfigRow ws, 27, "TOOL_CASE_INDEX_COL", 11, "Case set index column"
    WriteConfigRow ws, 28, "TOOL_CASE_NAME_COL", 12, "Case set name column"
    WriteConfigRow ws, 29, "TOOL_CASE_PATH_COL", 19, "Case set result path column"
    WriteConfigRow ws, 30, "TOOL_GID_INDEX_COL", 21, "GID index column"
    WriteConfigRow ws, 31, "TOOL_GID_NAME_COL", 22, "GID name column"
    WriteConfigRow ws, 32, "TOOL_GID_PATH_WRITE_COL", 26, "GID path output column"
    WriteConfigRow ws, 33, "TOOL_GID_CLEAR_TO_COL", 98, "Last column to clear for GID list"
    WriteConfigRow ws, 34, "DATA_FIELD_WIDTH", 16, "Fixed width of values in GID data section"

    WriteConfigRow ws, 35, "TOOL_SELECTED_EX_CELL", "O3", "Selected EX index cell"
    WriteConfigRow ws, 36, "EX_EXTENSION", "ex", "Template file extension"
    WriteConfigRow ws, 37, "DATE_FORMAT", "yyyy-mm-dd hh:nn:ss", "Display format for modified date"
End Sub

Private Sub WriteConfigRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal keyName As String, ByVal keyValue As Variant, ByVal description As String)
    ws.Cells(rowIndex, 1).Value = keyName
    ws.Cells(rowIndex, 2).Value = keyValue
    ws.Cells(rowIndex, 3).Value = description
End Sub
