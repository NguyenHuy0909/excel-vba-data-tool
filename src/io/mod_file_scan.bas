Attribute VB_Name = "mod_file_scan"
Option Explicit

'=========================================
' MODULE: mod_file_scan
' PURPOSE: Discover folders/files from Tool sheet inputs.
' PROJECT: GID Excel Tool
'=========================================

Public Sub FindTemplateFolders()
    Dim wsTool As Worksheet
    Dim fileSystem As Object, sourceFolder As Object, folderFile As Object
    Dim folderPath As String
    Dim outputRow As Long
    Dim firstRow As Long
    Dim idxCol As Long, nameCol As Long, pathCol As Long, dateCol As Long

    Set wsTool = GetWorksheetByConfig("TOOL_SHEET")
    folderPath = CStr(wsTool.Range(CStr(GetConfig("TOOL_FOLDER_CELL"))).Value)

    firstRow = GetConfigLong("TOOL_FIRST_ROW")
    idxCol = GetConfigLong("TOOL_EX_INDEX_COL")
    nameCol = GetConfigLong("TOOL_EX_NAME_COL")
    pathCol = GetConfigLong("TOOL_EX_PATH_COL")
    dateCol = GetConfigLong("TOOL_EX_DATE_COL")

    Application.ScreenUpdating = False
    ClearColumnFromRow wsTool, idxCol, firstRow
    ClearColumnFromRow wsTool, nameCol, firstRow
    ClearColumnFromRow wsTool, dateCol, firstRow

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set sourceFolder = fileSystem.GetFolder(folderPath)

    outputRow = firstRow
    For Each folderFile In sourceFolder.Files
        If LCase$(fileSystem.GetExtensionName(folderFile.Name)) = LCase$(CStr(GetConfig("EX_EXTENSION"))) Then
            wsTool.Cells(outputRow, idxCol).Value = outputRow - (firstRow - 1)
            wsTool.Cells(outputRow, nameCol).Value = folderFile.Name
            wsTool.Cells(outputRow, pathCol).Value = folderFile.Path
            wsTool.Cells(outputRow, dateCol).Value = Format(folderFile.DateLastModified, CStr(GetConfig("DATE_FORMAT")))
            outputRow = outputRow + 1
        End If
    Next folderFile
    Application.ScreenUpdating = True
End Sub

Public Sub FindCaseSets()
    Dim wsTool As Worksheet
    Dim fileSystem As Object, mainFolder As Object, subFolder As Object
    Dim folderPath As String, caseSetKeyword As String
    Dim firstRow As Long, selectedExRow As Long, outputRow As Long
    Dim idxCol As Long, nameCol As Long, pathCol As Long

    Set wsTool = GetWorksheetByConfig("TOOL_SHEET")
    folderPath = CStr(wsTool.Range(CStr(GetConfig("TOOL_FOLDER_CELL"))).Value)

    firstRow = GetConfigLong("TOOL_FIRST_ROW")
    idxCol = GetConfigLong("TOOL_CASE_INDEX_COL")
    nameCol = GetConfigLong("TOOL_CASE_NAME_COL")
    pathCol = GetConfigLong("TOOL_CASE_PATH_COL")

    Application.ScreenUpdating = False
    ClearColumnFromRow wsTool, idxCol, firstRow
    ClearColumnFromRow wsTool, nameCol, firstRow
    ClearColumnFromRow wsTool, pathCol, firstRow

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set mainFolder = fileSystem.GetFolder(folderPath)

    selectedExRow = firstRow + CLng(wsTool.Range(CStr(GetConfig("TOOL_SELECTED_EX_CELL"))).Value) - 1
    caseSetKeyword = Replace(CStr(wsTool.Cells(selectedExRow, GetConfigLong("TOOL_EX_NAME_COL")).Value), CStr(GetConfig("EX_EXTENSION")), vbNullString)

    outputRow = firstRow
    For Each subFolder In mainFolder.SubFolders
        If InStr(1, subFolder.Name, CStr(GetConfig("RPM_FOLDER_PATTERN")), vbTextCompare) > 0 Then
            If InStr(1, subFolder.Name, caseSetKeyword, vbTextCompare) > 0 Then
                wsTool.Cells(outputRow, idxCol).Value = outputRow - (firstRow - 1)
                wsTool.Cells(outputRow, nameCol).Value = subFolder.Name
                wsTool.Cells(outputRow, pathCol).Value = subFolder.Path & "\" & CStr(GetConfig("RESULT_FOLDER"))
                outputRow = outputRow + 1
            End If
        End If
    Next subFolder
    Application.ScreenUpdating = True
End Sub

Private Sub ClearColumnFromRow(ByVal ws As Worksheet, ByVal columnIndex As Long, ByVal firstRow As Long)
    ws.Range(ws.Cells(firstRow, columnIndex), ws.Cells(ws.Rows.Count, columnIndex)).ClearContents
End Sub
