Attribute VB_Name = "ui_buttons"
Option Explicit

'=========================================
' MODULE: ui_buttons
' PURPOSE: UI macro entry points.
' PROJECT: GID Excel Tool
'=========================================

Public Sub BrowseFolder()
    Dim wsTool As Worksheet
    Dim folderPath As String

    LoadConfig
    Set wsTool = GetWorksheetByConfig("TOOL_SHEET")

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    wsTool.Range(CStr(GetConfig("TOOL_FOLDER_CELL"))).Value = folderPath
End Sub

Public Sub FindExFiles()
    LoadConfig
    FindTemplateFolders
End Sub

Public Sub FindCaseSetFolders()
    LoadConfig
    FindCaseSets
End Sub

Public Sub FindGidFiles()
    Dim wsTool As Worksheet
    Dim caseSetList As Variant, nodeIdList As Variant, dofList As Variant
    Dim caseSetItem As Variant
    Dim outputRow As Long

    LoadConfig
    Set wsTool = GetWorksheetByConfig("TOOL_SHEET")

    caseSetList = ParseInputTokens(CStr(wsTool.Range(CStr(GetConfig("TOOL_CASESET_INPUT"))).Value))
    nodeIdList = ParseInputTokens(CStr(wsTool.Range(CStr(GetConfig("TOOL_NODE_INPUT"))).Value))
    dofList = ParseInputTokens(CStr(wsTool.Range(CStr(GetConfig("TOOL_DOF_INPUT"))).Value))

    If Not HasArrayItems(caseSetList) Or Not HasArrayItems(nodeIdList) Or Not HasArrayItems(dofList) Then
        MsgBox "Please check Case Set / Node ID / DoF inputs.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    wsTool.Range(wsTool.Cells(GetConfigLong("TOOL_FIRST_ROW"), GetConfigLong("TOOL_GID_INDEX_COL")), wsTool.Cells(wsTool.Rows.Count, GetConfigLong("TOOL_GID_CLEAR_TO_COL"))).ClearContents

    outputRow = GetConfigLong("TOOL_FIRST_ROW")
    For Each caseSetItem In caseSetList
        outputRow = WriteMatchedGidFilesFromCaseSet(wsTool, CLng(caseSetItem), nodeIdList, dofList, outputRow)
    Next caseSetItem

    Application.ScreenUpdating = True
End Sub

Public Sub ImportHeaderGID()
    LoadConfig
    WriteDataToSheet
End Sub

Public Sub ConvertMillisecondsToSeconds()
    LoadConfig
    ConvertUnitsToSI
End Sub

Public Sub ClearDataSheet()
    LoadConfig
    mod_output.ClearDataSheet
End Sub

' Backward-compatible entry points
Public Sub Browser()
    BrowseFolder
End Sub

Public Sub find_ex()
    FindExFiles
End Sub

Public Sub FindCaseSet()
    FindCaseSetFolders
End Sub

Public Sub find_GID()
    FindGidFiles
End Sub

Public Sub Convert_ms()
    ConvertMillisecondsToSeconds
End Sub

Public Sub Clear_data()
    ClearDataSheet
End Sub

Private Function WriteMatchedGidFilesFromCaseSet(ByVal wsTool As Worksheet, ByVal caseSetIndexValue As Long, ByVal nodeIdList As Variant, ByVal dofList As Variant, ByVal startRow As Long) As Long
    Dim fileSystem As Object, resultFolder As Object, folderFile As Object
    Dim folderPath As String
    Dim caseSetRow As Long

    caseSetRow = GetConfigLong("TOOL_FIRST_ROW") + caseSetIndexValue - 1
    folderPath = CStr(wsTool.Cells(caseSetRow, GetConfigLong("TOOL_CASE_PATH_COL")).Value)

    If folderPath = vbNullString Then
        MsgBox "Result is not exist! Please check case set.", vbCritical
        WriteMatchedGidFilesFromCaseSet = startRow
        Exit Function
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set resultFolder = fileSystem.GetFolder(folderPath)

    WriteMatchedGidFilesFromCaseSet = startRow

    For Each folderFile In resultFolder.Files
        If InStr(1, folderFile.Name, CStr(GetConfig("GID_FILE_MARKER")), vbTextCompare) > 0 Then
            If LCase$(fileSystem.GetExtensionName(folderFile.Name)) = LCase$(CStr(GetConfig("GID_EXTENSION"))) Then
                If IsGidFileMatchingNodeAndDof(folderFile.Name, nodeIdList, dofList) Then
                    wsTool.Cells(WriteMatchedGidFilesFromCaseSet, GetConfigLong("TOOL_GID_INDEX_COL")).Value = WriteMatchedGidFilesFromCaseSet - (GetConfigLong("TOOL_FIRST_ROW") - 1)
                    wsTool.Cells(WriteMatchedGidFilesFromCaseSet, GetConfigLong("TOOL_GID_NAME_COL")).Value = folderFile.Name
                    wsTool.Cells(WriteMatchedGidFilesFromCaseSet, GetConfigLong("TOOL_GID_PATH_WRITE_COL")).Value = folderFile.Path
                    WriteMatchedGidFilesFromCaseSet = WriteMatchedGidFilesFromCaseSet + 1
                End If
            End If
        End If
    Next folderFile
End Function

Private Function IsGidFileMatchingNodeAndDof(ByVal fileName As String, ByVal nodeIdList As Variant, ByVal dofList As Variant) As Boolean
    Dim nodeItem As Variant, dofItem As Variant
    Dim nodeDofKey As String

    For Each nodeItem In nodeIdList
        For Each dofItem In dofList
            nodeDofKey = CStr(nodeItem) & "-" & CStr(dofItem)
            If InStr(1, fileName, nodeDofKey, vbTextCompare) > 0 Then
                IsGidFileMatchingNodeAndDof = True
                Exit Function
            End If
        Next dofItem
    Next nodeItem
End Function
