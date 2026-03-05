Attribute VB_Name = "ui_buttons"
Option Explicit

'=========================================
' MODULE: ui_buttons
'
' PURPOSE
' Handle workbook UI actions and workflow orchestration.
'
' MAIN RESPONSIBILITIES
' - Button handlers
' - Read Tool sheet inputs
' - Write Data/Tool sheet outputs
' - Orchestrate IO/Parser/Core modules
'
' DEPENDENCIES
' - gid_file_reader, header_parser, data_parser, result_extractor, helpers
'
' PROJECT NAME
' GID Excel Tool
'=========================================

Private Const TOOL_FIRST_ROW As Long = 5
Private Const DATA_HEADER_ROW As Long = 6

Public Sub BrowseFolder()
    Dim folderPath As String
    Dim wsTool As Worksheet

    Set wsTool = ThisWorkbook.Sheets("Tool")

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    wsTool.Range("C1").Value = folderPath
End Sub

Public Sub FindExFiles()
    Dim fileSystem As Object
    Dim sourceFolder As Object
    Dim folderFile As Object
    Dim outputRow As Long
    Dim wsTool As Worksheet
    Dim folderPath As String

    Set wsTool = ThisWorkbook.Sheets("Tool")
    folderPath = CStr(wsTool.Range("C1").Value)

    Application.ScreenUpdating = False

    ClearExListArea wsTool

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set sourceFolder = fileSystem.GetFolder(folderPath)

    outputRow = TOOL_FIRST_ROW

    For Each folderFile In sourceFolder.Files
        If LCase$(fileSystem.GetExtensionName(folderFile.Name)) = "ex" Then
            WriteExFileRow wsTool, outputRow, folderFile
            outputRow = outputRow + 1
        End If
    Next folderFile

    Application.ScreenUpdating = True
End Sub

Public Sub FindCaseSetFolders()
    Dim fileSystem As Object
    Dim mainFolder As Object
    Dim subFolder As Object
    Dim caseSetKeyword As String
    Dim outputRow As Long
    Dim selectedExRow As Long
    Dim wsTool As Worksheet
    Dim folderPath As String

    Set wsTool = ThisWorkbook.Sheets("Tool")
    folderPath = CStr(wsTool.Range("C1").Value)

    Application.ScreenUpdating = False

    ClearCaseSetListArea wsTool

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set mainFolder = fileSystem.GetFolder(folderPath)

    selectedExRow = TOOL_FIRST_ROW + CLng(wsTool.Range("O3").Value) - 1
    caseSetKeyword = Replace(CStr(wsTool.Cells(selectedExRow, 2).Value), "ex", vbNullString)

    outputRow = TOOL_FIRST_ROW

    For Each subFolder In mainFolder.SubFolders
        If InStr(1, subFolder.Name, "rpm", vbTextCompare) > 0 Then
            If InStr(1, subFolder.Name, caseSetKeyword, vbTextCompare) > 0 Then
                WriteCaseSetRow wsTool, outputRow, subFolder
                outputRow = outputRow + 1
            End If
        End If
    Next subFolder

    Application.ScreenUpdating = True
End Sub

Public Sub FindGidFiles()
    Dim caseSetList As Variant
    Dim nodeIdList As Variant
    Dim dofList As Variant
    Dim caseSetItem As Variant
    Dim wsTool As Worksheet
    Dim outputRow As Long

    Set wsTool = ThisWorkbook.Sheets("Tool")

    caseSetList = ParseInputTokens(CStr(wsTool.Range("X1").Value))
    nodeIdList = ParseInputTokens(CStr(wsTool.Range("X2").Value))
    dofList = ParseInputTokens(CStr(wsTool.Range("X3").Value))

    If Not HasArrayItems(caseSetList) Or Not HasArrayItems(nodeIdList) Or Not HasArrayItems(dofList) Then
        MsgBox "Please check Case Set / Node ID / DoF inputs.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False

    wsTool.Range(wsTool.Cells(TOOL_FIRST_ROW, 21), wsTool.Cells(wsTool.Rows.Count, 98)).ClearContents

    outputRow = TOOL_FIRST_ROW

    For Each caseSetItem In caseSetList
        outputRow = WriteMatchedGidFilesFromCaseSet(wsTool, CLng(caseSetItem), nodeIdList, dofList, outputRow)
    Next caseSetItem

    Application.ScreenUpdating = True
End Sub

Public Sub ImportHeaderGID()
    Dim wsTool As Worksheet
    Dim wsData As Worksheet
    Dim rowIndex As Long
    Dim lastFileRow As Long
    Dim filePath As String
    Dim rowHeader As Long
    Dim headerBuffers(1 To 2) As String
    Dim firstOutputColumn As Long
    Dim startColumn As Long

    Set wsTool = ThisWorkbook.Sheets("Tool")
    Set wsData = ThisWorkbook.Sheets("Data")

    Application.ScreenUpdating = False

    firstOutputColumn = GetFirstOutputColumnFromDataSheet(wsData)
    lastFileRow = wsTool.Cells(wsTool.Rows.Count, "Z").End(xlUp).Row

    For rowIndex = TOOL_FIRST_ROW To lastFileRow
        filePath = CStr(wsTool.Range("Z" & rowIndex).Value)

        If Not GetFileExistsFromPath(filePath) Then
            MsgBox "Not found *.GID file path. Please check Load Folder Path", vbCritical
            Application.ScreenUpdating = True
            Exit Sub
        End If

        rowHeader = DATA_HEADER_ROW
        startColumn = GetNextHeaderColumnFromDataSheet(wsData, DATA_HEADER_ROW)

        ReadHeaderBuffers filePath, headerBuffers
        WriteHeaderBuffersToSheet wsData, startColumn, rowHeader, headerBuffers
        ImportGidData filePath, wsData, startColumn, rowHeader

        FilterResultColumns wsData, wsTool
        FilterResultColumns wsData, wsTool
    Next rowIndex

    RemoveDuplicateFirstOutputColumn wsData, wsTool
    AddResultTitles wsData, wsTool, firstOutputColumn

    Application.ScreenUpdating = True
    Sheets("Data").Select
End Sub

Public Sub ConvertMillisecondsToSeconds()
    Dim wsData As Worksheet
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim unitRange As Range
    Dim dataRange As Range

    Set wsData = ThisWorkbook.Sheets("Data")

    lastColumn = wsData.Cells(8, wsData.Columns.Count).End(xlToLeft).Column
    lastRow = wsData.Cells(wsData.Rows.Count, lastColumn).End(xlUp).Row

    Set dataRange = wsData.Range(wsData.Cells(8, 2), wsData.Cells(lastRow, lastColumn))
    Set unitRange = wsData.Range(wsData.Cells(7, 2), wsData.Cells(7, lastColumn))

    dataRange.Value = wsData.Evaluate(dataRange.Address & "*0.001")
    unitRange.Replace What:="mm/s^2", Replacement:="[m/s^2]", LookAt:=xlPart

    Sheets("Data").Select
End Sub

Public Sub ClearDataSheet()
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("Data").Cells.ClearContents
    Application.ScreenUpdating = True
End Sub

' Backward-compatible entry points for existing workbook button assignments.
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

Private Function GetFirstOutputColumnFromDataSheet(ByVal wsData As Worksheet) As Long
    Dim lastColumn As Long

    If Application.WorksheetFunction.CountA(wsData.Cells) = 0 Then
        GetFirstOutputColumnFromDataSheet = 2
        Exit Function
    End If

    lastColumn = wsData.Cells(DATA_HEADER_ROW, wsData.Columns.Count).End(xlToLeft).Column
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

Private Function WriteMatchedGidFilesFromCaseSet(ByVal wsTool As Worksheet, ByVal caseSetIndexValue As Long, ByVal nodeIdList As Variant, ByVal dofList As Variant, ByVal startRow As Long) As Long
    Dim fileSystem As Object
    Dim resultFolder As Object
    Dim folderFile As Object
    Dim folderPath As String
    Dim caseSetRow As Long

    caseSetRow = TOOL_FIRST_ROW + caseSetIndexValue - 1
    folderPath = CStr(wsTool.Range("S" & caseSetRow).Value)

    If folderPath = vbNullString Then
        MsgBox "Result is not exist! Please check case set.", vbCritical
        WriteMatchedGidFilesFromCaseSet = startRow
        Exit Function
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set resultFolder = fileSystem.GetFolder(folderPath)

    WriteMatchedGidFilesFromCaseSet = startRow

    For Each folderFile In resultFolder.Files
        If InStr(1, folderFile.Name, "abs_GID", vbTextCompare) > 0 Then
            If IsGidFileMatchingNodeAndDof(folderFile.Name, nodeIdList, dofList) Then
                wsTool.Cells(WriteMatchedGidFilesFromCaseSet, 21).Value = WriteMatchedGidFilesFromCaseSet - 4
                wsTool.Cells(WriteMatchedGidFilesFromCaseSet, 22).Value = folderFile.Name
                wsTool.Cells(WriteMatchedGidFilesFromCaseSet, 26).Value = folderFile.Path
                WriteMatchedGidFilesFromCaseSet = WriteMatchedGidFilesFromCaseSet + 1
            End If
        End If
    Next folderFile
End Function

Private Function IsGidFileMatchingNodeAndDof(ByVal fileName As String, ByVal nodeIdList As Variant, ByVal dofList As Variant) As Boolean
    Dim nodeItem As Variant
    Dim dofItem As Variant
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

Private Sub ClearExListArea(ByVal wsTool As Worksheet)
    wsTool.Range(wsTool.Cells(TOOL_FIRST_ROW, 1), wsTool.Cells(wsTool.Rows.Count, 1)).ClearContents
    wsTool.Range(wsTool.Cells(TOOL_FIRST_ROW, 2), wsTool.Cells(wsTool.Rows.Count, 2)).ClearContents
    wsTool.Range(wsTool.Cells(TOOL_FIRST_ROW, 9), wsTool.Cells(wsTool.Rows.Count, 9)).ClearContents
End Sub

Private Sub ClearCaseSetListArea(ByVal wsTool As Worksheet)
    wsTool.Range(wsTool.Cells(TOOL_FIRST_ROW, 11), wsTool.Cells(wsTool.Rows.Count, 11)).ClearContents
    wsTool.Range(wsTool.Cells(TOOL_FIRST_ROW, 12), wsTool.Cells(wsTool.Rows.Count, 12)).ClearContents
    wsTool.Range(wsTool.Cells(TOOL_FIRST_ROW, 19), wsTool.Cells(wsTool.Rows.Count, 19)).ClearContents
End Sub

Private Sub WriteExFileRow(ByVal wsTool As Worksheet, ByVal rowIndex As Long, ByVal folderFile As Object)
    wsTool.Cells(rowIndex, 1).Value = rowIndex - 4
    wsTool.Cells(rowIndex, 2).Value = folderFile.Name
    wsTool.Cells(rowIndex, 3).Value = folderFile.Path
    wsTool.Cells(rowIndex, 9).Value = Format(folderFile.DateLastModified, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Sub WriteCaseSetRow(ByVal wsTool As Worksheet, ByVal rowIndex As Long, ByVal subFolder As Object)
    wsTool.Cells(rowIndex, 11).Value = rowIndex - 4
    wsTool.Cells(rowIndex, 12).Value = subFolder.Name
    wsTool.Cells(rowIndex, 19).Value = subFolder.Path & "\results"
End Sub
