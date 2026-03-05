Attribute VB_Name = "result_extractor"
Option Explicit

'=========================================
' MODULE: result_extractor
' PURPOSE: Extract and organize requested response results.
' PROJECT: GID Excel Tool
'=========================================

Public Function GetSelectedOutputsFromToolSheet(ByVal wsTool As Worksheet) As Variant
    On Error GoTo ERR_HANDLER
    GetSelectedOutputsFromToolSheet = GetValuesFromRange(wsTool, CStr(GetConfig("TOOL_OUTPUT_RANGE")))
    Exit Function
ERR_HANDLER:
    ErrorHandler "GetSelectedOutputsFromToolSheet"
End Function

Public Sub FilterResultColumns(ByVal wsData As Worksheet, ByVal wsTool As Worksheet)
    On Error GoTo ERR_HANDLER

    Dim selectedOutputs As Variant
    Dim keepColumns As Variant
    Dim lastColumn As Long
    Dim currentColumn As Long
    Dim headerRow As Long

    DebugLog "Start FilterResultColumns"

    headerRow = GetConfigLong("HEADER_ROW")
    selectedOutputs = GetSelectedOutputsFromToolSheet(wsTool)

    If Not HasArrayItems(selectedOutputs) Then
        MsgBox "Check output!", vbCritical
        Exit Sub
    End If

    keepColumns = FindColumnsByText(wsData, headerRow, selectedOutputs)
    If Not HasArrayItems(keepColumns) Then Exit Sub

    lastColumn = wsData.Cells(headerRow, wsData.Columns.Count).End(xlToLeft).Column

    For currentColumn = lastColumn To 1 Step -1
        If Not IsColumnInList(currentColumn, keepColumns) Then wsData.Columns(currentColumn).Delete
    Next currentColumn

    DebugLog "End FilterResultColumns"
    Exit Sub

ERR_HANDLER:
    ErrorHandler "FilterResultColumns"
End Sub

Public Sub RemoveDuplicateFirstOutputColumn(ByVal wsData As Worksheet, ByVal wsTool As Worksheet)
    On Error GoTo ERR_HANDLER

    Dim selectedOutputs As Variant
    Dim lastColumn As Long
    Dim currentColumn As Long
    Dim firstOutputName As String
    Dim headerRow As Long

    DebugLog "Start RemoveDuplicateFirstOutputColumn"

    headerRow = GetConfigLong("HEADER_ROW")
    selectedOutputs = GetSelectedOutputsFromToolSheet(wsTool)
    If Not HasArrayItems(selectedOutputs) Then Exit Sub

    firstOutputName = CStr(selectedOutputs(LBound(selectedOutputs)))
    lastColumn = wsData.Cells(headerRow, wsData.Columns.Count).End(xlToLeft).Column

    For currentColumn = lastColumn To 2 Step -1
        If StrComp(CStr(wsData.Cells(headerRow, currentColumn).Value), firstOutputName, vbTextCompare) = 0 Then
            wsData.Columns(currentColumn).Delete
        End If
    Next currentColumn

    DebugLog "End RemoveDuplicateFirstOutputColumn"
    Exit Sub

ERR_HANDLER:
    ErrorHandler "RemoveDuplicateFirstOutputColumn"
End Sub

Public Sub AddResultTitles(ByVal wsData As Worksheet, ByVal wsTool As Worksheet, ByVal firstOutputColumn As Long)
    On Error GoTo ERR_HANDLER

    Dim nodeIdItems As Variant, dofItems As Variant, caseSetItems As Variant, outputItems As Variant
    Dim folderPathParts As Variant, modeNameParts As Variant
    Dim resultFolderPath As String, modeName As String, rpmName As String
    Dim caseItem As Variant, nodeItem As Variant, dofItem As Variant
    Dim casePathRow As Long
    Dim outputCount As Long, dofCount As Long
    Dim outputColumn As Long, nodeTitleColumn As Long, modelTitleColumn As Long
    Dim caseCol As String

    DebugLog "Start AddResultTitles"

    caseSetItems = ParseInputTokens(CStr(wsTool.Range(CStr(GetConfig("TOOL_CASESET_INPUT"))).Value))
    nodeIdItems = ParseInputTokens(CStr(wsTool.Range(CStr(GetConfig("TOOL_NODE_INPUT"))).Value))
    dofItems = ParseInputTokens(CStr(wsTool.Range(CStr(GetConfig("TOOL_DOF_INPUT"))).Value))
    outputItems = GetSelectedOutputsFromToolSheet(wsTool)

    If Not HasArrayItems(caseSetItems) Or Not HasArrayItems(nodeIdItems) Or Not HasArrayItems(dofItems) Or Not HasArrayItems(outputItems) Then
        MsgBox "Please check Case Set / Node ID / DoF / Output inputs.", vbCritical
        Exit Sub
    End If

    outputCount = UBound(outputItems) - LBound(outputItems)
    dofCount = UBound(dofItems) - LBound(dofItems) + 1
    caseCol = CStr(GetConfig("TOOL_CASESET_RANGE_COL"))

    wsData.Cells(2, 1).Value = "Model"
    wsData.Cells(3, 1).Value = "RPM"
    wsData.Cells(4, 1).Value = "Node ID"
    wsData.Cells(5, 1).Value = "Dof"

    outputColumn = firstOutputColumn
    nodeTitleColumn = firstOutputColumn
    modelTitleColumn = firstOutputColumn

    For Each caseItem In caseSetItems
        casePathRow = CLng(caseItem) + (GetConfigLong("TOOL_FIRST_ROW") - 1)
        resultFolderPath = CStr(wsTool.Range(caseCol & casePathRow).Value)
        folderPathParts = Split(resultFolderPath, "\")

        If UBound(folderPathParts) >= 1 Then
            modeName = CStr(folderPathParts(UBound(folderPathParts) - 1))
        Else
            modeName = resultFolderPath
        End If

        modeNameParts = Split(modeName, ".")
        modeName = CStr(modeNameParts(LBound(modeNameParts)))
        rpmName = CStr(modeNameParts(UBound(modeNameParts)))

        wsData.Cells(2, modelTitleColumn).Value = modeName
        wsData.Cells(3, modelTitleColumn).Value = rpmName

        For Each nodeItem In nodeIdItems
            wsData.Cells(4, nodeTitleColumn).Value = CStr(nodeItem)

            For Each dofItem In dofItems
                wsData.Cells(5, outputColumn).Value = CStr(dofItem)
                outputColumn = outputColumn + outputCount
            Next dofItem

            nodeTitleColumn = nodeTitleColumn + outputCount * dofCount
        Next nodeItem

        modelTitleColumn = outputColumn
    Next caseItem

    DebugLog "End AddResultTitles"
    Exit Sub

ERR_HANDLER:
    ErrorHandler "AddResultTitles"
End Sub
