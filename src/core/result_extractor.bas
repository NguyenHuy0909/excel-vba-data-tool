Attribute VB_Name = "result_extractor"
Option Explicit

'=========================================
' MODULE: result_extractor
'
' PURPOSE
' Extract and organize requested response results.
'
' MAIN RESPONSIBILITIES
' - Resolve selected output list
' - Filter result columns
' - Remove duplicate first output column
' - Build title rows (Model/RPM/Node ID/Dof)
'
' DEPENDENCIES
' - helpers module
'
' PROJECT NAME
' GID Excel Tool
'=========================================

Private Const DATA_HEADER_ROW As Long = 6

Public Function GetSelectedOutputsFromToolSheet(ByVal wsTool As Worksheet) As Variant
    GetSelectedOutputsFromToolSheet = GetValuesFromRow(wsTool, 1, 31, 35)
End Function

Public Sub FilterResultColumns(ByVal wsData As Worksheet, ByVal wsTool As Worksheet)
    Dim selectedOutputs As Variant
    Dim keepColumns As Variant
    Dim lastColumn As Long
    Dim currentColumn As Long

    selectedOutputs = GetSelectedOutputsFromToolSheet(wsTool)

    If Not HasArrayItems(selectedOutputs) Then
        MsgBox "Check output!", vbCritical
        Exit Sub
    End If

    keepColumns = FindColumnsByText(wsData, DATA_HEADER_ROW, selectedOutputs)
    If Not HasArrayItems(keepColumns) Then Exit Sub

    lastColumn = wsData.Cells(DATA_HEADER_ROW, wsData.Columns.Count).End(xlToLeft).Column

    For currentColumn = lastColumn To 1 Step -1
        If Not IsColumnInList(currentColumn, keepColumns) Then
            wsData.Columns(currentColumn).Delete
        End If
    Next currentColumn
End Sub

Public Sub RemoveDuplicateFirstOutputColumn(ByVal wsData As Worksheet, ByVal wsTool As Worksheet)
    Dim selectedOutputs As Variant
    Dim lastColumn As Long
    Dim currentColumn As Long
    Dim firstOutputName As String

    selectedOutputs = GetSelectedOutputsFromToolSheet(wsTool)
    If Not HasArrayItems(selectedOutputs) Then Exit Sub

    firstOutputName = CStr(selectedOutputs(LBound(selectedOutputs)))
    lastColumn = wsData.Cells(DATA_HEADER_ROW, wsData.Columns.Count).End(xlToLeft).Column

    For currentColumn = lastColumn To 2 Step -1
        If StrComp(CStr(wsData.Cells(DATA_HEADER_ROW, currentColumn).Value), firstOutputName, vbTextCompare) = 0 Then
            wsData.Columns(currentColumn).Delete
        End If
    Next currentColumn
End Sub

Public Sub AddResultTitles(ByVal wsData As Worksheet, ByVal wsTool As Worksheet, ByVal firstOutputColumn As Long)
    Dim nodeIdItems As Variant
    Dim dofItems As Variant
    Dim caseSetItems As Variant
    Dim outputItems As Variant
    Dim folderPathParts As Variant
    Dim modeNameParts As Variant
    Dim resultFolderPath As String
    Dim modeName As String
    Dim rpmName As String
    Dim caseItem As Variant
    Dim nodeItem As Variant
    Dim dofItem As Variant
    Dim casePathRow As Long
    Dim outputCount As Long
    Dim dofCount As Long
    Dim outputColumn As Long
    Dim nodeTitleColumn As Long
    Dim modelTitleColumn As Long

    caseSetItems = ParseInputTokens(CStr(wsTool.Range("X1").Value))
    nodeIdItems = ParseInputTokens(CStr(wsTool.Range("X2").Value))
    dofItems = ParseInputTokens(CStr(wsTool.Range("X3").Value))
    outputItems = GetSelectedOutputsFromToolSheet(wsTool)

    If Not HasArrayItems(caseSetItems) Or Not HasArrayItems(nodeIdItems) Or Not HasArrayItems(dofItems) Or Not HasArrayItems(outputItems) Then
        MsgBox "Please check Case Set / Node ID / DoF / Output inputs.", vbCritical
        Exit Sub
    End If

    outputCount = UBound(outputItems) - LBound(outputItems)
    dofCount = UBound(dofItems) - LBound(dofItems) + 1

    wsData.Cells(2, 1).Value = "Model"
    wsData.Cells(3, 1).Value = "RPM"
    wsData.Cells(4, 1).Value = "Node ID"
    wsData.Cells(5, 1).Value = "Dof"

    outputColumn = firstOutputColumn
    nodeTitleColumn = firstOutputColumn
    modelTitleColumn = firstOutputColumn

    For Each caseItem In caseSetItems
        casePathRow = CLng(caseItem) + 4
        resultFolderPath = CStr(wsTool.Range("S" & casePathRow).Value)
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
End Sub

Private Function IsColumnInList(ByVal columnIndex As Long, ByVal keepColumns As Variant) As Boolean
    Dim keepIndex As Long

    For keepIndex = LBound(keepColumns) To UBound(keepColumns)
        If columnIndex = CLng(keepColumns(keepIndex)) Then
            IsColumnInList = True
            Exit Function
        End If
    Next keepIndex
End Function
