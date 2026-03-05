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
