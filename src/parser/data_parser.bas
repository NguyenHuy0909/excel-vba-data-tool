Attribute VB_Name = "data_parser"
Option Explicit

'=========================================
' MODULE: data_parser
' PURPOSE: Parse GID data block values.
' PROJECT: GID Excel Tool
'=========================================

Public Sub ImportGidData(ByVal filePath As String, ByVal wsData As Worksheet, ByVal startColumn As Long, ByRef rowHeader As Long)
    On Error GoTo ERR_HANDLER

    Dim textStream As Object
    Dim lineText As String
    Dim lineIndex As Long
    Dim dataStartLine As Long
    Dim importedRows As Long

    DebugLog "Start ImportGidData"
    SetCurrentFileContext filePath

    Set textStream = GetReadTextStreamFromFile(filePath)

    Do Until textStream.AtEndOfStream
        lineText = textStream.ReadLine
        lineIndex = lineIndex + 1

        If dataStartLine = 0 Then
            If InStr(1, lineText, "END", vbBinaryCompare) > 0 Then
                dataStartLine = lineIndex + 1
                DebugLog "Detected data start line: " & CStr(dataStartLine)
            End If
        ElseIf lineIndex >= dataStartLine Then
            WriteFixedWidthValuesToRow wsData, rowHeader, startColumn, lineText
            rowHeader = rowHeader + 1
            importedRows = importedRows + 1
        End If
    Loop

    textStream.Close
    DebugLog "Imported data rows: " & CStr(importedRows)
    DebugLog "End ImportGidData"
    Exit Sub

ERR_HANDLER:
    On Error Resume Next
    If Not textStream Is Nothing Then textStream.Close
    ErrorHandler "ImportGidData"
End Sub

Private Sub WriteFixedWidthValuesToRow(ByVal wsData As Worksheet, ByVal rowIndex As Long, ByVal startColumn As Long, ByVal lineText As String)
    On Error GoTo ERR_HANDLER

    Dim targetColumn As Long
    Dim fieldWidth As Long
    Dim valueCount As Long

    targetColumn = startColumn
    fieldWidth = GetConfigLong("DATA_FIELD_WIDTH")

    Do While Len(lineText) > 0
        wsData.Cells(rowIndex, targetColumn).Value = Left$(lineText, fieldWidth)
        lineText = Mid$(lineText, fieldWidth + 1)
        targetColumn = targetColumn + 1
        valueCount = valueCount + 1
    Loop

    If rowIndex = GetConfigLong("HEADER_ROW") + 2 Then
        DebugLog "First data row field count: " & CStr(valueCount)
    End If
    Exit Sub

ERR_HANDLER:
    ErrorHandler "WriteFixedWidthValuesToRow"
End Sub
