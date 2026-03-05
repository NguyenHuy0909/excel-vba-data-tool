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
        End If
    Loop

    textStream.Close
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

    targetColumn = startColumn
    fieldWidth = GetConfigLong("DATA_FIELD_WIDTH")

    Do While Len(lineText) > 0
        wsData.Cells(rowIndex, targetColumn).Value = Left$(lineText, fieldWidth)
        lineText = Mid$(lineText, fieldWidth + 1)
        targetColumn = targetColumn + 1
    Loop
    Exit Sub

ERR_HANDLER:
    ErrorHandler "WriteFixedWidthValuesToRow"
End Sub
