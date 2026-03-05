Attribute VB_Name = "data_parser"
Option Explicit

'=========================================
' MODULE: data_parser
'
' PURPOSE
' Parse GID data block values.
'
' MAIN RESPONSIBILITIES
' - Detect data start after END
' - Parse fixed-width numeric fields
' - Write parsed values to Data worksheet
'
' DEPENDENCIES
' - gid_file_reader.GetReadTextStreamFromFile
'
' PROJECT NAME
' GID Excel Tool
'=========================================

Private Const DATA_FIELD_WIDTH As Long = 16

Public Sub ImportGidData(ByVal filePath As String, ByVal wsData As Worksheet, ByVal startColumn As Long, ByRef rowHeader As Long)
    Dim textStream As Object
    Dim lineText As String
    Dim lineIndex As Long
    Dim dataStartLine As Long

    Set textStream = GetReadTextStreamFromFile(filePath)

    Do Until textStream.AtEndOfStream
        lineText = textStream.ReadLine
        lineIndex = lineIndex + 1

        If dataStartLine = 0 Then
            If InStr(1, lineText, "END", vbBinaryCompare) > 0 Then
                dataStartLine = lineIndex + 1
            End If
        ElseIf lineIndex >= dataStartLine Then
            WriteFixedWidthValuesToRow wsData, rowHeader, startColumn, lineText
            rowHeader = rowHeader + 1
        End If
    Loop

    textStream.Close
End Sub

Private Sub WriteFixedWidthValuesToRow(ByVal wsData As Worksheet, ByVal rowIndex As Long, ByVal startColumn As Long, ByVal lineText As String)
    Dim targetColumn As Long

    targetColumn = startColumn

    Do While Len(lineText) > 0
        wsData.Cells(rowIndex, targetColumn).Value = Left$(lineText, DATA_FIELD_WIDTH)
        lineText = Mid$(lineText, DATA_FIELD_WIDTH + 1)
        targetColumn = targetColumn + 1
    Loop
End Sub
