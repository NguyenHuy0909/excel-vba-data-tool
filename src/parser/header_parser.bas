Attribute VB_Name = "header_parser"
Option Explicit

'=========================================
' MODULE: header_parser
'
' PURPOSE
' Parse GID header information.
'
' MAIN RESPONSIBILITIES
' - Parse CHANNEL line
' - Parse UNIT line
' - Merge continuation lines (&)
' - Write parsed headers to Data worksheet
'
' DEPENDENCIES
' - gid_file_reader.GetReadTextStreamFromFile
'
' PROJECT NAME
' GID Excel Tool
'=========================================

Public Sub ReadHeaderBuffers(ByVal filePath As String, ByRef headerBuffers() As String)
    Dim textStream As Object
    Dim currentLine As String

    headerBuffers(1) = vbNullString
    headerBuffers(2) = vbNullString

    Set textStream = GetReadTextStreamFromFile(filePath)

    Do Until textStream.AtEndOfStream
        currentLine = textStream.ReadLine

        If InStr(1, currentLine, "CHANNEL", vbBinaryCompare) > 0 Then
            headerBuffers(1) = GetCombinedHeaderLineFromStream(currentLine, textStream)
        ElseIf InStr(1, currentLine, "UNIT", vbBinaryCompare) > 0 Then
            headerBuffers(2) = GetCombinedHeaderLineFromStream(currentLine, textStream)
        End If
    Loop

    textStream.Close
End Sub

Public Function GetCombinedHeaderLineFromStream(ByVal firstLine As String, ByVal textStream As Object) As String
    Dim combinedLine As String
    Dim nextLine As String

    combinedLine = Replace(firstLine, "&", vbNullString)

    Do While InStr(1, firstLine, "&", vbBinaryCompare) > 0 And Not textStream.AtEndOfStream
        nextLine = textStream.ReadLine
        combinedLine = combinedLine & Replace(nextLine, "&", vbNullString)
        firstLine = nextLine
    Loop

    GetCombinedHeaderLineFromStream = combinedLine
End Function

Public Sub WriteHeaderBuffersToSheet(ByVal wsData As Worksheet, ByVal startColumn As Long, ByRef rowHeader As Long, ByRef headerBuffers() As String)
    Dim headerIndex As Long
    Dim targetColumn As Long
    Dim tokenIndex As Long
    Dim headerTokens As Variant

    For headerIndex = 1 To 2
        If headerBuffers(headerIndex) <> vbNullString Then
            headerTokens = Split(headerBuffers(headerIndex), "'")
            targetColumn = startColumn

            For tokenIndex = 2 To UBound(headerTokens) Step 2
                wsData.Cells(rowHeader, targetColumn).Value = headerTokens(tokenIndex)
                targetColumn = targetColumn + 1
            Next tokenIndex

            rowHeader = rowHeader + 1
        End If
    Next headerIndex
End Sub
