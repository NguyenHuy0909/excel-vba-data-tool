Attribute VB_Name = "header_parser"
Option Explicit

'=========================================
' MODULE: header_parser
' PURPOSE: Parse GID header information.
' PROJECT: GID Excel Tool
'=========================================

Public Sub ReadHeaderBuffers(ByVal filePath As String, ByRef headerBuffers() As String)
    On Error GoTo ERR_HANDLER

    Dim textStream As Object
    Dim currentLine As String

    DebugLog "Start ReadHeaderBuffers"
    SetCurrentFileContext filePath

    headerBuffers(1) = vbNullString
    headerBuffers(2) = vbNullString

    Set textStream = GetReadTextStreamFromFile(filePath)
    DebugLog "Opening GID file: " & filePath

    Do Until textStream.AtEndOfStream
        currentLine = textStream.ReadLine

        If InStr(1, currentLine, "CHANNEL", vbBinaryCompare) > 0 Then
            headerBuffers(1) = GetCombinedHeaderLineFromStream(currentLine, textStream)
            DebugLog "Read CHANNEL header"
        ElseIf InStr(1, currentLine, "UNIT", vbBinaryCompare) > 0 Then
            headerBuffers(2) = GetCombinedHeaderLineFromStream(currentLine, textStream)
            DebugLog "Read UNIT header"
        End If
    Loop

    textStream.Close
    DebugLog "End ReadHeaderBuffers"
    Exit Sub

ERR_HANDLER:
    On Error Resume Next
    If Not textStream Is Nothing Then textStream.Close
    ErrorHandler "ReadHeaderBuffers"
End Sub

Public Function GetCombinedHeaderLineFromStream(ByVal firstLine As String, ByVal textStream As Object) As String
    On Error GoTo ERR_HANDLER

    Dim combinedLine As String
    Dim nextLine As String

    combinedLine = Replace(firstLine, "&", vbNullString)

    Do While InStr(1, firstLine, "&", vbBinaryCompare) > 0 And Not textStream.AtEndOfStream
        nextLine = textStream.ReadLine
        combinedLine = combinedLine & Replace(nextLine, "&", vbNullString)
        firstLine = nextLine
    Loop

    GetCombinedHeaderLineFromStream = combinedLine
    Exit Function

ERR_HANDLER:
    ErrorHandler "GetCombinedHeaderLineFromStream"
End Function

Public Sub WriteHeaderBuffersToSheet(ByVal wsData As Worksheet, ByVal startColumn As Long, ByRef rowHeader As Long, ByRef headerBuffers() As String)
    On Error GoTo ERR_HANDLER

    Dim headerIndex As Long
    Dim targetColumn As Long
    Dim tokenIndex As Long
    Dim headerTokens As Variant

    DebugLog "Start WriteHeaderBuffersToSheet"

    For headerIndex = 1 To 2
        If headerBuffers(headerIndex) <> vbNullString Then
            headerTokens = Split(headerBuffers(headerIndex), "'")
            targetColumn = startColumn
            DebugLog "Header token count: " & CStr(UBound(headerTokens) + 1)

            For tokenIndex = 2 To UBound(headerTokens) Step 2
                wsData.Cells(rowHeader, targetColumn).Value = headerTokens(tokenIndex)
                targetColumn = targetColumn + 1
            Next tokenIndex

            rowHeader = rowHeader + 1
        End If
    Next headerIndex

    DebugLog "End WriteHeaderBuffersToSheet"
    Exit Sub

ERR_HANDLER:
    ErrorHandler "WriteHeaderBuffersToSheet"
End Sub
