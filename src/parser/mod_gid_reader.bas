Attribute VB_Name = "mod_gid_reader"
Option Explicit

'=========================================
' MODULE: mod_gid_reader
' PURPOSE: Read GID header/data sections.
' PROJECT: GID Excel Tool
'=========================================

Public Sub ReadGIDHeader(ByVal filePath As String, ByVal wsData As Worksheet, ByVal startColumn As Long, ByRef rowHeader As Long)
    On Error GoTo ERR_HANDLER

    Dim headerBuffers(1 To 2) As String

    DebugLog "Start ReadGIDHeader"
    SetCurrentFileContext filePath

    ReadHeaderBuffers filePath, headerBuffers
    WriteHeaderBuffersToSheet wsData, startColumn, rowHeader, headerBuffers

    DebugLog "End ReadGIDHeader"
    Exit Sub

ERR_HANDLER:
    ErrorHandler "ReadGIDHeader"
End Sub

Public Sub ReadGIDData(ByVal filePath As String, ByVal wsData As Worksheet, ByVal startColumn As Long, ByRef rowHeader As Long)
    On Error GoTo ERR_HANDLER

    DebugLog "Start ReadGIDData"
    SetCurrentFileContext filePath

    ImportGidData filePath, wsData, startColumn, rowHeader

    DebugLog "End ReadGIDData"
    Exit Sub

ERR_HANDLER:
    ErrorHandler "ReadGIDData"
End Sub
