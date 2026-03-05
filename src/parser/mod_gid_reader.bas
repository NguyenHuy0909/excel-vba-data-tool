Attribute VB_Name = "mod_gid_reader"
Option Explicit

'=========================================
' MODULE: mod_gid_reader
' PURPOSE: Read GID header/data sections.
' PROJECT: GID Excel Tool
'=========================================

Public Sub ReadGIDHeader(ByVal filePath As String, ByVal wsData As Worksheet, ByVal startColumn As Long, ByRef rowHeader As Long)
    Dim headerBuffers(1 To 2) As String
    ReadHeaderBuffers filePath, headerBuffers
    WriteHeaderBuffersToSheet wsData, startColumn, rowHeader, headerBuffers
End Sub

Public Sub ReadGIDData(ByVal filePath As String, ByVal wsData As Worksheet, ByVal startColumn As Long, ByRef rowHeader As Long)
    ImportGidData filePath, wsData, startColumn, rowHeader
End Sub
