Attribute VB_Name = "gid_file_reader"
Option Explicit

'=========================================
' MODULE: gid_file_reader
' PURPOSE: Handle file reading operations for GID files.
' PROJECT: GID Excel Tool
'=========================================

Public Function GetFileExistsFromPath(ByVal filePath As String) As Boolean
    On Error GoTo ERR_HANDLER

    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    GetFileExistsFromPath = fileSystem.FileExists(filePath)
    Exit Function

ERR_HANDLER:
    ErrorHandler "GetFileExistsFromPath"
End Function

Public Function GetFolderExistsFromPath(ByVal folderPath As String) As Boolean
    On Error GoTo ERR_HANDLER

    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    GetFolderExistsFromPath = fileSystem.FolderExists(folderPath)
    Exit Function

ERR_HANDLER:
    ErrorHandler "GetFolderExistsFromPath"
End Function

Public Function GetReadTextStreamFromFile(ByVal filePath As String) As Object
    On Error GoTo ERR_HANDLER

    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set GetReadTextStreamFromFile = fileSystem.OpenTextFile(filePath, 1)
    Exit Function

ERR_HANDLER:
    ErrorHandler "GetReadTextStreamFromFile"
End Function

Public Function GetDataStartLineFromGidFile(ByVal filePath As String) As Long
    On Error GoTo ERR_HANDLER

    Dim textStream As Object
    Dim lineText As String
    Dim lineIndex As Long

    Set textStream = GetReadTextStreamFromFile(filePath)

    Do Until textStream.AtEndOfStream
        lineText = textStream.ReadLine
        lineIndex = lineIndex + 1

        If InStr(1, lineText, "END", vbBinaryCompare) > 0 Then
            GetDataStartLineFromGidFile = lineIndex + 1
            Exit Do
        End If
    Loop

    textStream.Close
    Exit Function

ERR_HANDLER:
    On Error Resume Next
    If Not textStream Is Nothing Then textStream.Close
    ErrorHandler "GetDataStartLineFromGidFile"
End Function
