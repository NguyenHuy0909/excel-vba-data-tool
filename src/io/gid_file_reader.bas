Attribute VB_Name = "gid_file_reader"
Option Explicit

'=========================================
' MODULE: gid_file_reader
'
' PURPOSE
' Handle file reading operations for GID files.
'
' MAIN RESPONSIBILITIES
' - Check file/folder existence
' - Open text stream from GID file
' - Locate END marker position for data section
'
' DEPENDENCIES
' - Scripting.FileSystemObject (late binding)
'
' PROJECT NAME
' GID Excel Tool
'=========================================

Public Function GetFileExistsFromPath(ByVal filePath As String) As Boolean
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    GetFileExistsFromPath = fileSystem.FileExists(filePath)
End Function

Public Function GetFolderExistsFromPath(ByVal folderPath As String) As Boolean
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    GetFolderExistsFromPath = fileSystem.FolderExists(folderPath)
End Function

Public Function GetReadTextStreamFromFile(ByVal filePath As String) As Object
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set GetReadTextStreamFromFile = fileSystem.OpenTextFile(filePath, 1)
End Function

Public Function GetDataStartLineFromGidFile(ByVal filePath As String) As Long
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
End Function
