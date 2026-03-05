Attribute VB_Name = "mod_debug"
Option Explicit

'=========================================
' MODULE: mod_debug
' PURPOSE: Centralized debugging and error reporting utilities.
' PROJECT: GID Excel Tool
'=========================================

Public CurrentFileName As String
Public CurrentRPM As String
Public CurrentNode As String
Public CurrentComponent As String

' Writes trace information to Immediate Window and log file.
Public Sub DebugLog(ByVal msg As String)
    Dim logLine As String

    logLine = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & GetCallingProcedureName() & " | " & msg
    Debug.Print logLine
    WriteLogFile logLine
End Sub

' Centralized error reporting with current processing context.
Public Sub ErrorHandler(ByVal procName As String)
    Dim errorMessage As String

    errorMessage = "ERROR in " & procName & vbCrLf & _
                   "Error Number: " & CStr(Err.Number) & vbCrLf & _
                   "Description: " & Err.Description & vbCrLf & _
                   "File: " & SafeContextValue(CurrentFileName) & vbCrLf & _
                   "RPM: " & SafeContextValue(CurrentRPM) & vbCrLf & _
                   "Node: " & SafeContextValue(CurrentNode) & vbCrLf & _
                   "Component: " & SafeContextValue(CurrentComponent)

    Debug.Print errorMessage
    WriteLogFile Replace(errorMessage, vbCrLf, " | ")
    MsgBox errorMessage, vbCritical, "Processing Error"
End Sub

' Appends one line to tool_debug_log.txt in workbook folder.
Public Sub WriteLogFile(ByVal msg As String)
    Dim logPath As String
    Dim fileHandle As Integer

    On Error GoTo EXIT_SUB

    logPath = ThisWorkbook.Path & Application.PathSeparator & "tool_debug_log.txt"
    fileHandle = FreeFile
    Open logPath For Append As #fileHandle
    Print #fileHandle, msg
    Close #fileHandle

EXIT_SUB:
    On Error Resume Next
    If fileHandle > 0 Then Close #fileHandle
End Sub

Public Sub SetCurrentFileContext(ByVal fileName As String, Optional ByVal rpm As String = "", Optional ByVal node As String = "", Optional ByVal component As String = "")
    If fileName <> vbNullString Then CurrentFileName = fileName
    If rpm <> vbNullString Then CurrentRPM = rpm
    If node <> vbNullString Then CurrentNode = node
    If component <> vbNullString Then CurrentComponent = component
End Sub

Public Sub ClearCurrentContext()
    CurrentFileName = vbNullString
    CurrentRPM = vbNullString
    CurrentNode = vbNullString
    CurrentComponent = vbNullString
End Sub

Private Function SafeContextValue(ByVal value As String) As String
    If Trim$(value) = vbNullString Then
        SafeContextValue = "N/A"
    Else
        SafeContextValue = value
    End If
End Function

Private Function GetCallingProcedureName() As String
    ' VBA has no direct stack API; this placeholder keeps output format stable.
    GetCallingProcedureName = "Trace"
End Function
