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

Private Const DEBUG_SHEET_NAME As String = "DEBUG_LOG"

' Writes trace information to Immediate Window, log file and DEBUG_LOG worksheet.
Public Sub DebugLog(ByVal msg As String)
    Dim logLine As String

    logLine = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & GetCallingProcedureName() & " | " & msg
    Debug.Print logLine
    WriteLogFile logLine
    WriteDebugSheetRow "INFO", msg
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
    WriteDebugSheetRow "ERROR", Replace(errorMessage, vbCrLf, " | ")
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

' One-click cleanup for both text log and in-workbook debug grid.
Public Sub ResetDebugLog()
    On Error GoTo EXIT_SUB

    Dim ws As Worksheet
    Dim logPath As String

    Set ws = EnsureDebugSheet()
    ws.Rows("2:" & ws.Rows.Count).ClearContents

    logPath = ThisWorkbook.Path & Application.PathSeparator & "tool_debug_log.txt"
    If Len(Dir$(logPath)) > 0 Then Kill logPath

EXIT_SUB:
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

Private Sub WriteDebugSheetRow(ByVal levelName As String, ByVal message As String)
    On Error GoTo EXIT_SUB

    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = EnsureDebugSheet()
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = levelName
    ws.Cells(nextRow, 3).Value = message
    ws.Cells(nextRow, 4).Value = SafeContextValue(CurrentFileName)
    ws.Cells(nextRow, 5).Value = SafeContextValue(CurrentRPM)
    ws.Cells(nextRow, 6).Value = SafeContextValue(CurrentNode)
    ws.Cells(nextRow, 7).Value = SafeContextValue(CurrentComponent)

EXIT_SUB:
End Sub

Private Function EnsureDebugSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DEBUG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = DEBUG_SHEET_NAME
    End If

    If Trim$(CStr(ws.Cells(1, 1).Value)) = vbNullString Then
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Level"
        ws.Cells(1, 3).Value = "Message"
        ws.Cells(1, 4).Value = "File"
        ws.Cells(1, 5).Value = "RPM"
        ws.Cells(1, 6).Value = "Node"
        ws.Cells(1, 7).Value = "Component"
    End If

    Set EnsureDebugSheet = ws
End Function

Private Function SafeContextValue(ByVal value As String) As String
    If Trim$(value) = vbNullString Then
        SafeContextValue = "N/A"
    Else
        SafeContextValue = value
    End If
End Function

Private Function GetCallingProcedureName() As String
    GetCallingProcedureName = "Trace"
End Function
