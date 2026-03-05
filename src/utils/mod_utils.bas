Attribute VB_Name = "mod_utils"
Option Explicit

'=========================================
' MODULE: mod_utils
' PURPOSE: Shared utility helpers.
' PROJECT: GID Excel Tool
'=========================================

Public Function ParseInputTokens(ByVal rawValue As String) As Variant
    Dim normalizedValue As String
    Dim rawTokens() As String
    Dim resultTokens() As String
    Dim sourceIndex As Long, targetIndex As Long

    normalizedValue = Trim$(rawValue)
    normalizedValue = Replace(normalizedValue, ",", " ")

    Do While InStr(normalizedValue, "  ") > 0
        normalizedValue = Replace(normalizedValue, "  ", " ")
    Loop

    If Len(normalizedValue) = 0 Then
        ParseInputTokens = Array()
        Exit Function
    End If

    rawTokens = Split(normalizedValue, " ")

    For sourceIndex = LBound(rawTokens) To UBound(rawTokens)
        If Trim$(rawTokens(sourceIndex)) <> vbNullString Then
            targetIndex = targetIndex + 1
            ReDim Preserve resultTokens(1 To targetIndex)
            resultTokens(targetIndex) = Trim$(rawTokens(sourceIndex))
        End If
    Next sourceIndex

    If targetIndex = 0 Then
        ParseInputTokens = Array()
    Else
        ParseInputTokens = resultTokens
    End If
End Function

Public Function HasArrayItems(ByVal values As Variant) As Boolean
    On Error GoTo EmptyArray
    If IsArray(values) Then HasArrayItems = (UBound(values) >= LBound(values))
    Exit Function
EmptyArray:
    HasArrayItems = False
End Function

Public Function GetValuesFromRange(ByVal ws As Worksheet, ByVal rangeAddress As String) As Variant
    Dim rng As Range
    Dim c As Range
    Dim result() As Variant
    Dim count As Long

    Set rng = ws.Range(rangeAddress)
    For Each c In rng.Cells
        If Not IsEmpty(c.Value) Then
            count = count + 1
            ReDim Preserve result(1 To count)
            result(count) = c.Value
        End If
    Next c

    If count = 0 Then
        GetValuesFromRange = Array()
    Else
        GetValuesFromRange = result
    End If
End Function

Public Function FindColumnsByText(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal keywords As Variant) As Variant
    Dim lastColumn As Long, columnIndex As Long, keywordIndex As Long
    Dim matchedColumns() As Long, matchCount As Long

    lastColumn = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For columnIndex = 1 To lastColumn
        For keywordIndex = LBound(keywords) To UBound(keywords)
            If InStr(1, CStr(ws.Cells(headerRow, columnIndex).Value), CStr(keywords(keywordIndex)), vbTextCompare) > 0 Then
                matchCount = matchCount + 1
                ReDim Preserve matchedColumns(1 To matchCount)
                matchedColumns(matchCount) = columnIndex
                Exit For
            End If
        Next keywordIndex
    Next columnIndex

    If matchCount = 0 Then
        FindColumnsByText = Array()
    Else
        FindColumnsByText = matchedColumns
    End If
End Function

Public Function GetWorksheetByConfig(ByVal configKey As String) As Worksheet
    Set GetWorksheetByConfig = ThisWorkbook.Worksheets(CStr(GetConfig(configKey)))
End Function

Public Function GetConfigLong(ByVal keyName As String) As Long
    GetConfigLong = CLng(GetConfig(keyName))
End Function

Public Function IsColumnInList(ByVal columnIndex As Long, ByVal keepColumns As Variant) As Boolean
    Dim i As Long
    For i = LBound(keepColumns) To UBound(keepColumns)
        If columnIndex = CLng(keepColumns(i)) Then
            IsColumnInList = True
            Exit Function
        End If
    Next i
End Function
