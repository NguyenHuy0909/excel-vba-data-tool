Attribute VB_Name = "helpers"
Option Explicit

'=========================================
' MODULE: helpers
'
' PURPOSE
' Provide reusable utility helpers.
'
' MAIN RESPONSIBILITIES
' - Parse token list from input text
' - Validate arrays
' - Read non-empty row values
' - Find columns by header keywords
'
' DEPENDENCIES
' - Excel object model
'
' PROJECT NAME
' GID Excel Tool
'=========================================

Public Function ParseInputTokens(ByVal rawValue As String) As Variant
    Dim normalizedValue As String
    Dim rawTokens() As String
    Dim resultTokens() As String
    Dim sourceIndex As Long
    Dim targetIndex As Long

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

    If IsArray(values) Then
        HasArrayItems = (UBound(values) >= LBound(values))
    End If
    Exit Function

EmptyArray:
    HasArrayItems = False
End Function

Public Function GetValuesFromRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colStart As Long, ByVal colEnd As Long) As Variant
    Dim collectedValues() As Variant
    Dim columnIndex As Long
    Dim valueCount As Long

    For columnIndex = colStart To colEnd
        If Not IsEmpty(ws.Cells(rowIndex, columnIndex).Value) Then
            valueCount = valueCount + 1
            ReDim Preserve collectedValues(1 To valueCount)
            collectedValues(valueCount) = ws.Cells(rowIndex, columnIndex).Value
        End If
    Next columnIndex

    If valueCount > 0 Then
        GetValuesFromRow = collectedValues
    Else
        GetValuesFromRow = Array()
    End If
End Function

Public Function FindColumnsByText(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal keywords As Variant) As Variant
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim keywordIndex As Long
    Dim matchedColumns() As Long
    Dim matchCount As Long

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

    If matchCount > 0 Then
        FindColumnsByText = matchedColumns
    Else
        FindColumnsByText = Array()
    End If
End Function
