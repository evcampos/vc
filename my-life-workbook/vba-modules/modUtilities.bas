Attribute VB_Name = "modUtilities"
Option Explicit

' =========================================================================
' MODULE: modUtilities
' PURPOSE: Common utility functions for [MY LIFE] workbook
' =========================================================================

' =========================================================================
' FUNCTION: ClearWorksheetData
' PURPOSE: Clear all data from a worksheet except headers
' PARAMETERS:
'   ws - Worksheet object
'   headerRow - Row number of headers (default: 1)
' =========================================================================
Public Sub ClearWorksheetData(ws As Worksheet, Optional headerRow As Long = 1)
    On Error GoTo ErrorHandler

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow > headerRow Then
        ws.Rows(headerRow + 1 & ":" & lastRow).Delete
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error clearing worksheet " & ws.Name & ": " & Err.Description, _
           vbCritical, "Clear Data Error"
End Sub

' =========================================================================
' FUNCTION: GetLastRow
' PURPOSE: Get last row with data in a worksheet
' PARAMETERS:
'   ws - Worksheet object
'   col - Column number to check (default: 1)
' RETURNS: Long - Last row number
' =========================================================================
Public Function GetLastRow(ws As Worksheet, Optional col As Long = 1) As Long
    On Error Resume Next
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If Err.Number <> 0 Then GetLastRow = 1
End Function

' =========================================================================
' FUNCTION: GetLastCol
' PURPOSE: Get last column with data in a worksheet
' PARAMETERS:
'   ws - Worksheet object
'   row - Row number to check (default: 1)
' RETURNS: Long - Last column number
' =========================================================================
Public Function GetLastCol(ws As Worksheet, Optional row As Long = 1) As Long
    On Error Resume Next
    GetLastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
    If Err.Number <> 0 Then GetLastCol = 1
End Function

' =========================================================================
' FUNCTION: FindColumnIndex
' PURPOSE: Find column index by header name
' PARAMETERS:
'   ws - Worksheet object
'   headerName - Name of header to find
'   headerRow - Row containing headers (default: 1)
' RETURNS: Long - Column index or 0 if not found
' =========================================================================
Public Function FindColumnIndex(ws As Worksheet, headerName As String, _
                               Optional headerRow As Long = 1) As Long
    On Error GoTo ErrorHandler

    Dim lastCol As Long
    Dim i As Long

    lastCol = GetLastCol(ws, headerRow)

    For i = 1 To lastCol
        If Trim(UCase(ws.Cells(headerRow, i).Value)) = Trim(UCase(headerName)) Then
            FindColumnIndex = i
            Exit Function
        End If
    Next i

    FindColumnIndex = 0
    Exit Function

ErrorHandler:
    FindColumnIndex = 0
End Function

' =========================================================================
' FUNCTION: IsValidDate
' PURPOSE: Check if value is a valid date
' PARAMETERS:
'   value - Value to check
' RETURNS: Boolean
' =========================================================================
Public Function IsValidDate(value As Variant) As Boolean
    On Error Resume Next
    IsValidDate = IsDate(value)
End Function

' =========================================================================
' FUNCTION: ParseBrazilianDate
' PURPOSE: Parse Brazilian date format (DD/MM/YYYY)
' PARAMETERS:
'   dateStr - Date string
' RETURNS: Date or Empty if invalid
' =========================================================================
Public Function ParseBrazilianDate(dateStr As String) As Variant
    On Error Resume Next

    Dim dateParts() As String
    Dim day As Integer, month As Integer, year As Integer

    dateStr = Trim(dateStr)

    ' Try direct conversion first
    If IsDate(dateStr) Then
        ParseBrazilianDate = CDate(dateStr)
        Exit Function
    End If

    ' Parse DD/MM/YYYY
    If InStr(dateStr, "/") > 0 Then
        dateParts = Split(dateStr, "/")
        If UBound(dateParts) = 2 Then
            day = CInt(dateParts(0))
            month = CInt(dateParts(1))
            year = CInt(dateParts(2))
            ParseBrazilianDate = DateSerial(year, month, day)
            Exit Function
        End If
    End If

    ParseBrazilianDate = Empty
End Function

' =========================================================================
' FUNCTION: CleanNumericValue
' PURPOSE: Clean and convert numeric strings to Double
' PARAMETERS:
'   value - Value to clean
' RETURNS: Double
' =========================================================================
Public Function CleanNumericValue(value As Variant) As Double
    On Error Resume Next

    Dim cleanValue As String

    If IsNumeric(value) Then
        CleanNumericValue = CDbl(value)
        Exit Function
    End If

    ' Remove currency symbols and thousand separators
    cleanValue = CStr(value)
    cleanValue = Replace(cleanValue, "R$", "")
    cleanValue = Replace(cleanValue, "$", "")
    cleanValue = Replace(cleanValue, "€", "")
    cleanValue = Replace(cleanValue, ".", "") ' Brazilian thousand separator
    cleanValue = Replace(cleanValue, ",", ".") ' Brazilian decimal separator
    cleanValue = Trim(cleanValue)

    If IsNumeric(cleanValue) Then
        CleanNumericValue = CDbl(cleanValue)
    Else
        CleanNumericValue = 0
    End If
End Function

' =========================================================================
' FUNCTION: ShowProgressBar
' PURPOSE: Display progress in status bar
' PARAMETERS:
'   current - Current item
'   total - Total items
'   operation - Operation description
' =========================================================================
Public Sub ShowProgressBar(current As Long, total As Long, operation As String)
    Dim percentage As Double

    If total > 0 Then
        percentage = (current / total) * 100
        Application.StatusBar = operation & ": " & Format(percentage, "0.0") & _
                               "% (" & current & " of " & total & ")"
    End If
End Sub

' =========================================================================
' FUNCTION: ResetStatusBar
' PURPOSE: Clear status bar
' =========================================================================
Public Sub ResetStatusBar()
    Application.StatusBar = False
End Sub

' =========================================================================
' FUNCTION: FileExists
' PURPOSE: Check if file exists
' PARAMETERS:
'   filePath - Full path to file
' RETURNS: Boolean
' =========================================================================
Public Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
End Function

' =========================================================================
' FUNCTION: CreateNamedRange
' PURPOSE: Create or update a named range
' PARAMETERS:
'   rangeName - Name for the range
'   rangeAddress - Address of the range
'   ws - Worksheet containing the range
' =========================================================================
Public Sub CreateNamedRange(rangeName As String, rangeAddress As String, ws As Worksheet)
    On Error Resume Next

    ' Delete existing name if it exists
    ThisWorkbook.Names(rangeName).Delete

    ' Create new name
    ThisWorkbook.Names.Add Name:=rangeName, _
                          RefersTo:="='" & ws.Name & "'!" & rangeAddress
End Sub

' =========================================================================
' FUNCTION: SafeDivide
' PURPOSE: Divide with zero-check
' PARAMETERS:
'   numerator - Numerator
'   denominator - Denominator
' RETURNS: Double (0 if division by zero)
' =========================================================================
Public Function SafeDivide(numerator As Double, denominator As Double) As Double
    If denominator = 0 Then
        SafeDivide = 0
    Else
        SafeDivide = numerator / denominator
    End If
End Function

' =========================================================================
' FUNCTION: StandardizeText
' PURPOSE: Standardize text for comparison (uppercase, trim, remove extra spaces)
' PARAMETERS:
'   text - Text to standardize
' RETURNS: String
' =========================================================================
Public Function StandardizeText(text As String) As String
    Dim result As String
    result = Trim(UCase(text))
    ' Remove multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    StandardizeText = result
End Function
