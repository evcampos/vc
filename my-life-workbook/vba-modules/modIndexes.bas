Attribute VB_Name = "modIndexes"
Option Explicit

' =========================================================================
' MODULE: modIndexes
' PURPOSE: Manage and update financial indexes with cumulative factors
' =========================================================================

' =========================================================================
' SUB: UpdateAllIndexes
' PURPOSE: Update all financial indexes from official sources
' NOTE: MacOS Excel has limitations with web queries
'       This implementation provides structure for manual updates
'       and automatic cumulative factor calculation
' =========================================================================
Public Sub UpdateAllIndexes()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Calculate cumulative factors for all existing index data
    Call CalculateCumulativeFactors

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "Index update completed in " & Format(Timer - startTime, "0.0") & " seconds." & vbCrLf & _
           vbCrLf & "Note: Please ensure index values are up to date.", _
           vbInformation, "Update Complete"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error updating indexes: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: CalculateCumulativeFactors
' PURPOSE: Calculate cumulative factors for all indexes
' LOGIC: Cumulative factor = Product of (1 + daily rate) from first date
' =========================================================================
Private Sub CalculateCumulativeFactors()
    On Error GoTo ErrorHandler

    Dim wsIndexes As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentIndex As String
    Dim previousIndex As String
    Dim currentDate As Date
    Dim previousDate As Date
    Dim indexValue As Double
    Dim previousFactor As Double
    Dim currentFactor As Double
    Dim daysDiff As Long
    Dim dailyRate As Double

    Set wsIndexes = ThisWorkbook.Worksheets(WS_INDEXES)
    lastRow = GetLastRow(wsIndexes)

    ' Sort data by index name and date
    Call SortIndexData

    ' Initialize
    previousIndex = ""
    previousFactor = 1#

    ' Process each row
    For i = 2 To lastRow
        ShowProgressBar i - 1, lastRow - 1, "Calculating Cumulative Factors"

        currentIndex = Trim(wsIndexes.Cells(i, 1).Value)
        currentDate = wsIndexes.Cells(i, 2).Value
        indexValue = wsIndexes.Cells(i, 3).Value

        ' Check if this is a new index
        If currentIndex <> previousIndex Then
            ' Starting a new index, reset base factor
            currentFactor = 1#
            previousIndex = currentIndex
            previousDate = currentDate
            previousFactor = 1#
        Else
            ' Calculate factor based on previous period
            daysDiff = currentDate - previousDate

            If daysDiff > 0 Then
                ' Annual rate converted to daily and compounded
                dailyRate = (indexValue / 100) / 252 ' Business days in year
                currentFactor = previousFactor * ((1 + dailyRate) ^ daysDiff)
            Else
                currentFactor = previousFactor
            End If

            previousDate = currentDate
            previousFactor = currentFactor
        End If

        ' Write cumulative factor
        wsIndexes.Cells(i, 4).Value = currentFactor

    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error calculating cumulative factors: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: SortIndexData
' PURPOSE: Sort index data by index name and date
' =========================================================================
Private Sub SortIndexData()
    On Error Resume Next

    Dim wsIndexes As Worksheet
    Dim lastRow As Long
    Dim sortRange As Range

    Set wsIndexes = ThisWorkbook.Worksheets(WS_INDEXES)
    lastRow = GetLastRow(wsIndexes)

    If lastRow <= 1 Then Exit Sub

    Set sortRange = wsIndexes.Range("A1:D" & lastRow)

    wsIndexes.Sort.SortFields.Clear

    ' Sort by index name first
    wsIndexes.Sort.SortFields.Add Key:=wsIndexes.Range("A2:A" & lastRow), _
                                 SortOn:=xlSortOnValues, _
                                 Order:=xlAscending

    ' Then by date
    wsIndexes.Sort.SortFields.Add Key:=wsIndexes.Range("B2:B" & lastRow), _
                                 SortOn:=xlSortOnValues, _
                                 Order:=xlAscending

    With wsIndexes.Sort
        .SetRange sortRange
        .Header = xlYes
        .Apply
    End With
End Sub

' =========================================================================
' SUB: AddIndexEntry
' PURPOSE: Manually add a new index entry
' PARAMETERS:
'   indexName - Name of index (CDI, SELIC, etc.)
'   entryDate - Date of index value
'   indexValue - Index value (percentage)
' =========================================================================
Public Sub AddIndexEntry(indexName As String, entryDate As Date, indexValue As Double)
    On Error GoTo ErrorHandler

    Dim wsIndexes As Worksheet
    Dim lastRow As Long

    Set wsIndexes = ThisWorkbook.Worksheets(WS_INDEXES)
    lastRow = GetLastRow(wsIndexes) + 1

    ' Add entry
    wsIndexes.Cells(lastRow, 1).Value = UCase(Trim(indexName))
    wsIndexes.Cells(lastRow, 2).Value = entryDate
    wsIndexes.Cells(lastRow, 3).Value = indexValue
    wsIndexes.Cells(lastRow, 4).Value = 0 ' Will be calculated

    ' Recalculate all cumulative factors
    Call CalculateCumulativeFactors

    MsgBox "Index entry added successfully.", vbInformation, "Entry Added"

    Exit Sub

ErrorHandler:
    MsgBox "Error adding index entry: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: ImportIndexFromCSV
' PURPOSE: Import index data from CSV file
' PARAMETERS:
'   filePath - Path to CSV file
'   indexName - Name of index
' FORMAT: CSV should have Date, Value columns
' =========================================================================
Public Sub ImportIndexFromCSV(filePath As String, indexName As String)
    On Error GoTo ErrorHandler

    Dim wsIndexes As Worksheet
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim i As Long

    If Not FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbExclamation, "File Not Found"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Open CSV
    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(1)

    Set wsIndexes = ThisWorkbook.Worksheets(WS_INDEXES)
    targetRow = GetLastRow(wsIndexes) + 1

    lastRow = GetLastRow(sourceWs)

    ' Import data
    For i = 2 To lastRow ' Skip header
        ShowProgressBar i - 1, lastRow - 1, "Importing " & indexName

        wsIndexes.Cells(targetRow, 1).Value = UCase(Trim(indexName))
        wsIndexes.Cells(targetRow, 2).Value = sourceWs.Cells(i, 1).Value ' Date
        wsIndexes.Cells(targetRow, 3).Value = CleanNumericValue(sourceWs.Cells(i, 2).Value) ' Value
        wsIndexes.Cells(targetRow, 4).Value = 0 ' Cumulative factor (to be calculated)

        targetRow = targetRow + 1
    Next i

    sourceWb.Close SaveChanges:=False

    ' Calculate cumulative factors
    Call CalculateCumulativeFactors

    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "Index import completed. " & (lastRow - 1) & " entries imported.", _
           vbInformation, "Import Complete"

    Exit Sub

ErrorHandler:
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error importing index: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: InitializeIndexStructure
' PURPOSE: Create initial index structure with headers
' =========================================================================
Public Sub InitializeIndexStructure()
    On Error Resume Next

    Dim wsIndexes As Worksheet
    Set wsIndexes = ThisWorkbook.Worksheets(WS_INDEXES)

    ' Clear existing data
    wsIndexes.Cells.Clear

    ' Write headers
    wsIndexes.Cells(1, 1).Value = "Index Name"
    wsIndexes.Cells(1, 2).Value = "Date"
    wsIndexes.Cells(1, 3).Value = "Index Value (%)"
    wsIndexes.Cells(1, 4).Value = "Cumulative Factor"

    ' Format headers
    With wsIndexes.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Set column widths
    wsIndexes.Columns("A:A").ColumnWidth = 15
    wsIndexes.Columns("B:B").ColumnWidth = 12
    wsIndexes.Columns("C:C").ColumnWidth = 15
    wsIndexes.Columns("D:D").ColumnWidth = 18

    ' Format columns
    wsIndexes.Columns("B:B").NumberFormat = "DD/MM/YYYY"
    wsIndexes.Columns("C:C").NumberFormat = "0.00"
    wsIndexes.Columns("D:D").NumberFormat = "0.000000"

    MsgBox "Index structure initialized.", vbInformation, "Initialization Complete"
End Sub

' =========================================================================
' FUNCTION: GetLatestIndexValue
' PURPOSE: Get the most recent value for a specific index
' PARAMETERS:
'   idxType - Index type
' RETURNS: Double - Latest index value
' =========================================================================
Public Function GetLatestIndexValue(idxType As IndexType) As Double
    On Error Resume Next

    Dim wsIndexes As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim indexName As String
    Dim latestDate As Date
    Dim latestValue As Double

    Set wsIndexes = ThisWorkbook.Worksheets(WS_INDEXES)
    lastRow = GetLastRow(wsIndexes)

    indexName = GetIndexName(idxType)
    latestDate = 0
    latestValue = 0

    For i = 2 To lastRow
        If Trim(wsIndexes.Cells(i, 1).Value) = indexName Then
            If wsIndexes.Cells(i, 2).Value > latestDate Then
                latestDate = wsIndexes.Cells(i, 2).Value
                latestValue = wsIndexes.Cells(i, 3).Value
            End If
        End If
    Next i

    GetLatestIndexValue = latestValue
End Function
