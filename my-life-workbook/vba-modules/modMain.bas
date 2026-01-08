Attribute VB_Name = "modMain"
Option Explicit

' =========================================================================
' MODULE: modMain
' PURPOSE: Main orchestration and user interface entry points
' =========================================================================

' =========================================================================
' SUB: InitializeWorkbook
' PURPOSE: Initialize the workbook with all structure and headers
' =========================================================================
Public Sub InitializeWorkbook()
    On Error GoTo ErrorHandler

    Dim response As VbMsgBoxResult

    response = MsgBox("This will initialize the [MY LIFE] workbook structure." & vbCrLf & _
                     "Any existing data will be preserved, but headers will be set." & vbCrLf & vbCrLf & _
                     "Continue?", vbQuestion + vbYesNo, "Initialize Workbook")

    If response = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Initialize each worksheet
    Call InitializeFilesPathsSheet
    Call InitializeFilesStructureSheet
    Call InitializeBanksSheet
    Call InitializeCardsSheet
    Call InitializeInvestmentsSheet
    Call InitializeOPUSSheet
    Call InitializeDebtsSheet
    Call InitializeIndexStructure ' From modIndexes
    Call InitializeCategoriesSheet
    Call InitializeDashboardSheet

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Workbook initialization complete!", vbInformation, "Initialization Complete"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error initializing workbook: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: RunFullImport
' PURPOSE: Run complete import process for all data sources
' =========================================================================
Public Sub RunFullImport()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    Dim response As VbMsgBoxResult

    response = MsgBox("This will import data from all configured sources." & vbCrLf & _
                     "Make sure all file paths are configured in FILES PATHS sheet." & vbCrLf & vbCrLf & _
                     "Continue?", vbQuestion + vbYesNo, "Full Import")

    If response = vbNo Then Exit Sub

    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Validate structure first
    If Not ValidateWorkbookStructure() Then
        MsgBox "Workbook structure is invalid. Please run InitializeWorkbook first.", _
               vbCritical, "Structure Error"
        GoTo CleanUp
    End If

    ' Import all data
    Call ImportAllBanks
    Call ImportAllCards
    Call ImportInvestments

    ' Classify transactions
    Call ClassifyAllTransactions

    ' Update indexes
    Call UpdateAllIndexes

    ' Update capital costs
    Call UpdateDebtValues
    Call UpdateOPUSValues

    ' Refresh dashboard
    Call RefreshDashboard

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "Full import completed in " & Format(Timer - startTime, "0.0") & " seconds.", _
           vbInformation, "Import Complete"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error during import: " & Err.Description, vbCritical, "Import Error"
End Sub

' =========================================================================
' SUB: RunQuickRefresh
' PURPOSE: Quick refresh of calculations and dashboard (no imports)
' =========================================================================
Public Sub RunQuickRefresh()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Update calculations
    Call CalculateCumulativeFactors
    Call UpdateDebtValues
    Call UpdateOPUSValues

    ' Refresh dashboard
    Call RefreshDashboard

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "Quick refresh complete!", vbInformation, "Refresh Complete"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error during refresh: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' Worksheet Initialization Subroutines
' =========================================================================

Private Sub InitializeFilesPathsSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(WS_FILES_PATHS)

    ws.Cells.Clear

    ' Headers
    ws.Cells(1, 1).Value = "Source"
    ws.Cells(1, 2).Value = "File Path"

    ' Default entries
    ws.Cells(2, 1).Value = "ITAU_BANK"
    ws.Cells(3, 1).Value = "NUBANK_BANK"
    ws.Cells(4, 1).Value = "C6_BANK"
    ws.Cells(5, 1).Value = "BB_BANK"
    ws.Cells(6, 1).Value = "ITAU_CARD"
    ws.Cells(7, 1).Value = "NUBANK_CARD"
    ws.Cells(8, 1).Value = "C6_CARD"
    ws.Cells(9, 1).Value = "INVESTMENTS"
    ws.Cells(10, 1).Value = "OPUS"
    ws.Cells(11, 1).Value = "DEBTS"

    ' Format
    With ws.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns("A:A").ColumnWidth = 20
    ws.Columns("B:B").ColumnWidth = 50
End Sub

Private Sub InitializeFilesStructureSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(WS_FILES_STRUCTURE)

    ws.Cells.Clear

    ' Headers
    ws.Cells(1, 1).Value = "Source Type"
    ws.Cells(1, 2).Value = "Column Name"
    ws.Cells(1, 3).Value = "Column Index"
    ws.Cells(1, 4).Value = "Data Type"
    ws.Cells(1, 5).Value = "Required"

    ' Format
    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns("A:E").AutoFit
End Sub

Private Sub InitializeBanksSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(WS_BANKS)

    If GetLastRow(ws) <= 1 Then ws.Cells.Clear

    ' Headers
    ws.Cells(1, 1).Value = "Bank"
    ws.Cells(1, 2).Value = "Date"
    ws.Cells(1, 3).Value = "Description"
    ws.Cells(1, 4).Value = "Value"
    ws.Cells(1, 5).Value = "Category"
    ws.Cells(1, 6).Value = "Subcategory"
    ws.Cells(1, 7).Value = "Import Timestamp"
    ws.Cells(1, 8).Value = "Correlation ID"
    ws.Cells(1, 9).Value = "Correlation Status"

    ' Format
    With ws.Range("A1:I1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns("A:I").AutoFit
End Sub

Private Sub InitializeCardsSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(WS_CARDS)

    If GetLastRow(ws) <= 1 Then ws.Cells.Clear

    ' Headers
    ws.Cells(1, 1).Value = "Bank"
    ws.Cells(1, 2).Value = "Card Number"
    ws.Cells(1, 3).Value = "Purchase Date"
    ws.Cells(1, 4).Value = "Category (Raw)"
    ws.Cells(1, 5).Value = "Description"
    ws.Cells(1, 6).Value = "Installment"
    ws.Cells(1, 7).Value = "Value"
    ws.Cells(1, 8).Value = "Category"
    ws.Cells(1, 9).Value = "Subcategory"
    ws.Cells(1, 10).Value = "Import Timestamp"

    ' Format
    With ws.Range("A1:J1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns("A:J").AutoFit
End Sub

Private Sub InitializeInvestmentsSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(WS_INVESTMENTS)

    If GetLastRow(ws) <= 1 Then ws.Cells.Clear

    ' Headers
    ws.Cells(1, 1).Value = "Institution"
    ws.Cells(1, 2).Value = "Date"
    ws.Cells(1, 3).Value = "Description"
    ws.Cells(1, 4).Value = "Value"
    ws.Cells(1, 5).Value = "Category"
    ws.Cells(1, 6).Value = "Subcategory"
    ws.Cells(1, 7).Value = "Correlation ID"
    ws.Cells(1, 8).Value = "Correlation Status"
    ws.Cells(1, 9).Value = "Import Timestamp"

    ' Format
    With ws.Range("A1:I1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns("A:I").AutoFit
End Sub

Private Sub InitializeOPUSSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(WS_OPUS)

    If GetLastRow(ws) <= 1 Then ws.Cells.Clear

    ' Headers
    ws.Cells(1, 1).Value = "Company"
    ws.Cells(1, 2).Value = "Investment Cost"
    ws.Cells(1, 3).Value = "Capital Cost (%)"
    ws.Cells(1, 4).Value = "Updated Cost"
    ws.Cells(1, 5).Value = "Start Date"
    ws.Cells(1, 6).Value = "Currency"
    ws.Cells(1, 7).Value = "Prior Management Value (USD)"
    ws.Cells(1, 8).Value = "Accumulated Value"

    ' Format
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns("A:H").AutoFit
End Sub

Private Sub InitializeDebtsSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(WS_DEBTS)

    If GetLastRow(ws) <= 1 Then ws.Cells.Clear

    ' Headers
    ws.Cells(1, 1).Value = "Creditor"
    ws.Cells(1, 2).Value = "Interest Rate (%)"
    ws.Cells(1, 3).Value = "Amount Paid"
    ws.Cells(1, 4).Value = "Updated Amount"
    ws.Cells(1, 5).Value = "Currency"
    ws.Cells(1, 6).Value = "Start Date"

    ' Format
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns("A:F").AutoFit
End Sub

Private Sub InitializeCategoriesSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(WS_CATEGORIES)

    If GetLastRow(ws) <= 1 Then ws.Cells.Clear

    ' Headers
    ws.Cells(1, 1).Value = "Category"
    ws.Cells(1, 2).Value = "Subcategory"
    ws.Cells(1, 3).Value = "Keywords / Mapping Rules"
    ws.Cells(1, 4).Value = "Date Added"

    ' Format
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns("A:D").AutoFit

    ' Add some example categories
    ws.Cells(2, 1).Value = "Food"
    ws.Cells(2, 2).Value = "Restaurants"
    ws.Cells(2, 3).Value = "RESTAURANT|IFOOD|RAPPI"

    ws.Cells(3, 1).Value = "Food"
    ws.Cells(3, 2).Value = "Groceries"
    ws.Cells(3, 3).Value = "SUPERMARKET|GROCERY|MERCADO"

    ws.Cells(4, 1).Value = "Transportation"
    ws.Cells(4, 2).Value = "Uber/Taxi"
    ws.Cells(4, 3).Value = "UBER|99|TAXI"

    ws.Cells(5, 1).Value = "Transportation"
    ws.Cells(5, 2).Value = "Gas"
    ws.Cells(5, 3).Value = "POSTO|GAS|COMBUSTIVEL"
End Sub

Private Sub InitializeDashboardSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(WS_DASHBOARD)

    ws.Cells.Clear

    ' Title
    ws.Cells(1, 1).Value = "[MY LIFE] - Executive Dashboard"
    ws.Range("A1:F1").Merge
    With ws.Range("A1")
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(68, 114, 196)
    End With

    ' Filters section
    ws.Cells(3, 1).Value = "Filters:"
    ws.Cells(3, 1).Font.Bold = True

    ws.Cells(4, 1).Value = "Year:"
    ws.Cells(4, 2).Value = Year(Date)

    ws.Cells(5, 1).Value = "Month:"
    ws.Cells(5, 2).Value = "All"

    ws.Cells(6, 1).Value = "Institution:"
    ws.Cells(6, 2).Value = "All"

    ws.Cells(7, 1).Value = "Currency:"
    ws.Cells(7, 2).Value = "All"

    ' KPIs section
    ws.Cells(10, 1).Value = "Executive KPIs:"
    ws.Cells(10, 1).Font.Bold = True
    ws.Cells(10, 1).Font.Size = 14

    ws.Cells(11, 1).Value = "Total Income:"
    ws.Cells(11, 1).Font.Bold = True

    ws.Cells(12, 1).Value = "Total Expenses:"
    ws.Cells(12, 1).Font.Bold = True

    ws.Cells(13, 1).Value = "Balance:"
    ws.Cells(13, 1).Font.Bold = True

    ' Consolidated tables headers
    ws.Cells(15, 1).Value = "Consolidated Data Tables"
    ws.Cells(15, 1).Font.Bold = True
    ws.Cells(15, 1).Font.Size = 14

    ws.Cells(17, 1).Value = "Note: Run 'Refresh Dashboard' to populate data"
    ws.Cells(17, 1).Font.Italic = True

    ws.Columns("A:F").AutoFit
End Sub

' =========================================================================
' FUNCTION: GetOrCreateSheet
' PURPOSE: Get worksheet or create if doesn't exist
' =========================================================================
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws
End Function
