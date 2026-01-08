Attribute VB_Name = "modDashboard"
Option Explicit

' =========================================================================
' MODULE: modDashboard
' PURPOSE: Dashboard data aggregation and KPI calculations
' =========================================================================

' =========================================================================
' SUB: RefreshDashboard
' PURPOSE: Refresh all dashboard data and KPIs
' =========================================================================
Public Sub RefreshDashboard()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Clear existing dashboard data (except filters)
    Call ClearDashboardData

    ' Update consolidated tables
    Call UpdateConsolidatedCash
    Call UpdateConsolidatedCards
    Call UpdateConsolidatedTransactions
    Call UpdateConsolidatedNetDebts

    ' Update KPIs
    Call UpdateKPIs

    ' Update named ranges
    Call UpdateNamedRanges

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "Dashboard refreshed in " & Format(Timer - startTime, "0.0") & " seconds.", _
           vbInformation, "Refresh Complete"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error refreshing dashboard: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: ClearDashboardData
' PURPOSE: Clear dashboard data tables (preserve filters and structure)
' =========================================================================
Private Sub ClearDashboardData()
    On Error Resume Next

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets(WS_DASHBOARD)

    ' Clear consolidated tables - adjust ranges as needed
    ' This is a placeholder - actual ranges will depend on dashboard layout
    ' For now, we'll clear specific named ranges

    wsDash.Range("A20:G1000").ClearContents ' Consolidated Cash area
    wsDash.Range("I20:O1000").ClearContents ' Consolidated Cards area
    wsDash.Range("A1020:G2000").ClearContents ' Consolidated Transactions area
    wsDash.Range("I1020:O2000").ClearContents ' Consolidated Net Debts area

End Sub

' =========================================================================
' SUB: UpdateConsolidatedCash
' PURPOSE: Populate Consolidate Cash table
' =========================================================================
Private Sub UpdateConsolidatedCash()
    On Error GoTo ErrorHandler

    Dim wsBanks As Worksheet
    Dim wsInvestments As Worksheet
    Dim wsDash As Worksheet
    Dim lastRowBanks As Long
    Dim lastRowInv As Long
    Dim i As Long
    Dim dashRow As Long

    Set wsBanks = ThisWorkbook.Worksheets(WS_BANKS)
    Set wsInvestments = ThisWorkbook.Worksheets(WS_INVESTMENTS)
    Set wsDash = ThisWorkbook.Worksheets(WS_DASHBOARD)

    dashRow = 21 ' Start row for consolidated cash (adjust as needed)

    ' Write header
    wsDash.Cells(20, 1).Value = "Bank"
    wsDash.Cells(20, 2).Value = "Month"
    wsDash.Cells(20, 3).Value = "Year"
    wsDash.Cells(20, 4).Value = "Kind"
    wsDash.Cells(20, 5).Value = "Total Inflows"
    wsDash.Cells(20, 6).Value = "Total Outflows"
    wsDash.Cells(20, 7).Value = "Net Value"

    ' Process BANKS data
    lastRowBanks = GetLastRow(wsBanks)

    For i = 2 To lastRowBanks
        ShowProgressBar i - 1, lastRowBanks - 1, "Updating Consolidated Cash"

        wsDash.Cells(dashRow, 1).Value = wsBanks.Cells(i, 1).Value ' Bank
        wsDash.Cells(dashRow, 2).Value = Month(wsBanks.Cells(i, 2).Value) ' Month
        wsDash.Cells(dashRow, 3).Value = Year(wsBanks.Cells(i, 2).Value) ' Year
        wsDash.Cells(dashRow, 4).Value = "BANK"

        ' Inflows (positive values)
        If wsBanks.Cells(i, 4).Value > 0 Then
            wsDash.Cells(dashRow, 5).Value = wsBanks.Cells(i, 4).Value
            wsDash.Cells(dashRow, 6).Value = 0
        Else
            wsDash.Cells(dashRow, 5).Value = 0
            wsDash.Cells(dashRow, 6).Value = Abs(wsBanks.Cells(i, 4).Value)
        End If

        wsDash.Cells(dashRow, 7).Value = wsBanks.Cells(i, 4).Value ' Net

        dashRow = dashRow + 1
    Next i

    ' Process INVESTMENTS data
    lastRowInv = GetLastRow(wsInvestments)

    For i = 2 To lastRowInv
        wsDash.Cells(dashRow, 1).Value = wsInvestments.Cells(i, 1).Value ' Institution
        wsDash.Cells(dashRow, 2).Value = Month(wsInvestments.Cells(i, 2).Value) ' Month
        wsDash.Cells(dashRow, 3).Value = Year(wsInvestments.Cells(i, 2).Value) ' Year
        wsDash.Cells(dashRow, 4).Value = "INVESTMENT"

        If wsInvestments.Cells(i, 4).Value > 0 Then
            wsDash.Cells(dashRow, 5).Value = wsInvestments.Cells(i, 4).Value
            wsDash.Cells(dashRow, 6).Value = 0
        Else
            wsDash.Cells(dashRow, 5).Value = 0
            wsDash.Cells(dashRow, 6).Value = Abs(wsInvestments.Cells(i, 4).Value)
        End If

        wsDash.Cells(dashRow, 7).Value = wsInvestments.Cells(i, 4).Value

        dashRow = dashRow + 1
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error updating consolidated cash: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: UpdateConsolidatedCards
' PURPOSE: Populate Consolidate Cards table
' =========================================================================
Private Sub UpdateConsolidatedCards()
    On Error GoTo ErrorHandler

    Dim wsCards As Worksheet
    Dim wsDash As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dashRow As Long

    Set wsCards = ThisWorkbook.Worksheets(WS_CARDS)
    Set wsDash = ThisWorkbook.Worksheets(WS_DASHBOARD)

    dashRow = 21 ' Start row for consolidated cards

    ' Write header (column I onwards)
    wsDash.Cells(20, 9).Value = "Bank"
    wsDash.Cells(20, 10).Value = "Month"
    wsDash.Cells(20, 11).Value = "Year"
    wsDash.Cells(20, 12).Value = "Kind"
    wsDash.Cells(20, 13).Value = "Total Value"

    lastRow = GetLastRow(wsCards)

    For i = 2 To lastRow
        ShowProgressBar i - 1, lastRow - 1, "Updating Consolidated Cards"

        wsDash.Cells(dashRow, 9).Value = wsCards.Cells(i, 1).Value ' Bank
        wsDash.Cells(dashRow, 10).Value = Month(wsCards.Cells(i, 3).Value) ' Month
        wsDash.Cells(dashRow, 11).Value = Year(wsCards.Cells(i, 3).Value) ' Year
        wsDash.Cells(dashRow, 12).Value = "CARDS"
        wsDash.Cells(dashRow, 13).Value = wsCards.Cells(i, 7).Value ' Value

        dashRow = dashRow + 1
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error updating consolidated cards: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: UpdateConsolidatedTransactions
' PURPOSE: Populate Consolidate Transactions table
' =========================================================================
Private Sub UpdateConsolidatedTransactions()
    On Error GoTo ErrorHandler

    Dim wsBanks As Worksheet
    Dim wsCards As Worksheet
    Dim wsDash As Worksheet
    Dim lastRowBanks As Long
    Dim lastRowCards As Long
    Dim i As Long
    Dim dashRow As Long

    Set wsBanks = ThisWorkbook.Worksheets(WS_BANKS)
    Set wsCards = ThisWorkbook.Worksheets(WS_CARDS)
    Set wsDash = ThisWorkbook.Worksheets(WS_DASHBOARD)

    dashRow = 1021 ' Start row for consolidated transactions

    ' Write header
    wsDash.Cells(1020, 1).Value = "Bank"
    wsDash.Cells(1020, 2).Value = "Month"
    wsDash.Cells(1020, 3).Value = "Year"
    wsDash.Cells(1020, 4).Value = "Category"
    wsDash.Cells(1020, 5).Value = "Total Value"

    ' Process banks
    lastRowBanks = GetLastRow(wsBanks)

    For i = 2 To lastRowBanks
        If Trim(wsBanks.Cells(i, 5).Value) <> "" And _
           Trim(wsBanks.Cells(i, 5).Value) <> "UNCLASSIFIED" Then

            wsDash.Cells(dashRow, 1).Value = wsBanks.Cells(i, 1).Value ' Bank
            wsDash.Cells(dashRow, 2).Value = Month(wsBanks.Cells(i, 2).Value)
            wsDash.Cells(dashRow, 3).Value = Year(wsBanks.Cells(i, 2).Value)
            wsDash.Cells(dashRow, 4).Value = wsBanks.Cells(i, 5).Value ' Category
            wsDash.Cells(dashRow, 5).Value = wsBanks.Cells(i, 4).Value ' Value

            dashRow = dashRow + 1
        End If
    Next i

    ' Process cards
    lastRowCards = GetLastRow(wsCards)

    For i = 2 To lastRowCards
        If Trim(wsCards.Cells(i, 8).Value) <> "" And _
           Trim(wsCards.Cells(i, 8).Value) <> "UNCLASSIFIED" Then

            wsDash.Cells(dashRow, 1).Value = wsCards.Cells(i, 1).Value ' Bank
            wsDash.Cells(dashRow, 2).Value = Month(wsCards.Cells(i, 3).Value)
            wsDash.Cells(dashRow, 3).Value = Year(wsCards.Cells(i, 3).Value)
            wsDash.Cells(dashRow, 4).Value = wsCards.Cells(i, 8).Value ' Category
            wsDash.Cells(dashRow, 5).Value = wsCards.Cells(i, 7).Value ' Value

            dashRow = dashRow + 1
        End If
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error updating consolidated transactions: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: UpdateConsolidatedNetDebts
' PURPOSE: Populate Consolidate Net Debts table
' =========================================================================
Private Sub UpdateConsolidatedNetDebts()
    On Error GoTo ErrorHandler

    Dim wsDebts As Worksheet
    Dim wsDash As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dashRow As Long

    Set wsDebts = ThisWorkbook.Worksheets(WS_DEBTS)
    Set wsDash = ThisWorkbook.Worksheets(WS_DASHBOARD)

    dashRow = 1021 ' Start row

    ' Write header (column I onwards)
    wsDash.Cells(1020, 9).Value = "Creditor"
    wsDash.Cells(1020, 10).Value = "Month"
    wsDash.Cells(1020, 11).Value = "Year"
    wsDash.Cells(1020, 12).Value = "Opening Balance"
    wsDash.Cells(1020, 13).Value = "Payments"
    wsDash.Cells(1020, 14).Value = "Updated Balance"

    lastRow = GetLastRow(wsDebts)

    For i = 2 To lastRow
        wsDash.Cells(dashRow, 9).Value = wsDebts.Cells(i, 1).Value ' Creditor
        wsDash.Cells(dashRow, 10).Value = Month(Date) ' Current month
        wsDash.Cells(dashRow, 11).Value = Year(Date) ' Current year
        wsDash.Cells(dashRow, 12).Value = wsDebts.Cells(i, 3).Value ' Amount Paid
        wsDash.Cells(dashRow, 13).Value = 0 ' Payments (to be calculated from transactions)
        wsDash.Cells(dashRow, 14).Value = wsDebts.Cells(i, 4).Value ' Updated Amount

        dashRow = dashRow + 1
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error updating consolidated debts: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: UpdateKPIs
' PURPOSE: Calculate and update main KPIs
' =========================================================================
Private Sub UpdateKPIs()
    On Error GoTo ErrorHandler

    Dim wsBanks As Worksheet
    Dim wsCards As Worksheet
    Dim wsDash As Worksheet
    Dim lastRowBanks As Long
    Dim lastRowCards As Long
    Dim i As Long
    Dim totalIncome As Double
    Dim totalExpenses As Double
    Dim balance As Double

    Set wsBanks = ThisWorkbook.Worksheets(WS_BANKS)
    Set wsCards = ThisWorkbook.Worksheets(WS_CARDS)
    Set wsDash = ThisWorkbook.Worksheets(WS_DASHBOARD)

    totalIncome = 0
    totalExpenses = 0

    ' Calculate from banks
    lastRowBanks = GetLastRow(wsBanks)
    For i = 2 To lastRowBanks
        If wsBanks.Cells(i, 4).Value > 0 Then
            totalIncome = totalIncome + wsBanks.Cells(i, 4).Value
        Else
            totalExpenses = totalExpenses + Abs(wsBanks.Cells(i, 4).Value)
        End If
    Next i

    ' Calculate from cards
    lastRowCards = GetLastRow(wsCards)
    For i = 2 To lastRowCards
        totalExpenses = totalExpenses + Abs(wsCards.Cells(i, 7).Value)
    Next i

    balance = totalIncome - totalExpenses

    ' Write to dashboard (adjust cell positions as needed)
    wsDash.Range("B2").Value = totalIncome
    wsDash.Range("B3").Value = totalExpenses
    wsDash.Range("B4").Value = balance

    Exit Sub

ErrorHandler:
    MsgBox "Error updating KPIs: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: UpdateNamedRanges
' PURPOSE: Update all named ranges for dashboard
' =========================================================================
Private Sub UpdateNamedRanges()
    On Error Resume Next

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets(WS_DASHBOARD)

    ' Create named ranges for KPIs
    CreateNamedRange "Total_Income", "B2", wsDash
    CreateNamedRange "Total_Expenses", "B3", wsDash
    CreateNamedRange "Balance", "B4", wsDash

    ' Create named ranges for data tables (adjust as needed)
    CreateNamedRange "Consolidated_Cash", "A20:G1000", wsDash
    CreateNamedRange "Consolidated_Cards", "I20:M1000", wsDash
    CreateNamedRange "Consolidated_Transactions", "A1020:E2000", wsDash
    CreateNamedRange "Consolidated_Debts", "I1020:N2000", wsDash

End Sub
