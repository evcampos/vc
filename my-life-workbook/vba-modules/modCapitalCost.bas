Attribute VB_Name = "modCapitalCost"
Option Explicit

' =========================================================================
' MODULE: modCapitalCost
' PURPOSE: Calculate capital costs using index-based adjustments
' =========================================================================

' =========================================================================
' FUNCTION: CalculateCapitalCost
' PURPOSE: Calculate capital cost for a given amount and period
' PARAMETERS:
'   principal - Initial amount
'   startDate - Start date
'   endDate - End date (defaults to today)
'   indexType - Index to use (CDI, SELIC, IPCA, USD/BRL, FED_FUNDS)
'   currency - Currency (BRL or USD)
' RETURNS: Double - Updated amount
' =========================================================================
Public Function CalculateCapitalCost(principal As Double, _
                                    startDate As Date, _
                                    Optional endDate As Date, _
                                    Optional idxType As IndexType = IDX_CDI, _
                                    Optional cur As Currency = CUR_BRL) As Double
    On Error GoTo ErrorHandler

    Dim wsIndexes As Worksheet
    Dim startFactor As Double
    Dim endFactor As Double
    Dim multiplier As Double

    ' Default end date to today
    If endDate = 0 Then endDate = Date

    ' Validate dates
    If startDate > endDate Then
        CalculateCapitalCost = principal
        Exit Function
    End If

    Set wsIndexes = ThisWorkbook.Worksheets(WS_INDEXES)

    ' Get cumulative factors for start and end dates
    startFactor = GetCumulativeFactor(idxType, startDate)
    endFactor = GetCumulativeFactor(idxType, endDate)

    If startFactor = 0 Or endFactor = 0 Then
        ' If factors not available, return principal without adjustment
        CalculateCapitalCost = principal
        Exit Function
    End If

    ' Calculate multiplier
    multiplier = endFactor / startFactor

    ' Apply to principal
    CalculateCapitalCost = principal * multiplier

    Exit Function

ErrorHandler:
    MsgBox "Error calculating capital cost: " & Err.Description, vbCritical, "Error"
    CalculateCapitalCost = principal
End Function

' =========================================================================
' FUNCTION: GetCumulativeFactor
' PURPOSE: Get cumulative index factor for a specific date
' PARAMETERS:
'   idxType - Index type
'   targetDate - Date to get factor for
' RETURNS: Double - Cumulative factor
' =========================================================================
Private Function GetCumulativeFactor(idxType As IndexType, targetDate As Date) As Double
    On Error GoTo ErrorHandler

    Dim wsIndexes As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim indexName As String
    Dim indexDate As Date
    Dim factor As Double
    Dim closestDate As Date
    Dim closestFactor As Double
    Dim minDiff As Long

    Set wsIndexes = ThisWorkbook.Worksheets(WS_INDEXES)
    lastRow = GetLastRow(wsIndexes)

    indexName = GetIndexName(idxType)
    minDiff = 999999
    closestFactor = 0

    ' Search for the date or closest earlier date
    For i = 2 To lastRow
        If Trim(wsIndexes.Cells(i, 1).Value) = indexName Then
            indexDate = wsIndexes.Cells(i, 2).Value

            ' Exact match
            If indexDate = targetDate Then
                GetCumulativeFactor = wsIndexes.Cells(i, 4).Value ' Cumulative factor column
                Exit Function
            End If

            ' Find closest earlier date
            If indexDate < targetDate Then
                If Abs(targetDate - indexDate) < minDiff Then
                    minDiff = Abs(targetDate - indexDate)
                    closestFactor = wsIndexes.Cells(i, 4).Value
                End If
            End If
        End If
    Next i

    ' Return closest factor
    GetCumulativeFactor = closestFactor

    Exit Function

ErrorHandler:
    GetCumulativeFactor = 0
End Function

' =========================================================================
' FUNCTION: GetIndexValue
' PURPOSE: Get index value for a specific date
' PARAMETERS:
'   idxType - Index type
'   targetDate - Date to get value for
' RETURNS: Double - Index value
' =========================================================================
Public Function GetIndexValue(idxType As IndexType, targetDate As Date) As Double
    On Error GoTo ErrorHandler

    Dim wsIndexes As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim indexName As String
    Dim indexDate As Date
    Dim minDiff As Long
    Dim closestValue As Double

    Set wsIndexes = ThisWorkbook.Worksheets(WS_INDEXES)
    lastRow = GetLastRow(wsIndexes)

    indexName = GetIndexName(idxType)
    minDiff = 999999
    closestValue = 0

    For i = 2 To lastRow
        If Trim(wsIndexes.Cells(i, 1).Value) = indexName Then
            indexDate = wsIndexes.Cells(i, 2).Value

            ' Exact match
            If indexDate = targetDate Then
                GetIndexValue = wsIndexes.Cells(i, 3).Value ' Index value column
                Exit Function
            End If

            ' Find closest date
            If Abs(targetDate - indexDate) < minDiff Then
                minDiff = Abs(targetDate - indexDate)
                closestValue = wsIndexes.Cells(i, 3).Value
            End If
        End If
    Next i

    GetIndexValue = closestValue

    Exit Function

ErrorHandler:
    GetIndexValue = 0
End Function

' =========================================================================
' SUB: UpdateDebtValues
' PURPOSE: Update all debts with current values based on capital cost
' =========================================================================
Public Sub UpdateDebtValues()
    On Error GoTo ErrorHandler

    Dim wsDebts As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim originalAmount As Double
    Dim interestRate As Double
    Dim currency As String
    Dim updatedAmount As Double
    Dim idxType As IndexType
    Dim cur As Currency
    Dim startDate As Date

    Set wsDebts = ThisWorkbook.Worksheets(WS_DEBTS)
    lastRow = GetLastRow(wsDebts)

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        ShowProgressBar i - 1, lastRow - 1, "Updating Debt Values"

        originalAmount = wsDebts.Cells(i, 3).Value ' Amount Paid column
        currency = UCase(Trim(wsDebts.Cells(i, 5).Value)) ' Currency column
        startDate = wsDebts.Cells(i, 6).Value ' Start date column

        ' Determine index based on currency
        If currency = "BRL" Then
            cur = CUR_BRL
            idxType = IDX_CDI ' Use CDI for BRL debts
        ElseIf currency = "USD" Then
            cur = CUR_USD
            idxType = IDX_FED_FUNDS ' Use Fed Funds for USD debts
        Else
            ' Default to BRL
            cur = CUR_BRL
            idxType = IDX_CDI
        End If

        ' Calculate updated amount
        updatedAmount = CalculateCapitalCost(originalAmount, startDate, Date, idxType, cur)

        ' Write updated amount
        wsDebts.Cells(i, 4).Value = updatedAmount ' Updated Amount column

    Next i

    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "Debt values updated successfully.", vbInformation, "Update Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error updating debt values: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: UpdateOPUSValues
' PURPOSE: Update OPUS investment values with capital cost
' =========================================================================
Public Sub UpdateOPUSValues()
    On Error GoTo ErrorHandler

    Dim wsOPUS As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim investmentCost As Double
    Dim capitalCost As Double
    Dim updatedCost As Double
    Dim startDate As Date
    Dim currency As String
    Dim idxType As IndexType
    Dim cur As Currency

    Set wsOPUS = ThisWorkbook.Worksheets(WS_OPUS)
    lastRow = GetLastRow(wsOPUS)

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        ShowProgressBar i - 1, lastRow - 1, "Updating OPUS Values"

        investmentCost = wsOPUS.Cells(i, 2).Value ' Investment Cost
        startDate = wsOPUS.Cells(i, 5).Value ' Start Date
        currency = UCase(Trim(wsOPUS.Cells(i, 6).Value)) ' Currency

        ' Determine index
        If currency = "BRL" Then
            idxType = IDX_CDI
            cur = CUR_BRL
        Else
            idxType = IDX_FED_FUNDS
            cur = CUR_USD
        End If

        ' Calculate updated cost
        updatedCost = CalculateCapitalCost(investmentCost, startDate, Date, idxType, cur)

        ' Write updated cost
        wsOPUS.Cells(i, 4).Value = updatedCost ' Updated Cost column

    Next i

    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "OPUS values updated successfully.", vbInformation, "Update Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error updating OPUS values: " & Err.Description, vbCritical, "Error"
End Sub
