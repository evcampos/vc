Attribute VB_Name = "modImportInvestments"
Option Explicit

' =========================================================================
' MODULE: modImportInvestments
' PURPOSE: Import investment transactions and correlate with bank movements
' =========================================================================

' =========================================================================
' SUB: ImportInvestments
' PURPOSE: Import investment transactions
' =========================================================================
Public Sub ImportInvestments()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    Dim filePath As String
    Dim ws As Worksheet
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetRow As Long

    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Get file path from configuration
    filePath = GetFilePath(DS_INVESTMENTS)

    If filePath = "" Then
        MsgBox "File path not configured for INVESTMENTS", vbExclamation, "Configuration Missing"
        GoTo CleanUp
    End If

    If Not FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbExclamation, "File Not Found"
        GoTo CleanUp
    End If

    ' Open source file
    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(1)

    ' Get target worksheet
    Set ws = ThisWorkbook.Worksheets(WS_INVESTMENTS)

    ' Find next available row
    targetRow = GetLastRow(ws) + 1

    ' Get last row of source data
    lastRow = GetLastRow(sourceWs)

    ' Import data (same structure as BANKS)
    For i = 2 To lastRow
        ShowProgressBar i - 1, lastRow - 1, "Importing Investments"

        ws.Cells(targetRow, 1).Value = sourceWs.Cells(i, 1).Value ' Bank/Institution
        ws.Cells(targetRow, 2).Value = sourceWs.Cells(i, 2).Value ' Date
        ws.Cells(targetRow, 3).Value = sourceWs.Cells(i, 3).Value ' Description
        ws.Cells(targetRow, 4).Value = CleanNumericValue(sourceWs.Cells(i, 4).Value) ' Value
        ws.Cells(targetRow, 5).Value = "" ' Category
        ws.Cells(targetRow, 6).Value = "" ' Subcategory
        ws.Cells(targetRow, 7).Value = "" ' Correlation ID (to be filled)
        ws.Cells(targetRow, 8).Value = "" ' Correlation Status
        ws.Cells(targetRow, 9).Value = Now ' Import timestamp

        targetRow = targetRow + 1
    Next i

    sourceWb.Close SaveChanges:=False

    ' After import, run correlation
    Call CorrelateInvestmentsWithBanks

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "Investment import completed in " & Format(Timer - startTime, "0.0") & " seconds.", _
           vbInformation, "Import Complete"

    Exit Sub

ErrorHandler:
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error importing investments: " & Err.Description, vbCritical, "Import Error"
End Sub

' =========================================================================
' SUB: CorrelateInvestmentsWithBanks
' PURPOSE: Correlate investment movements with bank transactions
' LOGIC:
'   - Investment Application: Withdrawal from BANK → Deposit into INVESTMENT
'   - Investment Redemption: Withdrawal from INVESTMENT → Deposit into BANK
' =========================================================================
Public Sub CorrelateInvestmentsWithBanks()
    On Error GoTo ErrorHandler

    Dim wsInvestments As Worksheet
    Dim wsBanks As Worksheet
    Dim lastRowInv As Long
    Dim lastRowBank As Long
    Dim i As Long, j As Long
    Dim invDate As Date, bankDate As Date
    Dim invValue As Double, bankValue As Double
    Dim invDesc As String, bankDesc As String
    Dim correlationID As String
    Dim matchFound As Boolean
    Dim tolerance As Double
    Dim dateTolerance As Long

    ' Tolerance settings
    tolerance = 0.01 ' Value tolerance (1 cent)
    dateTolerance = 3 ' Days tolerance for date matching

    Set wsInvestments = ThisWorkbook.Worksheets(WS_INVESTMENTS)
    Set wsBanks = ThisWorkbook.Worksheets(WS_BANKS)

    lastRowInv = GetLastRow(wsInvestments)
    lastRowBank = GetLastRow(wsBanks)

    Application.ScreenUpdating = False

    ' Loop through investment transactions
    For i = 2 To lastRowInv
        ShowProgressBar i - 1, lastRowInv - 1, "Correlating Investments"

        ' Skip if already correlated
        If Trim(wsInvestments.Cells(i, 7).Value) <> "" Then
            GoTo NextInvestment
        End If

        invDate = wsInvestments.Cells(i, 2).Value
        invValue = wsInvestments.Cells(i, 4).Value
        invDesc = UCase(Trim(wsInvestments.Cells(i, 3).Value))

        matchFound = False

        ' Look for corresponding bank transaction
        For j = 2 To lastRowBank
            ' Skip if bank transaction already correlated
            If Trim(wsBanks.Cells(j, 8).Value) <> "" Then
                GoTo NextBank
            End If

            bankDate = wsBanks.Cells(j, 2).Value
            bankValue = wsBanks.Cells(j, 4).Value
            bankDesc = UCase(Trim(wsBanks.Cells(j, 3).Value))

            ' Check for matching criteria
            ' 1. Date within tolerance
            ' 2. Opposite sign values (investment out = bank in, or vice versa)
            ' 3. Similar amounts (within tolerance)

            If Abs(invDate - bankDate) <= dateTolerance Then
                ' Check if values are opposite and similar in magnitude
                If Abs(Abs(invValue) - Abs(bankValue)) <= tolerance Then
                    ' Check if signs are opposite (one positive, one negative)
                    If (invValue > 0 And bankValue < 0) Or (invValue < 0 And bankValue > 0) Then
                        ' Match found!
                        correlationID = "CORR-" & Format(invDate, "YYYYMMDD") & "-" & i & "-" & j

                        ' Update investment record
                        wsInvestments.Cells(i, 7).Value = correlationID
                        wsInvestments.Cells(i, 8).Value = "MATCHED"

                        ' Update bank record (add correlation columns if needed)
                        ' Assuming column 8 is for correlation ID in banks
                        wsBanks.Cells(j, 8).Value = correlationID
                        wsBanks.Cells(j, 9).Value = "MATCHED_INV"

                        matchFound = True
                        Exit For
                    End If
                End If
            End If

NextBank:
        Next j

        ' If no match found, mark as unmatched
        If Not matchFound Then
            wsInvestments.Cells(i, 8).Value = "UNMATCHED"
        End If

NextInvestment:
    Next i

    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "Investment correlation completed.", vbInformation, "Correlation Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error correlating investments: " & Err.Description, vbCritical, "Correlation Error"
End Sub

' =========================================================================
' FUNCTION: GetCorrelationBalance
' PURPOSE: Calculate balance of correlated vs uncorrelated transactions
' RETURNS: Double - Net uncorrelated value (should be close to zero)
' =========================================================================
Public Function GetCorrelationBalance() As Double
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalUnmatched As Double

    Set ws = ThisWorkbook.Worksheets(WS_INVESTMENTS)
    lastRow = GetLastRow(ws)

    totalUnmatched = 0

    For i = 2 To lastRow
        If Trim(ws.Cells(i, 8).Value) = "UNMATCHED" Then
            totalUnmatched = totalUnmatched + ws.Cells(i, 4).Value
        End If
    Next i

    GetCorrelationBalance = totalUnmatched
End Function
