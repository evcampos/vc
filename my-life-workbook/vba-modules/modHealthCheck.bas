Attribute VB_Name = "modHealthCheck"
Option Explicit

' =========================================================================
' MODULE: modHealthCheck
' PURPOSE: Comprehensive health check and validation system
' =========================================================================

' Structure to hold health check results
Private Type HealthCheckResult
    checkName As String
    status As String ' PASS, FAIL, WARNING
    message As String
    details As String
End Type

' =========================================================================
' SUB: RunFullHealthCheck
' PURPOSE: Execute all health checks and display results
' =========================================================================
Public Sub RunFullHealthCheck()
    On Error GoTo ErrorHandler

    Dim results() As HealthCheckResult
    Dim resultCount As Integer
    Dim i As Integer
    Dim report As String
    Dim passCount As Integer
    Dim failCount As Integer
    Dim warningCount As Integer

    resultCount = 0
    ReDim results(1 To 100) ' Initial size

    Application.ScreenUpdating = False

    ' Run all checks
    Call CheckWorkbookStructure(results, resultCount)
    Call CheckImportedData(results, resultCount)
    Call CheckClassification(results, resultCount)
    Call CheckCorrelation(results, resultCount)
    Call CheckIndexData(results, resultCount)
    Call CheckDataIntegrity(results, resultCount)

    Application.ScreenUpdating = True

    ' Generate report
    report = "========================================" & vbCrLf
    report = report & "HEALTH CHECK REPORT" & vbCrLf
    report = report & "Generated: " & Format(Now, "YYYY-MM-DD HH:MM:SS") & vbCrLf
    report = report & "========================================" & vbCrLf & vbCrLf

    passCount = 0
    failCount = 0
    warningCount = 0

    For i = 1 To resultCount
        Select Case results(i).status
            Case "PASS": passCount = passCount + 1
            Case "FAIL": failCount = failCount + 1
            Case "WARNING": warningCount = warningCount + 1
        End Select

        report = report & "[" & results(i).status & "] " & results(i).checkName & vbCrLf
        report = report & "    " & results(i).message & vbCrLf

        If results(i).details <> "" Then
            report = report & "    Details: " & results(i).details & vbCrLf
        End If

        report = report & vbCrLf
    Next i

    report = report & "========================================" & vbCrLf
    report = report & "SUMMARY:" & vbCrLf
    report = report & "  PASSED:   " & passCount & vbCrLf
    report = report & "  WARNINGS: " & warningCount & vbCrLf
    report = report & "  FAILED:   " & failCount & vbCrLf
    report = report & "========================================" & vbCrLf

    ' Display report
    MsgBox report, vbInformation, "Health Check Results"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error running health check: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: CheckWorkbookStructure
' PURPOSE: Verify all required worksheets exist
' =========================================================================
Private Sub CheckWorkbookStructure(ByRef results() As HealthCheckResult, ByRef count As Integer)
    On Error Resume Next

    If ValidateWorkbookStructure() Then
        count = count + 1
        results(count).checkName = "Workbook Structure"
        results(count).status = "PASS"
        results(count).message = "All required worksheets exist"
        results(count).details = ""
    Else
        count = count + 1
        results(count).checkName = "Workbook Structure"
        results(count).status = "FAIL"
        results(count).message = "One or more required worksheets are missing"
        results(count).details = ""
    End If
End Sub

' =========================================================================
' SUB: CheckImportedData
' PURPOSE: Verify data has been imported
' =========================================================================
Private Sub CheckImportedData(ByRef results() As HealthCheckResult, ByRef count As Integer)
    On Error Resume Next

    Dim wsBanks As Worksheet
    Dim wsCards As Worksheet
    Dim wsInvestments As Worksheet
    Dim bankRows As Long
    Dim cardRows As Long
    Dim invRows As Long

    Set wsBanks = ThisWorkbook.Worksheets(WS_BANKS)
    Set wsCards = ThisWorkbook.Worksheets(WS_CARDS)
    Set wsInvestments = ThisWorkbook.Worksheets(WS_INVESTMENTS)

    bankRows = GetLastRow(wsBanks) - 1
    cardRows = GetLastRow(wsCards) - 1
    invRows = GetLastRow(wsInvestments) - 1

    count = count + 1
    results(count).checkName = "Imported Data - BANKS"

    If bankRows > 0 Then
        results(count).status = "PASS"
        results(count).message = bankRows & " transactions imported"
    Else
        results(count).status = "WARNING"
        results(count).message = "No bank transactions imported"
    End If
    results(count).details = ""

    count = count + 1
    results(count).checkName = "Imported Data - CARDS"

    If cardRows > 0 Then
        results(count).status = "PASS"
        results(count).message = cardRows & " transactions imported"
    Else
        results(count).status = "WARNING"
        results(count).message = "No card transactions imported"
    End If
    results(count).details = ""

    count = count + 1
    results(count).checkName = "Imported Data - INVESTMENTS"

    If invRows > 0 Then
        results(count).status = "PASS"
        results(count).message = invRows & " transactions imported"
    Else
        results(count).status = "WARNING"
        results(count).message = "No investment transactions imported"
    End If
    results(count).details = ""
End Sub

' =========================================================================
' SUB: CheckClassification
' PURPOSE: Verify transaction classification status
' =========================================================================
Private Sub CheckClassification(ByRef results() As HealthCheckResult, ByRef count As Integer)
    On Error Resume Next

    Dim wsBanks As Worksheet
    Dim wsCards As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim unclassifiedBanks As Long
    Dim unclassifiedCards As Long
    Dim totalBanks As Long
    Dim totalCards As Long

    Set wsBanks = ThisWorkbook.Worksheets(WS_BANKS)
    Set wsCards = ThisWorkbook.Worksheets(WS_CARDS)

    unclassifiedBanks = 0
    unclassifiedCards = 0

    ' Check banks
    lastRow = GetLastRow(wsBanks)
    totalBanks = lastRow - 1

    For i = 2 To lastRow
        If Trim(wsBanks.Cells(i, 5).Value) = "UNCLASSIFIED" Or _
           Trim(wsBanks.Cells(i, 5).Value) = "" Then
            unclassifiedBanks = unclassifiedBanks + 1
        End If
    Next i

    ' Check cards
    lastRow = GetLastRow(wsCards)
    totalCards = lastRow - 1

    For i = 2 To lastRow
        If Trim(wsCards.Cells(i, 8).Value) = "UNCLASSIFIED" Or _
           Trim(wsCards.Cells(i, 8).Value) = "" Then
            unclassifiedCards = unclassifiedCards + 1
        End If
    Next i

    count = count + 1
    results(count).checkName = "Transaction Classification"

    If unclassifiedBanks = 0 And unclassifiedCards = 0 Then
        results(count).status = "PASS"
        results(count).message = "All transactions classified"
        results(count).details = ""
    ElseIf unclassifiedBanks + unclassifiedCards < (totalBanks + totalCards) * 0.1 Then
        results(count).status = "WARNING"
        results(count).message = "Some transactions unclassified"
        results(count).details = unclassifiedBanks & " banks, " & unclassifiedCards & " cards"
    Else
        results(count).status = "FAIL"
        results(count).message = "Many transactions unclassified"
        results(count).details = unclassifiedBanks & " banks (" & _
                                Format(unclassifiedBanks / totalBanks * 100, "0.0") & "%), " & _
                                unclassifiedCards & " cards (" & _
                                Format(unclassifiedCards / totalCards * 100, "0.0") & "%)"
    End If
End Sub

' =========================================================================
' SUB: CheckCorrelation
' PURPOSE: Verify investment-bank correlation status
' =========================================================================
Private Sub CheckCorrelation(ByRef results() As HealthCheckResult, ByRef count As Integer)
    On Error Resume Next

    Dim wsInvestments As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim unmatched As Long
    Dim total As Long
    Dim balance As Double

    Set wsInvestments = ThisWorkbook.Worksheets(WS_INVESTMENTS)
    lastRow = GetLastRow(wsInvestments)
    total = lastRow - 1

    unmatched = 0

    For i = 2 To lastRow
        If Trim(wsInvestments.Cells(i, 8).Value) = "UNMATCHED" Then
            unmatched = unmatched + 1
        End If
    Next i

    balance = GetCorrelationBalance()

    count = count + 1
    results(count).checkName = "Investment Correlation"

    If total = 0 Then
        results(count).status = "WARNING"
        results(count).message = "No investment transactions to correlate"
        results(count).details = ""
    ElseIf unmatched = 0 Then
        results(count).status = "PASS"
        results(count).message = "All investments correlated"
        results(count).details = "Balance: " & Format(balance, "#,##0.00")
    ElseIf unmatched < total * 0.2 Then
        results(count).status = "WARNING"
        results(count).message = unmatched & " of " & total & " investments unmatched"
        results(count).details = "Balance: " & Format(balance, "#,##0.00")
    Else
        results(count).status = "FAIL"
        results(count).message = "Many investments unmatched"
        results(count).details = unmatched & " of " & total & " (" & _
                                Format(unmatched / total * 100, "0.0") & "%), Balance: " & _
                                Format(balance, "#,##0.00")
    End If
End Sub

' =========================================================================
' SUB: CheckIndexData
' PURPOSE: Verify index data availability and freshness
' =========================================================================
Private Sub CheckIndexData(ByRef results() As HealthCheckResult, ByRef count As Integer)
    On Error Resume Next

    Dim wsIndexes As Worksheet
    Dim lastRow As Long
    Dim lastDate As Date
    Dim daysSinceUpdate As Long

    Set wsIndexes = ThisWorkbook.Worksheets(WS_INDEXES)
    lastRow = GetLastRow(wsIndexes)

    count = count + 1
    results(count).checkName = "Index Data Availability"

    If lastRow <= 1 Then
        results(count).status = "FAIL"
        results(count).message = "No index data available"
        results(count).details = "Please update index data"
    Else
        ' Find most recent date
        lastDate = wsIndexes.Cells(2, 2).Value
        Dim i As Long
        For i = 3 To lastRow
            If wsIndexes.Cells(i, 2).Value > lastDate Then
                lastDate = wsIndexes.Cells(i, 2).Value
            End If
        Next i

        daysSinceUpdate = Date - lastDate

        If daysSinceUpdate <= 7 Then
            results(count).status = "PASS"
            results(count).message = "Index data is current"
            results(count).details = "Last update: " & Format(lastDate, "YYYY-MM-DD")
        ElseIf daysSinceUpdate <= 30 Then
            results(count).status = "WARNING"
            results(count).message = "Index data may be outdated"
            results(count).details = "Last update: " & Format(lastDate, "YYYY-MM-DD") & _
                                   " (" & daysSinceUpdate & " days ago)"
        Else
            results(count).status = "FAIL"
            results(count).message = "Index data is outdated"
            results(count).details = "Last update: " & Format(lastDate, "YYYY-MM-DD") & _
                                   " (" & daysSinceUpdate & " days ago)"
        End If
    End If
End Sub

' =========================================================================
' SUB: CheckDataIntegrity
' PURPOSE: Check for orphan records and inconsistencies
' =========================================================================
Private Sub CheckDataIntegrity(ByRef results() As HealthCheckResult, ByRef count As Integer)
    On Error Resume Next

    Dim wsBanks As Worksheet
    Dim wsCards As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim invalidDates As Long
    Dim invalidValues As Long

    Set wsBanks = ThisWorkbook.Worksheets(WS_BANKS)
    Set wsCards = ThisWorkbook.Worksheets(WS_CARDS)

    invalidDates = 0
    invalidValues = 0

    ' Check banks
    lastRow = GetLastRow(wsBanks)
    For i = 2 To lastRow
        If Not IsDate(wsBanks.Cells(i, 2).Value) Then
            invalidDates = invalidDates + 1
        End If
        If Not IsNumeric(wsBanks.Cells(i, 4).Value) Then
            invalidValues = invalidValues + 1
        End If
    Next i

    ' Check cards
    lastRow = GetLastRow(wsCards)
    For i = 2 To lastRow
        If Not IsDate(wsCards.Cells(i, 3).Value) Then
            invalidDates = invalidDates + 1
        End If
        If Not IsNumeric(wsCards.Cells(i, 7).Value) Then
            invalidValues = invalidValues + 1
        End If
    Next i

    count = count + 1
    results(count).checkName = "Data Integrity"

    If invalidDates = 0 And invalidValues = 0 Then
        results(count).status = "PASS"
        results(count).message = "All data is valid"
        results(count).details = ""
    Else
        results(count).status = "FAIL"
        results(count).message = "Data integrity issues found"
        results(count).details = "Invalid dates: " & invalidDates & ", Invalid values: " & invalidValues
    End If
End Sub
