Attribute VB_Name = "modImportCards"
Option Explicit

' =========================================================================
' MODULE: modImportCards
' PURPOSE: Import and normalize credit card transaction data
' =========================================================================

' =========================================================================
' SUB: ImportAllCards
' PURPOSE: Import transactions from all card sources
' =========================================================================
Public Sub ImportAllCards()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Import each card
    ImportCard DS_ITAU_CARD, "ITAU"
    ImportCard DS_NUBANK_CARD, "NUBANK"
    ImportCard DS_C6_CARD, "C6"

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "All card imports completed in " & Format(Timer - startTime, "0.0") & " seconds.", _
           vbInformation, "Import Complete"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error importing cards: " & Err.Description, vbCritical, "Import Error"
End Sub

' =========================================================================
' SUB: ImportCard
' PURPOSE: Import transactions from a specific card
' PARAMETERS:
'   sourceType - DataSource enum value
'   cardName - Display name of card
' =========================================================================
Private Sub ImportCard(sourceType As DataSource, cardName As String)
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim ws As Worksheet
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetRow As Long
    Dim cardNumber As String
    Dim purchaseDate As Date
    Dim category As String
    Dim description As String
    Dim installment As String
    Dim value As Double

    ' Get file path from configuration
    filePath = GetFilePath(sourceType)

    If filePath = "" Then
        MsgBox "File path not configured for " & cardName, vbExclamation, "Configuration Missing"
        Exit Sub
    End If

    If Not FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbExclamation, "File Not Found"
        Exit Sub
    End If

    ' Open source file
    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(1)

    ' Get target worksheet
    Set ws = ThisWorkbook.Worksheets(WS_CARDS)

    ' Find next available row
    targetRow = GetLastRow(ws) + 1

    ' Get last row of source data
    lastRow = GetLastRow(sourceWs)

    ' Import data
    For i = 2 To lastRow ' Skip header
        ShowProgressBar i - 1, lastRow - 1, "Importing " & cardName

        ' Extract data from source (adjust column indices based on FILES STRUCTURE)
        cardNumber = sourceWs.Cells(i, GetCardNumberColumn(sourceType)).Value
        purchaseDate = sourceWs.Cells(i, GetCardDateColumn(sourceType)).Value
        category = sourceWs.Cells(i, GetCardCategoryColumn(sourceType)).Value
        description = sourceWs.Cells(i, GetCardDescriptionColumn(sourceType)).Value
        installment = sourceWs.Cells(i, GetCardInstallmentColumn(sourceType)).Value
        value = CleanNumericValue(sourceWs.Cells(i, GetCardValueColumn(sourceType)).Value)

        ' Write standardized data
        ws.Cells(targetRow, 1).Value = cardName ' Bank
        ws.Cells(targetRow, 2).Value = cardNumber ' Card Number
        ws.Cells(targetRow, 3).Value = purchaseDate ' Purchase Date
        ws.Cells(targetRow, 4).Value = category ' Category (raw)
        ws.Cells(targetRow, 5).Value = description ' Description
        ws.Cells(targetRow, 6).Value = installment ' Installment
        ws.Cells(targetRow, 7).Value = value ' Value
        ws.Cells(targetRow, 8).Value = "" ' Classified Category (to be filled)
        ws.Cells(targetRow, 9).Value = "" ' Classified Subcategory (to be filled)
        ws.Cells(targetRow, 10).Value = Now ' Import timestamp

        targetRow = targetRow + 1
    Next i

    ' Close source file
    sourceWb.Close SaveChanges:=False

    Exit Sub

ErrorHandler:
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    MsgBox "Error importing " & cardName & ": " & Err.Description, vbCritical, "Import Error"
End Sub

' =========================================================================
' Column mapping functions for card imports
' These should eventually read from FILES STRUCTURE sheet
' =========================================================================

Private Function GetCardNumberColumn(sourceType As DataSource) As Long
    Select Case sourceType
        Case DS_ITAU_CARD: GetCardNumberColumn = 1
        Case DS_NUBANK_CARD: GetCardNumberColumn = 1
        Case DS_C6_CARD: GetCardNumberColumn = 1
        Case Else: GetCardNumberColumn = 1
    End Select
End Function

Private Function GetCardDateColumn(sourceType As DataSource) As Long
    Select Case sourceType
        Case DS_ITAU_CARD: GetCardDateColumn = 2
        Case DS_NUBANK_CARD: GetCardDateColumn = 2
        Case DS_C6_CARD: GetCardDateColumn = 2
        Case Else: GetCardDateColumn = 2
    End Select
End Function

Private Function GetCardCategoryColumn(sourceType As DataSource) As Long
    Select Case sourceType
        Case DS_ITAU_CARD: GetCardCategoryColumn = 3
        Case DS_NUBANK_CARD: GetCardCategoryColumn = 3
        Case DS_C6_CARD: GetCardCategoryColumn = 3
        Case Else: GetCardCategoryColumn = 3
    End Select
End Function

Private Function GetCardDescriptionColumn(sourceType As DataSource) As Long
    Select Case sourceType
        Case DS_ITAU_CARD: GetCardDescriptionColumn = 4
        Case DS_NUBANK_CARD: GetCardDescriptionColumn = 4
        Case DS_C6_CARD: GetCardDescriptionColumn = 4
        Case Else: GetCardDescriptionColumn = 4
    End Select
End Function

Private Function GetCardInstallmentColumn(sourceType As DataSource) As Long
    Select Case sourceType
        Case DS_ITAU_CARD: GetCardInstallmentColumn = 5
        Case DS_NUBANK_CARD: GetCardInstallmentColumn = 5
        Case DS_C6_CARD: GetCardInstallmentColumn = 5
        Case Else: GetCardInstallmentColumn = 5
    End Select
End Function

Private Function GetCardValueColumn(sourceType As DataSource) As Long
    Select Case sourceType
        Case DS_ITAU_CARD: GetCardValueColumn = 6
        Case DS_NUBANK_CARD: GetCardValueColumn = 6
        Case DS_C6_CARD: GetCardValueColumn = 6
        Case Else: GetCardValueColumn = 6
    End Select
End Function
