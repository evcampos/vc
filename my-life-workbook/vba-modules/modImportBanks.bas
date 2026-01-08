Attribute VB_Name = "modImportBanks"
Option Explicit

' =========================================================================
' MODULE: modImportBanks
' PURPOSE: Import and normalize bank transaction data
' =========================================================================

' =========================================================================
' SUB: ImportAllBanks
' PURPOSE: Import transactions from all bank sources
' =========================================================================
Public Sub ImportAllBanks()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Import each bank
    ImportBank DS_ITAU_BANK, "ITAU"
    ImportBank DS_NUBANK_BANK, "NUBANK"
    ImportBank DS_C6_BANK, "C6"
    ImportBank DS_BB_BANK, "BANCO DO BRASIL"

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "All bank imports completed in " & Format(Timer - startTime, "0.0") & " seconds.", _
           vbInformation, "Import Complete"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error importing banks: " & Err.Description, vbCritical, "Import Error"
End Sub

' =========================================================================
' SUB: ImportBank
' PURPOSE: Import transactions from a specific bank
' PARAMETERS:
'   sourceType - DataSource enum value
'   bankName - Display name of bank
' =========================================================================
Private Sub ImportBank(sourceType As DataSource, bankName As String)
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim ws As Worksheet
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetRow As Long

    ' Get file path from configuration
    filePath = GetFilePath(sourceType)

    If filePath = "" Then
        MsgBox "File path not configured for " & bankName, vbExclamation, "Configuration Missing"
        Exit Sub
    End If

    If Not FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbExclamation, "File Not Found"
        Exit Sub
    End If

    ' Open source file
    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(1) ' Assume first sheet

    ' Validate structure
    If Not ValidateBankStructure(sourceWs) Then
        sourceWb.Close SaveChanges:=False
        MsgBox "Invalid file structure for " & bankName, vbExclamation, "Validation Failed"
        Exit Sub
    End If

    ' Get target worksheet
    Set ws = ThisWorkbook.Worksheets(WS_BANKS)

    ' Find next available row
    targetRow = GetLastRow(ws) + 1

    ' Get last row of source data
    lastRow = GetLastRow(sourceWs)

    ' Import data
    For i = 2 To lastRow ' Skip header
        ShowProgressBar i - 1, lastRow - 1, "Importing " & bankName

        ' Write standardized data
        ws.Cells(targetRow, 1).Value = bankName ' Bank
        ws.Cells(targetRow, 2).Value = sourceWs.Cells(i, GetBankDateColumn(sourceType)).Value ' Date
        ws.Cells(targetRow, 3).Value = sourceWs.Cells(i, GetBankDescriptionColumn(sourceType)).Value ' Description
        ws.Cells(targetRow, 4).Value = CleanNumericValue(sourceWs.Cells(i, GetBankValueColumn(sourceType)).Value) ' Value
        ws.Cells(targetRow, 5).Value = "" ' Category (to be classified)
        ws.Cells(targetRow, 6).Value = "" ' Subcategory (to be classified)
        ws.Cells(targetRow, 7).Value = Now ' Import timestamp

        targetRow = targetRow + 1
    Next i

    ' Close source file
    sourceWb.Close SaveChanges:=False

    Exit Sub

ErrorHandler:
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    MsgBox "Error importing " & bankName & ": " & Err.Description, vbCritical, "Import Error"
End Sub

' =========================================================================
' FUNCTION: ValidateBankStructure
' PURPOSE: Validate that source file has required columns
' PARAMETERS:
'   ws - Source worksheet
' RETURNS: Boolean
' =========================================================================
Private Function ValidateBankStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler

    ' Basic validation - ensure at least 3 columns exist
    If GetLastCol(ws) < 3 Then
        ValidateBankStructure = False
        Exit Function
    End If

    ' Ensure there's data beyond header
    If GetLastRow(ws) < 2 Then
        ValidateBankStructure = False
        Exit Function
    End If

    ValidateBankStructure = True
    Exit Function

ErrorHandler:
    ValidateBankStructure = False
End Function

' =========================================================================
' FUNCTION: GetBankDateColumn
' PURPOSE: Get date column index for each bank
' PARAMETERS:
'   sourceType - DataSource enum
' RETURNS: Long - Column index
' =========================================================================
Private Function GetBankDateColumn(sourceType As DataSource) As Long
    ' This should be configured in FILES STRUCTURE sheet
    ' For now, using defaults
    Select Case sourceType
        Case DS_ITAU_BANK: GetBankDateColumn = 1
        Case DS_NUBANK_BANK: GetBankDateColumn = 1
        Case DS_C6_BANK: GetBankDateColumn = 1
        Case DS_BB_BANK: GetBankDateColumn = 1
        Case Else: GetBankDateColumn = 1
    End Select
End Function

' =========================================================================
' FUNCTION: GetBankDescriptionColumn
' PURPOSE: Get description column index for each bank
' =========================================================================
Private Function GetBankDescriptionColumn(sourceType As DataSource) As Long
    Select Case sourceType
        Case DS_ITAU_BANK: GetBankDescriptionColumn = 2
        Case DS_NUBANK_BANK: GetBankDescriptionColumn = 2
        Case DS_C6_BANK: GetBankDescriptionColumn = 2
        Case DS_BB_BANK: GetBankDescriptionColumn = 2
        Case Else: GetBankDescriptionColumn = 2
    End Select
End Function

' =========================================================================
' FUNCTION: GetBankValueColumn
' PURPOSE: Get value column index for each bank
' =========================================================================
Private Function GetBankValueColumn(sourceType As DataSource) As Long
    Select Case sourceType
        Case DS_ITAU_BANK: GetBankValueColumn = 3
        Case DS_NUBANK_BANK: GetBankValueColumn = 3
        Case DS_C6_BANK: GetBankValueColumn = 3
        Case DS_BB_BANK: GetBankValueColumn = 3
        Case Else: GetBankValueColumn = 3
    End Select
End Function
