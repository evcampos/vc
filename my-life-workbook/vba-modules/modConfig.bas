Attribute VB_Name = "modConfig"
Option Explicit

' =========================================================================
' MODULE: modConfig
' PURPOSE: Central configuration management for [MY LIFE] workbook
' DESCRIPTION: Manages file paths, constants, and global settings
' =========================================================================

' Worksheet name constants
Public Const WS_FILES_PATHS As String = "FILES PATHS"
Public Const WS_FILES_STRUCTURE As String = "FILES STRUCTURE"
Public Const WS_BANKS As String = "BANKS"
Public Const WS_CARDS As String = "CARDS"
Public Const WS_INVESTMENTS As String = "INVESTMENTS"
Public Const WS_OPUS As String = "OPUS"
Public Const WS_DEBTS As String = "DEBTS"
Public Const WS_INDEXES As String = "INDEXES"
Public Const WS_CATEGORIES As String = "CATEGORIES"
Public Const WS_DASHBOARD As String = "DASHBOARD"

' Source types
Public Enum DataSource
    DS_ITAU_BANK = 1
    DS_NUBANK_BANK = 2
    DS_C6_BANK = 3
    DS_BB_BANK = 4
    DS_ITAU_CARD = 5
    DS_NUBANK_CARD = 6
    DS_C6_CARD = 7
    DS_INVESTMENTS = 8
    DS_OPUS = 9
    DS_DEBTS = 10
End Enum

' Index types
Public Enum IndexType
    IDX_CDI = 1
    IDX_SELIC = 2
    IDX_IPCA = 3
    IDX_USD_BRL = 4
    IDX_FED_FUNDS = 5
End Enum

' Currency types
Public Enum Currency
    CUR_BRL = 1
    CUR_USD = 2
End Enum

' =========================================================================
' FUNCTION: GetFilePath
' PURPOSE: Retrieve file path from FILES PATHS sheet
' PARAMETERS:
'   sourceType - DataSource enum value
' RETURNS: String - File path or empty string if not found
' =========================================================================
Public Function GetFilePath(sourceType As DataSource) As String
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sourceName As String

    Set ws = ThisWorkbook.Worksheets(WS_FILES_PATHS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Get source name from enum
    sourceName = GetSourceName(sourceType)

    ' Search for source in column A and return path from column B
    For i = 2 To lastRow ' Assuming row 1 is header
        If Trim(ws.Cells(i, 1).Value) = sourceName Then
            GetFilePath = Trim(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i

    GetFilePath = ""
    Exit Function

ErrorHandler:
    MsgBox "Error retrieving file path for " & sourceName & ": " & Err.Description, _
           vbCritical, "Configuration Error"
    GetFilePath = ""
End Function

' =========================================================================
' FUNCTION: GetSourceName
' PURPOSE: Convert DataSource enum to readable name
' =========================================================================
Public Function GetSourceName(sourceType As DataSource) As String
    Select Case sourceType
        Case DS_ITAU_BANK: GetSourceName = "ITAU_BANK"
        Case DS_NUBANK_BANK: GetSourceName = "NUBANK_BANK"
        Case DS_C6_BANK: GetSourceName = "C6_BANK"
        Case DS_BB_BANK: GetSourceName = "BB_BANK"
        Case DS_ITAU_CARD: GetSourceName = "ITAU_CARD"
        Case DS_NUBANK_CARD: GetSourceName = "NUBANK_CARD"
        Case DS_C6_CARD: GetSourceName = "C6_CARD"
        Case DS_INVESTMENTS: GetSourceName = "INVESTMENTS"
        Case DS_OPUS: GetSourceName = "OPUS"
        Case DS_DEBTS: GetSourceName = "DEBTS"
        Case Else: GetSourceName = "UNKNOWN"
    End Select
End Function

' =========================================================================
' FUNCTION: GetIndexName
' PURPOSE: Convert IndexType enum to readable name
' =========================================================================
Public Function GetIndexName(idxType As IndexType) As String
    Select Case idxType
        Case IDX_CDI: GetIndexName = "CDI"
        Case IDX_SELIC: GetIndexName = "SELIC"
        Case IDX_IPCA: GetIndexName = "IPCA"
        Case IDX_USD_BRL: GetIndexName = "USD/BRL"
        Case IDX_FED_FUNDS: GetIndexName = "FED_FUNDS"
        Case Else: GetIndexName = "UNKNOWN"
    End Select
End Function

' =========================================================================
' FUNCTION: ValidateWorkbookStructure
' PURPOSE: Ensure all required worksheets exist
' RETURNS: Boolean - True if valid, False otherwise
' =========================================================================
Public Function ValidateWorkbookStructure() As Boolean
    On Error GoTo ErrorHandler

    Dim requiredSheets As Variant
    Dim sheet As Variant
    Dim ws As Worksheet
    Dim found As Boolean

    requiredSheets = Array(WS_FILES_PATHS, WS_FILES_STRUCTURE, WS_BANKS, _
                          WS_CARDS, WS_INVESTMENTS, WS_OPUS, WS_DEBTS, _
                          WS_INDEXES, WS_CATEGORIES, WS_DASHBOARD)

    For Each sheet In requiredSheets
        found = False
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = sheet Then
                found = True
                Exit For
            End If
        Next ws

        If Not found Then
            MsgBox "Required worksheet '" & sheet & "' not found!", _
                   vbCritical, "Structure Validation Failed"
            ValidateWorkbookStructure = False
            Exit Function
        End If
    Next sheet

    ValidateWorkbookStructure = True
    Exit Function

ErrorHandler:
    MsgBox "Error validating workbook structure: " & Err.Description, _
           vbCritical, "Validation Error"
    ValidateWorkbookStructure = False
End Function

' =========================================================================
' FUNCTION: GetCurrencySymbol
' PURPOSE: Get currency symbol from Currency enum
' =========================================================================
Public Function GetCurrencySymbol(cur As Currency) As String
    Select Case cur
        Case CUR_BRL: GetCurrencySymbol = "BRL"
        Case CUR_USD: GetCurrencySymbol = "USD"
        Case Else: GetCurrencySymbol = "UNKNOWN"
    End Select
End Function
