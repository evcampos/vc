Attribute VB_Name = "modClassification"
Option Explicit

' =========================================================================
' MODULE: modClassification
' PURPOSE: Transaction classification using exact and fuzzy matching
' =========================================================================

' Structure to hold classification mapping
Private Type CategoryMapping
    keyword As String
    category As String
    subcategory As String
    matchType As String ' EXACT or PARTIAL
End Type

' =========================================================================
' SUB: ClassifyAllTransactions
' PURPOSE: Classify all transactions in BANKS and CARDS sheets
' =========================================================================
Public Sub ClassifyAllTransactions()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Classify bank transactions
    Call ClassifyBankTransactions

    ' Classify card transactions
    Call ClassifyCardTransactions

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar

    MsgBox "Classification completed in " & Format(Timer - startTime, "0.0") & " seconds.", _
           vbInformation, "Classification Complete"

    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ResetStatusBar
    MsgBox "Error classifying transactions: " & Err.Description, vbCritical, "Classification Error"
End Sub

' =========================================================================
' SUB: ClassifyBankTransactions
' PURPOSE: Classify transactions in BANKS worksheet
' =========================================================================
Private Sub ClassifyBankTransactions()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsCategories As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim description As String
    Dim category As String
    Dim subcategory As String
    Dim matchType As String

    Set ws = ThisWorkbook.Worksheets(WS_BANKS)
    Set wsCategories = ThisWorkbook.Worksheets(WS_CATEGORIES)

    lastRow = GetLastRow(ws)

    For i = 2 To lastRow
        ShowProgressBar i - 1, lastRow - 1, "Classifying Bank Transactions"

        ' Skip if already classified
        If Trim(ws.Cells(i, 5).Value) <> "" Then
            GoTo NextTransaction
        End If

        description = Trim(ws.Cells(i, 3).Value)

        ' Try to classify
        If ClassifyTransaction(description, category, subcategory, matchType) Then
            ws.Cells(i, 5).Value = category
            ws.Cells(i, 6).Value = subcategory
        Else
            ' Mark as unclassified
            ws.Cells(i, 5).Value = "UNCLASSIFIED"
            ws.Cells(i, 6).Value = "PENDING"
        End If

NextTransaction:
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error classifying bank transactions: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: ClassifyCardTransactions
' PURPOSE: Classify transactions in CARDS worksheet
' =========================================================================
Private Sub ClassifyCardTransactions()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim description As String
    Dim category As String
    Dim subcategory As String
    Dim matchType As String

    Set ws = ThisWorkbook.Worksheets(WS_CARDS)

    lastRow = GetLastRow(ws)

    For i = 2 To lastRow
        ShowProgressBar i - 1, lastRow - 1, "Classifying Card Transactions"

        ' Skip if already classified
        If Trim(ws.Cells(i, 8).Value) <> "" Then
            GoTo NextTransaction
        End If

        description = Trim(ws.Cells(i, 5).Value) ' Card description column

        ' Try to classify
        If ClassifyTransaction(description, category, subcategory, matchType) Then
            ws.Cells(i, 8).Value = category
            ws.Cells(i, 9).Value = subcategory
        Else
            ' Mark as unclassified
            ws.Cells(i, 8).Value = "UNCLASSIFIED"
            ws.Cells(i, 9).Value = "PENDING"
        End If

NextTransaction:
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error classifying card transactions: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' FUNCTION: ClassifyTransaction
' PURPOSE: Classify a single transaction using exact and fuzzy matching
' PARAMETERS:
'   description - Transaction description
'   category - Output category
'   subcategory - Output subcategory
'   matchType - Output match type (EXACT or PARTIAL)
' RETURNS: Boolean - True if classified, False if not
' =========================================================================
Private Function ClassifyTransaction(description As String, _
                                    ByRef category As String, _
                                    ByRef subcategory As String, _
                                    ByRef matchType As String) As Boolean
    On Error GoTo ErrorHandler

    Dim wsCategories As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim keyword As String
    Dim stdDescription As String
    Dim stdKeyword As String

    Set wsCategories = ThisWorkbook.Worksheets(WS_CATEGORIES)
    lastRow = GetLastRow(wsCategories)

    stdDescription = StandardizeText(description)

    ' First pass: Exact match
    For i = 2 To lastRow
        keyword = Trim(wsCategories.Cells(i, 3).Value) ' Keywords column
        stdKeyword = StandardizeText(keyword)

        If stdDescription = stdKeyword Then
            category = Trim(wsCategories.Cells(i, 1).Value)
            subcategory = Trim(wsCategories.Cells(i, 2).Value)
            matchType = "EXACT"
            ClassifyTransaction = True
            Exit Function
        End If
    Next i

    ' Second pass: Partial/fuzzy match
    For i = 2 To lastRow
        keyword = Trim(wsCategories.Cells(i, 3).Value)
        stdKeyword = StandardizeText(keyword)

        ' Check if keyword is contained in description
        If InStr(1, stdDescription, stdKeyword, vbTextCompare) > 0 Then
            category = Trim(wsCategories.Cells(i, 1).Value)
            subcategory = Trim(wsCategories.Cells(i, 2).Value)
            matchType = "PARTIAL"
            ClassifyTransaction = True
            Exit Function
        End If

        ' Also check reverse (description contained in keyword)
        If InStr(1, stdKeyword, stdDescription, vbTextCompare) > 0 Then
            category = Trim(wsCategories.Cells(i, 1).Value)
            subcategory = Trim(wsCategories.Cells(i, 2).Value)
            matchType = "PARTIAL"
            ClassifyTransaction = True
            Exit Function
        End If
    Next i

    ' No match found
    ClassifyTransaction = False
    Exit Function

ErrorHandler:
    ClassifyTransaction = False
End Function

' =========================================================================
' SUB: AddCategoryMapping
' PURPOSE: Add new category mapping from user input
' PARAMETERS:
'   description - Transaction description to map
'   category - Category to assign
'   subcategory - Subcategory to assign
' =========================================================================
Public Sub AddCategoryMapping(description As String, category As String, subcategory As String)
    On Error GoTo ErrorHandler

    Dim wsCategories As Worksheet
    Dim lastRow As Long

    Set wsCategories = ThisWorkbook.Worksheets(WS_CATEGORIES)
    lastRow = GetLastRow(wsCategories) + 1

    wsCategories.Cells(lastRow, 1).Value = category
    wsCategories.Cells(lastRow, 2).Value = subcategory
    wsCategories.Cells(lastRow, 3).Value = description
    wsCategories.Cells(lastRow, 4).Value = Now ' Date added

    MsgBox "Category mapping added successfully.", vbInformation, "Mapping Added"

    Exit Sub

ErrorHandler:
    MsgBox "Error adding category mapping: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' SUB: ShowUnclassifiedTransactions
' PURPOSE: Generate report of unclassified transactions
' =========================================================================
Public Sub ShowUnclassifiedTransactions()
    On Error GoTo ErrorHandler

    Dim wsBanks As Worksheet
    Dim wsCards As Worksheet
    Dim lastRowBanks As Long
    Dim lastRowCards As Long
    Dim i As Long
    Dim unclassifiedList As String
    Dim count As Long

    Set wsBanks = ThisWorkbook.Worksheets(WS_BANKS)
    Set wsCards = ThisWorkbook.Worksheets(WS_CARDS)

    lastRowBanks = GetLastRow(wsBanks)
    lastRowCards = GetLastRow(wsCards)

    unclassifiedList = "UNCLASSIFIED TRANSACTIONS:" & vbCrLf & vbCrLf

    count = 0

    ' Check banks
    unclassifiedList = unclassifiedList & "BANKS:" & vbCrLf
    For i = 2 To lastRowBanks
        If Trim(wsBanks.Cells(i, 5).Value) = "UNCLASSIFIED" Then
            unclassifiedList = unclassifiedList & "- " & wsBanks.Cells(i, 3).Value & vbCrLf
            count = count + 1
        End If
    Next i

    ' Check cards
    unclassifiedList = unclassifiedList & vbCrLf & "CARDS:" & vbCrLf
    For i = 2 To lastRowCards
        If Trim(wsCards.Cells(i, 8).Value) = "UNCLASSIFIED" Then
            unclassifiedList = unclassifiedList & "- " & wsCards.Cells(i, 5).Value & vbCrLf
            count = count + 1
        End If
    Next i

    If count = 0 Then
        MsgBox "All transactions are classified!", vbInformation, "Classification Status"
    Else
        MsgBox unclassifiedList & vbCrLf & "Total: " & count & " unclassified transactions", _
               vbInformation, "Unclassified Transactions"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error showing unclassified transactions: " & Err.Description, vbCritical, "Error"
End Sub
