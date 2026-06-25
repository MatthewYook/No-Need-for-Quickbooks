' =============================================
' MODULE: Core_Accounting
' Handles journal entries, trial balance, and financial statements
' =============================================

Option Explicit

Public Const JOURNAL_SHEET As String = "Journal"
Public Const CHART_SHEET As String = "ChartOfAccounts"
Public Const PNL_SHEET As String = "P&L"
Public Const BALANCE_SHEET As String = "BalanceSheet"

Public Sub AddJournalEntry(accountName As String, debit As Double, credit As Double, description As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    
    On Error GoTo ErrorHandler
    Set ws = GetWorksheet(JOURNAL_SHEET)
    ValidateJournalEntry accountName, debit, credit
    
    nextRow = GetLastRow(ws, 1) + 1
    ws.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd")
    ws.Cells(nextRow, 2).Value = Trim(accountName)
    ws.Cells(nextRow, 3).Value = debit
    ws.Cells(nextRow, 4).Value = credit
    ws.Cells(nextRow, 5).Value = Trim(description)
    Exit Sub

ErrorHandler:
    MsgBox "AddJournalEntry failed: " & Err.Description, vbExclamation, "Core Accounting"
End Sub

Public Sub UpdateTrialBalance()
    Dim wsChart As Worksheet
    Dim wsJournal As Worksheet
    Dim chartRow As Long
    Dim accountName As String
    Dim totalDebit As Double
    Dim totalCredit As Double
    
    On Error GoTo ErrorHandler
    Set wsChart = GetWorksheet(CHART_SHEET)
    Set wsJournal = GetWorksheet(JOURNAL_SHEET)
    
    For chartRow = 2 To GetLastRow(wsChart, 1)
        accountName = Trim(wsChart.Cells(chartRow, 1).Value)
        If Len(accountName) > 0 Then
            totalDebit = Application.WorksheetFunction.SumIfs(wsJournal.Range("C:C"), wsJournal.Range("B:B"), accountName)
            totalCredit = Application.WorksheetFunction.SumIfs(wsJournal.Range("D:D"), wsJournal.Range("B:B"), accountName)
            wsChart.Cells(chartRow, 2).Value = totalDebit - totalCredit
        Else
            wsChart.Cells(chartRow, 2).ClearContents
        End If
    Next chartRow
    Exit Sub

ErrorHandler:
    MsgBox "UpdateTrialBalance failed: " & Err.Description, vbExclamation, "Core Accounting"
End Sub

Public Sub UpdateFinancialStatements()
    Dim wsChart As Worksheet
    Dim wsPnL As Worksheet
    Dim wsBalance As Worksheet
    Dim chartRow As Long
    Dim accountType As String
    Dim accountBalance As Double
    Dim revenueTotal As Double
    Dim expenseTotal As Double
    Dim assetTotal As Double
    Dim liabilityTotal As Double
    Dim equityTotal As Double
    Dim revenueRow As Long
    Dim expenseRow As Long
    Dim assetRow As Long
    Dim liabilityRow As Long
    Dim equityRow As Long

    On Error GoTo ErrorHandler
    Set wsChart = GetWorksheet(CHART_SHEET)
    Set wsPnL = GetWorksheet(PNL_SHEET)
    Set wsBalance = GetWorksheet(BALANCE_SHEET)
    
    ClearWorksheet wsPnL
    ClearWorksheet wsBalance
    
    wsPnL.Cells(1, 1).Value = "Profit and Loss Statement"
    wsPnL.Cells(2, 1).Value = "As of " & Format(Date, "mmmm dd, yyyy")
    wsPnL.Cells(4, 1).Value = "Revenue"
    wsPnL.Cells(6, 1).Value = "Expenses"
    wsPnL.Cells(8, 1).Value = "Totals"
    
    wsBalance.Cells(1, 1).Value = "Balance Sheet"
    wsBalance.Cells(2, 1).Value = "As of " & Format(Date, "mmmm dd, yyyy")
    wsBalance.Cells(4, 1).Value = "Assets"
    wsBalance.Cells(4, 3).Value = "Liabilities"
    wsBalance.Cells(4, 5).Value = "Equity"
    
    revenueRow = 5
    expenseRow = 7
    assetRow = 5
    liabilityRow = 5
    equityRow = 5
    
    For chartRow = 2 To GetLastRow(wsChart, 1)
        accountType = Trim(UCase(wsChart.Cells(chartRow, 3).Value))
        accountBalance = wsChart.Cells(chartRow, 2).Value
        
        If Len(accountType) = 0 Then GoTo ContinueLoop
        
        Select Case accountType
            Case "REVENUE"
                wsPnL.Cells(revenueRow, 1).Value = wsChart.Cells(chartRow, 1).Value
                wsPnL.Cells(revenueRow, 2).Value = accountBalance
                revenueTotal = revenueTotal + accountBalance
                revenueRow = revenueRow + 1
            Case "EXPENSE"
                wsPnL.Cells(expenseRow, 1).Value = wsChart.Cells(chartRow, 1).Value
                wsPnL.Cells(expenseRow, 2).Value = accountBalance
                expenseTotal = expenseTotal + accountBalance
                expenseRow = expenseRow + 1
            Case "ASSET"
                wsBalance.Cells(assetRow, 1).Value = wsChart.Cells(chartRow, 1).Value
                wsBalance.Cells(assetRow, 2).Value = accountBalance
                assetTotal = assetTotal + accountBalance
                assetRow = assetRow + 1
            Case "LIABILITY"
                wsBalance.Cells(liabilityRow, 3).Value = wsChart.Cells(chartRow, 1).Value
                wsBalance.Cells(liabilityRow, 4).Value = accountBalance
                liabilityTotal = liabilityTotal + accountBalance
                liabilityRow = liabilityRow + 1
            Case "EQUITY"
                wsBalance.Cells(equityRow, 5).Value = wsChart.Cells(chartRow, 1).Value
                wsBalance.Cells(equityRow, 6).Value = accountBalance
                equityTotal = equityTotal + accountBalance
                equityRow = equityRow + 1
        End Select
ContinueLoop:
    Next chartRow
    
    wsPnL.Cells(revenueRow + 1, 1).Value = "Total Revenue"
    wsPnL.Cells(revenueRow + 1, 2).Value = revenueTotal
    wsPnL.Cells(expenseRow + 1, 1).Value = "Total Expenses"
    wsPnL.Cells(expenseRow + 1, 2).Value = expenseTotal
    wsPnL.Cells(expenseRow + 3, 1).Value = "Net Income"
    wsPnL.Cells(expenseRow + 3, 2).Value = revenueTotal - expenseTotal
    
    wsBalance.Cells(assetRow + 1, 1).Value = "Total Assets"
    wsBalance.Cells(assetRow + 1, 2).Value = assetTotal
    wsBalance.Cells(liabilityRow + 1, 3).Value = "Total Liabilities"
    wsBalance.Cells(liabilityRow + 1, 4).Value = liabilityTotal
    wsBalance.Cells(equityRow + 1, 5).Value = "Total Equity"
    wsBalance.Cells(equityRow + 1, 6).Value = equityTotal
    wsBalance.Cells(equityRow + 3, 5).Value = "Liabilities + Equity"
    wsBalance.Cells(equityRow + 3, 6).Value = liabilityTotal + equityTotal
    
    FormatStatementRange wsPnL, 1, 1, expenseRow + 3, 2
    FormatStatementRange wsBalance, 1, 1, equityRow + 3, 6
    Exit Sub

ErrorHandler:
    MsgBox "UpdateFinancialStatements failed: " & Err.Description, vbExclamation, "Core Accounting"
End Sub

Public Sub RefreshAccounting()
    UpdateTrialBalance
    UpdateFinancialStatements
End Sub

Public Function IsBalanced(Optional tolerance As Double = 0.01) As Boolean
    Dim wsJournal As Worksheet
    Dim totalDebit As Double
    Dim totalCredit As Double
    
    On Error GoTo ErrorHandler
    Set wsJournal = GetWorksheet(JOURNAL_SHEET)
    totalDebit = Application.WorksheetFunction.Sum(wsJournal.Range("C:C"))
    totalCredit = Application.WorksheetFunction.Sum(wsJournal.Range("D:D"))
    IsBalanced = (Abs(totalDebit - totalCredit) <= tolerance)
    Exit Function

ErrorHandler:
    MsgBox "IsBalanced failed: " & Err.Description, vbExclamation, "Core Accounting"
    IsBalanced = False
End Function

Private Function GetWorksheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetWorksheet Is Nothing Then
        Err.Raise vbObjectError + 513, "GetWorksheet", "Worksheet '" & sheetName & "' not found."
    End If
End Function

Private Function GetLastRow(ws As Worksheet, Optional columnIndex As Long = 1) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
End Function

Private Sub ClearWorksheet(ws As Worksheet)
    ws.Cells.ClearContents
    ws.Cells.ClearFormats
End Sub

Private Sub ValidateJournalEntry(accountName As String, debit As Double, credit As Double)
    If Len(Trim(accountName)) = 0 Then Err.Raise vbObjectError + 514, "ValidateJournalEntry", "Account name is required."
    If debit < 0 Or credit < 0 Then Err.Raise vbObjectError + 515, "ValidateJournalEntry", "Debit and credit amounts must be non-negative."
    If debit = 0 And credit = 0 Then Err.Raise vbObjectError + 516, "ValidateJournalEntry", "Journal entry must include a debit or credit amount."
    If debit > 0 And credit > 0 Then Err.Raise vbObjectError + 517, "ValidateJournalEntry", "Journal entry cannot contain both debit and credit amounts."
End Sub

Private Sub FormatStatementRange(ws As Worksheet, topRow As Long, leftCol As Long, bottomRow As Long, rightCol As Long)
    With ws.Range(ws.Cells(topRow, leftCol), ws.Cells(bottomRow, rightCol))
        .Columns.AutoFit
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
    End With
End Sub