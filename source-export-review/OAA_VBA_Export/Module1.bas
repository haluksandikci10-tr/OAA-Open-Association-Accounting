Attribute VB_Name = "Module1"
Option Explicit

'==============================================================================
' OAA v2 - clsExpenseService
' Service layer for expense operations.
'==============================================================================

Private mRepository As clsExpenseRepository
Private mLedgerService As clsLedgerService

Private Sub Class_Initialize()
    Set mRepository = New clsExpenseRepository
    Set mLedgerService = New clsLedgerService
End Sub

Private Sub Class_Terminate()
    Set mLedgerService = Nothing
    Set mRepository = Nothing
End Sub

'------------------------------------------------------------------------------
' ADD EXPENSE
' Creates an expense record without a linked ledger entry.
'------------------------------------------------------------------------------

Public Function AddExpense(ByVal Expense As clsExpense) As Long
    On Error GoTo ErrorHandler

    AddExpense = 0

    If Expense Is Nothing Then Exit Function
    If Not ValidateExpense(Expense) Then Exit Function

    AddExpense = mRepository.Add(Expense)
    Exit Function

ErrorHandler:
    AddExpense = 0
End Function

'------------------------------------------------------------------------------
' ADD EXPENSE WITH LEDGER
' Creates the expense and its linked ledger entry.
'------------------------------------------------------------------------------

Public Function AddExpenseWithLedger(ByVal Expense As clsExpense) As Boolean
    Dim ExpenseID As Long
    Dim LedgerEntry As clsLedgerEntry
    Dim LedgerAdded As Boolean

    On Error GoTo ErrorHandler

    AddExpenseWithLedger = False
    ExpenseID = 0
    LedgerAdded = False

    If Expense Is Nothing Then Exit Function
    If Not ValidateExpense(Expense) Then Exit Function

    ExpenseID = mRepository.Add(Expense)
    If ExpenseID <= 0 Then Exit Function

    Set LedgerEntry = New clsLedgerEntry

    LedgerEntry.ExpenseID = ExpenseID
    LedgerEntry.EntryDate = Expense.EntryDate
    LedgerEntry.DocumentNo = Expense.DocumentNo
    LedgerEntry.ReceiptInvoice = Expense.DocumentNo
    LedgerEntry.Description = Expense.Description
    LedgerEntry.IncomeType = vbNullString
    LedgerEntry.ExpenseType = Expense.ExpenseType
    LedgerEntry.IncomeAmount = CCur(0)
    LedgerEntry.ExpenseAmount = Expense.ExpenseAmount
    LedgerEntry.CashAmount = Expense.CashAmount
    LedgerEntry.BankAmount = Expense.BankAmount
    LedgerEntry.BalanceAmount = _
        mLedgerService.GetCurrentBalance - Expense.ExpenseAmount

    LedgerAdded = mLedgerService.AddEntry(LedgerEntry)

    If Not LedgerAdded Then
        On Error Resume Next
        Call mRepository.DeleteByID(ExpenseID)
        On Error GoTo ErrorHandler
        GoTo CleanExit
    End If

    If Not mLedgerService.RecalculateBalances() Then
        ' Keep the records rather than deleting one side and leaving
        ' an inconsistent relationship; report failure to the caller.
        AddExpenseWithLedger = False
        GoTo CleanExit
    End If

    AddExpenseWithLedger = True

CleanExit:
    Set LedgerEntry = Nothing
    Exit Function

ErrorHandler:
    AddExpenseWithLedger = False

    On Error Resume Next
    If ExpenseID > 0 And Not LedgerAdded Then
        Call mRepository.DeleteByID(ExpenseID)
    End If
    On Error GoTo 0

    Set LedgerEntry = Nothing
End Function

'------------------------------------------------------------------------------
' GET EXPENSE
'------------------------------------------------------------------------------

Public Function GetExpense(ByVal ExpenseID As Long) As clsExpense
    On Error GoTo ErrorHandler

    Set GetExpense = Nothing
    If ExpenseID <= 0 Then Exit Function

    Set GetExpense = mRepository.GetByID(ExpenseID)
    Exit Function

ErrorHandler:
    Set GetExpense = Nothing
End Function

'------------------------------------------------------------------------------
' GET ALL EXPENSES
'------------------------------------------------------------------------------

Public Function GetAllExpenses() As Collection
    On Error GoTo ErrorHandler

    Set GetAllExpenses = mRepository.GetAll
    Exit Function

ErrorHandler:
    Set GetAllExpenses = Nothing
End Function

'------------------------------------------------------------------------------
' UPDATE EXPENSE
' Updates the expense and its linked ledger entry, if present.
'------------------------------------------------------------------------------

Public Function UpdateExpense(ByVal Expense As clsExpense) As Boolean
    Dim ExistingExpense As clsExpense
    Dim LedgerEntry As clsLedgerEntry
    Dim OriginalLedgerEntry As clsLedgerEntry
    Dim Entries As Collection
    Dim I As Long
    Dim FoundLedger As Boolean
    Dim ExpenseUpdated As Boolean
    Dim LedgerUpdated As Boolean

    On Error GoTo ErrorHandler

    UpdateExpense = False
    FoundLedger = False
    ExpenseUpdated = False
    LedgerUpdated = False

    If Expense Is Nothing Then Exit Function
    If Expense.ExpenseID <= 0 Then Exit Function
    If Not ValidateExpense(Expense) Then Exit Function

    Set ExistingExpense = mRepository.GetByID(Expense.ExpenseID)
    If ExistingExpense Is Nothing Then Exit Function

    Set Entries = mLedgerService.GetAllEntries
    If Entries Is Nothing Then Exit Function

    For I = 1 To Entries.Count
        Set LedgerEntry = Entries(I)
        If Not LedgerEntry Is Nothing Then
            If LedgerEntry.ExpenseID = Expense.ExpenseID Then
                FoundLedger = True
                Exit For
            End If
        End If
    Next I

    If FoundLedger Then
        Set OriginalLedgerEntry = New clsLedgerEntry
        OriginalLedgerEntry.LedgerID = LedgerEntry.LedgerID
        OriginalLedgerEntry.ExpenseID = LedgerEntry.ExpenseID
        OriginalLedgerEntry.EntryDate = LedgerEntry.EntryDate
        OriginalLedgerEntry.DocumentNo = LedgerEntry.DocumentNo
        OriginalLedgerEntry.ReceiptInvoice = LedgerEntry.ReceiptInvoice
        OriginalLedgerEntry.Description = LedgerEntry.Description
        OriginalLedgerEntry.IncomeType = LedgerEntry.IncomeType
        OriginalLedgerEntry.ExpenseType = LedgerEntry.ExpenseType
        OriginalLedgerEntry.IncomeAmount = LedgerEntry.IncomeAmount
        OriginalLedgerEntry.ExpenseAmount = LedgerEntry.ExpenseAmount
        OriginalLedgerEntry.CashAmount = LedgerEntry.CashAmount
        OriginalLedgerEntry.BankAmount = LedgerEntry.BankAmount
        OriginalLedgerEntry.BalanceAmount = LedgerEntry.BalanceAmount
    End If

    ExpenseUpdated = mRepository.UpdateByID(Expense)
    If Not ExpenseUpdated Then GoTo CleanExit

    If Not FoundLedger Then
        UpdateExpense = True
        GoTo CleanExit
    End If

    LedgerEntry.EntryDate = Expense.EntryDate
    LedgerEntry.DocumentNo = Expense.DocumentNo
    LedgerEntry.ReceiptInvoice = Expense.DocumentNo
    LedgerEntry.Description = Expense.Description
    LedgerEntry.IncomeType = vbNullString
    LedgerEntry.ExpenseType = Expense.ExpenseType
    LedgerEntry.IncomeAmount = CCur(0)
    LedgerEntry.ExpenseAmount = Expense.ExpenseAmount
    LedgerEntry.CashAmount = Expense.CashAmount
    LedgerEntry.BankAmount = Expense.BankAmount
    LedgerEntry.ExpenseID = Expense.ExpenseID

    If Not mLedgerService.UpdateEntry(LedgerEntry) Then GoTo RollbackExpense
    LedgerUpdated = True

    If Not mLedgerService.RecalculateBalances() Then GoTo RollbackExpense

    UpdateExpense = True
    GoTo CleanExit

RollbackExpense:
    On Error Resume Next
    If LedgerUpdated And Not OriginalLedgerEntry Is Nothing Then
        Call mLedgerService.UpdateEntry(OriginalLedgerEntry)
    End If
    If Not ExistingExpense Is Nothing Then
        Call mRepository.UpdateByID(ExistingExpense)
    End If
    Call mLedgerService.RecalculateBalances
    On Error GoTo ErrorHandler
    UpdateExpense = False
    GoTo CleanExit

CleanExit:
    Set Entries = Nothing
    Set OriginalLedgerEntry = Nothing
    Set LedgerEntry = Nothing
    Set ExistingExpense = Nothing
    Exit Function

ErrorHandler:
    UpdateExpense = False
    On Error Resume Next
    If LedgerUpdated And Not OriginalLedgerEntry Is Nothing Then
        Call mLedgerService.UpdateEntry(OriginalLedgerEntry)
    End If
    If ExpenseUpdated And Not ExistingExpense Is Nothing Then
        Call mRepository.UpdateByID(ExistingExpense)
    End If
    Call mLedgerService.RecalculateBalances
    On Error GoTo 0
    Set Entries = Nothing
    Set OriginalLedgerEntry = Nothing
    Set LedgerEntry = Nothing
    Set ExistingExpense = Nothing
End Function

'------------------------------------------------------------------------------
' DELETE EXPENSE
' Deletes the linked ledger entry and expense record.
'------------------------------------------------------------------------------

Public Function DeleteExpense(ByVal ExpenseID As Long) As Boolean
    Dim Expense As clsExpense
    Dim Entries As Collection
    Dim Entry As clsLedgerEntry
    Dim OriginalLedgerEntry As clsLedgerEntry
    Dim I As Long
    Dim LedgerFound As Boolean
    Dim LedgerDeleted As Boolean
    Dim ExpenseDeleted As Boolean
    Dim RestoreSucceeded As Boolean

    On Error GoTo ErrorHandler

    DeleteExpense = False
    LedgerFound = False
    LedgerDeleted = False
    ExpenseDeleted = False
    RestoreSucceeded = False

    If ExpenseID <= 0 Then Exit Function

    Set Expense = mRepository.GetByID(ExpenseID)
    If Expense Is Nothing Then Exit Function

    Set Entries = mLedgerService.GetAllEntries
    If Entries Is Nothing Then Exit Function

    For I = 1 To Entries.Count
        Set Entry = Entries(I)
        If Not Entry Is Nothing Then
            If Entry.ExpenseID = ExpenseID Then
                LedgerFound = True
                Set OriginalLedgerEntry = New clsLedgerEntry
                OriginalLedgerEntry.LedgerID = Entry.LedgerID
                OriginalLedgerEntry.ExpenseID = Entry.ExpenseID
                OriginalLedgerEntry.EntryDate = Entry.EntryDate
                OriginalLedgerEntry.DocumentNo = Entry.DocumentNo
                OriginalLedgerEntry.ReceiptInvoice = Entry.ReceiptInvoice
                OriginalLedgerEntry.Description = Entry.Description
                OriginalLedgerEntry.IncomeType = Entry.IncomeType
                OriginalLedgerEntry.ExpenseType = Entry.ExpenseType
                OriginalLedgerEntry.IncomeAmount = Entry.IncomeAmount
                OriginalLedgerEntry.ExpenseAmount = Entry.ExpenseAmount
                OriginalLedgerEntry.CashAmount = Entry.CashAmount
                OriginalLedgerEntry.BankAmount = Entry.BankAmount
                OriginalLedgerEntry.BalanceAmount = Entry.BalanceAmount
                Exit For
            End If
        End If
    Next I

    If LedgerFound Then
        If OriginalLedgerEntry.LedgerID <= 0 Then Exit Function
        LedgerDeleted = mLedgerService.DeleteEntry(OriginalLedgerEntry.LedgerID)
        If Not LedgerDeleted Then Exit Function
    End If

    ExpenseDeleted = mRepository.DeleteByID(ExpenseID)

    If Not ExpenseDeleted Then
        If LedgerDeleted Then
            On Error Resume Next
            RestoreSucceeded = mLedgerService.AddEntry(OriginalLedgerEntry)
            On Error GoTo ErrorHandler
        End If
        GoTo CleanExit
    End If

    If Not mLedgerService.RecalculateBalances() Then
        ' Best-effort rollback if recalculation fails.
        On Error Resume Next
        If LedgerDeleted And Not OriginalLedgerEntry Is Nothing Then
            RestoreSucceeded = mLedgerService.AddEntry(OriginalLedgerEntry)
        End If
        If ExpenseDeleted Then
            Call mRepository.Add(Expense)
        End If
        Call mLedgerService.RecalculateBalances
        On Error GoTo ErrorHandler
        GoTo CleanExit
    End If

    DeleteExpense = True

CleanExit:
    Set OriginalLedgerEntry = Nothing
    Set Entry = Nothing
    Set Entries = Nothing
    Set Expense = Nothing
    Exit Function

ErrorHandler:
    DeleteExpense = False
    On Error Resume Next
    If LedgerDeleted And Not ExpenseDeleted Then
        If Not OriginalLedgerEntry Is Nothing Then
            Call mLedgerService.AddEntry(OriginalLedgerEntry)
        End If
    End If
    Call mLedgerService.RecalculateBalances
    On Error GoTo 0
    Set OriginalLedgerEntry = Nothing
    Set Entry = Nothing
    Set Entries = Nothing
    Set Expense = Nothing
End Function

'------------------------------------------------------------------------------
' VALIDATE EXPENSE
'------------------------------------------------------------------------------

Private Function ValidateExpense(ByVal Expense As clsExpense) As Boolean
    On Error GoTo ErrorHandler

    ValidateExpense = False

    If Expense Is Nothing Then Exit Function
    If Expense.EntryDate <= 0 Then Exit Function
    If Trim$(Expense.ExpenseType) = vbNullString Then Exit Function
    If Expense.ExpenseAmount <= 0 Then Exit Function
    If Expense.CashAmount < 0 Then Exit Function
    If Expense.BankAmount < 0 Then Exit Function

    If Expense.CashAmount + Expense.BankAmount <> Expense.ExpenseAmount Then
        Exit Function
    End If

    ValidateExpense = True
    Exit Function

ErrorHandler:
    ValidateExpense = False
End Function
