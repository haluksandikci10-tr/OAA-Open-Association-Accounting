VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLedgerService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA - Open Association Accounting
'
' Class : clsLedgerService
' Layer : Service
'
' Purpose:
'   Provides the Service layer for ledger operations.
'
' Responsibilities:
'   - Manage Ledger entities
'   - Coordinate Repository access
'   - Manage LedgerID generation
'   - Calculate current ledger balance
'   - Perform basic ledger CRUD operations
'
'==============================================================================

Private mRepository As clsLedgerRepository


'==============================================================================
' CLASS INITIALIZE
'==============================================================================

Private Sub Class_Initialize()

    Set mRepository = New clsLedgerRepository

End Sub


'==============================================================================
' CLASS TERMINATE
'==============================================================================

Private Sub Class_Terminate()

    Set mRepository = Nothing

End Sub


'==============================================================================
' ADD ENTRY
'==============================================================================

Public Function AddEntry( _
    ByVal Entry As clsLedgerEntry) As Boolean

    On Error GoTo ErrorHandler

    AddEntry = False

    If Entry Is Nothing Then Exit Function

    If Entry.LedgerID <= 0 Then

        Entry.LedgerID = mRepository.GetNextLedgerID

        If Entry.LedgerID <= 0 Then Exit Function

    End If

    AddEntry = mRepository.Add(Entry)

    Exit Function

ErrorHandler:

    AddEntry = False

End Function


'==============================================================================
' GET ENTRY
'==============================================================================

Public Function GetEntry( _
    ByVal LedgerID As Long) As clsLedgerEntry

    On Error GoTo ErrorHandler

    Set GetEntry = Nothing

    If LedgerID <= 0 Then Exit Function

    Set GetEntry = mRepository.GetByID(LedgerID)

    Exit Function

ErrorHandler:

    Set GetEntry = Nothing

End Function


'==============================================================================
' GET ALL ENTRIES
'==============================================================================

Public Function GetAllEntries() As Collection

    On Error GoTo ErrorHandler

    Set GetAllEntries = mRepository.GetAll

    Exit Function

ErrorHandler:

    Set GetAllEntries = Nothing

End Function


'==============================================================================
' UPDATE ENTRY
'==============================================================================

Public Function UpdateEntry( _
    ByVal Entry As clsLedgerEntry) As Boolean

    Dim ExistingEntry As clsLedgerEntry

    On Error GoTo ErrorHandler

    UpdateEntry = False

    If Entry Is Nothing Then Exit Function
    If Entry.LedgerID <= 0 Then Exit Function

    Set ExistingEntry = mRepository.GetByID(Entry.LedgerID)

    If ExistingEntry Is Nothing Then Exit Function

    UpdateEntry = mRepository.UpdateByID(Entry)

    Set ExistingEntry = Nothing

    Exit Function

ErrorHandler:

    UpdateEntry = False

    Set ExistingEntry = Nothing

End Function


'==============================================================================
' DELETE ENTRY
'==============================================================================

Public Function DeleteEntry( _
    ByVal LedgerID As Long) As Boolean

    On Error GoTo ErrorHandler

    DeleteEntry = False

    If LedgerID <= 0 Then Exit Function

    DeleteEntry = mRepository.DeleteByID(LedgerID)

    Exit Function

ErrorHandler:

    DeleteEntry = False

End Function


'==============================================================================
' RECALCULATE BALANCES
'
' Recalculates BalanceAmount in the current repository order.
' The existing record order is preserved.
'==============================================================================

Public Function RecalculateBalances() As Boolean

    Dim Entries As Collection
    Dim Entry As clsLedgerEntry
    Dim RunningBalance As Currency
    Dim I As Long

    On Error GoTo ErrorHandler

    RecalculateBalances = False
    RunningBalance = 0@

    Set Entries = mRepository.GetAll

    If Entries Is Nothing Then Exit Function

    For I = 1 To Entries.Count

        Set Entry = Entries(I)

        If Not Entry Is Nothing Then

            RunningBalance = _
                RunningBalance + _
                Entry.IncomeAmount - _
                Entry.ExpenseAmount

            Entry.BalanceAmount = RunningBalance

            If Not mRepository.UpdateByID(Entry) Then
                Exit Function
            End If

        End If

    Next I

    RecalculateBalances = True

    Exit Function

ErrorHandler:

    RecalculateBalances = False

End Function


'==============================================================================
' GET CURRENT BALANCE
'
' Calculates total income minus total expense.
' Does not rely on stored BalanceAmount values.
'==============================================================================

Public Function GetCurrentBalance() As Currency

    Dim Entries As Collection
    Dim Entry As clsLedgerEntry

    Dim TotalIncome As Currency
    Dim TotalExpense As Currency

    Dim CurrentBalance As Currency
    Dim I As Long

    On Error GoTo ErrorHandler

    GetCurrentBalance = 0@

    TotalIncome = 0@
    TotalExpense = 0@

    Set Entries = mRepository.GetAll

    If Entries Is Nothing Then Exit Function

    For I = 1 To Entries.Count

        Set Entry = Entries(I)

        If Not Entry Is Nothing Then

            If Entry.IncomeAmount > 0 Then

                TotalIncome = _
                    TotalIncome + Entry.IncomeAmount

            End If

            If Entry.ExpenseAmount > 0 Then

                TotalExpense = _
                    TotalExpense + Entry.ExpenseAmount

            End If

        End If

    Next I

    CurrentBalance = TotalIncome - TotalExpense

    GetCurrentBalance = CurrentBalance

    Exit Function

ErrorHandler:

    GetCurrentBalance = 0@

End Function

