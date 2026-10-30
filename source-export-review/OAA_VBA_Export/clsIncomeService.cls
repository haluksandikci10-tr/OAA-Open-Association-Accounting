VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIncomeService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'==============================================================================
' OAA v2
' Service : clsIncomeService
'
' Sorumluluklar:
'   - Gelir is kurallarini dogrulamak
'   - Gelir Repository islemlerini yonetmek
'   - Gelir hareketlerini Isletme Defterine baglamak
'   - Gelir ve defter islemlerinde veri tutarliligini korumak
'
' Repository yalnizca veri erisiminden sorumludur.
' Is kurallari ve servisler arasi koordinasyon bu siniftadir.
'==============================================================================

Private mRepository As clsIncomeRepository
Private mLedgerService As clsLedgerService

'==============================================================================
' CLASS INITIALIZE / TERMINATE
'==============================================================================

Private Sub Class_Initialize()

    Set mRepository = New clsIncomeRepository
    Set mLedgerService = New clsLedgerService

End Sub

Private Sub Class_Terminate()

    Set mLedgerService = Nothing
    Set mRepository = Nothing

End Sub

'==============================================================================
' ADD INCOME
' Yalnizca TBL_INCOME kaydi olusturur.
' Isletme defteri entegrasyonu icin AddIncomeWithLedger kullanilmalidir.
'==============================================================================

Public Function AddIncome(ByVal Income As clsIncome) As Long

    On Error GoTo ErrorHandler

    AddIncome = 0

    If Income Is Nothing Then Exit Function
    If Not ValidateIncome(Income) Then Exit Function

    AddIncome = mRepository.Add(Income)

    Exit Function

ErrorHandler:

    AddIncome = 0

End Function

'==============================================================================
' ADD INCOME WITH LEDGER
' Gelir ve defter kaydini IncomeID ile iliskili olarak olusturur.
'==============================================================================

Public Function AddIncomeWithLedger(ByVal Income As clsIncome) As Boolean

    Dim NewIncomeID As Long
    Dim LedgerEntry As clsLedgerEntry
    Dim LedgerAdded As Boolean

    On Error GoTo ErrorHandler

    AddIncomeWithLedger = False

    If Income Is Nothing Then Exit Function
    If Not ValidateIncome(Income) Then Exit Function

    NewIncomeID = mRepository.Add(Income)
    If NewIncomeID <= 0 Then Exit Function

    Income.IncomeID = NewIncomeID

    Set LedgerEntry = New clsLedgerEntry

    With LedgerEntry
        .IncomeID = NewIncomeID
        .ExpenseID = 0
        .EntryDate = Income.EntryDate
        .DocumentNo = Income.DocumentNo
        .ReceiptInvoice = Income.DocumentNo
        .Description = Income.Description
        .IncomeType = Income.IncomeType
        .ExpenseType = vbNullString
        .IncomeAmount = Income.IncomeAmount
        .ExpenseAmount = 0@
        .CashAmount = Income.CashAmount
        .BankAmount = Income.BankAmount
        .BalanceAmount = mLedgerService.GetCurrentBalance + Income.IncomeAmount
    End With

    LedgerAdded = mLedgerService.AddEntry(LedgerEntry)
    If Not LedgerAdded Then GoTo RollbackIncome

    If Not mLedgerService.RecalculateBalances Then GoTo RollbackLedgerAndIncome

    AddIncomeWithLedger = True
    GoTo CleanExit

RollbackLedgerAndIncome:

    If LedgerEntry.LedgerID > 0 Then
        Call mLedgerService.DeleteEntry(LedgerEntry.LedgerID)
    End If

RollbackIncome:

    If NewIncomeID > 0 Then
        Call mRepository.DeleteByID(NewIncomeID)
    End If

    Call mLedgerService.RecalculateBalances
    GoTo CleanExit

ErrorHandler:

    On Error Resume Next

    If Not LedgerEntry Is Nothing Then
        If LedgerEntry.LedgerID > 0 Then
            Call mLedgerService.DeleteEntry(LedgerEntry.LedgerID)
        End If
    End If

    If NewIncomeID > 0 Then
        Call mRepository.DeleteByID(NewIncomeID)
    End If

    Call mLedgerService.RecalculateBalances

    On Error GoTo 0
    AddIncomeWithLedger = False

CleanExit:

    Set LedgerEntry = Nothing

End Function

'==============================================================================
' GET INCOME
'==============================================================================

Public Function GetIncome(ByVal IncomeID As Long) As clsIncome

    On Error GoTo ErrorHandler

    Set GetIncome = Nothing
    If IncomeID <= 0 Then Exit Function

    Set GetIncome = mRepository.GetByID(IncomeID)
    Exit Function

ErrorHandler:

    Set GetIncome = Nothing

End Function

'==============================================================================
' GET ALL INCOMES
'==============================================================================

Public Function GetAllIncomes() As Collection

    On Error GoTo ErrorHandler

    Set GetAllIncomes = mRepository.GetAll
    Exit Function

ErrorHandler:

    Set GetAllIncomes = Nothing

End Function

'==============================================================================
' GET NEXT INCOME ID
'==============================================================================

Public Function GetNextIncomeID() As Long

    On Error GoTo ErrorHandler

    GetNextIncomeID = mRepository.GetNextIncomeID
    Exit Function

ErrorHandler:

    GetNextIncomeID = 0

End Function

'==============================================================================
' UPDATE INCOME
' Gelir kaydi ve IncomeID ile bagli defter hareketi birlikte guncellenir.
'==============================================================================

Public Function UpdateIncome(ByVal Income As clsIncome) As Boolean

    Dim OldIncome As clsIncome
    Dim OldLedgerEntry As clsLedgerEntry
    Dim NewLedgerEntry As clsLedgerEntry
    Dim LedgerMatchCount As Long
    Dim UpdateSucceeded As Boolean

    On Error GoTo ErrorHandler

    UpdateIncome = False

    If Income Is Nothing Then Exit Function
    If Income.IncomeID <= 0 Then Exit Function
    If Not ValidateIncome(Income) Then Exit Function

    Set OldIncome = mRepository.GetByID(Income.IncomeID)
    If OldIncome Is Nothing Then Exit Function

    Set OldLedgerEntry = FindIncomeLedgerEntry(Income.IncomeID, LedgerMatchCount)

    ' Belirsiz veya eksik iliski uzerinde guncelleme yapilmaz.
    If LedgerMatchCount <> 1 Then GoTo CleanExit
    If OldLedgerEntry Is Nothing Then GoTo CleanExit

    Set NewLedgerEntry = New clsLedgerEntry

    With NewLedgerEntry
        .LedgerID = OldLedgerEntry.LedgerID
        .IncomeID = Income.IncomeID
        .ExpenseID = 0
        .EntryDate = Income.EntryDate
        .DocumentNo = Income.DocumentNo
        .ReceiptInvoice = Income.DocumentNo
        .Description = Income.Description
        .IncomeType = Income.IncomeType
        .ExpenseType = vbNullString
        .IncomeAmount = Income.IncomeAmount
        .ExpenseAmount = 0@
        .CashAmount = Income.CashAmount
        .BankAmount = Income.BankAmount
        .BalanceAmount = OldLedgerEntry.BalanceAmount
    End With

    If Not mRepository.UpdateByID(Income) Then GoTo CleanExit

    If Not mLedgerService.UpdateEntry(NewLedgerEntry) Then
        Call mRepository.UpdateByID(OldIncome)
        GoTo CleanExit
    End If

    If Not mLedgerService.RecalculateBalances Then
        Call mLedgerService.UpdateEntry(OldLedgerEntry)
        Call mRepository.UpdateByID(OldIncome)
        Call mLedgerService.RecalculateBalances
        GoTo CleanExit
    End If

    UpdateSucceeded = True
    UpdateIncome = UpdateSucceeded

CleanExit:

    Set NewLedgerEntry = Nothing
    Set OldLedgerEntry = Nothing
    Set OldIncome = Nothing
    Exit Function

ErrorHandler:

    On Error Resume Next

    If Not OldLedgerEntry Is Nothing Then
        Call mLedgerService.UpdateEntry(OldLedgerEntry)
    End If

    If Not OldIncome Is Nothing Then
        Call mRepository.UpdateByID(OldIncome)
    End If

    Call mLedgerService.RecalculateBalances

    On Error GoTo 0

    UpdateIncome = False
    Resume CleanExit

End Function

'==============================================================================
' DELETE INCOME
' Iliskili tek defter hareketi ve gelir kaydi birlikte silinir.
' Eski/iliskisiz veya birden fazla iliskili hareket varsa silme reddedilir.
'==============================================================================

Public Function DeleteIncome(ByVal IncomeID As Long) As Boolean

    Dim OldIncome As clsIncome
    Dim CurrentIncome As clsIncome
    Dim OldLedgerEntry As clsLedgerEntry
    Dim LedgerMatchCount As Long
    Dim LedgerDeleted As Boolean
    Dim IncomeDeleted As Boolean

    On Error GoTo ErrorHandler

    DeleteIncome = False

    If IncomeID <= 0 Then Exit Function

    Set OldIncome = mRepository.GetByID(IncomeID)
    If OldIncome Is Nothing Then Exit Function

    Set OldLedgerEntry = FindIncomeLedgerEntry(IncomeID, LedgerMatchCount)

    ' Eski baglantisiz kayitlarda otomatik silme yapilmaz.
    If LedgerMatchCount <> 1 Then GoTo CleanExit
    If OldLedgerEntry Is Nothing Then GoTo CleanExit

    LedgerDeleted = mLedgerService.DeleteEntry(OldLedgerEntry.LedgerID)
    If Not LedgerDeleted Then GoTo CleanExit

    IncomeDeleted = mRepository.DeleteByID(IncomeID)

    If Not IncomeDeleted Then
        ' Silinen defter kaydini ayni LedgerID ile geri yuklemeyi dene.
        Call mLedgerService.AddEntry(OldLedgerEntry)
        Call mLedgerService.RecalculateBalances
        GoTo CleanExit
    End If

    If Not mLedgerService.RecalculateBalances Then
        ' Geri alma: gelir kaydini ve defter kaydini eski kimliklerle geri yukle.
        Call mRepository.Add(OldIncome)
        Call mLedgerService.AddEntry(OldLedgerEntry)
        Call mLedgerService.RecalculateBalances
        GoTo CleanExit
    End If

    DeleteIncome = True

CleanExit:

    Set OldLedgerEntry = Nothing
    Set OldIncome = Nothing
    Exit Function

ErrorHandler:

    On Error Resume Next

    If Not OldIncome Is Nothing Then
        Set CurrentIncome = mRepository.GetByID(IncomeID)
        If CurrentIncome Is Nothing Then
            Call mRepository.Add(OldIncome)
        End If
    End If

    If Not OldLedgerEntry Is Nothing Then
        If OldLedgerEntry.LedgerID > 0 Then
            If mLedgerService.GetEntry(OldLedgerEntry.LedgerID) Is Nothing Then
                Call mLedgerService.AddEntry(OldLedgerEntry)
            End If
        End If
    End If

    Call mLedgerService.RecalculateBalances
    On Error GoTo 0

    DeleteIncome = False
    Resume CleanExit

End Function

'==============================================================================
' FIND INCOME LEDGER ENTRY
' IncomeID uzerinden tam olarak bir defter hareketi bulunmalidir.
'==============================================================================

Private Function FindIncomeLedgerEntry( _
    ByVal IncomeID As Long, _
    ByRef MatchCount As Long) As clsLedgerEntry

    Dim Entries As Collection
    Dim Entry As clsLedgerEntry
    Dim FoundEntry As clsLedgerEntry
    Dim I As Long

    On Error GoTo ErrorHandler

    MatchCount = 0
    Set FindIncomeLedgerEntry = Nothing

    If IncomeID <= 0 Then Exit Function

    Set Entries = mLedgerService.GetAllEntries
    If Entries Is Nothing Then Exit Function

    For I = 1 To Entries.Count

        Set Entry = Entries(I)

        If Not Entry Is Nothing Then
            If Entry.IncomeID = IncomeID Then
                MatchCount = MatchCount + 1
                Set FoundEntry = Entry
            End If
        End If

    Next I

    If MatchCount = 1 Then
        Set FindIncomeLedgerEntry = FoundEntry
    End If

    Set FoundEntry = Nothing
    Set Entry = Nothing
    Set Entries = Nothing
    Exit Function

ErrorHandler:

    MatchCount = 0
    Set FindIncomeLedgerEntry = Nothing
    Set FoundEntry = Nothing
    Set Entry = Nothing
    Set Entries = Nothing

End Function

'==============================================================================
' VALIDATE INCOME
'==============================================================================

Private Function ValidateIncome(ByVal Income As clsIncome) As Boolean

    On Error GoTo ErrorHandler

    ValidateIncome = False

    If Income Is Nothing Then Exit Function
    If Income.EntryDate <= 0 Then Exit Function
    If Trim$(Income.IncomeType) = vbNullString Then Exit Function
    If Income.IncomeAmount <= 0 Then Exit Function
    If Income.CashAmount < 0 Then Exit Function
    If Income.BankAmount < 0 Then Exit Function

    If Income.CashAmount + Income.BankAmount <> Income.IncomeAmount Then
        Exit Function
    End If

    ValidateIncome = True
    Exit Function

ErrorHandler:

    ValidateIncome = False

End Function


