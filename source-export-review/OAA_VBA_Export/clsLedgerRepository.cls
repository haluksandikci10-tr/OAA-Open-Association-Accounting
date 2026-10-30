VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLedgerRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'==============================================================================
' OAA – Open Association Accounting
'
' Class : clsLedgerRepository
' Layer : Repository
' Purpose:
'   TBL_LEDGER tablosu üzerinde Ýþletme Defteri veri eriþimini gerçekleþtirir.
'
' Sorumluluk:
'   - Kayýt ekleme
'   - Kayýt bulma
'   - Tüm kayýtlarý getirme
'   - Sonraki LedgerID üretme
'   - Kayýt silme
'
' Repository iþ kuralý içermez.
' Ýþ kurallarý Service katmanýnda uygulanacaktýr.
'==============================================================================


Private Const SHEET_NAME As String = SHEET_LEDGER
Private Const TABLE_NAME As String = TABLE_LEDGER


'==============================================================================
' ADD
'==============================================================================

Public Function Add(ByVal Entry As clsLedgerEntry) As Boolean

    On Error GoTo ErrHandler

    Dim Table As ListObject
    Dim NewRow As ListRow

    If Entry Is Nothing Then
        Add = False
        Exit Function
    End If

    Set Table = GetTable(SHEET_NAME, TABLE_NAME)

    Set NewRow = Table.ListRows.Add

    With NewRow.Range

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_ID)).Value = _
            Entry.LedgerID

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_EXPENSE_ID)).Value = _
            Entry.ExpenseID

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_INCOME_ID)).Value = _
            Entry.IncomeID

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_ENTRY_DATE)).Value = _
            Entry.EntryDate

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_DOCUMENT_NO)).Value = _
            Entry.DocumentNo

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_RECEIPT_INVOICE)).Value = _
            Entry.ReceiptInvoice

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_DESCRIPTION)).Value = _
            Entry.Description

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_INCOME_TYPE)).Value = _
            Entry.IncomeType

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_EXPENSE_TYPE)).Value = _
            Entry.ExpenseType

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_INCOME_AMOUNT)).Value = _
            Entry.IncomeAmount

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_EXPENSE_AMOUNT)).Value = _
            Entry.ExpenseAmount

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_CASH_AMOUNT)).Value = _
            Entry.CashAmount

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_BANK_AMOUNT)).Value = _
            Entry.BankAmount

        .Cells(1, GetColumnIndex(Table, COL_LEDGER_BALANCE_AMOUNT)).Value = _
            Entry.BalanceAmount

    End With

    Add = True

    Exit Function

ErrHandler:

    Add = False

    Err.Raise Err.Number, _
              "clsLedgerRepository.Add", _
              Err.Description

End Function


'==============================================================================
' GET BY ID
'==============================================================================

Public Function GetByID(ByVal LedgerID As Long) As clsLedgerEntry

    On Error GoTo ErrHandler

    Dim Table As ListObject
    Dim RowIndex As Long
    Dim Entry As clsLedgerEntry

    Set Table = GetTable(SHEET_NAME, TABLE_NAME)

    RowIndex = FindLedgerRow(Table, LedgerID)

    If RowIndex = 0 Then
        Set GetByID = Nothing
        Exit Function
    End If

    Set Entry = CreateLedgerEntry(Table, RowIndex)

    Set GetByID = Entry

    Exit Function

ErrHandler:

    Set GetByID = Nothing

End Function


'==============================================================================
' GET ALL
'==============================================================================

Public Function GetAll() As Collection

    On Error GoTo ErrHandler

    Dim Result As New Collection
    Dim Table As ListObject
    Dim RowIndex As Long
    Dim Entry As clsLedgerEntry

    Set Table = GetTable(SHEET_NAME, TABLE_NAME)

    If Table.DataBodyRange Is Nothing Then
        Set GetAll = Result
        Exit Function
    End If

    For RowIndex = 1 To Table.ListRows.Count

        Set Entry = CreateLedgerEntry(Table, RowIndex)

        If Not Entry Is Nothing Then
            Result.Add Entry
        End If

    Next RowIndex

    Set GetAll = Result

    Exit Function

ErrHandler:

    Set GetAll = Nothing

End Function


'==============================================================================
' GET NEXT LEDGER ID
'==============================================================================

Public Function GetNextLedgerID() As Long

    On Error GoTo ErrHandler

    Dim Table As ListObject
    Dim RowIndex As Long
    Dim CurrentID As Long
    Dim MaxID As Long

    MaxID = 0

    Set Table = GetTable(SHEET_NAME, TABLE_NAME)

    If Table.DataBodyRange Is Nothing Then
        GetNextLedgerID = 1
        Exit Function
    End If

    For RowIndex = 1 To Table.ListRows.Count

        If IsNumeric( _
            Table.DataBodyRange.Cells( _
                RowIndex, _
                GetColumnIndex(Table, COL_LEDGER_ID) _
            ).Value _
        ) Then

            CurrentID = CLng( _
                Table.DataBodyRange.Cells( _
                    RowIndex, _
                    GetColumnIndex(Table, COL_LEDGER_ID) _
                ).Value _
            )

            If CurrentID > MaxID Then
                MaxID = CurrentID
            End If

        End If

    Next RowIndex

    GetNextLedgerID = MaxID + 1

    Exit Function

ErrHandler:

    GetNextLedgerID = 0

End Function


'==============================================================================
' UPDATE BY ID
' Mevcut Isletme Defteri kaydini LedgerID uzerinden gunceller.
'==============================================================================

Public Function UpdateByID(ByVal Entry As clsLedgerEntry) As Boolean

    On Error GoTo ErrHandler

    Dim Table As ListObject
    Dim RowIndex As Long
    Dim OriginalErrorNumber As Long
    Dim OriginalErrorDescription As String

    UpdateByID = False

    If Entry Is Nothing Then Exit Function
    If Entry.LedgerID <= 0 Then Exit Function

    Set Table = GetTable(SHEET_NAME, TABLE_NAME)
    RowIndex = FindLedgerRow(Table, Entry.LedgerID)

    If RowIndex = 0 Then Exit Function

    With Table.ListRows(RowIndex).Range
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_EXPENSE_ID)).Value = Entry.ExpenseID
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_INCOME_ID)).Value = Entry.IncomeID
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_ENTRY_DATE)).Value = Entry.EntryDate
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_DOCUMENT_NO)).Value = Entry.DocumentNo
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_RECEIPT_INVOICE)).Value = Entry.ReceiptInvoice
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_DESCRIPTION)).Value = Entry.Description
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_INCOME_TYPE)).Value = Entry.IncomeType
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_EXPENSE_TYPE)).Value = Entry.ExpenseType
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_INCOME_AMOUNT)).Value = Entry.IncomeAmount
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_EXPENSE_AMOUNT)).Value = Entry.ExpenseAmount
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_CASH_AMOUNT)).Value = Entry.CashAmount
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_BANK_AMOUNT)).Value = Entry.BankAmount
        .Cells(1, GetColumnIndex(Table, COL_LEDGER_BALANCE_AMOUNT)).Value = Entry.BalanceAmount
    End With

    UpdateByID = True
    Exit Function

ErrHandler:
    OriginalErrorNumber = Err.Number
    OriginalErrorDescription = Err.Description
    Err.Raise OriginalErrorNumber, "clsLedgerRepository.UpdateByID", OriginalErrorDescription

End Function


'==============================================================================
' DELETE BY ID
'==============================================================================

Public Function DeleteByID(ByVal LedgerID As Long) As Boolean

    On Error GoTo ErrHandler

    Dim Table As ListObject
    Dim RowIndex As Long

    Set Table = GetTable(SHEET_NAME, TABLE_NAME)

    RowIndex = FindLedgerRow(Table, LedgerID)

    If RowIndex = 0 Then
        DeleteByID = False
        Exit Function
    End If

    Table.ListRows(RowIndex).Delete

    DeleteByID = True

    Exit Function

ErrHandler:

    DeleteByID = False

    Err.Raise Err.Number, _
              "clsLedgerRepository.DeleteByID", _
              Err.Description

End Function


'==============================================================================
' FIND LEDGER ROW
'==============================================================================

Private Function FindLedgerRow( _
            ByVal Table As ListObject, _
            ByVal LedgerID As Long) As Long

    On Error GoTo ErrHandler

    Dim RowIndex As Long
    Dim Value As Variant

    FindLedgerRow = 0

    If Table.DataBodyRange Is Nothing Then
        Exit Function
    End If

    For RowIndex = 1 To Table.ListRows.Count

        Value = Table.DataBodyRange.Cells( _
                    RowIndex, _
                    GetColumnIndex(Table, COL_LEDGER_ID) _
                ).Value

        If IsNumeric(Value) Then

            If CLng(Value) = LedgerID Then
                FindLedgerRow = RowIndex
                Exit Function
            End If

        End If

    Next RowIndex

    Exit Function

ErrHandler:

    FindLedgerRow = 0

End Function


'==============================================================================
' CREATE ENTITY FROM TABLE ROW
'==============================================================================

Private Function CreateLedgerEntry( _
            ByVal Table As ListObject, _
            ByVal RowIndex As Long) As clsLedgerEntry

    On Error GoTo ErrHandler

    Dim Entry As clsLedgerEntry

    Set Entry = New clsLedgerEntry

    With Table.DataBodyRange

        If IsNumeric(.Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_ID)).Value) Then

            Entry.LedgerID = CLng( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_ID)).Value _
            )

        End If

        If IsNumeric(.Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_EXPENSE_ID)).Value) Then

            Entry.ExpenseID = CLng( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_EXPENSE_ID)).Value _
            )

        End If

        If IsNumeric(.Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_INCOME_ID)).Value) Then

            Entry.IncomeID = CLng( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_INCOME_ID)).Value _
            )

        End If

        If IsDate(.Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_ENTRY_DATE)).Value) Then

            Entry.EntryDate = CDate( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_ENTRY_DATE)).Value _
            )

        End If

        Entry.DocumentNo = CStr( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_DOCUMENT_NO)).Value _
        )

        Entry.ReceiptInvoice = CStr( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_RECEIPT_INVOICE)).Value _
        )

        Entry.Description = CStr( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_DESCRIPTION)).Value _
        )

        Entry.IncomeType = CStr( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_INCOME_TYPE)).Value _
        )

        Entry.ExpenseType = CStr( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_EXPENSE_TYPE)).Value _
        )

        If IsNumeric(.Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_INCOME_AMOUNT)).Value) Then

            Entry.IncomeAmount = CCur( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_INCOME_AMOUNT)).Value _
            )

        End If

        If IsNumeric(.Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_EXPENSE_AMOUNT)).Value) Then

            Entry.ExpenseAmount = CCur( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_EXPENSE_AMOUNT)).Value _
            )

        End If

        If IsNumeric(.Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_CASH_AMOUNT)).Value) Then

            Entry.CashAmount = CCur( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_CASH_AMOUNT)).Value _
            )

        End If

        If IsNumeric(.Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_BANK_AMOUNT)).Value) Then

            Entry.BankAmount = CCur( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_BANK_AMOUNT)).Value _
            )

        End If

        If IsNumeric(.Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_BALANCE_AMOUNT)).Value) Then

            Entry.BalanceAmount = CCur( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_LEDGER_BALANCE_AMOUNT)).Value _
            )

        End If

    End With

    Set CreateLedgerEntry = Entry

    Exit Function

ErrHandler:

    Set CreateLedgerEntry = Nothing

End Function


