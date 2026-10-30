VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExpenseRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Class : clsExpenseRepository
' Purpose: Gider verilerinin Excel tablosu uzerinden CRUD ilemleri
'
' Sheet  : Gider
' Table  : TBL_EXPENSE
'
' Sorumluluk:
'   - Gider kayytlaryna veri eriimi
'   - CRUD ilemleri
'   - ID uretimi
'
' Repository i kuraly uygulamaz.
'==============================================================================



'==============================================================================
' ADD
'==============================================================================

Public Function Add(ByVal Expense As clsExpense) As Long

    On Error GoTo ErrHandler

    Dim Table As ListObject
    Dim NewRow As ListRow

    Add = 0

    If Expense Is Nothing Then Exit Function

    Set Table = GetTable(SHEET_EXPENSE, TABLE_EXPENSE)

    '----------------------------------------------------------------------
    ' ExpenseID verilmemise yeni ID olutur.
    '----------------------------------------------------------------------

    If Expense.ExpenseID <= 0 Then

        Expense.ExpenseID = GetNextExpenseID

        If Expense.ExpenseID <= 0 Then
            Expense.ExpenseID = 0
            Exit Function
        End If

    End If

    '----------------------------------------------------------------------
    ' Yeni satyr
    '----------------------------------------------------------------------

    Set NewRow = Table.ListRows.Add

    With NewRow.Range

        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_ID)).Value = _
            Expense.ExpenseID

        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_ENTRY_DATE)).Value = _
            Expense.EntryDate

        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_DOCUMENT_NO)).Value = _
            Expense.DocumentNo

        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_TYPE)).Value = _
            Expense.ExpenseType

        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_DESCRIPTION)).Value = _
            Expense.Description

        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_AMOUNT)).Value = _
            Expense.ExpenseAmount

        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_CASH_AMOUNT)).Value = _
            Expense.CashAmount

        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_BANK_AMOUNT)).Value = _
            Expense.BankAmount

    End With

    Add = Expense.ExpenseID

    Exit Function


ErrHandler:

    Add = 0

    Err.Raise Err.Number, _
              "clsExpenseRepository.Add", _
              Err.Description

End Function


'==============================================================================
' GET BY ID
'==============================================================================

Public Function GetByID(ByVal ExpenseID As Long) As clsExpense

    On Error GoTo ErrHandler

    Dim Table As ListObject
    Dim RowIndex As Long
    Dim Expense As clsExpense

    Set GetByID = Nothing

    If ExpenseID <= 0 Then Exit Function

    Set Table = GetTable(SHEET_EXPENSE, TABLE_EXPENSE)

    RowIndex = FindExpenseRow(Table, ExpenseID)

    If RowIndex = 0 Then Exit Function

    Set Expense = CreateExpense(Table, RowIndex)

    Set GetByID = Expense

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
    Dim Expense As clsExpense

    Set Table = GetTable(SHEET_EXPENSE, TABLE_EXPENSE)

    If Table.DataBodyRange Is Nothing Then

        Set GetAll = Result
        Exit Function

    End If

    For RowIndex = 1 To Table.ListRows.Count

        Set Expense = CreateExpense(Table, RowIndex)

        If Not Expense Is Nothing Then
            Result.Add Expense
        End If

    Next RowIndex

    Set GetAll = Result

    Exit Function


ErrHandler:

    Set GetAll = Nothing

End Function


'==============================================================================
' GET NEXT EXPENSE ID
'==============================================================================

Public Function GetNextExpenseID() As Long

    On Error GoTo ErrHandler

    Dim Table As ListObject
    Dim RowIndex As Long

    Dim CurrentID As Long
    Dim MaxID As Long

    MaxID = 0

    Set Table = GetTable(SHEET_EXPENSE, TABLE_EXPENSE)

    If Table.DataBodyRange Is Nothing Then

        GetNextExpenseID = 1
        Exit Function

    End If

    For RowIndex = 1 To Table.ListRows.Count

        If IsNumeric( _
            Table.DataBodyRange.Cells( _
                RowIndex, _
                GetColumnIndex(Table, COL_EXPENSE_ID) _
            ).Value _
        ) Then

            CurrentID = CLng( _
                Table.DataBodyRange.Cells( _
                    RowIndex, _
                    GetColumnIndex(Table, COL_EXPENSE_ID) _
                ).Value _
            )

            If CurrentID > MaxID Then
                MaxID = CurrentID
            End If

        End If

    Next RowIndex

    GetNextExpenseID = MaxID + 1

    Exit Function


ErrHandler:

    GetNextExpenseID = 0

End Function




'==============================================================================
' UPDATE BY ID
'==============================================================================

Public Function UpdateByID(ByVal Expense As clsExpense) As Boolean

    On Error GoTo ErrHandler

    Dim Table As ListObject
    Dim RowIndex As Long

    UpdateByID = False

    If Expense Is Nothing Then Exit Function
    If Expense.ExpenseID <= 0 Then Exit Function

    Set Table = GetTable(SHEET_EXPENSE, TABLE_EXPENSE)
    RowIndex = FindExpenseRow(Table, Expense.ExpenseID)

    If RowIndex <= 0 Then Exit Function

    With Table.ListRows(RowIndex).Range
        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_ENTRY_DATE)).Value = Expense.EntryDate
        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_DOCUMENT_NO)).Value = Expense.DocumentNo
        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_TYPE)).Value = Expense.ExpenseType
        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_DESCRIPTION)).Value = Expense.Description
        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_AMOUNT)).Value = Expense.ExpenseAmount
        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_CASH_AMOUNT)).Value = Expense.CashAmount
        .Cells(1, GetColumnIndex(Table, COL_EXPENSE_BANK_AMOUNT)).Value = Expense.BankAmount
    End With

    UpdateByID = True
    Exit Function

ErrHandler:
    UpdateByID = False
    Err.Raise Err.Number, "clsExpenseRepository.UpdateByID", Err.Description

End Function


'==============================================================================
' DELETE BY ID
'==============================================================================

Public Function DeleteByID(ByVal ExpenseID As Long) As Boolean

    On Error GoTo ErrHandler

    Dim Table As ListObject
    Dim RowIndex As Long

    DeleteByID = False

    If ExpenseID <= 0 Then Exit Function

    Set Table = GetTable(SHEET_EXPENSE, TABLE_EXPENSE)

    RowIndex = FindExpenseRow(Table, ExpenseID)

    If RowIndex = 0 Then Exit Function

    Table.ListRows(RowIndex).Delete

    DeleteByID = True

    Exit Function


ErrHandler:

    DeleteByID = False

    Err.Raise Err.Number, _
              "clsExpenseRepository.DeleteByID", _
              Err.Description

End Function


'==============================================================================
' FIND EXPENSE ROW
'==============================================================================

Private Function FindExpenseRow( _
    ByVal Table As ListObject, _
    ByVal ExpenseID As Long _
) As Long

    On Error GoTo ErrHandler

    Dim RowIndex As Long
    Dim Value As Variant

    FindExpenseRow = 0

    If Table.DataBodyRange Is Nothing Then Exit Function

    For RowIndex = 1 To Table.ListRows.Count

        Value = Table.DataBodyRange.Cells( _
                    RowIndex, _
                    GetColumnIndex(Table, COL_EXPENSE_ID) _
                ).Value

        If IsNumeric(Value) Then

            If CLng(Value) = ExpenseID Then

                FindExpenseRow = RowIndex
                Exit Function

            End If

        End If

    Next RowIndex

    Exit Function


ErrHandler:

    FindExpenseRow = 0

End Function


'==============================================================================
' CREATE EXPENSE
'==============================================================================

Private Function CreateExpense( _
    ByVal Table As ListObject, _
    ByVal RowIndex As Long _
) As clsExpense

    On Error GoTo ErrHandler

    Dim Expense As clsExpense

    Set Expense = New clsExpense

    With Table.DataBodyRange

        '------------------------------------------------------------------
        ' ExpenseID
        '------------------------------------------------------------------

        If IsNumeric( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_ID)).Value _
        ) Then

            Expense.ExpenseID = CLng( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_ID)).Value _
            )

        End If

        '------------------------------------------------------------------
        ' EntryDate
        '------------------------------------------------------------------

        If IsDate( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_ENTRY_DATE)).Value _
        ) Then

            Expense.EntryDate = CDate( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_ENTRY_DATE)).Value _
            )

        End If

        '------------------------------------------------------------------
        ' DocumentNo
        '------------------------------------------------------------------

        Expense.DocumentNo = CStr( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_DOCUMENT_NO)).Value _
        )

        '------------------------------------------------------------------
        ' ExpenseType
        '------------------------------------------------------------------

        Expense.ExpenseType = CStr( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_TYPE)).Value _
        )

        '------------------------------------------------------------------
        ' Description
        '------------------------------------------------------------------

        Expense.Description = CStr( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_DESCRIPTION)).Value _
        )

        '------------------------------------------------------------------
        ' ExpenseAmount
        '------------------------------------------------------------------

        If IsNumeric( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_AMOUNT)).Value _
        ) Then

            Expense.ExpenseAmount = CCur( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_AMOUNT)).Value _
            )

        End If

        '------------------------------------------------------------------
        ' CashAmount
        '------------------------------------------------------------------

        If IsNumeric( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_CASH_AMOUNT)).Value _
        ) Then

            Expense.CashAmount = CCur( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_CASH_AMOUNT)).Value _
            )

        End If

        '------------------------------------------------------------------
        ' BankAmount
        '------------------------------------------------------------------

        If IsNumeric( _
            .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_BANK_AMOUNT)).Value _
        ) Then

            Expense.BankAmount = CCur( _
                .Cells(RowIndex, GetColumnIndex(Table, COL_EXPENSE_BANK_AMOUNT)).Value _
            )

        End If

    End With

    Set CreateExpense = Expense

    Exit Function


ErrHandler:

    Set CreateExpense = Nothing

End Function


