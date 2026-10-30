VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIncomeRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'==============================================================================
' OAA v2
' Repository : clsIncomeRepository
' Amaç       : TBL_INCOME veri eriþimi
'
' Repository sorumluluklarý:
'
'   - Gelir kaydý ekleme
'   - Gelir kaydý getirme
'   - Tüm gelir kayýtlarýný getirme
'   - Sonraki IncomeID oluþturma
'   - Gelir kaydý silme
'
' Ýþ kurallarý Service katmanýndadýr.
' Repository doðrudan kullanýcý arayüzü tarafýndan çaðrýlmaz.
'==============================================================================


'==============================================================================
' SABÝTLER
'==============================================================================

Private Const SHEET_NAME As String = SHEET_INCOME

Private Const TABLE_NAME As String = TABLE_INCOME


'==============================================================================
' ADD
'
' IncomeID Service tarafýndan verilmemiþse Repository oluþturur.
'==============================================================================

Public Function Add( _
    ByVal Income As clsIncome) As Long

    Dim Table As ListObject
    Dim NewRow As ListRow

    On Error GoTo ErrorHandler

    Add = 0


    '--------------------------------------------------------------------------
    ' ENTITY KONTROLÜ
    '--------------------------------------------------------------------------

    If Income Is Nothing Then Exit Function


    '--------------------------------------------------------------------------
    ' INCOME ID
    '--------------------------------------------------------------------------

    If Income.IncomeID <= 0 Then

        Income.IncomeID = _
            GetNextIncomeID

        If Income.IncomeID <= 0 Then

            Add = 0

            Exit Function

        End If

    End If


    '--------------------------------------------------------------------------
    ' TABLO
    '--------------------------------------------------------------------------

    Set Table = _
        GetIncomeTable


    If Table Is Nothing Then Exit Function


    '--------------------------------------------------------------------------
    ' YENÝ SATIR
    '--------------------------------------------------------------------------

    Set NewRow = _
        Table.ListRows.Add


    '--------------------------------------------------------------------------
    ' VERÝLERÝ YAZ
    '--------------------------------------------------------------------------

    With NewRow.Range

        .Cells(1, Table.ListColumns(COL_INCOME_ID).Index).Value = _
            Income.IncomeID

        .Cells(1, Table.ListColumns(COL_INCOME_ENTRY_DATE).Index).Value = _
            Income.EntryDate

        .Cells(1, Table.ListColumns(COL_INCOME_DOCUMENT_NO).Index).Value = _
            Income.DocumentNo

        .Cells(1, Table.ListColumns(COL_INCOME_TYPE).Index).Value = _
            Income.IncomeType

        .Cells(1, Table.ListColumns(COL_INCOME_DESCRIPTION).Index).Value = _
            Income.Description

        .Cells(1, Table.ListColumns(COL_INCOME_AMOUNT).Index).Value = _
            Income.IncomeAmount

        .Cells(1, Table.ListColumns(COL_INCOME_CASH_AMOUNT).Index).Value = _
            Income.CashAmount

        .Cells(1, Table.ListColumns(COL_INCOME_BANK_AMOUNT).Index).Value = _
            Income.BankAmount

    End With


    '--------------------------------------------------------------------------
    ' GERÇEK ID'YÝ DÖNDÜR
    '--------------------------------------------------------------------------

    Add = _
        Income.IncomeID


    Exit Function


ErrorHandler:

    Add = 0

    Err.Raise Err.Number, _
              "clsIncomeRepository.Add", _
              Err.Description

End Function


'==============================================================================
' GET BY ID
'==============================================================================

Public Function GetByID( _
    ByVal IncomeID As Long) As clsIncome

    Dim Table As ListObject
    Dim RowIndex As Long

    Dim Income As clsIncome

    On Error GoTo ErrorHandler


    Set GetByID = Nothing


    If IncomeID <= 0 Then Exit Function


    Set Table = _
        GetIncomeTable


    If Table Is Nothing Then Exit Function


    RowIndex = _
        FindIncomeRow( _
            Table, _
            IncomeID)


    If RowIndex = 0 Then Exit Function


    Set Income = _
        CreateIncome( _
            Table, _
            RowIndex)


    Set GetByID = _
        Income


    Exit Function


ErrorHandler:

    Set GetByID = Nothing

End Function


'==============================================================================
' GET ALL
'==============================================================================

Public Function GetAll() As Collection

    Dim Result As New Collection

    Dim Table As ListObject

    Dim RowIndex As Long

    Dim Income As clsIncome

    On Error GoTo ErrorHandler


    Set GetAll = Nothing


    Set Table = _
        GetIncomeTable


    If Table Is Nothing Then

        Set GetAll = Result

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' TABLO BOÞ
    '--------------------------------------------------------------------------

    If Table.DataBodyRange Is Nothing Then

        Set GetAll = Result

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' TÜM GELÝR KAYITLARI
    '--------------------------------------------------------------------------

    For RowIndex = 1 To Table.ListRows.Count

        Set Income = _
            CreateIncome( _
                Table, _
                RowIndex)


        If Not Income Is Nothing Then

            Result.Add Income

        End If

    Next RowIndex


    Set GetAll = _
        Result


    Exit Function


ErrorHandler:

    Set GetAll = Nothing

End Function


'==============================================================================
' GET NEXT INCOME ID
'
' Mevcut en büyük IncomeID bulunur.
'
' Sonraki ID:
'
'     Max IncomeID + 1
'
' Tablo boþsa:
'
'     1
'==============================================================================

Public Function GetNextIncomeID() As Long

    Dim Table As ListObject

    Dim RowIndex As Long

    Dim CurrentID As Long

    Dim MaxID As Long

    On Error GoTo ErrorHandler


    GetNextIncomeID = 0

    MaxID = 0


    Set Table = _
        GetIncomeTable


    If Table Is Nothing Then Exit Function


    '--------------------------------------------------------------------------
    ' TABLO BOÞ
    '--------------------------------------------------------------------------

    If Table.DataBodyRange Is Nothing Then

        GetNextIncomeID = 1

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' EN BÜYÜK ID
    '--------------------------------------------------------------------------

    For RowIndex = 1 To Table.ListRows.Count

        CurrentID = _
            NzLongValue( _
                Table.DataBodyRange.Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_INCOME_ID).Index _
                ).Value _
            )


        If CurrentID > MaxID Then

            MaxID = CurrentID

        End If

    Next RowIndex


    GetNextIncomeID = _
        MaxID + 1


    Exit Function


ErrorHandler:

    GetNextIncomeID = 0

End Function


'==============================================================================
' UPDATE BY ID
'
' Mevcut gelir kaydini IncomeID uzerinden gunceller.
' Is kurallari Service katmanindadir.
'==============================================================================

Public Function UpdateByID( _
    ByVal Income As clsIncome) As Boolean

    Dim Table As ListObject
    Dim RowIndex As Long

    On Error GoTo ErrorHandler

    UpdateByID = False

    If Income Is Nothing Then Exit Function

    If Income.IncomeID <= 0 Then Exit Function

    Set Table = GetIncomeTable

    If Table Is Nothing Then Exit Function

    RowIndex = FindIncomeRow(Table, Income.IncomeID)

    If RowIndex = 0 Then Exit Function

    With Table.DataBodyRange

        .Cells(RowIndex, Table.ListColumns(COL_INCOME_ENTRY_DATE).Index).Value = _
            Income.EntryDate

        .Cells(RowIndex, Table.ListColumns(COL_INCOME_DOCUMENT_NO).Index).Value = _
            Income.DocumentNo

        .Cells(RowIndex, Table.ListColumns(COL_INCOME_TYPE).Index).Value = _
            Income.IncomeType

        .Cells(RowIndex, Table.ListColumns(COL_INCOME_DESCRIPTION).Index).Value = _
            Income.Description

        .Cells(RowIndex, Table.ListColumns(COL_INCOME_AMOUNT).Index).Value = _
            Income.IncomeAmount

        .Cells(RowIndex, Table.ListColumns(COL_INCOME_CASH_AMOUNT).Index).Value = _
            Income.CashAmount

        .Cells(RowIndex, Table.ListColumns(COL_INCOME_BANK_AMOUNT).Index).Value = _
            Income.BankAmount

    End With

    UpdateByID = True

    Exit Function

ErrorHandler:

    UpdateByID = False

    Err.Raise Err.Number, _
              "clsIncomeRepository.UpdateByID", _
              Err.Description

End Function


'==============================================================================
' DELETE BY ID
'==============================================================================

Public Function DeleteByID( _
    ByVal IncomeID As Long) As Boolean

    Dim Table As ListObject

    Dim RowIndex As Long

    On Error GoTo ErrorHandler


    DeleteByID = False


    If IncomeID <= 0 Then Exit Function


    Set Table = _
        GetIncomeTable


    If Table Is Nothing Then Exit Function


    RowIndex = _
        FindIncomeRow( _
            Table, _
            IncomeID)


    If RowIndex = 0 Then Exit Function


    Table.ListRows(RowIndex).Delete


    DeleteByID = True


    Exit Function


ErrorHandler:

    DeleteByID = False

    Err.Raise Err.Number, _
              "clsIncomeRepository.DeleteByID", _
              Err.Description

End Function


'==============================================================================
' GET INCOME TABLE
'==============================================================================

Private Function GetIncomeTable() As ListObject

    On Error GoTo ErrorHandler


    Set GetIncomeTable = _
        GetTable( _
            SHEET_NAME, _
            TABLE_NAME)


    Exit Function


ErrorHandler:

    Set GetIncomeTable = Nothing

End Function


'==============================================================================
' FIND INCOME ROW
'==============================================================================

Private Function FindIncomeRow( _
    ByVal Table As ListObject, _
    ByVal IncomeID As Long) As Long

    Dim RowIndex As Long

    Dim CurrentID As Long

    On Error GoTo ErrorHandler


    FindIncomeRow = 0


    If Table Is Nothing Then Exit Function


    If Table.DataBodyRange Is Nothing Then Exit Function


    For RowIndex = 1 To Table.ListRows.Count

        CurrentID = _
            NzLongValue( _
                Table.DataBodyRange.Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_INCOME_ID).Index _
                ).Value _
            )


        If CurrentID = IncomeID Then

            FindIncomeRow = _
                RowIndex

            Exit Function

        End If

    Next RowIndex


    Exit Function


ErrorHandler:

    FindIncomeRow = 0

End Function


'==============================================================================
' CREATE INCOME
'
' Excel satýrýný clsIncome entity nesnesine dönüþtürür.
'==============================================================================

Private Function CreateIncome( _
    ByVal Table As ListObject, _
    ByVal RowIndex As Long) As clsIncome

    Dim Income As clsIncome

    On Error GoTo ErrorHandler


    Set Income = _
        New clsIncome


    With Table.DataBodyRange

        Income.IncomeID = _
            NzLongValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_INCOME_ID).Index _
                ).Value _
            )


        Income.EntryDate = _
            NzDateValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_INCOME_ENTRY_DATE).Index _
                ).Value _
            )


        Income.DocumentNo = _
            NzStringValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_INCOME_DOCUMENT_NO).Index _
                ).Value _
            )


        Income.IncomeType = _
            NzStringValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_INCOME_TYPE).Index _
                ).Value _
            )


        Income.Description = _
            NzStringValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_INCOME_DESCRIPTION).Index _
                ).Value _
            )


        Income.IncomeAmount = _
            NzCurrencyValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_INCOME_AMOUNT).Index _
                ).Value _
            )


        Income.CashAmount = _
            NzCurrencyValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_INCOME_CASH_AMOUNT).Index _
                ).Value _
            )


        Income.BankAmount = _
            NzCurrencyValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_INCOME_BANK_AMOUNT).Index _
                ).Value _
            )

    End With


    Set CreateIncome = _
        Income


    Exit Function


ErrorHandler:

    Set CreateIncome = Nothing

End Function


