VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReceiptRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Sýnýf   : clsReceiptRepository
' Amaç    : Tahsilat / Makbuz veri eriþimi
'
' Repository sorumluluðu:
'
'   - TBL_RECEIPTS veri eriþimi
'   - Makbuz ekleme
'   - Makbuz getirme
'   - Üyeye ait makbuzlarý getirme
'   - Makbuz silme
'   - Yeni ReceiptID üretme
'   - Yeni ReceiptNumber üretme
'
' Ýþ kurallarý Service katmanýndadýr.
'==============================================================================


'==============================================================================
' SABÝTLER
'==============================================================================

Private Const SHEET_NAME As String = SHEET_RECEIPTS

Private Const TABLE_NAME As String = TABLE_RECEIPTS


'==============================================================================
' ADD
'
' ÖNEMLÝ:
'
' Yeni Receipt nesnesi Form tarafýndan ReceiptID olmadan gelebilir.
'
' Bu durumda Repository:
'
'     GetNextReceiptID
'
' ile yeni ID üretir.
'
' Böylece:
'
'     TBL_RECEIPTS'e kayýt yazýlýr
'     +
'     gerçek ReceiptID döndürülür
'
' Ayný iþlem içinde ID'nin 0 dönmesi nedeniyle Service'in
' "Makbuz kaydý oluþturulamadý" demesi engellenir.
'==============================================================================

Public Function Add( _
    ByVal Receipt As clsReceipt) As Long

    Dim Table As ListObject
    Dim NewRow As ListRow

    On Error GoTo ErrorHandler

    Add = 0


    '--------------------------------------------------------------------------
    ' ENTITY
    '--------------------------------------------------------------------------

    If Receipt Is Nothing Then Exit Function


    '--------------------------------------------------------------------------
    ' RECEIPT ID
    '
    ' Form / Service tarafýndan ID verilmemiþse Repository oluþturur.
    '--------------------------------------------------------------------------

    If Receipt.ReceiptID <= 0 Then

        Receipt.ReceiptID = _
            GetNextReceiptID

        If Receipt.ReceiptID <= 0 Then

            Add = 0

            Exit Function

        End If

    End If


    '--------------------------------------------------------------------------
    ' RECEIPT TABLE
    '--------------------------------------------------------------------------

    Set Table = _
        GetReceiptTable


    If Table Is Nothing Then Exit Function


    '--------------------------------------------------------------------------
    ' NEW ROW
    '--------------------------------------------------------------------------

    Set NewRow = _
        Table.ListRows.Add


    '--------------------------------------------------------------------------
    ' WRITE DATA
    '--------------------------------------------------------------------------

    With NewRow.Range

        .Cells(1, Table.ListColumns(COL_RECEIPT_ID).Index).Value = _
            Receipt.ReceiptID

        .Cells(1, Table.ListColumns(COL_RECEIPT_SERIES).Index).Value = _
            Receipt.Series

        .Cells(1, Table.ListColumns(COL_RECEIPT_NUMBER).Index).Value = _
            Receipt.ReceiptNumber

        .Cells(1, Table.ListColumns(COL_RECEIPT_DATE).Index).Value = _
            Receipt.ReceiptDate

        .Cells(1, Table.ListColumns(COL_RECEIPT_MEMBER_ID).Index).Value = _
            Receipt.MemberID

        .Cells(1, Table.ListColumns(COL_RECEIPT_CHARGE_ID).Index).Value = _
            Receipt.ChargeID

        .Cells(1, Table.ListColumns(COL_RECEIPT_AMOUNT).Index).Value = _
            Receipt.Amount

        .Cells(1, Table.ListColumns(COL_RECEIPT_PAYMENT_METHOD).Index).Value = _
            Receipt.PaymentMethod

    End With


    '--------------------------------------------------------------------------
    ' GERÇEK RECEIPT ID'YÝ DÖNDÜR
    '--------------------------------------------------------------------------

    Add = _
        Receipt.ReceiptID


    Exit Function


ErrorHandler:

    Add = 0

    Err.Raise Err.Number, _
              "clsReceiptRepository.Add", _
              Err.Description

End Function


'==============================================================================
' GET BY ID
'==============================================================================

Public Function GetByID( _
    ByVal ReceiptID As Long) As clsReceipt

    Dim Table As ListObject
    Dim RowIndex As Long

    Dim Receipt As clsReceipt

    On Error GoTo ErrorHandler


    Set GetByID = Nothing


    If ReceiptID <= 0 Then Exit Function


    Set Table = _
        GetReceiptTable


    If Table Is Nothing Then Exit Function


    RowIndex = _
        FindReceiptRow( _
            Table, _
            ReceiptID)


    If RowIndex = 0 Then Exit Function


    Set Receipt = _
        CreateReceipt( _
            Table, _
            RowIndex)


    Set GetByID = _
        Receipt


    Exit Function


ErrorHandler:

    Set GetByID = Nothing

End Function


'==============================================================================
' GET BY MEMBER ID
'==============================================================================

Public Function GetByMemberID( _
    ByVal MemberID As Long) As Collection

    Dim Result As New Collection

    Dim Table As ListObject

    Dim RowIndex As Long

    Dim Receipt As clsReceipt

    On Error GoTo ErrorHandler


    Set GetByMemberID = Nothing


    If MemberID <= 0 Then

        Set GetByMemberID = _
            Result

        Exit Function

    End If


    Set Table = _
        GetReceiptTable


    If Table Is Nothing Then

        Set GetByMemberID = _
            Result

        Exit Function

    End If


    If Table.DataBodyRange Is Nothing Then

        Set GetByMemberID = _
            Result

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' ÜYEYE AÝT MAKBUZLAR
    '--------------------------------------------------------------------------

    For RowIndex = 1 To Table.ListRows.Count

        If NzLongValue( _
            Table.DataBodyRange.Cells( _
                RowIndex, _
                Table.ListColumns(COL_RECEIPT_MEMBER_ID).Index _
            ).Value _
        ) = MemberID Then


            Set Receipt = _
                CreateReceipt( _
                    Table, _
                    RowIndex)


            If Not Receipt Is Nothing Then

                Result.Add Receipt

            End If

        End If

    Next RowIndex


    Set GetByMemberID = _
        Result


    Exit Function


ErrorHandler:

    Set GetByMemberID = Nothing

End Function


'==============================================================================
' GET NEXT RECEIPT ID
'
' TBL_RECEIPTS içerisindeki mevcut en büyük ReceiptID bulunur.
'
' Sonraki ID:
'
'     Max ReceiptID + 1
'
' Tablo boþsa:
'
'     1
'==============================================================================

Public Function GetNextReceiptID() As Long

    Dim Table As ListObject

    Dim RowIndex As Long

    Dim CurrentID As Long

    Dim MaxID As Long

    On Error GoTo ErrorHandler


    GetNextReceiptID = 0

    MaxID = 0


    Set Table = _
        GetReceiptTable


    If Table Is Nothing Then Exit Function


    '--------------------------------------------------------------------------
    ' TABLO BOÞSA
    '--------------------------------------------------------------------------

    If Table.DataBodyRange Is Nothing Then

        GetNextReceiptID = 1

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' EN BÜYÜK RECEIPT ID
    '--------------------------------------------------------------------------

    For RowIndex = 1 To Table.ListRows.Count

        CurrentID = _
            NzLongValue( _
                Table.DataBodyRange.Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_ID).Index _
                ).Value _
            )


        If CurrentID > MaxID Then

            MaxID = CurrentID

        End If

    Next RowIndex


    GetNextReceiptID = _
        MaxID + 1


    Exit Function


ErrorHandler:

    GetNextReceiptID = 0

End Function


'==============================================================================
' GET NEXT RECEIPT NUMBER
'
' TBL_RECEIPTS içerisindeki mevcut en büyük ReceiptNumber bulunur.
'
' Sonraki numara:
'
'     Max ReceiptNumber + 1
'
' Tablo boþsa:
'
'     1
'
' Seri dikkate alýnmaz.
'==============================================================================

Public Function GetNextReceiptNumber() As Long

    Dim Table As ListObject

    Dim RowIndex As Long

    Dim CurrentNumber As Long

    Dim MaxNumber As Long

    On Error GoTo ErrorHandler


    GetNextReceiptNumber = 0

    MaxNumber = 0


    Set Table = _
        GetReceiptTable


    If Table Is Nothing Then Exit Function


    '--------------------------------------------------------------------------
    ' TABLO BOÞSA
    '--------------------------------------------------------------------------

    If Table.DataBodyRange Is Nothing Then

        GetNextReceiptNumber = 1

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' EN BÜYÜK MAKBUZ NUMARASI
    '--------------------------------------------------------------------------

    For RowIndex = 1 To Table.ListRows.Count

        CurrentNumber = _
            NzLongValue( _
                Table.DataBodyRange.Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_NUMBER).Index _
                ).Value _
            )


        If CurrentNumber > MaxNumber Then

            MaxNumber = CurrentNumber

        End If

    Next RowIndex


    '--------------------------------------------------------------------------
    ' SONRAKÝ MAKBUZ NUMARASI
    '--------------------------------------------------------------------------

    GetNextReceiptNumber = _
        MaxNumber + 1


    Exit Function


ErrorHandler:

    GetNextReceiptNumber = 0

End Function


'==============================================================================
' DELETE BY ID
'==============================================================================

Public Function DeleteByID( _
    ByVal ReceiptID As Long) As Boolean

    Dim Table As ListObject

    Dim RowIndex As Long

    On Error GoTo ErrorHandler


    DeleteByID = False


    If ReceiptID <= 0 Then Exit Function


    Set Table = _
        GetReceiptTable


    If Table Is Nothing Then Exit Function


    RowIndex = _
        FindReceiptRow( _
            Table, _
            ReceiptID)


    If RowIndex = 0 Then Exit Function


    Table.ListRows(RowIndex).Delete


    DeleteByID = True


    Exit Function


ErrorHandler:

    DeleteByID = False

    Err.Raise Err.Number, _
              "clsReceiptRepository.DeleteByID", _
              Err.Description

End Function


'==============================================================================
' GET RECEIPT TABLE
'==============================================================================

Private Function GetReceiptTable() As ListObject

    On Error GoTo ErrorHandler


    Set GetReceiptTable = _
        GetTable( _
            SHEET_NAME, _
            TABLE_NAME)


    Exit Function


ErrorHandler:

    Set GetReceiptTable = Nothing

End Function


'==============================================================================
' FIND RECEIPT ROW
'==============================================================================

Private Function FindReceiptRow( _
    ByVal Table As ListObject, _
    ByVal ReceiptID As Long) As Long

    Dim RowIndex As Long

    Dim CurrentID As Long

    On Error GoTo ErrorHandler


    FindReceiptRow = 0


    If Table Is Nothing Then Exit Function


    If Table.DataBodyRange Is Nothing Then Exit Function


    For RowIndex = 1 To Table.ListRows.Count

        CurrentID = _
            NzLongValue( _
                Table.DataBodyRange.Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_ID).Index _
                ).Value _
            )


        If CurrentID = ReceiptID Then

            FindReceiptRow = _
                RowIndex

            Exit Function

        End If

    Next RowIndex


    Exit Function


ErrorHandler:

    FindReceiptRow = 0

End Function


'==============================================================================
' CREATE RECEIPT
'==============================================================================

Private Function CreateReceipt( _
    ByVal Table As ListObject, _
    ByVal RowIndex As Long) As clsReceipt

    Dim Receipt As clsReceipt

    On Error GoTo ErrorHandler


    Set Receipt = _
        New clsReceipt


    With Table.DataBodyRange

        Receipt.ReceiptID = _
            NzLongValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_ID).Index _
                ).Value _
            )


        Receipt.Series = _
            NzStringValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_SERIES).Index _
                ).Value _
            )


        Receipt.ReceiptNumber = _
            NzLongValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_NUMBER).Index _
                ).Value _
            )


        Receipt.ReceiptDate = _
            NzDateValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_DATE).Index _
                ).Value _
            )


        Receipt.MemberID = _
            NzLongValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_MEMBER_ID).Index _
                ).Value _
            )


        Receipt.ChargeID = _
            NzLongValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_CHARGE_ID).Index _
                ).Value _
            )


        Receipt.Amount = _
            NzCurrencyValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_AMOUNT).Index _
                ).Value _
            )


        Receipt.PaymentMethod = _
            NzStringValue( _
                .Cells( _
                    RowIndex, _
                    Table.ListColumns(COL_RECEIPT_PAYMENT_METHOD).Index _
                ).Value _
            )

    End With


    Set CreateReceipt = _
        Receipt


    Exit Function


ErrorHandler:

    Set CreateReceipt = Nothing

End Function
