VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMemberChargeRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' clsMemberChargeRepository
'
' WF-003 - Üye Borç Hareketleri Repository
'
' Görevi:
'   TBL_MEMBER_CHARGES tablosuna veri eriþimini yönetir.
'
' Ýþ kurallarý burada uygulanmaz.
' Ýþ kurallarý:
'   clsMemberChargeService
'
' Entity:
'   clsMemberCharge
'==============================================================================



'==============================================================================
' Add
'
' Yeni üye borç hareketini tabloya ekler.
'
' Dönen deðer:
'   Oluþturulan ChargeID
'==============================================================================

Public Function Add(ByVal Charge As clsMemberCharge) As Long

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim NewRow As ListRow
    Dim NewChargeID As Long

    On Error GoTo ErrorHandler

    If Charge Is Nothing Then
        Err.Raise vbObjectError + 7101, _
                  "clsMemberChargeRepository.Add", _
                  "Charge nesnesi Nothing."
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_MEMBER_CHARGES)
    Set tbl = ws.ListObjects(TABLE_MEMBER_CHARGES)

    NewChargeID = GetNextChargeID(tbl)

    Charge.ChargeID = NewChargeID

    Set NewRow = tbl.ListRows.Add

    With NewRow.Range

        .Cells(1, tbl.ListColumns(COL_CHARGE_ID).Index).Value = _
            Charge.ChargeID

        .Cells(1, tbl.ListColumns(COL_CHARGE_MEMBER_ID).Index).Value = _
            Charge.MemberID

        .Cells(1, tbl.ListColumns(COL_CHARGE_TYPE).Index).Value = _
            Charge.ChargeType

        .Cells(1, tbl.ListColumns(COL_CHARGE_DATE).Index).Value = _
            Charge.ChargeDate

        .Cells(1, tbl.ListColumns(COL_CHARGE_FISCAL_YEAR).Index).Value = _
            Charge.FiscalYear

        .Cells(1, tbl.ListColumns(COL_CHARGE_DUE_MONTH).Index).Value = _
            Charge.DueMonth

        .Cells(1, tbl.ListColumns(COL_CHARGE_AMOUNT).Index).Value = _
            Charge.Amount

        .Cells(1, tbl.ListColumns(COL_CHARGE_ADJUSTMENT_TYPE).Index).Value = _
            Charge.AdjustmentType

        .Cells(1, tbl.ListColumns(COL_CHARGE_SOURCE_CHARGE_ID).Index).Value = _
            Charge.SourceChargeID

        .Cells(1, tbl.ListColumns(COL_CHARGE_DESCRIPTION).Index).Value = _
            Charge.Description

        .Cells(1, tbl.ListColumns(COL_CHARGE_DECISION_NO).Index).Value = _
            Charge.DecisionNo

        .Cells(1, tbl.ListColumns(COL_CHARGE_DECISION_DATE).Index).Value = _
            Charge.DecisionDate

        .Cells(1, tbl.ListColumns(COL_CHARGE_ENTRY_DATE).Index).Value = _
            Charge.EntryDate

        .Cells(1, tbl.ListColumns(COL_CHARGE_STATUS).Index).Value = _
            Charge.Status

    End With

    Add = NewChargeID

    Exit Function

ErrorHandler:

    Add = 0

    Err.Raise Err.Number, _
              "clsMemberChargeRepository.Add", _
              Err.Description

End Function
'==============================================================================
' DeleteByID
'
' ChargeID degerine gore tek bir borc hareketini siler.
'
' Gorevi:
'   Yalnizca Repository seviyesinde veri silme islemini gerceklestirmek.
'
' Is kurali uygulanmaz.
'
' Donen deger:
'   True  = Kayit silindi
'   False = Kayit bulunamadi
'==============================================================================

Public Function DeleteByID(ByVal ChargeID As Long) As Boolean

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim RowIndex As Long

    On Error GoTo ErrorHandler

    DeleteByID = False

    '--------------------------------------------------------------------------
    ' Gecersiz ChargeID
    '--------------------------------------------------------------------------

    If ChargeID <= 0 Then Exit Function

    '--------------------------------------------------------------------------
    ' Tabloyu al
    '--------------------------------------------------------------------------

    Set ws = ThisWorkbook.Worksheets(SHEET_MEMBER_CHARGES)

    Set tbl = ws.ListObjects(TABLE_MEMBER_CHARGES)

    '--------------------------------------------------------------------------
    ' Tablo bos ise silinecek kayit yok
    '--------------------------------------------------------------------------

    If tbl.DataBodyRange Is Nothing Then Exit Function

    '--------------------------------------------------------------------------
    ' ChargeID satirini bul
    '--------------------------------------------------------------------------

    RowIndex = FindChargeRow(tbl, ChargeID)

    If RowIndex <= 0 Then Exit Function

    '--------------------------------------------------------------------------
    ' Kaydi sil
    '--------------------------------------------------------------------------

    tbl.ListRows(RowIndex).Delete

    DeleteByID = True

    Exit Function

ErrorHandler:

    DeleteByID = False

    Err.Raise Err.Number, _
              "clsMemberChargeRepository.DeleteByID", _
              Err.Description

End Function

'==============================================================================
' GetByID
'
' ChargeID deðerine göre tek bir borç hareketi getirir.
'==============================================================================

Public Function GetByID(ByVal ChargeID As Long) As clsMemberCharge

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim RowIndex As Long
    Dim Charge As clsMemberCharge

    On Error GoTo ErrorHandler

    If ChargeID <= 0 Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SHEET_MEMBER_CHARGES)
    Set tbl = ws.ListObjects(TABLE_MEMBER_CHARGES)

    RowIndex = FindChargeRow(tbl, ChargeID)

    If RowIndex = 0 Then Exit Function

    Set Charge = New clsMemberCharge

    With tbl.ListRows(RowIndex).Range

        Charge.ChargeID = _
            NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_ID).Index).Value)

        Charge.MemberID = _
            NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_MEMBER_ID).Index).Value)

        Charge.ChargeType = _
            NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_TYPE).Index).Value)

        Charge.ChargeDate = _
            NzDateValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DATE).Index).Value)

        Charge.FiscalYear = _
            NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_FISCAL_YEAR).Index).Value)

        Charge.DueMonth = _
            NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DUE_MONTH).Index).Value)

        Charge.Amount = _
            NzCurrencyValue(.Cells(1, tbl.ListColumns(COL_CHARGE_AMOUNT).Index).Value)

        Charge.AdjustmentType = _
            NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_ADJUSTMENT_TYPE).Index).Value)

        Charge.SourceChargeID = _
            NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_SOURCE_CHARGE_ID).Index).Value)

        Charge.Description = _
            NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DESCRIPTION).Index).Value)

        Charge.DecisionNo = _
            NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DECISION_NO).Index).Value)

        Charge.DecisionDate = _
            NzDateValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DECISION_DATE).Index).Value)

        Charge.EntryDate = _
            NzDateValue(.Cells(1, tbl.ListColumns(COL_CHARGE_ENTRY_DATE).Index).Value)

        Charge.Status = _
            NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_STATUS).Index).Value)

    End With

    Set GetByID = Charge

    Exit Function

ErrorHandler:

    Set GetByID = Nothing

End Function


'==============================================================================
' GetByMemberID
'
' Belirli bir üyeye ait tüm borç hareketlerini Collection olarak döndürür.
'==============================================================================

Public Function GetByMemberID(ByVal MemberID As Long) As Collection

    Dim Result As Collection
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim I As Long

    Dim Charge As clsMemberCharge
    Dim RowRange As Range

    On Error GoTo ErrorHandler

    Set Result = New Collection

    If MemberID <= 0 Then
        Set GetByMemberID = Result
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_MEMBER_CHARGES)
    Set tbl = ws.ListObjects(TABLE_MEMBER_CHARGES)

    If tbl.DataBodyRange Is Nothing Then
        Set GetByMemberID = Result
        Exit Function
    End If

    For I = 1 To tbl.ListRows.Count

        Set RowRange = tbl.ListRows(I).Range

        If NzLongValue( _
            RowRange.Cells(1, tbl.ListColumns(COL_CHARGE_MEMBER_ID).Index).Value _
        ) = MemberID Then

            Set Charge = New clsMemberCharge

            With RowRange

                Charge.ChargeID = _
                    NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_ID).Index).Value)

                Charge.MemberID = _
                    NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_MEMBER_ID).Index).Value)

                Charge.ChargeType = _
                    NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_TYPE).Index).Value)

                Charge.ChargeDate = _
                    NzDateValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DATE).Index).Value)

                Charge.FiscalYear = _
                    NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_FISCAL_YEAR).Index).Value)

                Charge.DueMonth = _
                    NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DUE_MONTH).Index).Value)

                Charge.Amount = _
                    NzCurrencyValue(.Cells(1, tbl.ListColumns(COL_CHARGE_AMOUNT).Index).Value)

                Charge.AdjustmentType = _
                    NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_ADJUSTMENT_TYPE).Index).Value)

                Charge.SourceChargeID = _
                    NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_SOURCE_CHARGE_ID).Index).Value)

                Charge.Description = _
                    NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DESCRIPTION).Index).Value)

                Charge.DecisionNo = _
                    NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DECISION_NO).Index).Value)

                Charge.DecisionDate = _
                    NzDateValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DECISION_DATE).Index).Value)

                Charge.EntryDate = _
                    NzDateValue(.Cells(1, tbl.ListColumns(COL_CHARGE_ENTRY_DATE).Index).Value)

                Charge.Status = _
                    NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_STATUS).Index).Value)

            End With

            Result.Add Charge

        End If

    Next I

    Set GetByMemberID = Result

    Exit Function

ErrorHandler:

    Set GetByMemberID = Nothing

End Function


'==============================================================================
' ExistsNormalCharge
'
' Ayný üye + mali yýl + ay + borç türü için
' NORMAL tahakkuk olup olmadýðýný kontrol eder.
'
' Düzeltme hareketleri bu kontrolde dikkate alýnmaz.
'==============================================================================

Public Function ExistsNormalCharge( _
    ByVal MemberID As Long, _
    ByVal FiscalYear As Long, _
    ByVal DueMonth As Long, _
    ByVal ChargeType As String) As Boolean

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim I As Long

    Dim CurrentMemberID As Long
    Dim CurrentFiscalYear As Long
    Dim CurrentDueMonth As Long
    Dim CurrentChargeType As String
    Dim CurrentAdjustmentType As String
    Dim CurrentStatus As String

    On Error GoTo ErrorHandler

    ExistsNormalCharge = False

    If MemberID <= 0 Then Exit Function
    If FiscalYear <= 0 Then Exit Function
    If DueMonth <= 0 Or DueMonth > 12 Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SHEET_MEMBER_CHARGES)
    Set tbl = ws.ListObjects(TABLE_MEMBER_CHARGES)

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For I = 1 To tbl.ListRows.Count

        With tbl.ListRows(I).Range

            CurrentMemberID = _
                NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_MEMBER_ID).Index).Value)

            CurrentFiscalYear = _
                NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_FISCAL_YEAR).Index).Value)

            CurrentDueMonth = _
                NzLongValue(.Cells(1, tbl.ListColumns(COL_CHARGE_DUE_MONTH).Index).Value)

            CurrentChargeType = _
                NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_TYPE).Index).Value)

            CurrentAdjustmentType = _
                NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_ADJUSTMENT_TYPE).Index).Value)

            CurrentStatus = _
                NzStringValue(.Cells(1, tbl.ListColumns(COL_CHARGE_STATUS).Index).Value)

        End With

        If CurrentMemberID = MemberID _
           And CurrentFiscalYear = FiscalYear _
           And CurrentDueMonth = DueMonth _
           And StrComp(CurrentChargeType, ChargeType, vbTextCompare) = 0 _
           And StrComp(CurrentAdjustmentType, ADJUSTMENT_TYPE_NORMAL, vbTextCompare) = 0 _
           And StrComp(CurrentStatus, CHARGE_STATUS_ACTIVE, vbTextCompare) = 0 Then

            ExistsNormalCharge = True
            Exit Function

        End If

    Next I

    Exit Function

ErrorHandler:

    ExistsNormalCharge = False

End Function


'==============================================================================
' GetNextChargeID
'==============================================================================

Private Function GetNextChargeID(ByVal tbl As ListObject) As Long

    Dim Cell As Range
    Dim MaxID As Long
    Dim CurrentID As Long

    MaxID = 0

    If tbl.DataBodyRange Is Nothing Then
        GetNextChargeID = 1
        Exit Function
    End If

    For Each Cell In _
        tbl.ListColumns(COL_CHARGE_ID).DataBodyRange.Cells

        If IsNumeric(Cell.Value) Then

            CurrentID = CLng(Cell.Value)

            If CurrentID > MaxID Then
                MaxID = CurrentID
            End If

        End If

    Next Cell

    GetNextChargeID = MaxID + 1

End Function


'==============================================================================
' FindChargeRow
'==============================================================================

Private Function FindChargeRow( _
    ByVal tbl As ListObject, _
    ByVal ChargeID As Long) As Long

    Dim I As Long
    Dim Value As Variant

    FindChargeRow = 0

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For I = 1 To tbl.ListRows.Count

        Value = tbl.ListRows(I).Range.Cells( _
                    1, _
                    tbl.ListColumns(COL_CHARGE_ID).Index _
                ).Value

        If IsNumeric(Value) Then

            If CLng(Value) = ChargeID Then
                FindChargeRow = I
                Exit Function
            End If

        End If

    Next I

End Function
