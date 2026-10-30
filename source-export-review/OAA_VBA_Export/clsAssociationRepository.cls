VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssociationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Class : clsAssociationRepository
' Purpose : Association Repository
'==============================================================================

'------------------------------------------------------------------------------
' Association tablosunu döndürür.
'------------------------------------------------------------------------------
Private Function GetAssociationTable() As ListObject

    Set GetAssociationTable = _
        ThisWorkbook.Worksheets(SHEET_ASSOCIATION) _
                    .ListObjects(TABLE_ASSOCIATION)

End Function

'------------------------------------------------------------------------------
' Sütun numarasýný döndürür.
'------------------------------------------------------------------------------
Private Function Col(ByVal ColumnName As String) As Long

    Col = GetAssociationTable.ListColumns(ColumnName).Index

End Function
'------------------------------------------------------------------------------
' Satýr bilgisinden Association nesnesi oluþturur.
'------------------------------------------------------------------------------
Private Function CreateAssociationFromRow( _
                ByVal Row As Range) As clsAssociation

    Dim Association As clsAssociation

    Set Association = New clsAssociation

    With Association

        .AssociationID = CLng(Row.Cells(1, Col(COL_ASSOCIATION_ID)).Value)
        .AssociationCode = CStr(Row.Cells(1, Col(COL_ASSOCIATION_CODE)).Value)
        .AssociationName = CStr(Row.Cells(1, Col(COL_ASSOCIATION_NAME)).Value)
        .ShortName = CStr(Row.Cells(1, Col(COL_SHORT_NAME)).Value)

        .LogoFileName = CStr(Row.Cells(1, Col(COL_LOGO_FILENAME)).Value)

        .Address = CStr(Row.Cells(1, Col(COL_ADDRESS)).Value)
        .District = CStr(Row.Cells(1, Col(COL_DISTRICT)).Value)
        .City = CStr(Row.Cells(1, Col(COL_CITY)).Value)
        .PostalCode = CStr(Row.Cells(1, Col(COL_POSTAL_CODE)).Value)
        .Country = CStr(Row.Cells(1, Col(COL_COUNTRY)).Value)

        .Phone = CStr(Row.Cells(1, Col(COL_PHONE)).Value)
        .Mobile = CStr(Row.Cells(1, Col(COL_MOBILE)).Value)
        .Email = CStr(Row.Cells(1, Col(COL_EMAIL)).Value)
        .Website = CStr(Row.Cells(1, Col(COL_WEBSITE)).Value)

        .TaxOffice = CStr(Row.Cells(1, Col(COL_TAX_OFFICE)).Value)
        .TaxNumber = CStr(Row.Cells(1, Col(COL_TAX_NUMBER)).Value)
        .ClubNumber = CStr(Row.Cells(1, Col(COL_CLUB_NUMBER)).Value)

        .DefaultCurrency = CStr(Row.Cells(1, Col(COL_DEFAULT_CURRENCY)).Value)

        .FiscalYearStart = _
            CInt(Val(Row.Cells(1, Col(COL_FISCAL_YEAR_START)).Value))

        If IsDate(Row.Cells(1, Col(COL_CHARTER_DATE)).Value) Then
            .CharterDate = CDate(Row.Cells(1, Col(COL_CHARTER_DATE)).Value)
        End If

        If IsDate(Row.Cells(1, Col(COL_FOUNDATION_DATE)).Value) Then
            .FoundationDate = CDate(Row.Cells(1, Col(COL_FOUNDATION_DATE)).Value)
        End If

        If IsDate(Row.Cells(1, Col(COL_CREATED_DATE)).Value) Then
            .CreatedDate = CDate(Row.Cells(1, Col(COL_CREATED_DATE)).Value)
        End If

        If IsDate(Row.Cells(1, Col(COL_MODIFIED_DATE)).Value) Then
            .ModifiedDate = CDate(Row.Cells(1, Col(COL_MODIFIED_DATE)).Value)
        End If

        .Active = CBool(Row.Cells(1, Col(COL_ACTIVE)).Value)
        .Notes = CStr(Row.Cells(1, Col(COL_NOTES)).Value)

    End With

    Set CreateAssociationFromRow = Association

End Function
'------------------------------------------------------------------------------
' Dernek bilgilerini döndürür.
'------------------------------------------------------------------------------
Public Function GetAssociation() As clsAssociation

    Dim lo As ListObject

    Set lo = GetAssociationTable()

    If lo.ListRows.Count = 0 Then
        Set GetAssociation = Nothing
        Exit Function
    End If

    Set GetAssociation = CreateAssociationFromRow(lo.ListRows(1).Range)

End Function
'------------------------------------------------------------------------------
' Association kaydý var mý?
'------------------------------------------------------------------------------
Public Function Exists() As Boolean

    Exists = (GetAssociationTable().ListRows.Count > 0)

End Function
'------------------------------------------------------------------------------
' Association kaydýný ekler.
'------------------------------------------------------------------------------
Public Function Add(ByVal Association As clsAssociation) As Boolean

    On Error GoTo ErrorHandler

    Dim lo As ListObject
    Dim lr As ListRow

    If Association Is Nothing Then
        Err.Raise vbObjectError + 7401, _
                  "clsAssociationRepository.Add", _
                  "Association nesnesi Nothing."
    End If

    Set lo = GetAssociationTable()

    ' Ayný dosyada ikinci Association kaydý oluþturulamaz (tekil kayýt kuralý).
    If lo.ListRows.Count > 0 Then
        Add = False
        Exit Function
    End If

    Set lr = lo.ListRows.Add

    ' Tüm alanlarý ortak yordam ile yaz
    WriteAssociationToRow lr.Range, Association

    ' Sistem alanlarý
    lr.Range.Cells(1, Col(COL_CREATED_DATE)).Value = Now
    lr.Range.Cells(1, Col(COL_MODIFIED_DATE)).Value = Now
    lr.Range.Cells(1, Col(COL_ACTIVE)).Value = True

    Add = True

    Exit Function

ErrorHandler:

    Add = False

    Err.Raise Err.Number, _
              "clsAssociationRepository.Add", _
              Err.Description

End Function

'------------------------------------------------------------------------------
' Association kaydýný günceller.
'------------------------------------------------------------------------------
Public Function Update( _
            ByVal Association As clsAssociation) As Boolean

    On Error GoTo ErrorHandler

    Dim lo As ListObject

    If Association Is Nothing Then
        Err.Raise vbObjectError + 7402, _
                  "clsAssociationRepository.Update", _
                  "Association nesnesi Nothing."
    End If

    Set lo = GetAssociationTable()

    ' Kayýt yok: teknik hata deðil, False döner.
    If lo.ListRows.Count = 0 Then
        Update = False
        Exit Function
    End If

    WriteAssociationToRow lo.ListRows(1).Range, Association

    lo.ListRows(1).Range.Cells( _
        1, Col(COL_MODIFIED_DATE)).Value = Now

    Update = True

    Exit Function

ErrorHandler:

    Update = False

    Err.Raise Err.Number, _
              "clsAssociationRepository.Update", _
              Err.Description

End Function

'------------------------------------------------------------------------------
' Association nesnesini tablo satýrýna yazar.
'------------------------------------------------------------------------------
Private Sub WriteAssociationToRow( _
            ByVal Row As Range, _
            ByVal Association As clsAssociation)

    With Row

        .Cells(1, Col(COL_ASSOCIATION_ID)).Value = Association.AssociationID
        .Cells(1, Col(COL_ASSOCIATION_CODE)).Value = Association.AssociationCode
        .Cells(1, Col(COL_ASSOCIATION_NAME)).Value = Association.AssociationName
        .Cells(1, Col(COL_SHORT_NAME)).Value = Association.ShortName

        .Cells(1, Col(COL_LOGO_FILENAME)).Value = Association.LogoFileName

        .Cells(1, Col(COL_ADDRESS)).Value = Association.Address
        .Cells(1, Col(COL_DISTRICT)).Value = Association.District
        .Cells(1, Col(COL_CITY)).Value = Association.City
        .Cells(1, Col(COL_POSTAL_CODE)).Value = Association.PostalCode
        .Cells(1, Col(COL_COUNTRY)).Value = Association.Country

        .Cells(1, Col(COL_PHONE)).Value = Association.Phone
        .Cells(1, Col(COL_MOBILE)).Value = Association.Mobile
        .Cells(1, Col(COL_EMAIL)).Value = Association.Email
        .Cells(1, Col(COL_WEBSITE)).Value = Association.Website

        .Cells(1, Col(COL_TAX_OFFICE)).Value = Association.TaxOffice
        .Cells(1, Col(COL_TAX_NUMBER)).Value = Association.TaxNumber
        .Cells(1, Col(COL_CLUB_NUMBER)).Value = Association.ClubNumber

        .Cells(1, Col(COL_DEFAULT_CURRENCY)).Value = Association.DefaultCurrency
        .Cells(1, Col(COL_FISCAL_YEAR_START)).Value = Association.FiscalYearStart

        .Cells(1, Col(COL_CHARTER_DATE)).Value = Association.CharterDate
        .Cells(1, Col(COL_FOUNDATION_DATE)).Value = Association.FoundationDate

        .Cells(1, Col(COL_NOTES)).Value = Association.Notes

    End With

End Sub

