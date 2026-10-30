VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAssociationWizard 
   Caption         =   "Dernek Bilgileri"
   ClientHeight    =   10965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14955
   OleObjectBlob   =   "frmAssociationWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAssociationWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Service As clsAssociationService
'==============================================================================
' OAA v2
' Form : frmAssociation
' Purpose : Association Management
'==============================================================================

'------------------------------------------------------------------------------
' Form açýlýþý.
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

    Set Service = New clsAssociationService

    InitializeForm

End Sub
'------------------------------------------------------------------------------
' Form hazýrlýðý.
'------------------------------------------------------------------------------
Private Sub InitializeForm()

    LoadCurrencyList

    LoadFiscalYearList

    LoadCountryList

    If Service.Exists Then
        LoadAssociation
    Else
    cmbDefaultCurrency.Value = "TRY"
    cmbCountry.Value = "Türkiye"
    cmbFiscalYearStart.ListIndex = 0
    chkActive.Value = True
End If

    

End Sub
'------------------------------------------------------------------------------
' Association bilgilerini yükler.
'------------------------------------------------------------------------------
Private Sub LoadAssociation()

    Dim Association As clsAssociation

    Set Association = Service.GetAssociation()

    If Association Is Nothing Then Exit Sub

    With Association

        ' Sayfa-1
        txtAssociationName.Value = .AssociationName
        txtShortName.Value = .ShortName
        txtAssociationCode.Value = .AssociationCode
        txtRegistrationNumber.Value = .ClubNumber

        If .FoundationDate <> 0 Then
            txtFoundationDate.Value = Format(.FoundationDate, "dd/mm/yyyy")
        End If

        If .CharterDate <> 0 Then
            txtCharterDate.Value = Format(.CharterDate, "dd/mm/yyyy")
        End If

        ' Sayfa-2
        txtAddress.Value = .Address
        cmbCity.Value = .City
        cmbDistrict.Value = .District
        txtPostalCode.Value = .PostalCode
        cmbCountry.Value = .Country
        txtPhone.Value = .Phone
        txtMobile.Value = .Mobile
        txtEmail.Value = .Email
        txtWebsite.Value = .Website

       ' Sayfa-3
txtTaxOffice.Value = .TaxOffice
txtTaxNumber.Value = .TaxNumber
txtClubNumber.Value = .ClubNumber
txtNotes.Value = .Notes

        ' Sayfa-4
        cmbDefaultCurrency.Value = .DefaultCurrency

        If .FiscalYearStart >= 1 And .FiscalYearStart <= 12 Then
            cmbFiscalYearStart.ListIndex = .FiscalYearStart - 1
        End If

        chkActive.Value = .Active

    End With

End Sub
'------------------------------------------------------------------------------
' Formdan Entity oluþturur.
'------------------------------------------------------------------------------
Private Function ReadForm() As clsAssociation

    Dim Association As clsAssociation

    Set Association = New clsAssociation

    With Association

        ' Sayfa-1
        .AssociationName = Trim$(txtAssociationName.Value)
        .ShortName = Trim$(txtShortName.Value)
        .AssociationCode = Trim$(txtAssociationCode.Value)
        .ClubNumber = Trim$(txtRegistrationNumber.Value)

        If IsDate(txtFoundationDate.Value) Then
            .FoundationDate = CDate(txtFoundationDate.Value)
        End If

        If IsDate(txtCharterDate.Value) Then
            .CharterDate = CDate(txtCharterDate.Value)
        End If

        ' Sayfa-2
        .Address = Trim$(txtAddress.Value)
        .City = Trim$(cmbCity.Value)
        .District = Trim$(cmbDistrict.Value)
        .PostalCode = Trim$(txtPostalCode.Value)
        .Country = Trim$(cmbCountry.Value)
        .Phone = Trim$(txtPhone.Value)
        .Mobile = Trim$(txtMobile.Value)
        .Email = Trim$(txtEmail.Value)
        .Website = Trim$(txtWebsite.Value)

        ' Sayfa-3
        .TaxOffice = Trim$(txtTaxOffice.Value)
        .TaxNumber = Trim$(txtTaxNumber.Value)
        .Notes = Trim$(txtNotes.Value)

        ' Sayfa-4
        .DefaultCurrency = Trim$(cmbDefaultCurrency.Value)
        .FiscalYearStart = cmbFiscalYearStart.ListIndex + 1
        .Active = chkActive.Value

    End With

    Set ReadForm = Association

End Function

'------------------------------------------------------------------------------
' Association kaydeder.
'------------------------------------------------------------------------------
Private Sub SaveAssociation()

    Dim Association As clsAssociation
    Dim Result As Boolean

    Set Association = ReadForm()

    Result = Service.Save(Association)

    If Result Then

        MsgBox "Dernek bilgileri baþarýyla kaydedildi.", _
               vbInformation, _
               "OAA"

    Else

        MsgBox "Dernek bilgileri kaydedilemedi.", _
               vbExclamation, _
               "OAA"

    End If

End Sub
'------------------------------------------------------------------------------
' Para birimlerini yükler.
'------------------------------------------------------------------------------
Private Sub LoadCurrencyList()

    With cmbDefaultCurrency

        .Clear

        .AddItem "TRY"
        .AddItem "USD"
        .AddItem "EUR"

    End With

End Sub
'------------------------------------------------------------------------------
' Mali yýl aylarýný yükler.
'------------------------------------------------------------------------------
Private Sub LoadFiscalYearList()

    Dim Months As Variant
    Dim I As Long

    Months = Array( _
        "Ocak", _
        "Þubat", _
        "Mart", _
        "Nisan", _
        "Mayýs", _
        "Haziran", _
        "Temmuz", _
        "Aðustos", _
        "Eylül", _
        "Ekim", _
        "Kasým", _
        "Aralýk")

    cmbFiscalYearStart.Clear

    For I = LBound(Months) To UBound(Months)
        cmbFiscalYearStart.AddItem Months(I)
    Next I

End Sub
'------------------------------------------------------------------------------
' Ülke listesini yükler.
'------------------------------------------------------------------------------
Private Sub LoadCountryList()

    With cmbCountry

        .Clear

        .AddItem "Türkiye"
        .AddItem "KKTC"

    End With

End Sub
'------------------------------------------------------------------------------
' Formu temizler.
'------------------------------------------------------------------------------
Private Sub ClearForm()

    ' Sayfa-1
    txtAssociationName.Value = vbNullString
    txtShortName.Value = vbNullString
    txtAssociationCode.Value = vbNullString
    txtRegistrationNumber.Value = vbNullString
    txtFoundationDate.Value = vbNullString
    txtCharterDate.Value = vbNullString

    ' Logo
    imgLogo.Picture = LoadPicture()

    ' Sayfa-2
    txtAddress.Value = vbNullString
    cmbCity.Value = vbNullString
    cmbDistrict.Value = vbNullString
    txtPostalCode.Value = vbNullString
    cmbCountry.Value = "Türkiye"
    txtPhone.Value = vbNullString
    txtMobile.Value = vbNullString
    txtEmail.Value = vbNullString
    txtWebsite.Value = vbNullString

   ' Sayfa-3
txtTaxOffice.Value = vbNullString
txtTaxNumber.Value = vbNullString
txtClubNumber.Value = vbNullString
txtNotes.Value = vbNullString

    ' Sayfa-4
    cmbDefaultCurrency.Value = "TRY"
    cmbFiscalYearStart.ListIndex = 0
    chkActive.Value = True

    ' Ýlk sekmeye dön
    pgWizard.Value = 0


End Sub

'------------------------------------------------------------------------------
' Kaydet.
'------------------------------------------------------------------------------
Private Sub cmdSave_Click()

    SaveAssociation

End Sub

'------------------------------------------------------------------------------
' Temizle.
'------------------------------------------------------------------------------
Private Sub cmdClear_Click()

    ClearForm

End Sub

'------------------------------------------------------------------------------
' Kapat.
'------------------------------------------------------------------------------
Private Sub cmdClose_Click()

    Unload Me

End Sub
