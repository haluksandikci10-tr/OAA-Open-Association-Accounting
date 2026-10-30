VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssociation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Class : clsAssociation
' Purpose : Association Model
'==============================================================================

Private mAssociationID As Long
Private mAssociationCode As String
Private mAssociationName As String
Private mShortName As String
Private mLogoFileName As String
Private mAddress As String
Private mDistrict As String
Private mCity As String
Private mPostalCode As String
Private mCountry As String
Private mPhone As String
Private mMobile As String
Private mEmail As String
Private mWebsite As String
Private mTaxOffice As String
Private mTaxNumber As String
Private mClubNumber As String
Private mCharterDate As Date
Private mFoundationDate As Date
Private mDefaultCurrency As String
Private mFiscalYearStart As Integer
Private mNotes As String
Private mCreatedDate As Date
Private mModifiedDate As Date
Private mActive As Boolean
'------------------------------------------------------------------------------
' Dernek ID'sini döndürür.
'------------------------------------------------------------------------------
Public Property Get AssociationID() As Long

    AssociationID = mAssociationID

End Property

'------------------------------------------------------------------------------
' Dernek ID'sini deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let AssociationID(ByVal Value As Long)

    mAssociationID = Value

End Property

'------------------------------------------------------------------------------
' Dernek kodunu döndürür.
'------------------------------------------------------------------------------
Public Property Get AssociationCode() As String

    AssociationCode = mAssociationCode

End Property

'------------------------------------------------------------------------------
' Dernek kodunu deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let AssociationCode(ByVal Value As String)

    mAssociationCode = Value

End Property
'------------------------------------------------------------------------------
' Dernek adýný döndürür.
'------------------------------------------------------------------------------
Public Property Get AssociationName() As String

    AssociationName = mAssociationName

End Property

'------------------------------------------------------------------------------
' Dernek adýný deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let AssociationName(ByVal Value As String)

    mAssociationName = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' Dernek kýsa adýný döndürür.
'------------------------------------------------------------------------------
Public Property Get ShortName() As String

    ShortName = mShortName

End Property

'------------------------------------------------------------------------------
' Dernek kýsa adýný deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let ShortName(ByVal Value As String)

    mShortName = Trim$(Value)

End Property
'------------------------------------------------------------------------------
' Logo dosya adýný döndürür.
'------------------------------------------------------------------------------
Public Property Get LogoFileName() As String

    LogoFileName = mLogoFileName

End Property

'------------------------------------------------------------------------------
' Logo dosya adýný deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let LogoFileName(ByVal Value As String)

    mLogoFileName = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' Adresi döndürür.
'------------------------------------------------------------------------------
Public Property Get Address() As String

    Address = mAddress

End Property

'------------------------------------------------------------------------------
' Adresi deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let Address(ByVal Value As String)

    mAddress = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' Ýlçeyi döndürür.
'------------------------------------------------------------------------------
Public Property Get District() As String

    District = mDistrict

End Property

'------------------------------------------------------------------------------
' Ýlçeyi deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let District(ByVal Value As String)

    mDistrict = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' Ýli döndürür.
'------------------------------------------------------------------------------
Public Property Get City() As String

    City = mCity

End Property

'------------------------------------------------------------------------------
' Ýli deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let City(ByVal Value As String)

    mCity = Trim$(Value)

End Property
'------------------------------------------------------------------------------
' Posta kodunu döndürür.
'------------------------------------------------------------------------------
Public Property Get PostalCode() As String

    PostalCode = mPostalCode

End Property

'------------------------------------------------------------------------------
' Posta kodunu deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let PostalCode(ByVal Value As String)

    mPostalCode = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' Ülkeyi döndürür.
'------------------------------------------------------------------------------
Public Property Get Country() As String

    Country = mCountry

End Property

'------------------------------------------------------------------------------
' Ülkeyi deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let Country(ByVal Value As String)

    mCountry = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' Telefon numarasýný döndürür.
'------------------------------------------------------------------------------
Public Property Get Phone() As String

    Phone = mPhone

End Property

'------------------------------------------------------------------------------
' Telefon numarasýný deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let Phone(ByVal Value As String)

    mPhone = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' Cep telefonu numarasýný döndürür.
'------------------------------------------------------------------------------
Public Property Get Mobile() As String

    Mobile = mMobile

End Property

'------------------------------------------------------------------------------
' Cep telefonu numarasýný deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let Mobile(ByVal Value As String)

    mMobile = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' E-posta adresini döndürür.
'------------------------------------------------------------------------------
Public Property Get Email() As String

    Email = mEmail

End Property

'------------------------------------------------------------------------------
' E-posta adresini deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let Email(ByVal Value As String)

    mEmail = LCase$(Trim$(Value))

End Property

'------------------------------------------------------------------------------
' Web sitesini döndürür.
'------------------------------------------------------------------------------
Public Property Get Website() As String

    Website = mWebsite

End Property

'------------------------------------------------------------------------------
' Web sitesini deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let Website(ByVal Value As String)

    mWebsite = Trim$(Value)

End Property
'------------------------------------------------------------------------------
' Vergi dairesini döndürür.
'------------------------------------------------------------------------------
Public Property Get TaxOffice() As String

    TaxOffice = mTaxOffice

End Property

'------------------------------------------------------------------------------
' Vergi dairesini deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let TaxOffice(ByVal Value As String)

    mTaxOffice = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' Vergi numarasýný döndürür.
'------------------------------------------------------------------------------
Public Property Get TaxNumber() As String

    TaxNumber = mTaxNumber

End Property

'------------------------------------------------------------------------------
' Vergi numarasýný deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let TaxNumber(ByVal Value As String)

    mTaxNumber = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' Kulüp numarasýný döndürür.
'------------------------------------------------------------------------------
Public Property Get ClubNumber() As String

    ClubNumber = mClubNumber

End Property

'------------------------------------------------------------------------------
' Kulüp numarasýný deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let ClubNumber(ByVal Value As String)

    mClubNumber = Trim$(Value)

End Property

'------------------------------------------------------------------------------
' Varsayýlan para birimini döndürür.
'------------------------------------------------------------------------------
Public Property Get DefaultCurrency() As String

    DefaultCurrency = mDefaultCurrency

End Property


'------------------------------------------------------------------------------
' Varsayýlan para birimini deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let DefaultCurrency(ByVal Value As String)

    mDefaultCurrency = UCase$(Trim$(Value))

End Property
'------------------------------------------------------------------------------
' Mali yýl baþlangýç ayýný döndürür.
'------------------------------------------------------------------------------
Public Property Get FiscalYearStart() As Integer

    FiscalYearStart = mFiscalYearStart

End Property

'------------------------------------------------------------------------------
' Mali yýl baþlangýç ayýný deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let FiscalYearStart(ByVal Value As Integer)

    If Value >= 1 And Value <= 12 Then
        mFiscalYearStart = Value
    Else
        mFiscalYearStart = 1
    End If

End Property
'------------------------------------------------------------------------------
' Charter tarihini döndürür.
'------------------------------------------------------------------------------
Public Property Get CharterDate() As Date

    CharterDate = mCharterDate

End Property

'------------------------------------------------------------------------------
' Charter tarihini deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let CharterDate(ByVal Value As Date)

    mCharterDate = Value

End Property

'------------------------------------------------------------------------------
' Kuruluþ tarihini döndürür.
'------------------------------------------------------------------------------
Public Property Get FoundationDate() As Date

    FoundationDate = mFoundationDate

End Property

'------------------------------------------------------------------------------
' Kuruluþ tarihini deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let FoundationDate(ByVal Value As Date)

    mFoundationDate = Value

End Property

'------------------------------------------------------------------------------
' Oluþturulma tarihini döndürür.
'------------------------------------------------------------------------------
Public Property Get CreatedDate() As Date

    CreatedDate = mCreatedDate

End Property

'------------------------------------------------------------------------------
' Oluþturulma tarihini deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let CreatedDate(ByVal Value As Date)

    mCreatedDate = Value

End Property

'------------------------------------------------------------------------------
' Son güncelleme tarihini döndürür.
'------------------------------------------------------------------------------
Public Property Get ModifiedDate() As Date

    ModifiedDate = mModifiedDate

End Property

'------------------------------------------------------------------------------
' Son güncelleme tarihini deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let ModifiedDate(ByVal Value As Date)

    mModifiedDate = Value

End Property

'------------------------------------------------------------------------------
' Aktif durumunu döndürür.
'------------------------------------------------------------------------------
Public Property Get Active() As Boolean

    Active = mActive

End Property

'------------------------------------------------------------------------------
' Aktif durumunu deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let Active(ByVal Value As Boolean)

    mActive = Value

End Property

'------------------------------------------------------------------------------
' Açýklamayý döndürür.
'------------------------------------------------------------------------------
Public Property Get Notes() As String

    Notes = mNotes

End Property

'------------------------------------------------------------------------------
' Açýklamayý deðiþtirir.
'------------------------------------------------------------------------------
Public Property Let Notes(ByVal Value As String)

    mNotes = Trim$(Value)

End Property
