VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMemberService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Sýnýf   : clsMemberService
' Amaç    : Üye iþ kurallarýný ve servis iþlemlerini yönetmek
'
' Mimari:
'
' frmMemberCard
'       |
'       v
' clsMemberService
'       |
'       v
' clsMemberRepository
'       |
'       v
' tblMembers
'
' Ýletiþim modeli:
' GSM + E-Posta + Adres
'
' Telefon ve WhatsApp kullanýlmamaktadýr.
'==============================================================================


Private mRepository As clsMemberRepository
Private mLastResult As clsValidationResult


'==============================================================================
' SINIF BAÞLATMA
'==============================================================================

'------------------------------------------------------------------------------
' Service oluþturulduðunda Repository nesnesini oluþturur.
'------------------------------------------------------------------------------

Private Sub Class_Initialize()

    Set mRepository = New clsMemberRepository

    Set mLastResult = New clsValidationResult

    mLastResult.SetSuccess

End Sub


'==============================================================================
' SINIF KAPATMA
'==============================================================================

'------------------------------------------------------------------------------
' Service kapatýldýðýnda Repository nesnesini serbest býrakýr.
'------------------------------------------------------------------------------

Private Sub Class_Terminate()

    Set mRepository = Nothing

    Set mLastResult = Nothing

End Sub


'==============================================================================
' SON ÝÞLEM SONUCU
'==============================================================================

'------------------------------------------------------------------------------
' Son AddMember / UpdateMember iþleminin sonucunu döndürür.
' Service kullanýcýya mesaj göstermez; mesajý çaðýran form gösterir.
'------------------------------------------------------------------------------
Public Property Get LastResult() As clsValidationResult

    Set LastResult = mLastResult

End Property


'------------------------------------------------------------------------------
' Son iþlemi hata olarak iþaretler.
'------------------------------------------------------------------------------
Private Sub SetFailure( _
            ByVal ErrorCode As Long, _
            ByVal FieldName As String, _
            ByVal Message As String)

    mLastResult.SetError ErrorCode, FieldName, Message

End Sub


'==============================================================================
' YENÝ ÜYE EKLEME
'==============================================================================

'------------------------------------------------------------------------------
' Yeni üyeyi iþ kurallarýndan geçirerek Repository'ye gönderir.
'------------------------------------------------------------------------------

Public Function AddMember(ByVal Member As clsMember) As Boolean

    On Error GoTo ErrorHandler

    mLastResult.SetSuccess


    '----------------------------------------------------------------------
    ' Nesne kontrolü
    '----------------------------------------------------------------------

    If Member Is Nothing Then

        AddMember = False

        SetFailure _
            ERR_INVALID_VALUE, _
            vbNullString, _
            "Üye bilgisi boþ."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Üye adý
    '----------------------------------------------------------------------

    If Trim$(Member.FirstName) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_FIRST_NAME, _
            "Üyenin adý boþ býrakýlamaz."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Üye soyadý
    '----------------------------------------------------------------------

    If Trim$(Member.LastName) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_LAST_NAME, _
            "Üyenin soyadý boþ býrakýlamaz."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Üye durumu
    '----------------------------------------------------------------------

    If Trim$(Member.Status) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_STATUS, _
            "Üye durumu belirtilmelidir."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' GSM zorunlu
    '----------------------------------------------------------------------

    If Trim$(Member.Mobile) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_GSM, _
            "GSM numarasý boþ býrakýlamaz."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' E-Posta zorunlu
    '----------------------------------------------------------------------

    If Trim$(Member.Email) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_EMAIL, _
            "E-posta adresi boþ býrakýlamaz."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Adres zorunlu
    '----------------------------------------------------------------------

    If Trim$(Member.Address) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_ADDRESS, _
            "Adres boþ býrakýlamaz."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Ayrýlýþ tarihi (Ýstifa ise zorunlu)
    '----------------------------------------------------------------------

    If Member.Status = MEMBER_STATUS_RESIGNED And Member.LeaveDate <= 0 Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_LEAVE_DATE, _
            "Ýstifa eden üyenin ayrýlýþ tarihi girilmelidir."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' E-Posta mükerrerlik kontrolü
    '----------------------------------------------------------------------

    If mRepository.EmailExists(Member.Email) Then

        SetFailure _
            ERR_DUPLICATE_RECORD, _
            COL_MEMBER_EMAIL, _
            "Bu e-posta adresi kayýtlýdýr."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' GSM mükerrerlik kontrolü
    '----------------------------------------------------------------------

    If mRepository.GSMExists(Member.Mobile) Then

        SetFailure _
            ERR_DUPLICATE_RECORD, _
            COL_MEMBER_GSM, _
            "Bu GSM numarasý kayýtlýdýr."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Repository üzerinden kayýt
    '----------------------------------------------------------------------

    AddMember = mRepository.Add(Member)

    Exit Function


ErrorHandler:

    AddMember = False

    SetFailure _
        ERR_UNKNOWN, _
        vbNullString, _
        "Beklenmeyen hata: " & Err.Description

End Function


'==============================================================================
' ÜYE GETÝRME
'==============================================================================

'------------------------------------------------------------------------------
' Üye numarasýna göre üyeyi getirir.
'
' ÖNEMLÝ:
' Repository içerisindeki gerçek metodun adý GetByID'dir.
'------------------------------------------------------------------------------

Public Function GetMemberByID( _
    ByVal MemberID As Long) As clsMember

    On Error GoTo ErrorHandler


    '----------------------------------------------------------------------
    ' Geçersiz üye numarasý
    '----------------------------------------------------------------------

    If MemberID <= 0 Then

        Set GetMemberByID = Nothing

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Repository'deki doðru metod:
    '
    ' GetByID
    '----------------------------------------------------------------------

    Set GetMemberByID = _
        mRepository.GetByID(MemberID)


    Exit Function


ErrorHandler:

    Set GetMemberByID = Nothing

End Function


'==============================================================================
' YENÝ ÜYE NUMARASI
'==============================================================================

'------------------------------------------------------------------------------
' Bir sonraki kullanýlabilir üye numarasýný döndürür.
'------------------------------------------------------------------------------

Public Function GetNextMemberID() As Long

    On Error GoTo ErrorHandler


    GetNextMemberID = _
        mRepository.GetNextMemberID


    Exit Function


ErrorHandler:

    GetNextMemberID = 0

End Function


'==============================================================================
' ÜYE KONTROLÜ
'==============================================================================

'------------------------------------------------------------------------------
' Belirtilen üye numarasýnýn kayýtlý olup olmadýðýný kontrol eder.
'------------------------------------------------------------------------------

Public Function MemberExists( _
    ByVal MemberID As Long) As Boolean

    On Error GoTo ErrorHandler


    If MemberID <= 0 Then

        MemberExists = False

        Exit Function

    End If


    MemberExists = _
        mRepository.Exists(MemberID)


    Exit Function


ErrorHandler:

    MemberExists = False

End Function


'==============================================================================
' ÜYEYÝ PASÝF YAPMA
'==============================================================================

'------------------------------------------------------------------------------
' Üyeyi veritabanýndan silmez.
'
' Sadece:
'
' Status    = Pasif / Resigned
' LeaveDate = Ayrýlýþ tarihi
'
' olarak güncellenir.
'------------------------------------------------------------------------------

Public Function SetMemberPassive( _
    ByVal MemberID As Long, _
    ByVal LeaveDate As Date) As Boolean

    On Error GoTo ErrorHandler


    mLastResult.SetSuccess

    If MemberID <= 0 Then

        SetMemberPassive = False

        SetFailure _
            ERR_INVALID_VALUE, _
            COL_MEMBER_ID, _
            "Geçerli bir üye numarasý bulunamadý."

        Exit Function

    End If


    If LeaveDate <= 0 Then

        SetMemberPassive = False

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_LEAVE_DATE, _
            "Ýstifa eden üyenin ayrýlýþ tarihi girilmelidir."

        Exit Function

    End If


    SetMemberPassive = _
        mRepository.SetPassive( _
            MemberID, _
            LeaveDate)


    Exit Function


ErrorHandler:

    SetMemberPassive = False

    SetFailure _
        ERR_UNKNOWN, _
        vbNullString, _
        "Beklenmeyen hata: " & Err.Description

End Function


'==============================================================================
' ÜYE GÜNCELLEME
'==============================================================================

'------------------------------------------------------------------------------
' Mevcut üyenin bilgilerini günceller.
'------------------------------------------------------------------------------

Public Function UpdateMember( _
    ByVal Member As clsMember) As Boolean

    On Error GoTo ErrorHandler

    mLastResult.SetSuccess


    '----------------------------------------------------------------------
    ' Nesne kontrolü
    '----------------------------------------------------------------------

    If Member Is Nothing Then

        UpdateMember = False

        SetFailure _
            ERR_INVALID_VALUE, _
            vbNullString, _
            "Üye bilgisi boþ."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Üye numarasý
    '----------------------------------------------------------------------

    If Member.MemberID <= 0 Then

        SetFailure _
            ERR_INVALID_VALUE, _
            COL_MEMBER_ID, _
            "Geçerli bir üye numarasý bulunamadý."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Ad
    '----------------------------------------------------------------------

    If Trim$(Member.FirstName) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_FIRST_NAME, _
            "Üyenin adý boþ býrakýlamaz."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Soyad
    '----------------------------------------------------------------------

    If Trim$(Member.LastName) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_LAST_NAME, _
            "Üyenin soyadý boþ býrakýlamaz."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Durum
    '----------------------------------------------------------------------

    If Trim$(Member.Status) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_STATUS, _
            "Üye durumu belirtilmelidir."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' GSM
    '----------------------------------------------------------------------

    If Trim$(Member.Mobile) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_GSM, _
            "GSM numarasý boþ býrakýlamaz."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' E-Posta
    '----------------------------------------------------------------------

    If Trim$(Member.Email) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_EMAIL, _
            "E-posta adresi boþ býrakýlamaz."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Adres
    '----------------------------------------------------------------------

    If Trim$(Member.Address) = "" Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_ADDRESS, _
            "Adres boþ býrakýlamaz."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Ayrýlýþ tarihi (Ýstifa ise zorunlu)
    '----------------------------------------------------------------------

    If Member.Status = MEMBER_STATUS_RESIGNED And Member.LeaveDate <= 0 Then

        SetFailure _
            ERR_REQUIRED_FIELD, _
            COL_MEMBER_LEAVE_DATE, _
            "Ýstifa eden üyenin ayrýlýþ tarihi girilmelidir."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' E-Posta mükerrerlik kontrolü (üyenin kendi kaydý hariç)
    '----------------------------------------------------------------------

    If mRepository.EmailExists(Member.Email, Member.MemberID) Then

        SetFailure _
            ERR_DUPLICATE_RECORD, _
            COL_MEMBER_EMAIL, _
            "Bu e-posta adresi kayýtlýdýr."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' GSM mükerrerlik kontrolü (üyenin kendi kaydý hariç)
    '----------------------------------------------------------------------

    If mRepository.GSMExists(Member.Mobile, Member.MemberID) Then

        SetFailure _
            ERR_DUPLICATE_RECORD, _
            COL_MEMBER_GSM, _
            "Bu GSM numarasý kayýtlýdýr."

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Repository üzerinden güncelle
    '----------------------------------------------------------------------

    UpdateMember = _
        mRepository.Update(Member)


    Exit Function


ErrorHandler:

    UpdateMember = False

    SetFailure _
        ERR_UNKNOWN, _
        vbNullString, _
        "Beklenmeyen hata: " & Err.Description

End Function


'==============================================================================
' TÜM ÜYELERÝ GETÝR
'==============================================================================

'------------------------------------------------------------------------------
' Tüm üyeleri Collection olarak döndürür.
'------------------------------------------------------------------------------

Public Function GetAllMembers() As Collection

    On Error GoTo ErrorHandler


    Set GetAllMembers = _
        mRepository.GetAll


    Exit Function


ErrorHandler:

    Set GetAllMembers = Nothing

End Function
