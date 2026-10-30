VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMemberCard 
   Caption         =   "HS Yazýlým- Üye Kartý"
   ClientHeight    =   11820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11355
   OleObjectBlob   =   "frmMemberCard.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMemberCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Form    : frmMemberCard
' Amaç    : Üye kartý / üye bilgi giriþ ve güncelleme ekraný
'
' Ýletiþim bilgileri:
'   - GSM
'   - E-Posta
'   - Adres
'
' Telefon ve WhatsApp kullanýlmamaktadýr.
'==============================================================================

Private mService As clsMemberService

Private mIsNewRecord As Boolean

Private mCurrentMemberID As Long


'==============================================================================
' GENEL BÝLGÝLER
'==============================================================================

Private Sub fraGeneral_Click()

End Sub


'==============================================================================
' FORM BAÞLATMA
'==============================================================================

'------------------------------------------------------------------------------
' Form açýldýðýnda çalýþýr.
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

    Set mService = New clsMemberService

    InitializeControls

    StartNewRecord

End Sub


'------------------------------------------------------------------------------
' Form üzerindeki kontrolleri hazýrlar.
'------------------------------------------------------------------------------
Private Sub InitializeControls()

    Dim v As Variant

    With cboStatus

        .Clear

        For Each v In GetMemberStatusList()

            .AddItem v

        Next v

        .ListIndex = 0

    End With

    ClearForm

End Sub


'==============================================================================
' FORM TEMÝZLEME
'==============================================================================

'------------------------------------------------------------------------------
' Form üzerindeki tüm giriþ alanlarýný temizler.
'------------------------------------------------------------------------------
Private Sub ClearForm()

    '----------------------------------------------------------------------
    ' Genel bilgiler
    '----------------------------------------------------------------------

    txtMemberID.Value = ""

    txtFirstName.Value = ""

    txtLastName.Value = ""

    cboStatus.ListIndex = 0


    '----------------------------------------------------------------------
    ' Ýletiþim bilgileri
    '----------------------------------------------------------------------

    txtGSM.Value = ""

    txtEmail.Value = ""

    txtAddress.Value = ""


    '----------------------------------------------------------------------
    ' Tarih bilgileri
    '----------------------------------------------------------------------

    txtBirthDate.Value = ""

    txtMarriageDate.Value = ""

    txtJoinDate.Value = ""

    txtLeaveDate.Value = ""


    '----------------------------------------------------------------------
    ' Aile bilgileri
    '----------------------------------------------------------------------

    txtSpouseName.Value = ""

    txtNotes.Value = ""

End Sub


'==============================================================================
' YENÝ ÜYE
'==============================================================================

'------------------------------------------------------------------------------
' Yeni üye kaydý baþlatýr.
'------------------------------------------------------------------------------
Private Sub cmdNew_Click()

    StartNewRecord

    txtMemberID.Locked = True

    cmdSave.Enabled = True

    cmdUpdate.Enabled = False

    txtFirstName.SetFocus

End Sub


'------------------------------------------------------------------------------
' Yeni kayýt modunu baþlatýr.
'------------------------------------------------------------------------------
Private Sub StartNewRecord()

    ClearForm

    mIsNewRecord = True

    mCurrentMemberID = mService.GetNextMemberID

    txtMemberID.Value = Format$(mCurrentMemberID, "000000")

    cboStatus.Value = MEMBER_STATUS_ACTIVE

    cmdSave.Enabled = True

    cmdUpdate.Enabled = False

    txtMemberID.Locked = True

    txtFirstName.SetFocus

End Sub


'==============================================================================
' FORM › ÜYE NESNESÝ
'==============================================================================

'------------------------------------------------------------------------------
' Formdaki bilgileri clsMember nesnesine aktarýr.
'------------------------------------------------------------------------------
Private Function CreateMemberFromForm() As clsMember

    Dim Member As clsMember

    Dim ParsedDate As Date


    Set Member = New clsMember


    With Member

        '----------------------------------------------------------------------
        ' Üye numarasý
        '----------------------------------------------------------------------

        If Trim$(txtMemberID.Value) = "" Then

            Err.Raise _
                vbObjectError + 1000, _
                "frmMemberCard", _
                "Üye numarasý oluþturulamadý."

        End If

        .MemberID = CLng(txtMemberID.Value)


        '----------------------------------------------------------------------
        ' Genel bilgiler
        '----------------------------------------------------------------------

        .FirstName = Trim$(txtFirstName.Value)

        .LastName = Trim$(txtLastName.Value)

        .Status = Trim$(cboStatus.Value)


        '----------------------------------------------------------------------
        ' Ýletiþim bilgileri
        '
        ' Kesin iletiþim modeli:
        '
        ' GSM
        ' E-Posta
        ' Adres
        '
        ' Telefon ve WhatsApp kullanýlmamaktadýr.
        '----------------------------------------------------------------------

        .Mobile = Trim$(txtGSM.Value)

        .Email = Trim$(txtEmail.Value)

        .Address = Trim$(txtAddress.Value)


        '----------------------------------------------------------------------
        ' Aile bilgileri
        '----------------------------------------------------------------------

        .SpouseName = Trim$(txtSpouseName.Value)

        .Notes = Trim$(txtNotes.Value)


        '======================================================================
        ' TARÝH BÝLGÝLERÝ
        '======================================================================


        '----------------------------------------------------------------------
        ' Doðum Tarihi
        '----------------------------------------------------------------------

        If Trim$(txtBirthDate.Value) <> "" Then

            If Not TryParseDate( _
                txtBirthDate.Value, _
                ParsedDate) Then

                MsgBox _
                    "Doðum tarihi geçersiz." & vbCrLf & vbCrLf & _
                    "Kullanýlabilecek biçimler:" & vbCrLf & _
                    "GG.AA.YYYY" & vbCrLf & _
                    "Örnek: 02.01.1960", _
                    vbExclamation, _
                    "OAA - Tarih Bilgisi"

                txtBirthDate.SetFocus

                Exit Function

            End If

            .BirthDate = ParsedDate

        End If


        '----------------------------------------------------------------------
        ' Evlilik Tarihi
        '----------------------------------------------------------------------

        If Trim$(txtMarriageDate.Value) <> "" Then

            If Not TryParseDate( _
                txtMarriageDate.Value, _
                ParsedDate) Then

                MsgBox _
                    "Evlilik tarihi geçersiz." & vbCrLf & vbCrLf & _
                    "Kullanýlabilecek biçimler:" & vbCrLf & _
                    "GG.AA.YYYY" & vbCrLf & _
                    "Örnek: 26.12.1988", _
                    vbExclamation, _
                    "OAA - Tarih Bilgisi"

                txtMarriageDate.SetFocus

                Exit Function

            End If

            .MarriageDate = ParsedDate

        End If


        '----------------------------------------------------------------------
        ' Giriþ Tarihi
        '----------------------------------------------------------------------

        If Trim$(txtJoinDate.Value) <> "" Then

            If Not TryParseDate( _
                txtJoinDate.Value, _
                ParsedDate) Then

                MsgBox _
                    "Giriþ tarihi geçersiz." & vbCrLf & vbCrLf & _
                    "Kullanýlabilecek biçim:" & vbCrLf & _
                    "GG.AA.YYYY", _
                    vbExclamation, _
                    "OAA - Tarih Bilgisi"

                txtJoinDate.SetFocus

                Exit Function

            End If

            .JoinDate = ParsedDate

        End If


        '----------------------------------------------------------------------
        ' Ayrýlýþ Tarihi
        '----------------------------------------------------------------------

        If Trim$(txtLeaveDate.Value) <> "" Then

            If Not TryParseDate( _
                txtLeaveDate.Value, _
                ParsedDate) Then

                MsgBox _
                    "Ayrýlýþ tarihi geçersiz." & vbCrLf & vbCrLf & _
                    "Kullanýlabilecek biçim:" & vbCrLf & _
                    "GG.AA.YYYY", _
                    vbExclamation, _
                    "OAA - Tarih Bilgisi"

                txtLeaveDate.SetFocus

                Exit Function

            End If

            .LeaveDate = ParsedDate

        End If

    End With


    Set CreateMemberFromForm = Member

End Function


'==============================================================================
' FORM KONTROLLERÝ
'==============================================================================

'------------------------------------------------------------------------------
' Formdaki zorunlu alanlarý kontrol eder.
'------------------------------------------------------------------------------
Private Function ValidateForm() As Boolean

    '----------------------------------------------------------------------
    ' Ad
    '----------------------------------------------------------------------

    If Trim$(txtFirstName.Value) = "" Then

        MsgBox _
            "Lütfen Ad alanýný giriniz.", _
            vbExclamation, _
            "OAA - Üye Bilgileri"

        txtFirstName.SetFocus

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Soyad
    '----------------------------------------------------------------------

    If Trim$(txtLastName.Value) = "" Then

        MsgBox _
            "Lütfen Soyad alanýný giriniz.", _
            vbExclamation, _
            "OAA - Üye Bilgileri"

        txtLastName.SetFocus

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Durum
    '----------------------------------------------------------------------

    If Trim$(cboStatus.Value) = "" Then

        MsgBox _
            "Lütfen üye durumunu seçiniz.", _
            vbExclamation, _
            "OAA - Üye Bilgileri"

        cboStatus.SetFocus

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Ayrýlýþ tarihi (Ýstifa ise zorunlu)
    '----------------------------------------------------------------------

    If Trim$(cboStatus.Value) = MEMBER_STATUS_RESIGNED And _
       Trim$(txtLeaveDate.Value) = "" Then

        MsgBox _
            "Ýstifa eden üyenin ayrýlýþ tarihi girilmelidir.", _
            vbExclamation, _
            "OAA - Üye Bilgileri"

        txtLeaveDate.SetFocus

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' GSM
    '----------------------------------------------------------------------

    If Trim$(txtGSM.Value) = "" Then

        MsgBox _
            "Lütfen GSM numarasýný giriniz.", _
            vbExclamation, _
            "OAA - Ýletiþim Bilgileri"

        txtGSM.SetFocus

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' E-Posta
    '----------------------------------------------------------------------

    If Trim$(txtEmail.Value) = "" Then

        MsgBox _
            "Lütfen e-posta adresini giriniz.", _
            vbExclamation, _
            "OAA - Ýletiþim Bilgileri"

        txtEmail.SetFocus

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Adres
    '----------------------------------------------------------------------

    If Trim$(txtAddress.Value) = "" Then

        MsgBox _
            "Lütfen adres bilgisini giriniz.", _
            vbExclamation, _
            "OAA - Ýletiþim Bilgileri"

        txtAddress.SetFocus

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Tüm kontroller baþarýlý
    '----------------------------------------------------------------------

    ValidateForm = True

End Function


'==============================================================================
' ÜYE KAYDET
'==============================================================================

'------------------------------------------------------------------------------
' Yeni üyeyi kaydeder.
'------------------------------------------------------------------------------
Private Sub cmdSave_Click()

    Dim Member As clsMember

    If Not ValidateForm Then Exit Sub


    On Error GoTo ErrorHandler


    Set Member = CreateMemberFromForm()


    '----------------------------------------------------------------------
    ' Tarih doðrulamasý baþarýsýzsa nesne oluþturulmamýþ olabilir.
    '----------------------------------------------------------------------

    If Member Is Nothing Then Exit Sub


    '----------------------------------------------------------------------
    ' Service üzerinden kayýt
    '----------------------------------------------------------------------

    If mService.AddMember(Member) Then

        RefreshMemberList

        MsgBox _
            "Üye baþarýyla kaydedildi.", _
            vbInformation, _
            "OAA - Üye Kaydý"

        StartNewRecord

    Else

        ShowServiceFailure _
            "Kayýt sýrasýnda hata oluþtu.", _
            "OAA - Üye Kaydý"

    End If

    Exit Sub


ErrorHandler:

    MsgBox _
        "Üye kaydý sýrasýnda hata oluþtu." & vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbCritical, _
        "OAA - Hata"

End Sub


'==============================================================================
' ÜYE GÜNCELLE
'==============================================================================

'------------------------------------------------------------------------------
' Mevcut üyenin bilgilerini günceller.
'------------------------------------------------------------------------------
Private Sub cmdUpdate_Click()

    Dim Member As clsMember


    If mIsNewRecord Then

        MsgBox _
            "Yeni kayýt güncellenemez.", _
            vbExclamation, _
            "OAA - Üye Güncelleme"

        Exit Sub

    End If


    If Not ValidateForm Then Exit Sub


    On Error GoTo ErrorHandler


    Set Member = CreateMemberFromForm()


    If Member Is Nothing Then Exit Sub


    '----------------------------------------------------------------------
    ' Service üzerinden güncelleme
    '----------------------------------------------------------------------

    If mService.UpdateMember(Member) Then

        RefreshMemberList

        MsgBox _
            "Üye bilgileri güncellendi.", _
            vbInformation, _
            "OAA - Üye Güncelleme"

        LoadMember Member.MemberID

    Else

        ShowServiceFailure _
            "Güncelleme baþarýsýz.", _
            "OAA - Üye Güncelleme"

    End If

    Exit Sub


ErrorHandler:

    MsgBox _
        "Üye güncelleme sýrasýnda hata oluþtu." & vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbCritical, _
        "OAA - Hata"

End Sub


'==============================================================================
' ÜYE GETÝRME
'==============================================================================

'------------------------------------------------------------------------------
' Dýþarýdan verilen üye numarasýna göre üye kartýný açar.
'------------------------------------------------------------------------------
Public Sub OpenMember(ByVal MemberID As Long)

    LoadMember MemberID

    Me.Show

End Sub


'------------------------------------------------------------------------------
' Üyeyi Service katmanýndan alarak forma yükler.
'------------------------------------------------------------------------------
Public Sub LoadMember(ByVal MemberID As Long)

    Dim Member As clsMember


    On Error GoTo ErrorHandler


    Set Member = mService.GetMemberByID(MemberID)


    '----------------------------------------------------------------------
    ' Üye bulunamadý
    '----------------------------------------------------------------------

    If Member Is Nothing Then

        MsgBox _
            "Üye bulunamadý.", _
            vbExclamation, _
            "OAA - Üye Bilgileri"

        Exit Sub

    End If


    '----------------------------------------------------------------------
    ' Mevcut kayýt modu
    '----------------------------------------------------------------------

    mIsNewRecord = False

    mCurrentMemberID = Member.MemberID


    '----------------------------------------------------------------------
    ' Üye numarasý
    '----------------------------------------------------------------------

    txtMemberID.Value = _
        Format$(Member.MemberID, "000000")


    '----------------------------------------------------------------------
    ' Genel bilgiler
    '----------------------------------------------------------------------

    txtFirstName.Value = Member.FirstName

    txtLastName.Value = Member.LastName

    cboStatus.Value = Member.Status


    '----------------------------------------------------------------------
    ' Ýletiþim bilgileri
    '----------------------------------------------------------------------

    txtGSM.Value = Member.Mobile

    txtEmail.Value = Member.Email

    txtAddress.Value = Member.Address


    '----------------------------------------------------------------------
    ' Aile bilgileri
    '----------------------------------------------------------------------

    txtSpouseName.Value = Member.SpouseName

    txtNotes.Value = Member.Notes


    '----------------------------------------------------------------------
    ' Doðum tarihi
    '----------------------------------------------------------------------

    If Member.BirthDate <> 0 Then

        txtBirthDate.Value = _
            Format$(Member.BirthDate, "dd.mm.yyyy")

    Else

        txtBirthDate.Value = ""

    End If


    '----------------------------------------------------------------------
    ' Evlilik tarihi
    '----------------------------------------------------------------------

    If Member.MarriageDate <> 0 Then

        txtMarriageDate.Value = _
            Format$(Member.MarriageDate, "dd.mm.yyyy")

    Else

        txtMarriageDate.Value = ""

    End If


    '----------------------------------------------------------------------
    ' Giriþ tarihi
    '----------------------------------------------------------------------

    If Member.JoinDate > 0 Then

        txtJoinDate.Value = _
            Format$(Member.JoinDate, "dd.mm.yyyy")

    Else

        txtJoinDate.Value = ""

    End If


    '----------------------------------------------------------------------
    ' Ayrýlýþ tarihi
    '----------------------------------------------------------------------

    If Member.LeaveDate > 0 Then

        txtLeaveDate.Value = _
            Format$(Member.LeaveDate, "dd.mm.yyyy")

    Else

        txtLeaveDate.Value = ""

    End If


    '----------------------------------------------------------------------
    ' Buton durumlarý
    '----------------------------------------------------------------------

    cmdSave.Enabled = False

    cmdUpdate.Enabled = True

    txtMemberID.Locked = True

    Exit Sub


ErrorHandler:

    MsgBox _
        "Üye bilgileri forma yüklenirken hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbCritical, _
        "OAA - Hata"

End Sub


'==============================================================================
' SERVÝS HATASINI GÖSTER
'==============================================================================

'------------------------------------------------------------------------------
' Service iþlemi baþarýsýz olduðunda Service'in verdiði mesajý gösterir.
' Mesaj yoksa verilen varsayýlan mesaj kullanýlýr.
'------------------------------------------------------------------------------
Private Sub ShowServiceFailure( _
            ByVal DefaultMessage As String, _
            ByVal Title As String)

    Dim Result As clsValidationResult

    Set Result = mService.LastResult

    If Len(Result.Message) = 0 Then

        MsgBox DefaultMessage, vbCritical, Title

    ElseIf Result.ErrorCode = ERR_UNKNOWN Then

        MsgBox Result.Message, vbCritical, Title

    Else

        MsgBox Result.Message, vbExclamation, Title

    End If

End Sub


'==============================================================================
' FORM KAPATMA
'==============================================================================

'------------------------------------------------------------------------------
' Üye kartýný kapatýr.
'------------------------------------------------------------------------------
Private Sub cmdClose_Click()

    Unload Me

End Sub


'==============================================================================
' TARÝH DÖNÜÞTÜRME
'==============================================================================

'------------------------------------------------------------------------------
' Girilen tarihi güvenli þekilde Date deðerine dönüþtürür.
'
' Kabul edilen biçimler:
'
'   02.01.1960
'   02/01/1960
'   02-01-1960
'   02011960
'
' 8 haneli sayý GGAAYYYY olarak deðerlendirilir.
' Örneðin:
'
'   02011960 = 02.01.1960
'   26121988 = 26.12.1988
'
'------------------------------------------------------------------------------

Private Function TryParseDate( _
    ByVal TextValue As String, _
    ByRef ResultDate As Date) As Boolean

    Dim CleanValue As String

    Dim DayValue As Long
    Dim MonthValue As Long
    Dim YearValue As Long

    Dim TestDate As Date


    On Error GoTo ErrorHandler


    CleanValue = Trim$(TextValue)


    '----------------------------------------------------------------------
    ' Boþ deðer
    '----------------------------------------------------------------------

    If CleanValue = "" Then

        TryParseDate = False

        Exit Function

    End If


    '======================================================================
    ' 8 HANELÝ TARÝH
    '
    ' GGAAYYYY
    '
    ' Örnek:
    ' 02011960
    ' 26121988
    '======================================================================

    If Len(CleanValue) = 8 And IsNumeric(CleanValue) Then

        DayValue = CLng(Left$(CleanValue, 2))

        MonthValue = CLng(Mid$(CleanValue, 3, 2))

        YearValue = CLng(Right$(CleanValue, 4))


        '------------------------------------------------------------------
        ' Temel deðer kontrolleri
        '------------------------------------------------------------------

        If YearValue < 1900 Or YearValue > 2100 Then

            TryParseDate = False

            Exit Function

        End If


        If MonthValue < 1 Or MonthValue > 12 Then

            TryParseDate = False

            Exit Function

        End If


        If DayValue < 1 Or DayValue > 31 Then

            TryParseDate = False

            Exit Function

        End If


        '------------------------------------------------------------------
        ' Gerçek tarihi oluþtur.
        '----------------------------------------------------------------------

        TestDate = DateSerial( _
            YearValue, _
            MonthValue, _
            DayValue)


        '------------------------------------------------------------------
        ' DateSerial hatalý günleri otomatik ileri taþýyabilir.
        '
        ' Örneðin:
        ' 31.02.1960
        '
        ' Bu nedenle oluþan tarihi tekrar kontrol ediyoruz.
        '----------------------------------------------------------------------

        If Day(TestDate) <> DayValue Then

            TryParseDate = False

            Exit Function

        End If


        If Month(TestDate) <> MonthValue Then

            TryParseDate = False

            Exit Function

        End If


        If Year(TestDate) <> YearValue Then

            TryParseDate = False

            Exit Function

        End If


        ResultDate = TestDate

        TryParseDate = True

        Exit Function

    End If


    '======================================================================
    ' NORMAL TARÝH BÝÇÝMLERÝ
    '
    ' 02.01.1960
    ' 02/01/1960
    ' 02-01-1960
    '======================================================================

    If IsDate(CleanValue) Then

        ResultDate = CDate(CleanValue)

        TryParseDate = True

        Exit Function

    End If


    '----------------------------------------------------------------------
    ' Geçersiz tarih
    '----------------------------------------------------------------------

    TryParseDate = False

    Exit Function


ErrorHandler:

    TryParseDate = False

End Function


