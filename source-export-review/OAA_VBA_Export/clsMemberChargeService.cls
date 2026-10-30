VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMemberChargeService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' clsMemberChargeService
'
' WF-003 - Üye Borçlandýrma Ýþ Kurallarý
'
' Repository:
'   clsMemberChargeRepository
'
' Entity:
'   clsMemberCharge
'
' Üye:
'   clsMemberRepository
'
' ÖNEMLÝ:
'   Bu sýnýf doðrudan Excel hücrelerine yazmaz.
'   Tüm veri eriþimi Repository üzerinden yapýlýr.
'==============================================================================


Private mChargeRepository As clsMemberChargeRepository
Private mMemberRepository As clsMemberRepository


'==============================================================================
' Class Initialize
'==============================================================================

Private Sub Class_Initialize()

    Set mChargeRepository = New clsMemberChargeRepository
    Set mMemberRepository = New clsMemberRepository

End Sub


'==============================================================================
' AddOfficialCharge
'
' RESMÝ AÝDAT tahakkuku oluþturur.
'
' Kurallar:
'
' 1. Aktif üye:
'       Tahakkuk yapýlabilir.
'
' 2. Ýzinli üye:
'       RESMÝ aidat tahakkuku yapýlabilir.
'
' 3. Ýstifa üyesi:
'       Ýstifa tarihinden dönem sonuna kadar tahakkuk yapýlabilir.
'
' 4. Dönem sonundan sonra istifa/pasif üyeye yeni resmi aidat verilmez.
'
' 5. Ayný üye + mali yýl + ay + RESMI için
'       ikinci NORMAL tahakkuk oluþturulamaz.
'
' periodEndDate:
'       Ýlgili mali dönemin son günü.
'==============================================================================

Public Function AddOfficialCharge( _
    ByVal MemberID As Long, _
    ByVal FiscalYear As Long, _
    ByVal DueMonth As Long, _
    ByVal ChargeDate As Date, _
    ByVal Amount As Currency, _
    ByVal PeriodEndDate As Date, _
    Optional ByVal DecisionNo As String = vbNullString, _
    Optional ByVal DecisionDate As Date = 0, _
    Optional ByVal Description As String = vbNullString) As Long

    Dim Member As clsMember
    Dim Charge As clsMemberCharge

    On Error GoTo ErrorHandler

    '----------------------------------------------------------------------
    ' Temel kontroller
    '----------------------------------------------------------------------

    ValidateBasicChargeParameters _
        MemberID, _
        FiscalYear, _
        DueMonth, _
        ChargeDate, _
        Amount, _
        PeriodEndDate

    '----------------------------------------------------------------------
    ' Üyeyi getir
    '----------------------------------------------------------------------

    Set Member = mMemberRepository.GetByID(MemberID)

    If Member Is Nothing Then

        Err.Raise vbObjectError + 7201, _
                  "clsMemberChargeService.AddOfficialCharge", _
                  "Üye bulunamadý. Üye No: " & MemberID

    End If

    '----------------------------------------------------------------------
    ' Üyenin resmi aidata uygunluðu
    '----------------------------------------------------------------------

    If Not CanReceiveOfficialCharge( _
        Member, _
        ChargeDate, _
        PeriodEndDate) Then

        Err.Raise vbObjectError + 7202, _
                  "clsMemberChargeService.AddOfficialCharge", _
                  "Üye bu dönem için resmi aidata borçlandýrýlamaz."

    End If

    '----------------------------------------------------------------------
    ' Mükerrer tahakkuk kontrolü
    '----------------------------------------------------------------------

    If mChargeRepository.ExistsNormalCharge( _
        MemberID, _
        FiscalYear, _
        DueMonth, _
        CHARGE_TYPE_OFFICIAL) Then

        Err.Raise vbObjectError + 7203, _
                  "clsMemberChargeService.AddOfficialCharge", _
                  "Bu üye için ayný döneme ait resmi aidat zaten tahakkuk edilmiþ."

    End If

    '----------------------------------------------------------------------
    ' Entity oluþtur
    '----------------------------------------------------------------------

    Set Charge = New clsMemberCharge

    Charge.MemberID = MemberID
    Charge.ChargeType = CHARGE_TYPE_OFFICIAL
    Charge.ChargeDate = ChargeDate
    Charge.FiscalYear = FiscalYear
    Charge.DueMonth = DueMonth
    Charge.Amount = Amount
    Charge.AdjustmentType = ADJUSTMENT_TYPE_NORMAL
    Charge.SourceChargeID = 0
    Charge.Description = Description
    Charge.DecisionNo = DecisionNo
    Charge.DecisionDate = DecisionDate
    Charge.EntryDate = Now
    Charge.Status = CHARGE_STATUS_ACTIVE

    '----------------------------------------------------------------------
    ' Repository'ye kaydet
    '----------------------------------------------------------------------

    AddOfficialCharge = mChargeRepository.Add(Charge)

    Exit Function

ErrorHandler:

    AddOfficialCharge = 0

    Err.Raise Err.Number, _
              "clsMemberChargeService.AddOfficialCharge", _
              Err.Description

End Function


'==============================================================================
' AddMeetingCharge
'
' TOPLANTI AÝDATI tahakkuku oluþturur.
'
' Kural:
'   Ýzinli üye toplantý aidatýna katýlmaz.
'
' Bu nedenle:
'   AKTÝF  -> uygulanýr
'   ÝZÝNLÝ -> reddedilir
'   ÝSTÝFA -> reddedilir
'
' Ayný üye + dönem + TOPLANTI için ikinci NORMAL kayýt oluþturulamaz.
'==============================================================================

Public Function AddMeetingCharge( _
    ByVal MemberID As Long, _
    ByVal FiscalYear As Long, _
    ByVal DueMonth As Long, _
    ByVal ChargeDate As Date, _
    ByVal Amount As Currency, _
    Optional ByVal DecisionNo As String = vbNullString, _
    Optional ByVal DecisionDate As Date = 0, _
    Optional ByVal Description As String = vbNullString) As Long

    Dim Member As clsMember
    Dim Charge As clsMemberCharge

    On Error GoTo ErrorHandler

    ValidateBasicChargeParameters _
        MemberID, _
        FiscalYear, _
        DueMonth, _
        ChargeDate, _
        Amount, _
        DateSerial(FiscalYear, 12, 31)

    Set Member = mMemberRepository.GetByID(MemberID)

    If Member Is Nothing Then

        Err.Raise vbObjectError + 7211, _
                  "clsMemberChargeService.AddMeetingCharge", _
                  "Üye bulunamadý. Üye No: " & MemberID

    End If

    '----------------------------------------------------------------------
    ' Toplantý aidatý yalnýzca aktif üyeye
    '----------------------------------------------------------------------

    If StrComp(Trim$(Member.Status), MEMBER_STATUS_ACTIVE, vbTextCompare) <> 0 Then

        Err.Raise vbObjectError + 7212, _
                  "clsMemberChargeService.AddMeetingCharge", _
                  "Ýzinli veya istifa durumundaki üyeye toplantý aidatý uygulanamaz."

    End If

    '----------------------------------------------------------------------
    ' Mükerrer kontrol
    '----------------------------------------------------------------------

    If mChargeRepository.ExistsNormalCharge( _
        MemberID, _
        FiscalYear, _
        DueMonth, _
        CHARGE_TYPE_MEETING) Then

        Err.Raise vbObjectError + 7213, _
                  "clsMemberChargeService.AddMeetingCharge", _
                  "Bu üye için ayný döneme ait toplantý aidatý zaten tahakkuk edilmiþ."

    End If

    '----------------------------------------------------------------------
    ' Entity
    '----------------------------------------------------------------------

    Set Charge = New clsMemberCharge

    Charge.MemberID = MemberID
    Charge.ChargeType = CHARGE_TYPE_MEETING
    Charge.ChargeDate = ChargeDate
    Charge.FiscalYear = FiscalYear
    Charge.DueMonth = DueMonth
    Charge.Amount = Amount
    Charge.AdjustmentType = ADJUSTMENT_TYPE_NORMAL
    Charge.SourceChargeID = 0
    Charge.Description = Description
    Charge.DecisionNo = DecisionNo
    Charge.DecisionDate = DecisionDate
    Charge.EntryDate = Now
    Charge.Status = CHARGE_STATUS_ACTIVE

    AddMeetingCharge = mChargeRepository.Add(Charge)

    Exit Function

ErrorHandler:

    AddMeetingCharge = 0

    Err.Raise Err.Number, _
              "clsMemberChargeService.AddMeetingCharge", _
              Err.Description

End Function


'==============================================================================
' AddOfficialChargeAdjustment
'
' Mevcut RESMI aidat tahakkukuna sonradan yapýlan ARTIS kaydýný oluþturur.
'
' ÖNEMLÝ:
'   Orijinal tahakkuk deðiþtirilmez.
'
'   Örneðin:
'
'       Orijinal aidat       : 500 TL
'       Sonradan artýþ        : 200 TL
'       Yeni toplam          : 700 TL
'
'   Sisteme:
'
'       NORMAL kayýt          : 500 TL
'       ARTIS kayýt           : 200 TL
'
'   olarak iki ayrý hareket yazýlýr.
'
' SourceChargeID:
'   Artýþ kaydýnýn hangi orijinal tahakkuktan kaynaklandýðýný gösterir.
'==============================================================================

Public Function AddOfficialChargeAdjustment( _
    ByVal SourceChargeID As Long, _
    ByVal AdjustmentDate As Date, _
    ByVal AdjustmentAmount As Currency, _
    Optional ByVal DecisionNo As String = vbNullString, _
    Optional ByVal DecisionDate As Date = 0, _
    Optional ByVal Description As String = vbNullString) As Long

    Dim SourceCharge As clsMemberCharge
    Dim AdjustmentCharge As clsMemberCharge
    Dim Member As clsMember

    On Error GoTo ErrorHandler

    If SourceChargeID <= 0 Then
        Err.Raise vbObjectError + 7241, "clsMemberChargeService.AddOfficialChargeAdjustment", "Geçersiz kaynak tahakkuk numarasý."
    End If

    If AdjustmentDate <= 0 Then
        Err.Raise vbObjectError + 7242, "clsMemberChargeService.AddOfficialChargeAdjustment", "Artýþ tarihi girilmelidir."
    End If

    If AdjustmentAmount <= 0 Then
        Err.Raise vbObjectError + 7243, "clsMemberChargeService.AddOfficialChargeAdjustment", "Artýþ tutarý sýfýrdan büyük olmalýdýr."
    End If

    Set SourceCharge = mChargeRepository.GetByID(SourceChargeID)

    If SourceCharge Is Nothing Then
        Err.Raise vbObjectError + 7244, "clsMemberChargeService.AddOfficialChargeAdjustment", "Kaynak tahakkuk bulunamadý. ChargeID: " & SourceChargeID
    End If

    If StrComp(Trim$(SourceCharge.ChargeType), CHARGE_TYPE_OFFICIAL, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 7245, "clsMemberChargeService.AddOfficialChargeAdjustment", "Artýþ yalnýzca resmi aidat tahakkuklarýna uygulanabilir."
    End If

    If StrComp(Trim$(SourceCharge.AdjustmentType), ADJUSTMENT_TYPE_NORMAL, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 7246, "clsMemberChargeService.AddOfficialChargeAdjustment", "Artýþ kaynaðý NORMAL bir resmi aidat tahakkuku olmalýdýr."
    End If

    If StrComp(Trim$(SourceCharge.Status), CHARGE_STATUS_ACTIVE, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 7247, "clsMemberChargeService.AddOfficialChargeAdjustment", "Pasif bir tahakkuka artýþ uygulanamaz."
    End If

    Set Member = mMemberRepository.GetByID(SourceCharge.MemberID)

    If Member Is Nothing Then
        Err.Raise vbObjectError + 7248, "clsMemberChargeService.AddOfficialChargeAdjustment", "Kaynak tahakkukun üyesi bulunamadý. Üye No: " & SourceCharge.MemberID
    End If

    If Not CanReceiveOfficialCharge(Member, AdjustmentDate, DateSerial(SourceCharge.FiscalYear, 12, 31)) Then
        Err.Raise vbObjectError + 7249, "clsMemberChargeService.AddOfficialChargeAdjustment", "Üye bu tarihte resmi aidat artýþýna uygun deðildir."
    End If

    Set AdjustmentCharge = New clsMemberCharge
    AdjustmentCharge.MemberID = SourceCharge.MemberID
    AdjustmentCharge.ChargeType = CHARGE_TYPE_OFFICIAL
    AdjustmentCharge.ChargeDate = AdjustmentDate
    AdjustmentCharge.FiscalYear = SourceCharge.FiscalYear
    AdjustmentCharge.DueMonth = SourceCharge.DueMonth
    AdjustmentCharge.Amount = AdjustmentAmount
    AdjustmentCharge.AdjustmentType = ADJUSTMENT_TYPE_INCREASE
    AdjustmentCharge.SourceChargeID = SourceChargeID
    AdjustmentCharge.Description = Description
    AdjustmentCharge.DecisionNo = DecisionNo
    AdjustmentCharge.DecisionDate = DecisionDate
    AdjustmentCharge.EntryDate = Now
    AdjustmentCharge.Status = CHARGE_STATUS_ACTIVE

    AddOfficialChargeAdjustment = mChargeRepository.Add(AdjustmentCharge)
    Exit Function

ErrorHandler:
    AddOfficialChargeAdjustment = 0
    Err.Raise Err.Number, "clsMemberChargeService.AddOfficialChargeAdjustment", Err.Description
End Function


'==============================================================================
' AddBulkOfficialCharges
'
' Seçilen mali yýl ve ay için toplu RESMI aidat tahakkuku oluþturur.
'
' Kurallar:
'
'   AKTÝF üye
'       -> RESMI aidat oluþturulur.
'
'   ÝZÝNLÝ üye
'       -> RESMI aidat oluþturulur.
'
'   ÝSTÝFA üye
'       -> AddOfficialCharge içindeki dönem/istifa kurallarý uygulanýr.
'
'   Dönem sonundan sonra tahakkuk uygun deðilse
'       -> kayýt oluþturulmaz.
'
'   Ayný ay için NORMAL RESMI tahakkuk zaten varsa
'       -> tekrar kayýt oluþturulmaz.
'
' Dönen deðer:
'   Baþarýyla oluþturulan tahakkuk sayýsý.
'
' ByRef:
'   CreatedCount -> oluþturulan kayýt sayýsý
'   SkippedCount -> tahakkuk oluþturulmayan üye sayýsý
'   ErrorCount   -> beklenmeyen hata sayýsý
'==============================================================================

Public Function AddBulkOfficialCharges( _
    ByVal FiscalYear As Long, _
    ByVal DueMonth As Long, _
    ByVal ChargeDate As Date, _
    ByVal Amount As Currency, _
    ByVal PeriodEndDate As Date, _
    Optional ByVal DecisionNo As String = vbNullString, _
    Optional ByVal DecisionDate As Date = 0, _
    Optional ByVal Description As String = vbNullString, _
    Optional ByRef CreatedCount As Long = 0, _
    Optional ByRef SkippedCount As Long = 0, _
    Optional ByRef ErrorCount As Long = 0) As Long

    Dim Members As Collection
    Dim Member As clsMember

    Dim ChargeID As Long

    On Error GoTo FatalError

    '----------------------------------------------------------------------
    ' Baþlangýç deðerleri
    '----------------------------------------------------------------------

    CreatedCount = 0
    SkippedCount = 0
    ErrorCount = 0

    '----------------------------------------------------------------------
    ' Temel parametre kontrolleri
    '----------------------------------------------------------------------

    If FiscalYear <= 0 Then

        Err.Raise vbObjectError + 7351, _
                  "clsMemberChargeService.AddBulkOfficialCharges", _
                  "Geçersiz mali yýl."

    End If

    If DueMonth < 1 Or DueMonth > 12 Then

        Err.Raise vbObjectError + 7352, _
                  "clsMemberChargeService.AddBulkOfficialCharges", _
                  "Aidat ayý 1-12 arasýnda olmalýdýr."

    End If

    If ChargeDate <= 0 Then

        Err.Raise vbObjectError + 7353, _
                  "clsMemberChargeService.AddBulkOfficialCharges", _
                  "Tahakkuk tarihi girilmelidir."

    End If

    If Amount <= 0 Then

        Err.Raise vbObjectError + 7354, _
                  "clsMemberChargeService.AddBulkOfficialCharges", _
                  "Aidat tutarý sýfýrdan büyük olmalýdýr."

    End If

    If PeriodEndDate <= 0 Then

        Err.Raise vbObjectError + 7355, _
                  "clsMemberChargeService.AddBulkOfficialCharges", _
                  "Dönem sonu tarihi girilmelidir."

    End If

    '----------------------------------------------------------------------
    ' Üyeleri getir
    '----------------------------------------------------------------------

    Set Members = mMemberRepository.GetAll

    If Members Is Nothing Then

        Err.Raise vbObjectError + 7356, _
                  "clsMemberChargeService.AddBulkOfficialCharges", _
                  "Üye listesi alýnamadý."

    End If

    '----------------------------------------------------------------------
    ' TÜM ÜYELERÝ ÝÞLE
    '----------------------------------------------------------------------

    For Each Member In Members

        ChargeID = 0

        '------------------------------------------------------------------
        ' Ayný ay için NORMAL RESMI tahakkuk zaten varsa tekrar oluþturma.
        '------------------------------------------------------------------

        If mChargeRepository.ExistsNormalCharge( _
                Member.MemberID, _
                FiscalYear, _
                DueMonth, _
                CHARGE_TYPE_OFFICIAL) Then

            SkippedCount = SkippedCount + 1

        Else

            '--------------------------------------------------------------
            ' Her üyeyi mevcut AddOfficialCharge iþ kuralýndan geçir.
            '
            ' Böylece:
            '
            '   Aktif
            '   Ýzinli
            '   Ýstifa
            '
            ' kurallarý tek bir yerde uygulanmýþ olur.
            '--------------------------------------------------------------

            On Error Resume Next

            Err.Clear

            ChargeID = AddOfficialCharge( _
                            Member.MemberID, _
                            FiscalYear, _
                            DueMonth, _
                            ChargeDate, _
                            Amount, _
                            PeriodEndDate, _
                            DecisionNo, _
                            DecisionDate, _
                            Description)

            If Err.Number <> 0 Then

                '----------------------------------------------------------
                ' AddOfficialCharge'ýn reddettiði üye toplu iþlemden
                ' atlanýr.
                '
                ' Örneðin:
                '   Ýstifa etmiþ ve dönem sonrasýna çýkmýþ üye.
                '----------------------------------------------------------

                SkippedCount = SkippedCount + 1

                Err.Clear

            ElseIf ChargeID > 0 Then

                CreatedCount = CreatedCount + 1

            Else

                SkippedCount = SkippedCount + 1

            End If

            On Error GoTo FatalError

        End If

    Next Member

    '----------------------------------------------------------------------
    ' Fonksiyon sonucu
    '----------------------------------------------------------------------

    AddBulkOfficialCharges = CreatedCount

    Exit Function


'==============================================================================
' BEKLENMEYEN HATA
'==============================================================================

FatalError:

    ErrorCount = ErrorCount + 1

    AddBulkOfficialCharges = CreatedCount

    Err.Raise Err.Number, _
              "clsMemberChargeService.AddBulkOfficialCharges", _
              Err.Description

End Function


'==============================================================================
' GetMemberBalance
'
' Belirli bir üyenin aktif borç hareketlerinin toplamýný döndürür.
'
' Repository:
'   Üyenin tüm borç hareketlerini getirir.
'
' Service:
'   Ýþ kuralýný uygular ve aktif borçlarý toplar.
'==============================================================================

Public Function GetMemberBalance(ByVal MemberID As Long) As Currency

    Dim Charges As Collection
    Dim Charge As clsMemberCharge

    Dim Balance As Currency

    On Error GoTo ErrorHandler

    Balance = 0@

    If MemberID <= 0 Then
        GetMemberBalance = 0@
        Exit Function
    End If

    '----------------------------------------------------------------------
    ' Üyenin tüm borç hareketlerini Repository'den al
    '----------------------------------------------------------------------

    Set Charges = mChargeRepository.GetByMemberID(MemberID)

    If Charges Is Nothing Then

        GetMemberBalance = 0@
        Exit Function

    End If

    '----------------------------------------------------------------------
    ' Aktif borç hareketlerini topla
    '----------------------------------------------------------------------

    For Each Charge In Charges

        If Not Charge Is Nothing Then

            If StrComp( _
                Trim$(Charge.Status), _
                CHARGE_STATUS_ACTIVE, _
                vbTextCompare) = 0 Then

                Balance = Balance + Charge.Amount

            End If

        End If

    Next Charge

    GetMemberBalance = Balance

    Exit Function

ErrorHandler:

    GetMemberBalance = 0@

End Function


'==============================================================================
' AddOtherCharge
'
' SAÝR BORÇ oluþturur.
'
' Kullaným:
'   Toplu borç veya kýsmi borç olabilir.
'
' Amount:
'   Kullanýcýnýn girdiði tutardýr.
'
' Ýstifa eden üye:
'   Sair borcun ChargeDate deðeri istifa tarihinden önce veya
'   istifa tarihine eþit olmalýdýr.
'
' Ýstifa tarihinden sonra yeni sair borç oluþturulamaz.
'==============================================================================

Public Function AddOtherCharge( _
    ByVal MemberID As Long, _
    ByVal ChargeDate As Date, _
    ByVal Amount As Currency, _
    ByVal Description As String, _
    Optional ByVal DecisionNo As String = vbNullString, _
    Optional ByVal DecisionDate As Date = 0) As Long

    Dim Member As clsMember
    Dim Charge As clsMemberCharge

    On Error GoTo ErrorHandler

    '----------------------------------------------------------------------
    ' Temel kontroller
    '----------------------------------------------------------------------

    If MemberID <= 0 Then

        Err.Raise vbObjectError + 7221, _
                  "clsMemberChargeService.AddOtherCharge", _
                  "Geçersiz Üye No."

    End If

    If ChargeDate <= 0 Then

        Err.Raise vbObjectError + 7222, _
                  "clsMemberChargeService.AddOtherCharge", _
                  "Borç tarihi girilmelidir."

    End If

    If Amount <= 0 Then

        Err.Raise vbObjectError + 7223, _
                  "clsMemberChargeService.AddOtherCharge", _
                  "Sair borç tutarý sýfýrdan büyük olmalýdýr."

    End If

    If Len(Trim$(Description)) = 0 Then

        Err.Raise vbObjectError + 7224, _
                  "clsMemberChargeService.AddOtherCharge", _
                  "Sair borç açýklamasý girilmelidir."

    End If

    '----------------------------------------------------------------------
    ' Üye
    '----------------------------------------------------------------------

    Set Member = mMemberRepository.GetByID(MemberID)

    If Member Is Nothing Then

        Err.Raise vbObjectError + 7225, _
                  "clsMemberChargeService.AddOtherCharge", _
                  "Üye bulunamadý. Üye No: " & MemberID

    End If

    '----------------------------------------------------------------------
    ' Ýstifa kontrolü
    '
    ' Ýstifa tarihinden sonra yeni sair borç oluþturulamaz.
    '----------------------------------------------------------------------

    If StrComp(Trim$(Member.Status), MEMBER_STATUS_RESIGNED, vbTextCompare) = 0 Then

        If Member.LeaveDate <= 0 Then

            Err.Raise vbObjectError + 7226, _
                      "clsMemberChargeService.AddOtherCharge", _
                      "Üyenin istifa tarihi bulunamadý."

        End If

        If ChargeDate > Member.LeaveDate Then

            Err.Raise vbObjectError + 7227, _
                      "clsMemberChargeService.AddOtherCharge", _
                      "Ýstifa tarihinden sonra yeni sair borç oluþturulamaz."

        End If

    End If

    '----------------------------------------------------------------------
    ' Entity
    '----------------------------------------------------------------------

    Set Charge = New clsMemberCharge

    Charge.MemberID = MemberID
    Charge.ChargeType = CHARGE_TYPE_OTHER
    Charge.ChargeDate = ChargeDate
    Charge.FiscalYear = Year(ChargeDate)
    Charge.DueMonth = Month(ChargeDate)
    Charge.Amount = Amount
    Charge.AdjustmentType = ADJUSTMENT_TYPE_NORMAL
    Charge.SourceChargeID = 0
    Charge.Description = Description
    Charge.DecisionNo = DecisionNo
    Charge.DecisionDate = DecisionDate
    Charge.EntryDate = Now
    Charge.Status = CHARGE_STATUS_ACTIVE

    AddOtherCharge = mChargeRepository.Add(Charge)

    Exit Function

ErrorHandler:

    AddOtherCharge = 0

    Err.Raise Err.Number, _
              "clsMemberChargeService.AddOtherCharge", _
              Err.Description

End Function


'==============================================================================
' CanReceiveOfficialCharge
'
' RESMI aidat için üye durumunu kontrol eder.
'
' Kurallar:
'
'   AKTÝF:
'       RESMI aidat alabilir.
'
'   ÝZÝNLÝ:
'       RESMI aidat alabilir.
'
'   ÝSTÝFA:
'       Yalnýzca istifa tarihine kadar RESMI aidat alabilir.
'
'       ChargeDate <= LeaveDate
'
'       Ýstifa tarihinden sonraki hiçbir tarihe yeni RESMI aidat
'       oluþturulamaz.
'==============================================================================

Private Function CanReceiveOfficialCharge( _
    ByVal Member As clsMember, _
    ByVal ChargeDate As Date, _
    ByVal PeriodEndDate As Date) As Boolean

    Dim MemberStatus As String

    CanReceiveOfficialCharge = False

    If Member Is Nothing Then Exit Function

    If ChargeDate <= 0 Then Exit Function

    MemberStatus = Trim$(Member.Status)


    '----------------------------------------------------------------------
    ' AKTÝF ÜYE
    '----------------------------------------------------------------------

    If StrComp( _
        MemberStatus, _
        MEMBER_STATUS_ACTIVE, _
        vbTextCompare) = 0 Then

        CanReceiveOfficialCharge = True
        Exit Function

    End If


    '----------------------------------------------------------------------
    ' ÝZÝNLÝ ÜYE
    '
    ' Ýzinli üye resmi aidattan muaf deðildir.
    '----------------------------------------------------------------------

    If StrComp( _
        MemberStatus, _
        MEMBER_STATUS_LEAVE, _
        vbTextCompare) = 0 Then

        CanReceiveOfficialCharge = True
        Exit Function

    End If


    '----------------------------------------------------------------------
    ' ÝSTÝFA EDEN ÜYE
    '
    ' SADECE ÝSTÝFA TARÝHÝNE KADAR
    '----------------------------------------------------------------------

    If StrComp( _
        MemberStatus, _
        MEMBER_STATUS_RESIGNED, _
        vbTextCompare) = 0 Then

        If Member.LeaveDate <= 0 Then

            CanReceiveOfficialCharge = False
            Exit Function

        End If


        '--------------------------------------------------------------
        ' Ýstifa tarihi dahil
        '--------------------------------------------------------------

        If ChargeDate <= Member.LeaveDate Then

            CanReceiveOfficialCharge = True

        Else

            CanReceiveOfficialCharge = False

        End If

        Exit Function

    End If

End Function

'==============================================================================
' ValidateBasicChargeParameters
'
' Resmi ve toplantý aidatý tahakkuklarýnda kullanýlan temel parametre
' kontrolleri.
'==============================================================================

Private Sub ValidateBasicChargeParameters( _
    ByVal MemberID As Long, _
    ByVal FiscalYear As Long, _
    ByVal DueMonth As Long, _
    ByVal ChargeDate As Date, _
    ByVal Amount As Currency, _
    ByVal PeriodEndDate As Date)

    '----------------------------------------------------------------------
    ' Üye
    '----------------------------------------------------------------------

    If MemberID <= 0 Then

        Err.Raise vbObjectError + 7231, _
                  "clsMemberChargeService", _
                  "Geçersiz Üye No."

    End If


    '----------------------------------------------------------------------
    ' Mali yýl
    '----------------------------------------------------------------------

    If FiscalYear <= 0 Then

        Err.Raise vbObjectError + 7232, _
                  "clsMemberChargeService", _
                  "Geçersiz mali yýl."

    End If


    '----------------------------------------------------------------------
    ' Aidat ayý
    '----------------------------------------------------------------------

    If DueMonth < 1 Or DueMonth > 12 Then

        Err.Raise vbObjectError + 7233, _
                  "clsMemberChargeService", _
                  "Aidat ayý 1 ile 12 arasýnda olmalýdýr."

    End If


    '----------------------------------------------------------------------
    ' Tahakkuk tarihi
    '----------------------------------------------------------------------

    If ChargeDate <= 0 Then

        Err.Raise vbObjectError + 7234, _
                  "clsMemberChargeService", _
                  "Tahakkuk tarihi girilmelidir."

    End If


    '----------------------------------------------------------------------
    ' Tutar
    '----------------------------------------------------------------------

    If Amount <= 0 Then

        Err.Raise vbObjectError + 7235, _
                  "clsMemberChargeService", _
                  "Borç tutarý sýfýrdan büyük olmalýdýr."

    End If


    '----------------------------------------------------------------------
    ' Dönem sonu
    '----------------------------------------------------------------------

    If PeriodEndDate <= 0 Then

        Err.Raise vbObjectError + 7236, _
                  "clsMemberChargeService", _
                  "Dönem sonu tarihi girilmelidir."

    End If


    '----------------------------------------------------------------------
    ' Tarih sýralamasý
    '----------------------------------------------------------------------

    If PeriodEndDate < ChargeDate Then

        Err.Raise vbObjectError + 7237, _
                  "clsMemberChargeService", _
                  "Dönem sonu tarihi tahakkuk tarihinden önce olamaz."

    End If

End Sub
