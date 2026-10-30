VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReceiptService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Sýnýf   : clsReceiptService
' Amaç    : Tahsilat / Makbuz iþ kurallarý
'
' MÝMARÝ
'
'       frmReceipts
'            |
'            v
'       clsReceiptService
'            |
'       +----+---------------------------+
'       |            |                   |
'       v            v                   v
' clsMemberRepo  clsChargeRepo    clsReceiptRepository
'                                      |
'                                      v
'                                 TBL_RECEIPTS
'
'            |
'            v
'       clsLedgerService
'            |
'            v
'        TBL_LEDGER
'
' ÖNEMLÝ:
'
'   UserForm Repository'lere doðrudan eriþmez.
'
'   Kullanýcý:
'       MemberID görmez.
'       ChargeID görmez.
'
'   Üye adý üzerinden seçim yapýlýr.
'
'   Gerçekleþen tahsilat:
'
'       Receipt
'           +
'       Ledger gelir hareketi
'
'   olarak kaydedilir.
'
'==============================================================================


Private mReceiptRepository As clsReceiptRepository
Private mMemberRepository As clsMemberRepository
Private mChargeRepository As clsMemberChargeRepository
Private mLedgerService As clsLedgerService


'==============================================================================
' CLASS INITIALIZE
'==============================================================================

Private Sub Class_Initialize()

    Set mReceiptRepository = New clsReceiptRepository
    Set mMemberRepository = New clsMemberRepository
    Set mChargeRepository = New clsMemberChargeRepository
    Set mLedgerService = New clsLedgerService

End Sub


'==============================================================================
' CLASS TERMINATE
'==============================================================================

Private Sub Class_Terminate()

    Set mReceiptRepository = Nothing
    Set mMemberRepository = Nothing
    Set mChargeRepository = Nothing
    Set mLedgerService = Nothing

End Sub


'==============================================================================
' GET MEMBERS
'==============================================================================

Public Function GetMembers() As Collection

    On Error GoTo ErrorHandler

    Set GetMembers = _
        mMemberRepository.GetAll

    Exit Function

ErrorHandler:

    Set GetMembers = Nothing

End Function


'==============================================================================
' ADD RECEIPT
'
' Ýþ akýþý:
'
'   1. Tahsilat doðrulamalarý
'   2. Receipt kaydý
'   3. Ledger Entity oluþturma
'   4. Ledger kaydý
'
' Receipt kaydý baþarýlý olduktan sonra Ledger kaydý baþarýsýz olursa,
' bu iþlemde oluþturulan Receipt geri alýnýr.
'
'==============================================================================

Public Function AddReceipt( _
    ByVal Receipt As clsReceipt) As Long

    Dim Member As clsMember
    Dim Charge As clsMemberCharge
    Dim LedgerEntry As clsLedgerEntry

    Dim OutstandingAmount As Currency
    Dim LedgerBalance As Currency

    Dim ReceiptID As Long
    Dim LedgerCreated As Boolean

    Dim ErrorNumber As Long
    Dim ErrorDescription As String
    Dim ErrorSource As String

    On Error GoTo ErrorHandler

    AddReceipt = 0
    LedgerCreated = False


    '--------------------------------------------------------------------------
    ' ENTITY
    '--------------------------------------------------------------------------

    If Receipt Is Nothing Then

        Err.Raise _
            vbObjectError + 8101, _
            "clsReceiptService.AddReceipt", _
            "Makbuz nesnesi boþ olamaz."

    End If


    '--------------------------------------------------------------------------
    ' MEMBER ID
    '--------------------------------------------------------------------------

    If Receipt.MemberID <= 0 Then

        Err.Raise _
            vbObjectError + 8102, _
            "clsReceiptService.AddReceipt", _
            "Geçersiz Üye No."

    End If


    '--------------------------------------------------------------------------
    ' ÜYE
    '--------------------------------------------------------------------------

    Set Member = _
        mMemberRepository.GetByID(Receipt.MemberID)

    If Member Is Nothing Then

        Err.Raise _
            vbObjectError + 8103, _
            "clsReceiptService.AddReceipt", _
            "Üye bulunamadý. Üye No: " & _
            Receipt.MemberID

    End If


    '--------------------------------------------------------------------------
    ' CHARGE ID
    '--------------------------------------------------------------------------

    If Receipt.ChargeID <= 0 Then

        Err.Raise _
            vbObjectError + 8104, _
            "clsReceiptService.AddReceipt", _
            "Geçerli bir ChargeID girilmelidir."

    End If


    '--------------------------------------------------------------------------
    ' BORÇ
    '--------------------------------------------------------------------------

    Set Charge = _
        mChargeRepository.GetByID(Receipt.ChargeID)

    If Charge Is Nothing Then

        Err.Raise _
            vbObjectError + 8105, _
            "clsReceiptService.AddReceipt", _
            "Borç kaydý bulunamadý. ChargeID: " & _
            Receipt.ChargeID

    End If


    '--------------------------------------------------------------------------
    ' ÜYE - BORÇ ÝLÝÞKÝSÝ
    '--------------------------------------------------------------------------

    If Charge.MemberID <> Receipt.MemberID Then

        Err.Raise _
            vbObjectError + 8106, _
            "clsReceiptService.AddReceipt", _
            "Tahsilatýn üyesi ile borç kaydýnýn üyesi farklý."

    End If


    '--------------------------------------------------------------------------
    ' BORÇ DURUMU
    '--------------------------------------------------------------------------

    If StrComp( _
        Trim$(Charge.Status), _
        CHARGE_STATUS_ACTIVE, _
        vbTextCompare) <> 0 Then

        Err.Raise _
            vbObjectError + 8107, _
            "clsReceiptService.AddReceipt", _
            "Pasif bir borç için tahsilat yapýlamaz."

    End If


    '--------------------------------------------------------------------------
    ' MAKBUZ SERÝSÝ
    '--------------------------------------------------------------------------

    If Len(Trim$(Receipt.Series)) = 0 Then

        Err.Raise _
            vbObjectError + 8108, _
            "clsReceiptService.AddReceipt", _
            "Makbuz serisi boþ býrakýlamaz."

    End If


    '--------------------------------------------------------------------------
    ' MAKBUZ NUMARASI
    '--------------------------------------------------------------------------

    If Receipt.ReceiptNumber <= 0 Then

    Receipt.ReceiptNumber = _
        mReceiptRepository.GetNextReceiptNumber

    If Receipt.ReceiptNumber <= 0 Then

        Err.Raise _
            vbObjectError + 8109, _
            "clsReceiptService.AddReceipt", _
            "Makbuz numarasý oluþturulamadý."

    End If

End If

    '--------------------------------------------------------------------------
    ' TARÝH
    '--------------------------------------------------------------------------

    If Receipt.ReceiptDate <= 0 Then

        Err.Raise _
            vbObjectError + 8110, _
            "clsReceiptService.AddReceipt", _
            "Makbuz tarihi girilmelidir."

    End If


    '--------------------------------------------------------------------------
    ' TUTAR
    '--------------------------------------------------------------------------

    If Receipt.Amount <= 0 Then

        Err.Raise _
            vbObjectError + 8111, _
            "clsReceiptService.AddReceipt", _
            "Tahsilat tutarý sýfýrdan büyük olmalýdýr."

    End If


    '--------------------------------------------------------------------------
    ' ÖDEME YÖNTEMÝ
    '--------------------------------------------------------------------------

    If Len(Trim$(Receipt.PaymentMethod)) = 0 Then

        Err.Raise _
            vbObjectError + 8112, _
            "clsReceiptService.AddReceipt", _
            "Ödeme yöntemi girilmelidir."

    End If


    '--------------------------------------------------------------------------
    ' SEÇÝLEN BORCUN KALANI
    '--------------------------------------------------------------------------

    OutstandingAmount = _
        GetChargeOutstandingAmount( _
            Receipt.ChargeID)


    If OutstandingAmount <= 0 Then

        Err.Raise _
            vbObjectError + 8113, _
            "clsReceiptService.AddReceipt", _
            "Bu borcun tahsil edilecek bakiyesi kalmamýþtýr."

    End If


    '--------------------------------------------------------------------------
    ' FAZLA TAHSÝLAT
    '--------------------------------------------------------------------------

    If Receipt.Amount > OutstandingAmount Then

        Err.Raise _
            vbObjectError + 8114, _
            "clsReceiptService.AddReceipt", _
            "Tahsilat tutarý kalan borçtan büyük olamaz." & _
            vbCrLf & _
            "Kalan borç: " & _
            Format$(OutstandingAmount, "#,##0.00")

    End If


    '--------------------------------------------------------------------------
    ' RECEIPT REPOSITORY
    '--------------------------------------------------------------------------

    ReceiptID = _
        mReceiptRepository.Add(Receipt)


    If ReceiptID <= 0 Then

        Err.Raise _
            vbObjectError + 8115, _
            "clsReceiptService.AddReceipt", _
            "Makbuz kaydý oluþturulamadý."

    End If


    '==========================================================================
    ' LEDGER ENTITY
    '==========================================================================

    Set LedgerEntry = New clsLedgerEntry


    '--------------------------------------------------------------------------
    ' TARÝH
    '--------------------------------------------------------------------------

    LedgerEntry.EntryDate = _
        Receipt.ReceiptDate


    '--------------------------------------------------------------------------
    ' BELGE NO
    '--------------------------------------------------------------------------

    LedgerEntry.DocumentNo = _
        Trim$(Receipt.Series) & "-" & _
        CStr(Receipt.ReceiptNumber)


    '--------------------------------------------------------------------------
    ' MAKBUZ / FATURA
    '--------------------------------------------------------------------------

    LedgerEntry.ReceiptInvoice = _
        Trim$(Receipt.Series) & "-" & _
        CStr(Receipt.ReceiptNumber)


    '--------------------------------------------------------------------------
    ' AÇIKLAMA
    '
    ' clsMember içinde doðrulanmýþ FirstName / LastName alanlarý kullanýlýr.
    ' FullName gibi varsayýmsal bir property kullanýlmaz.
    '--------------------------------------------------------------------------

    LedgerEntry.Description = _
        Trim$(Member.FirstName) & " " & _
        Trim$(Member.LastName) & _
        " - " & _
        Trim$(Charge.ChargeType)


    '--------------------------------------------------------------------------
    ' GELÝR TÜRÜ
    '--------------------------------------------------------------------------

    LedgerEntry.IncomeType = _
        Trim$(Charge.ChargeType) & " Tahsilatý"


    '--------------------------------------------------------------------------
    ' GÝDER TÜRÜ
    '--------------------------------------------------------------------------

    LedgerEntry.ExpenseType = _
        vbNullString


    '--------------------------------------------------------------------------
    ' GELÝR
    '--------------------------------------------------------------------------

    LedgerEntry.IncomeAmount = _
        Receipt.Amount


    '--------------------------------------------------------------------------
    ' GÝDER
    '--------------------------------------------------------------------------

    LedgerEntry.ExpenseAmount = _
        0@


    '--------------------------------------------------------------------------
    ' KASA / BANKA
    '
    ' Formdaki ödeme yöntemleri:
    '
    '   Nakit
    '   Banka
    '   Kredi Kartý
    '
    ' Nakit       -> CashAmount
    ' Banka       -> BankAmount
    ' Kredi Kartý -> BankAmount
    '
    '--------------------------------------------------------------------------

    LedgerEntry.CashAmount = 0@
    LedgerEntry.BankAmount = 0@


   Select Case True

    Case StrComp( _
        Trim$(Receipt.PaymentMethod), _
        "Nakit", _
        vbTextCompare) = 0

        LedgerEntry.CashAmount = _
            Receipt.Amount


    Case StrComp( _
        Trim$(Receipt.PaymentMethod), _
        "Banka", _
        vbTextCompare) = 0

        LedgerEntry.BankAmount = _
            Receipt.Amount


    Case StrComp( _
        Trim$(Receipt.PaymentMethod), _
        "Kredi Kartý", _
        vbTextCompare) = 0

        LedgerEntry.BankAmount = _
            Receipt.Amount


    Case Else

        Err.Raise _
            vbObjectError + 8116, _
            "clsReceiptService.AddReceipt", _
            "Geçersiz ödeme yöntemi: " & _
            Receipt.PaymentMethod

End Select


    '==========================================================================
    ' LEDGER BAKÝYESÝ
    '==========================================================================

    LedgerBalance = _
        mLedgerService.GetCurrentBalance


    '--------------------------------------------------------------------------
    ' TAHSÝLAT GELÝR OLDUÐU ÝÇÝN BAKÝYE ARTAR
    '--------------------------------------------------------------------------

    LedgerEntry.BalanceAmount = _
        LedgerBalance + Receipt.Amount


    '==========================================================================
    ' LEDGER'A KAYDET
    '==========================================================================

    If Not mLedgerService.AddEntry(LedgerEntry) Then

        Err.Raise _
            vbObjectError + 8117, _
            "clsReceiptService.AddReceipt", _
            "Tahsilat kaydedildi ancak Ýþletme Defteri kaydý oluþturulamadý."

    End If


    LedgerCreated = True


    '--------------------------------------------------------------------------
    ' BAÞARILI
    '--------------------------------------------------------------------------

    AddReceipt = ReceiptID

    Exit Function


'==============================================================================
' ERROR HANDLER
'==============================================================================

ErrorHandler:

    '--------------------------------------------------------------------------
    ' HATA BÝLGÝSÝNÝ CLEANUP ÖNCESÝ SAKLA
    '--------------------------------------------------------------------------

    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    ErrorSource = Err.Source


    '--------------------------------------------------------------------------
    ' RECEIPT OLUÞTU FAKAT LEDGER OLUÞMADIYSA ROLLBACK
    '
    ' Yalnýzca bu iþlemde oluþturulan Receipt silinir.
    '
    ' Mevcut gerçek tahsilat kayýtlarýna dokunulmaz.
    '--------------------------------------------------------------------------

    If ReceiptID > 0 Then

        If Not LedgerCreated Then

            On Error Resume Next

            mReceiptRepository.DeleteByID ReceiptID

            On Error GoTo 0

        End If

    End If


    '--------------------------------------------------------------------------
    ' SONUÇ
    '--------------------------------------------------------------------------

    AddReceipt = 0


    '--------------------------------------------------------------------------
    ' ORÝJÝNAL HATAYI ÜST KATMANA AKTAR
    '--------------------------------------------------------------------------

    Err.Raise _
        ErrorNumber, _
        ErrorSource, _
        ErrorDescription

End Function


'==============================================================================
' GET RECEIPT
'==============================================================================

Public Function GetReceipt( _
    ByVal ReceiptID As Long) As clsReceipt

    On Error GoTo ErrorHandler

    Set GetReceipt = Nothing

    If ReceiptID <= 0 Then Exit Function

    Set GetReceipt = _
        mReceiptRepository.GetByID(ReceiptID)

    Exit Function

ErrorHandler:

    Set GetReceipt = Nothing

End Function


'==============================================================================
' GET MEMBER RECEIPTS
'==============================================================================

Public Function GetMemberReceipts( _
    ByVal MemberID As Long) As Collection

    On Error GoTo ErrorHandler

    Set GetMemberReceipts = Nothing

    If MemberID <= 0 Then Exit Function

    Set GetMemberReceipts = _
        mReceiptRepository.GetByMemberID(MemberID)

    Exit Function

ErrorHandler:

    Set GetMemberReceipts = Nothing

End Function


'==============================================================================
' GET OPEN CHARGES FOR MEMBER
'==============================================================================

Public Function GetOpenChargesForMember( _
    ByVal MemberID As Long) As Collection

    Dim Result As Collection
    Dim Charges As Collection
    Dim Charge As clsMemberCharge

    Dim OutstandingAmount As Currency

    Dim I As Long

    On Error GoTo ErrorHandler

    Set Result = New Collection

    Set GetOpenChargesForMember = Result

    If MemberID <= 0 Then Exit Function


    Set Charges = _
        mChargeRepository.GetByMemberID(MemberID)

    If Charges Is Nothing Then Exit Function


    For I = 1 To Charges.Count

        Set Charge = Charges(I)

        If Not Charge Is Nothing Then

            If StrComp( _
                Trim$(Charge.Status), _
                CHARGE_STATUS_ACTIVE, _
                vbTextCompare) = 0 Then

                OutstandingAmount = _
                    GetChargeOutstandingAmount( _
                        Charge.ChargeID)

                If OutstandingAmount > 0 Then

                    Result.Add Charge

                End If

            End If

        End If

    Next I


    Set GetOpenChargesForMember = Result

    Exit Function


ErrorHandler:

    Set GetOpenChargesForMember = Nothing

End Function


'==============================================================================
' GET MEMBER OUTSTANDING BALANCE
'
' ÜYENÝN TOPLAM GERÇEK KALAN BORCU
'
' Formül:
'
'   Aktif borç hareketlerinin toplamý
'             -
'   Üyenin yaptýðý tahsilatlar
'
'   = Gerçek kalan borç
'
' Mevcut GetMemberBalance davranýþý deðiþtirilmez.
'==============================================================================

Public Function GetMemberOutstandingBalance( _
    ByVal MemberID As Long) As Currency

    Dim Charges As Collection
    Dim Receipts As Collection

    Dim Charge As clsMemberCharge
    Dim Receipt As clsReceipt

    Dim TotalCharges As Currency
    Dim TotalReceipts As Currency

    Dim OutstandingBalance As Currency

    On Error GoTo ErrorHandler

    GetMemberOutstandingBalance = 0@


    If MemberID <= 0 Then Exit Function


    '--------------------------------------------------------------------------
    ' AKTÝF BORÇLAR
    '--------------------------------------------------------------------------

    Set Charges = _
        mChargeRepository.GetByMemberID(MemberID)

    If Not Charges Is Nothing Then

        For Each Charge In Charges

            If Not Charge Is Nothing Then

                If StrComp( _
                    Trim$(Charge.Status), _
                    CHARGE_STATUS_ACTIVE, _
                    vbTextCompare) = 0 Then

                    TotalCharges = _
                        TotalCharges + Charge.Amount

                End If

            End If

        Next Charge

    End If


    '--------------------------------------------------------------------------
    ' TAHSÝLATLAR
    '--------------------------------------------------------------------------

    Set Receipts = _
        mReceiptRepository.GetByMemberID(MemberID)

    If Not Receipts Is Nothing Then

        For Each Receipt In Receipts

            If Not Receipt Is Nothing Then

                If Receipt.Amount > 0 Then

                    TotalReceipts = _
                        TotalReceipts + Receipt.Amount

                End If

            End If

        Next Receipt

    End If


    '--------------------------------------------------------------------------
    ' GERÇEK KALAN BORÇ
    '--------------------------------------------------------------------------

    OutstandingBalance = _
        TotalCharges - TotalReceipts


    If OutstandingBalance < 0 Then

        OutstandingBalance = 0@

    End If


    GetMemberOutstandingBalance = _
        OutstandingBalance

    Exit Function


ErrorHandler:

    GetMemberOutstandingBalance = 0@

End Function


'==============================================================================
' GET CHARGE RECEIPTS
'==============================================================================

Private Function GetChargeReceipts( _
    ByVal ChargeID As Long) As Collection

    Dim Result As Collection
    Dim MemberReceipts As Collection

    Dim Receipt As clsReceipt
    Dim Charge As clsMemberCharge

    Dim I As Long

    On Error GoTo ErrorHandler

    Set Result = New Collection

    Set GetChargeReceipts = Result

    If ChargeID <= 0 Then Exit Function


    Set Charge = _
        mChargeRepository.GetByID(ChargeID)

    If Charge Is Nothing Then Exit Function


    Set MemberReceipts = _
        mReceiptRepository.GetByMemberID( _
            Charge.MemberID)

    If MemberReceipts Is Nothing Then Exit Function


    For I = 1 To MemberReceipts.Count

        Set Receipt = MemberReceipts(I)

        If Not Receipt Is Nothing Then

            If Receipt.ChargeID = ChargeID Then

                Result.Add Receipt

            End If

        End If

    Next I


    Set GetChargeReceipts = Result

    Exit Function


ErrorHandler:

    Set GetChargeReceipts = Nothing

End Function


'==============================================================================
' GET CHARGE OUTSTANDING AMOUNT
'==============================================================================

Public Function GetChargeOutstandingAmount( _
    ByVal ChargeID As Long) As Currency

    Dim Charge As clsMemberCharge
    Dim Receipts As Collection

    Dim Receipt As clsReceipt

    Dim TotalCollected As Currency
    Dim OutstandingAmount As Currency

    On Error GoTo ErrorHandler

    GetChargeOutstandingAmount = 0@


    If ChargeID <= 0 Then Exit Function


    Set Charge = _
        mChargeRepository.GetByID(ChargeID)

    If Charge Is Nothing Then Exit Function


    '--------------------------------------------------------------------------
    ' SADECE AKTÝF BORÇ
    '--------------------------------------------------------------------------

    If StrComp( _
        Trim$(Charge.Status), _
        CHARGE_STATUS_ACTIVE, _
        vbTextCompare) <> 0 Then

        Exit Function

    End If


    OutstandingAmount = _
        Charge.Amount


    '--------------------------------------------------------------------------
    ' BU BORCA AÝT TAHSÝLATLAR
    '--------------------------------------------------------------------------

    Set Receipts = _
        GetChargeReceipts(ChargeID)

    If Not Receipts Is Nothing Then

        For Each Receipt In Receipts

            If Not Receipt Is Nothing Then

                If Receipt.Amount > 0 Then

                    TotalCollected = _
                        TotalCollected + Receipt.Amount

                End If

            End If

        Next Receipt

    End If


    '--------------------------------------------------------------------------
    ' KALAN
    '--------------------------------------------------------------------------

    OutstandingAmount = _
        OutstandingAmount - TotalCollected


    If OutstandingAmount < 0 Then

        OutstandingAmount = 0@

    End If


    GetChargeOutstandingAmount = _
        OutstandingAmount

    Exit Function


ErrorHandler:

    GetChargeOutstandingAmount = 0@

End Function


'==============================================================================
' DELETE RECEIPT
'==============================================================================

Public Function DeleteReceipt( _
    ByVal ReceiptID As Long) As Boolean

    On Error GoTo ErrorHandler

    DeleteReceipt = False

    If ReceiptID <= 0 Then Exit Function

    DeleteReceipt = _
        mReceiptRepository.DeleteByID(ReceiptID)

    Exit Function


ErrorHandler:

    DeleteReceipt = False

End Function
