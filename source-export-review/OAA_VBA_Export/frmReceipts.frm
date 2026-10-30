VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReceipts 
   Caption         =   "Tahsilat / Makbuz Giriþi"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   OleObjectBlob   =   "frmReceipts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Form    : frmReceipts
' Amaç    : Tahsilat / Makbuz Giriþ Ekraný
'
' TEMEL UI KURALI
'
' Kullanýcý iþlemleri ÜYE ADI seçilerek yapar.
'
' Kullanýcý:
'   - MemberID girmez.
'   - ChargeID girmez.
'
' Üye seçimi:
'   cmbMember
'
' Borç seçimi:
'   cmbCharge
'
' Teknik ID deðerleri ComboBox'larýn gizli ilk kolonlarýnda tutulur.
'
'
' KALAN BORÇ KURALI
'
'   txtOutstanding
'
'   = Seçilen üyenin TOPLAM GERÇEK KALAN BORCU
'
'   Aktif borçlar toplamý
'       -
'   Üyenin toplam tahsilatlarý
'       =
'   Üyenin gerçek kalan borcu
'
'
' BORÇ KURALI
'
'   cmbCharge
'
'   = Seçilen üyeye ait açýk borçlar.
'
'   Her borcun kendi kalan tutarý ComboBox açýklamasýnda gösterilir.
'
'
' MÝMARÝ
'
'       frmReceipts
'            |
'            v
'       clsReceiptService
'            |
'       +----+---------------------------+
'       |                                |
'       v                                v
' clsMemberRepository          clsMemberChargeRepository
'
'       clsReceiptRepository
'
' Form Repository'lere doðrudan eriþmez.
'==============================================================================


Private mService As clsReceiptService


Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

'==============================================================================
' FORM INITIALIZE
'==============================================================================

Private Sub UserForm_Initialize()

    On Error GoTo ErrorHandler

    Set mService = New clsReceiptService

    InitializeForm

    Exit Sub

ErrorHandler:

    MsgBox _
        "Tahsilat ekraný baþlatýlamadý." & vbCrLf & vbCrLf & _
        "Hata No : " & Err.Number & vbCrLf & _
        "Açýklama: " & Err.Description, _
        vbCritical, _
        "OAA - Tahsilat"

End Sub


'==============================================================================
' FORM INITIALIZATION
'==============================================================================

Private Sub InitializeForm()

    '--------------------------------------------------------------------------
    ' ÜYE COMBOBOX
    '--------------------------------------------------------------------------

    With cmbMember

        .Clear

        .ColumnCount = 2

        .ColumnWidths = "0 pt;250 pt"

        .BoundColumn = 1

        .Style = fmStyleDropDownList

        .MatchEntry = fmMatchEntryComplete

    End With


    '--------------------------------------------------------------------------
    ' BORÇ COMBOBOX
    '--------------------------------------------------------------------------

    With cmbCharge

        .Clear

        .ColumnCount = 2

        .ColumnWidths = "0 pt;300 pt"

        .BoundColumn = 1

        .Style = fmStyleDropDownList

        .MatchEntry = fmMatchEntryComplete

    End With


    '--------------------------------------------------------------------------
    ' KALAN BORÇ
    '
    ' Bu alan ÜYENÝN TOPLAM GERÇEK KALAN BORCUNU gösterir.
    ' Kullanýcý deðiþtiremez.
    '--------------------------------------------------------------------------

    With txtOutstanding

        .Value = ""

        .Locked = True

        .Enabled = True

        .TabStop = False

    End With


    '--------------------------------------------------------------------------
    ' TAHSÝLAT TUTARI
    '--------------------------------------------------------------------------

    With txtAmount

        .Value = ""

        .Locked = False

        .Enabled = True

        .TabStop = True

    End With


    '--------------------------------------------------------------------------
    ' ÖDEME YÖNTEMÝ
    '--------------------------------------------------------------------------

    InitializePaymentMethods


    '--------------------------------------------------------------------------
    ' MAKBUZ SERÝSÝ
    '--------------------------------------------------------------------------

    With txtSeries

        .Value = "A"

        .Locked = False

        .Enabled = True

    End With


    '--------------------------------------------------------------------------
    ' MAKBUZ NUMARASI
    '
    ' Makbuz numarasý artýk kullanýcý tarafýndan girilmez.
    '
    ' Numara clsReceiptService / clsReceiptRepository tarafýndan
    ' otomatik oluþturulur.
    '--------------------------------------------------------------------------

    With txtReceiptNumber

        .Value = ""

        .MaxLength = 0

        .Locked = True

        .Enabled = True

        .TabStop = False

    End With


    '--------------------------------------------------------------------------
    ' TAHSÝLAT TARÝHÝ
    '--------------------------------------------------------------------------

    With txtReceiptDate

        .Value = Format$(Date, "dd.mm.yyyy")

        .MaxLength = 0

        .Locked = False

        .Enabled = True

        .TabStop = True

    End With


    '--------------------------------------------------------------------------
    ' ÜYE LÝSTESÝ
    '--------------------------------------------------------------------------

    LoadMembers


    '--------------------------------------------------------------------------
    ' BAÞLANGIÇ
    '--------------------------------------------------------------------------

    cmbMember.ListIndex = -1

    cmbCharge.Clear

    txtOutstanding.Value = ""

End Sub


'==============================================================================
' ÖDEME YÖNTEMLERÝ
'==============================================================================

Private Sub InitializePaymentMethods()

    With cmbPaymentMethod

        .Clear

        .AddItem "Nakit"

        .AddItem "Banka"

        .AddItem "Kredi Kartý"

        .ListIndex = -1

    End With

End Sub


'==============================================================================
' ÜYE LÝSTESÝNÝ YÜKLE
'
' Service üzerinden GetMembers kullanýlýr.
'
' ComboBox:
'
'   Kolon 0 = MemberID   -> GÝZLÝ
'   Kolon 1 = FullName   -> GÖRÜNÜR
'
' Kullanýcý yalnýzca üye adýný görür.
'==============================================================================

Private Sub LoadMembers()

    Dim Members As Collection
    Dim Member As clsMember

    Dim I As Long
    Dim RowIndex As Long

    On Error GoTo ErrorHandler


    '--------------------------------------------------------------------------
    ' SERVICE
    '--------------------------------------------------------------------------

    Set Members = _
        mService.GetMembers


    If Members Is Nothing Then Exit Sub

    If Members.Count = 0 Then Exit Sub


    '--------------------------------------------------------------------------
    ' COMBOBOX
    '--------------------------------------------------------------------------

    With cmbMember

        .Clear

        For I = 1 To Members.Count

            Set Member = Members(I)

            If Not Member Is Nothing Then

                '----------------------------------------------------------------
                ' Gizli MemberID
                '----------------------------------------------------------------

                .AddItem CStr(Member.MemberID)

                RowIndex = .ListCount - 1


                '----------------------------------------------------------------
                ' Görünen Üye Adý
                '----------------------------------------------------------------

                .List(RowIndex, 1) = Member.FullName

            End If

        Next I

    End With


    Set Member = Nothing
    Set Members = Nothing

    Exit Sub


ErrorHandler:

    Set Member = Nothing
    Set Members = Nothing

    MsgBox _
        "Üye listesi yüklenemedi." & vbCrLf & vbCrLf & _
        "Hata No : " & Err.Number & vbCrLf & _
        "Açýklama: " & Err.Description, _
        vbCritical, _
        "OAA - Tahsilat"

End Sub


'==============================================================================
' ÜYE DEÐÝÞTÝÐÝNDE
'
' 1. Üyenin toplam gerçek kalan borcu hesaplanýr.
' 2. Üyeye ait açýk borçlar yüklenir.
'==============================================================================

Private Sub cmbMember_Change()

    Dim MemberID As Long

    Dim MemberOutstanding As Currency

    On Error GoTo ErrorHandler


    '--------------------------------------------------------------------------
    ' Önce eski seçimleri temizle
    '--------------------------------------------------------------------------

    cmbCharge.Clear

    txtOutstanding.Value = ""


    '--------------------------------------------------------------------------
    ' Üye seçilmemiþse
    '--------------------------------------------------------------------------

    If cmbMember.ListIndex < 0 Then Exit Sub


    '--------------------------------------------------------------------------
    ' Gizli MemberID
    '--------------------------------------------------------------------------

    If Not IsNumeric(cmbMember.Value) Then Exit Sub

    MemberID = CLng(cmbMember.Value)


    If MemberID <= 0 Then Exit Sub


    '--------------------------------------------------------------------------
    ' ÜYENÝN TOPLAM GERÇEK KALAN BORCU
    '--------------------------------------------------------------------------

    MemberOutstanding = _
        mService.GetMemberOutstandingBalance( _
            MemberID)


    txtOutstanding.Value = _
        Format$(MemberOutstanding, "#,##0.00")


    '--------------------------------------------------------------------------
    ' AÇIK BORÇLAR
    '--------------------------------------------------------------------------

    LoadOpenCharges MemberID

    Exit Sub


ErrorHandler:

    txtOutstanding.Value = ""

    MsgBox _
        "Üye bakiyesi yüklenemedi." & vbCrLf & vbCrLf & _
        "Hata No : " & Err.Number & vbCrLf & _
        "Açýklama: " & Err.Description, _
        vbCritical, _
        "OAA - Tahsilat"

End Sub


'==============================================================================
' AÇIK BORÇLARI YÜKLE
'
' Service:
'
'   GetOpenChargesForMember
'
' kullanýlýr.
'
' Kullanýcý ChargeID görmez.
'
' ComboBox görünümü:
'
'   Ocak 2027 - RESMÝ - Kalan: 750,00
'==============================================================================

Private Sub LoadOpenCharges( _
    ByVal MemberID As Long)

    Dim Charges As Collection
    Dim Charge As clsMemberCharge

    Dim OutstandingAmount As Currency

    Dim I As Long
    Dim RowIndex As Long

    Dim DisplayText As String

    On Error GoTo ErrorHandler


    '--------------------------------------------------------------------------
    ' SERVICE
    '--------------------------------------------------------------------------

    Set Charges = _
        mService.GetOpenChargesForMember(MemberID)


    If Charges Is Nothing Then Exit Sub


    '--------------------------------------------------------------------------
    ' AÇIK BORÇ YOK
    '--------------------------------------------------------------------------

    If Charges.Count = 0 Then

        Exit Sub

    End If


    '--------------------------------------------------------------------------
    ' COMBOBOX
    '--------------------------------------------------------------------------

    With cmbCharge

        .Clear

        For I = 1 To Charges.Count

            Set Charge = Charges(I)

            If Not Charge Is Nothing Then

                '----------------------------------------------------------------
                ' SEÇÝLEN BORCUN KENDÝ KALANI
                '----------------------------------------------------------------

                OutstandingAmount = _
                    mService.GetChargeOutstandingAmount( _
                        Charge.ChargeID)


                If OutstandingAmount > 0 Then

                    '--------------------------------------------------------------
                    ' Kullanýcýya gösterilecek açýklama
                    '--------------------------------------------------------------

                    DisplayText = _
                        Format$(Charge.ChargeDate, "mmmm yyyy") & _
                        " - " & _
                        Trim$(Charge.ChargeType) & _
                        " - Kalan: " & _
                        Format$(OutstandingAmount, "#,##0.00")


                    '--------------------------------------------------------------
                    ' Gizli ChargeID
                    '--------------------------------------------------------------

                    .AddItem CStr(Charge.ChargeID)

                    RowIndex = .ListCount - 1


                    '--------------------------------------------------------------
                    ' Görünen borç açýklamasý
                    '--------------------------------------------------------------

                    .List(RowIndex, 1) = DisplayText

                End If

            End If

        Next I

    End With


    Set Charge = Nothing
    Set Charges = Nothing

    Exit Sub


ErrorHandler:

    Set Charge = Nothing
    Set Charges = Nothing

    MsgBox _
        "Açýk borçlar yüklenemedi." & vbCrLf & vbCrLf & _
        "Hata No : " & Err.Number & vbCrLf & _
        "Açýklama: " & Err.Description, _
        vbCritical, _
        "OAA - Tahsilat"

End Sub


'==============================================================================
' BORÇ SEÇÝLDÝÐÝNDE
'
' ÖNEMLÝ:
'
' txtOutstanding DEÐÝÞTÝRÝLMEZ.
'
' Çünkü txtOutstanding artýk ÜYENÝN TOPLAM KALAN BORCUDUR.
'
' Seçilen borcun kendi kalan tutarý ComboBox açýklamasýnda zaten gösterilir.
'==============================================================================

Private Sub cmbCharge_Change()

    '--------------------------------------------------------------------------
    ' Bu olayda toplam üye bakiyesi deðiþtirilmez.
    '--------------------------------------------------------------------------

End Sub


'==============================================================================
' KAYDET
'==============================================================================

Private Sub cmdSave_Click()

    Dim Receipt As clsReceipt

    Dim MemberID As Long
    Dim ChargeID As Long

    Dim ReceiptID As Long

    Dim ParsedDate As Date

    Dim Amount As Currency

    Dim ReceiptNumber As Long

    On Error GoTo ErrorHandler


    '--------------------------------------------------------------------------
    ' 1 - ÜYE
    '--------------------------------------------------------------------------

    If cmbMember.ListIndex < 0 Then

        MsgBox _
            "Lütfen üye seçiniz.", _
            vbExclamation, _
            "OAA - Tahsilat"

        cmbMember.SetFocus

        Exit Sub

    End If


    If Not IsNumeric(cmbMember.Value) Then

        MsgBox _
            "Seçilen üye bilgisi geçersiz.", _
            vbExclamation, _
            "OAA - Tahsilat"

        cmbMember.SetFocus

        Exit Sub

    End If


    MemberID = _
        CLng(cmbMember.Value)


    If MemberID <= 0 Then

        MsgBox _
            "Geçerli bir üye seçilmelidir.", _
            vbExclamation, _
            "OAA - Tahsilat"

        cmbMember.SetFocus

        Exit Sub

    End If


    '--------------------------------------------------------------------------
    ' 2 - BORÇ
    '--------------------------------------------------------------------------

    If cmbCharge.ListIndex < 0 Then

        MsgBox _
            "Lütfen tahsilat yapýlacak borcu seçiniz.", _
            vbExclamation, _
            "OAA - Tahsilat"

        cmbCharge.SetFocus

        Exit Sub

    End If


    If Not IsNumeric(cmbCharge.Value) Then

        MsgBox _
            "Seçilen borç bilgisi geçersiz.", _
            vbExclamation, _
            "OAA - Tahsilat"

        cmbCharge.SetFocus

        Exit Sub

    End If


    ChargeID = _
        CLng(cmbCharge.Value)


    If ChargeID <= 0 Then

        MsgBox _
            "Geçerli bir borç seçilmelidir.", _
            vbExclamation, _
            "OAA - Tahsilat"

        cmbCharge.SetFocus

        Exit Sub

    End If


    '--------------------------------------------------------------------------
    ' 3 - TAHSÝLAT TUTARI
    '--------------------------------------------------------------------------

    If Trim$(txtAmount.Value) = "" Then

        MsgBox _
            "Tahsilat tutarý girilmelidir.", _
            vbExclamation, _
            "OAA - Tahsilat"

        txtAmount.SetFocus

        Exit Sub

    End If


    If Not IsNumeric(txtAmount.Value) Then

        MsgBox _
            "Tahsilat tutarý sayýsal olmalýdýr.", _
            vbExclamation, _
            "OAA - Tahsilat"

        txtAmount.SetFocus

        Exit Sub

    End If


    Amount = _
        CCur(txtAmount.Value)


    If Amount <= 0 Then

        MsgBox _
            "Tahsilat tutarý sýfýrdan büyük olmalýdýr.", _
            vbExclamation, _
            "OAA - Tahsilat"

        txtAmount.SetFocus

        Exit Sub

    End If


    '--------------------------------------------------------------------------
    ' 4 - ÖDEME YÖNTEMÝ
    '--------------------------------------------------------------------------

    If cmbPaymentMethod.ListIndex < 0 Then

        MsgBox _
            "Lütfen ödeme yöntemini seçiniz.", _
            vbExclamation, _
            "OAA - Tahsilat"

        cmbPaymentMethod.SetFocus

        Exit Sub

    End If


    '--------------------------------------------------------------------------
    ' 5 - MAKBUZ SERÝSÝ
    '--------------------------------------------------------------------------

    If Len(Trim$(txtSeries.Value)) = 0 Then

        MsgBox _
            "Makbuz serisi boþ býrakýlamaz.", _
            vbExclamation, _
            "OAA - Tahsilat"

        txtSeries.SetFocus

        Exit Sub

    End If


    '--------------------------------------------------------------------------
    ' 6 - MAKBUZ NUMARASI
    '
    ' Makbuz numarasý kullanýcýdan alýnmaz.
    '
    ' Service katmaný:
    '
    '   ReceiptNumber <= 0
    '
    ' olduðunda Repository üzerinden sýradaki numarayý oluþturur.
    '--------------------------------------------------------------------------

    ReceiptNumber = 0


    '--------------------------------------------------------------------------
    ' 7 - TARÝH
    '--------------------------------------------------------------------------

    If Trim$(txtReceiptDate.Value) = "" Then

        MsgBox _
            "Tahsilat tarihi girilmelidir.", _
            vbExclamation, _
            "OAA - Tahsilat"

        txtReceiptDate.SetFocus

        Exit Sub

    End If


    If Not IsDate(txtReceiptDate.Value) Then

        MsgBox _
            "Tahsilat tarihi geçersiz.", _
            vbExclamation, _
            "OAA - Tahsilat"

        txtReceiptDate.SetFocus

        Exit Sub

    End If


    ParsedDate = _
        CDate(txtReceiptDate.Value)


    '--------------------------------------------------------------------------
    ' 8 - RECEIPT ENTITY
    '--------------------------------------------------------------------------

    Set Receipt = New clsReceipt


    With Receipt

        .Series = _
            Trim$(txtSeries.Value)

        .ReceiptNumber = _
            ReceiptNumber

        .ReceiptDate = _
            ParsedDate

        .MemberID = _
            MemberID

        .ChargeID = _
            ChargeID

        .Amount = _
            Amount

        .PaymentMethod = _
            Trim$(cmbPaymentMethod.Value)

    End With


    '--------------------------------------------------------------------------
    ' 9 - SERVICE
    '
    ' Nihai iþ kurallarý Service katmanýnda uygulanýr.
    '
    ' Özellikle:
    '
    '   - Üye var mý?
    '   - Borç var mý?
    '   - Borç bu üyeye mi ait?
    '   - Borç aktif mi?
    '   - Tutar geçerli mi?
    '   - Fazla tahsilat var mý?
    '   - Makbuz numarasý oluþturulabiliyor mu?
    '--------------------------------------------------------------------------

    ReceiptID = _
        mService.AddReceipt(Receipt)


    If ReceiptID <= 0 Then

        Err.Raise _
            vbObjectError + 8201, _
            "frmReceipts.cmdSave_Click", _
            "Tahsilat kaydý oluþturulamadý."

    End If


    '--------------------------------------------------------------------------
    ' 10 - BAÞARILI
    '
    ' ÖNEMLÝ:
    '
    ' Makbuz numarasý Service tarafýndan oluþturulduðu için
    ' Receipt nesnesindeki gerçek deðer kullanýlýr.
    '--------------------------------------------------------------------------

    MsgBox _
        "Tahsilat baþarýyla kaydedildi." & vbCrLf & vbCrLf & _
        "Makbuz No: " & _
        Trim$(Receipt.Series) & "-" & _
        CStr(Receipt.ReceiptNumber) & vbCrLf & _
        "Tutar: " & _
        Format$(Amount, "#,##0.00"), _
        vbInformation, _
        "OAA - Tahsilat"


    '--------------------------------------------------------------------------
    ' YENÝ ÝÞLEM
    '--------------------------------------------------------------------------

    ClearReceiptForm


    Set Receipt = Nothing

    Exit Sub


ErrorHandler:

    Set Receipt = Nothing

    MsgBox _
        "Tahsilat kaydedilemedi." & vbCrLf & vbCrLf & _
        "Hata No : " & Err.Number & vbCrLf & _
        "Açýklama: " & Err.Description, _
        vbCritical, _
        "OAA - Tahsilat"

End Sub


'==============================================================================
' FORMU TEMÝZLE
'==============================================================================

Private Sub ClearReceiptForm()

    cmbMember.ListIndex = -1

    cmbCharge.Clear

    txtOutstanding.Value = ""

    txtAmount.Value = ""

    cmbPaymentMethod.ListIndex = -1

    txtSeries.Value = "A"

    txtReceiptNumber.Value = ""

    txtReceiptDate.Value = _
        Format$(Date, "dd.mm.yyyy")


    cmbMember.SetFocus

End Sub


'==============================================================================
' ÝPTAL
'==============================================================================

Private Sub cmdCancel_Click()

    Unload Me

End Sub

