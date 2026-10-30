VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIncome 
   Caption         =   "Gelir Kaydý"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   OleObjectBlob   =   "frmIncome.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' FORM       : frmIncome
' AMAÇ       : Gelir kaydý oluþturma
'
' Mimari:
'
'   frmIncome
'       |
'       v
'   clsIncomeService
'       |
'       +----------------------+
'       |                      |
'       v                      v
' clsIncomeRepository     clsLedgerService
'       |                      |
'       v                      v
' TBL_INCOME              TBL_LEDGER
'
' Form doðrudan Excel tablolarýna eriþmez.
'==============================================================================


'==============================================================================
' FORM INITIALIZE
'==============================================================================

Private Sub UserForm_Initialize()

    On Error GoTo ErrorHandler

    '--------------------------------------------------------------------------
    ' FORM BAÞLANGIÇ DEÐERLERÝ
    '--------------------------------------------------------------------------

    Me.Caption = "Gelir Kaydý"


    '--------------------------------------------------------------------------
    ' TARÝH
    '--------------------------------------------------------------------------

    Me.txtEntryDate.Value = _
        Format$(Date, "dd.mm.yyyy")


    '--------------------------------------------------------------------------
    ' BELGE NO
    '--------------------------------------------------------------------------

    Me.txtDocumentNo.Value = ""


    '--------------------------------------------------------------------------
    ' GELÝR TÜRÜ
    '
    ' Burada sabit gelir türü eklenmiyor.
    ' Mevcut OAA lookup yapýsý daha sonra baðlanacaktýr.
    '--------------------------------------------------------------------------

    Me.cmbIncomeType.Clear


    '--------------------------------------------------------------------------
    ' AÇIKLAMA
    '--------------------------------------------------------------------------

    Me.txtDescription.Value = ""


    '--------------------------------------------------------------------------
    ' TUTARLAR
    '--------------------------------------------------------------------------

    Me.txtIncomeAmount.Value = ""

    Me.txtCashAmount.Value = ""

    Me.txtBankAmount.Value = ""


    '--------------------------------------------------------------------------
    ' ÝLK ODAK
    '--------------------------------------------------------------------------

    Me.txtEntryDate.SetFocus


    Exit Sub


ErrorHandler:

    MsgBox _
        "Gelir formu baþlatýlýrken hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbExclamation, _
        "OAA - Gelir"

End Sub


'==============================================================================
' KAYDET
'==============================================================================

Private Sub cmdSave_Click()

    Dim Income As clsIncome

    Dim IncomeService As clsIncomeService

    Dim Result As Boolean

    On Error GoTo ErrorHandler


    '--------------------------------------------------------------------------
    ' FORM DOÐRULAMALARI
    '--------------------------------------------------------------------------

    If Not ValidateForm Then Exit Sub


    '--------------------------------------------------------------------------
    ' INCOME ENTITY
    '--------------------------------------------------------------------------

    Set Income = _
        New clsIncome


    '--------------------------------------------------------------------------
    ' TARÝH
    '--------------------------------------------------------------------------

    Income.EntryDate = _
        CDate(Me.txtEntryDate.Value)


    '--------------------------------------------------------------------------
    ' BELGE NO
    '--------------------------------------------------------------------------

    Income.DocumentNo = _
        Trim$(Me.txtDocumentNo.Value)


    '--------------------------------------------------------------------------
    ' GELÝR TÜRÜ
    '--------------------------------------------------------------------------

    Income.IncomeType = _
        Trim$(Me.cmbIncomeType.Value)


    '--------------------------------------------------------------------------
    ' AÇIKLAMA
    '--------------------------------------------------------------------------

    Income.Description = _
        Trim$(Me.txtDescription.Value)


    '--------------------------------------------------------------------------
    ' GELÝR TUTARI
    '--------------------------------------------------------------------------

    Income.IncomeAmount = _
        ParseCurrency(Me.txtIncomeAmount.Value)


    '--------------------------------------------------------------------------
    ' KASA
    '--------------------------------------------------------------------------

    Income.CashAmount = _
        ParseCurrency(Me.txtCashAmount.Value)


    '--------------------------------------------------------------------------
    ' BANKA
    '--------------------------------------------------------------------------

    Income.BankAmount = _
        ParseCurrency(Me.txtBankAmount.Value)


    '--------------------------------------------------------------------------
    ' SERVICE
    '--------------------------------------------------------------------------

    Set IncomeService = _
        New clsIncomeService


    '--------------------------------------------------------------------------
    ' GELÝR + ÝÞLETME DEFTERÝ
    '--------------------------------------------------------------------------

    Result = _
        IncomeService.AddIncomeWithLedger(Income)


    If Not Result Then

        MsgBox _
            "Gelir kaydý oluþturulamadý." & _
            vbCrLf & vbCrLf & _
            "Girilen bilgileri kontrol ediniz.", _
            vbExclamation, _
            "OAA - Gelir"

        Exit Sub

    End If


    '--------------------------------------------------------------------------
    ' BAÞARILI
    '--------------------------------------------------------------------------

    MsgBox _
        "Gelir kaydý baþarýyla oluþturuldu." & _
        vbCrLf & vbCrLf & _
        "Gelir No: " & CStr(Income.IncomeID) & _
        vbCrLf & _
        "Tutar: " & FormatAmount(Income.IncomeAmount), _
        vbInformation, _
        "OAA - Gelir"


    '--------------------------------------------------------------------------
    ' FORMU TEMÝZLE
    '--------------------------------------------------------------------------

    ClearForm


    '--------------------------------------------------------------------------
    ' YENÝ KAYIT ÝÇÝN ODAK
    '--------------------------------------------------------------------------

    Me.txtEntryDate.SetFocus


    Exit Sub


ErrorHandler:

    MsgBox _
        "Gelir kaydý sýrasýnda hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbCritical, _
        "OAA - Gelir"

End Sub


'==============================================================================
' ÝPTAL
'==============================================================================

Private Sub cmdCancel_Click()

    Unload Me

End Sub


'==============================================================================
' FORM VALIDATION
'==============================================================================

Private Function ValidateForm() As Boolean

    Dim IncomeAmount As Currency

    Dim CashAmount As Currency

    Dim BankAmount As Currency

    On Error GoTo ErrorHandler


    ValidateForm = False


    '--------------------------------------------------------------------------
    ' TARÝH
    '--------------------------------------------------------------------------

    If Trim$(Me.txtEntryDate.Value) = "" Then

        MsgBox _
            "Tarih giriniz.", _
            vbExclamation, _
            "OAA - Gelir"

        Me.txtEntryDate.SetFocus

        Exit Function

    End If


    If Not IsDate(Me.txtEntryDate.Value) Then

        MsgBox _
            "Geçerli bir tarih giriniz.", _
            vbExclamation, _
            "OAA - Gelir"

        Me.txtEntryDate.SetFocus

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' GELÝR TÜRÜ
    '--------------------------------------------------------------------------

    If Trim$(Me.cmbIncomeType.Value) = "" Then

        MsgBox _
            "Gelir türü seçiniz.", _
            vbExclamation, _
            "OAA - Gelir"

        Me.cmbIncomeType.SetFocus

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' GELÝR TUTARI
    '--------------------------------------------------------------------------

    If Trim$(Me.txtIncomeAmount.Value) = "" Then

        MsgBox _
            "Gelir tutarý giriniz.", _
            vbExclamation, _
            "OAA - Gelir"

        Me.txtIncomeAmount.SetFocus

        Exit Function

    End If


    If Not IsNumeric(Me.txtIncomeAmount.Value) Then

        MsgBox _
            "Gelir tutarý sayýsal olmalýdýr.", _
            vbExclamation, _
            "OAA - Gelir"

        Me.txtIncomeAmount.SetFocus

        Exit Function

    End If


    IncomeAmount = _
        ParseCurrency(Me.txtIncomeAmount.Value)


    If IncomeAmount <= 0 Then

        MsgBox _
            "Gelir tutarý sýfýrdan büyük olmalýdýr.", _
            vbExclamation, _
            "OAA - Gelir"

        Me.txtIncomeAmount.SetFocus

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' KASA
    '--------------------------------------------------------------------------

    If Trim$(Me.txtCashAmount.Value) = "" Then

        Me.txtCashAmount.Value = "0"

    End If


    If Not IsNumeric(Me.txtCashAmount.Value) Then

        MsgBox _
            "Kasa tutarý sayýsal olmalýdýr.", _
            vbExclamation, _
            "OAA - Gelir"

        Me.txtCashAmount.SetFocus

        Exit Function

    End If


    CashAmount = _
        ParseCurrency(Me.txtCashAmount.Value)


    If CashAmount < 0 Then

        MsgBox _
            "Kasa tutarý negatif olamaz.", _
            vbExclamation, _
            "OAA - Gelir"

        Me.txtCashAmount.SetFocus

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' BANKA
    '--------------------------------------------------------------------------

    If Trim$(Me.txtBankAmount.Value) = "" Then

        Me.txtBankAmount.Value = "0"

    End If


    If Not IsNumeric(Me.txtBankAmount.Value) Then

        MsgBox _
            "Banka tutarý sayýsal olmalýdýr.", _
            vbExclamation, _
            "OAA - Gelir"

        Me.txtBankAmount.SetFocus

        Exit Function

    End If


    BankAmount = _
        ParseCurrency(Me.txtBankAmount.Value)


    If BankAmount < 0 Then

        MsgBox _
            "Banka tutarý negatif olamaz.", _
            vbExclamation, _
            "OAA - Gelir"

        Me.txtBankAmount.SetFocus

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' KASA + BANKA
    '--------------------------------------------------------------------------

    If CashAmount + BankAmount <> IncomeAmount Then

        MsgBox _
            "Kasa ve banka toplamý gelir tutarýna eþit olmalýdýr." & _
            vbCrLf & vbCrLf & _
            "Gelir Tutarý : " & FormatAmount(IncomeAmount) & _
            vbCrLf & _
            "Kasa         : " & FormatAmount(CashAmount) & _
            vbCrLf & _
            "Banka        : " & FormatAmount(BankAmount) & _
            vbCrLf & _
            "Toplam       : " & FormatAmount(CashAmount + BankAmount), _
            vbExclamation, _
            "OAA - Gelir"

        Me.txtCashAmount.SetFocus

        Exit Function

    End If


    '--------------------------------------------------------------------------
    ' BAÞARILI
    '--------------------------------------------------------------------------

    ValidateForm = True


    Exit Function


ErrorHandler:

    ValidateForm = False

    MsgBox _
        "Form doðrulamasý sýrasýnda hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbCritical, _
        "OAA - Gelir"

End Function


'==============================================================================
' FORM TEMÝZLE
'==============================================================================

Private Sub ClearForm()

    On Error GoTo ErrorHandler


    Me.txtEntryDate.Value = _
        Format$(Date, "dd.mm.yyyy")


    Me.txtDocumentNo.Value = ""


    Me.cmbIncomeType.ListIndex = -1


    Me.txtDescription.Value = ""


    Me.txtIncomeAmount.Value = ""


    Me.txtCashAmount.Value = ""


    Me.txtBankAmount.Value = ""


    Exit Sub


ErrorHandler:

    MsgBox _
        "Form temizlenirken hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbExclamation, _
        "OAA - Gelir"

End Sub


'==============================================================================
' CURRENCY PARSE
'
' TextBox deðerini Currency deðerine dönüþtürür.
'==============================================================================

Private Function ParseCurrency( _
    ByVal Value As String) As Currency

    On Error GoTo ErrorHandler


    Value = Trim$(Value)


    If Value = "" Then

        ParseCurrency = 0

        Exit Function

    End If


    ParseCurrency = _
        CCur(Value)


    Exit Function


ErrorHandler:

    ParseCurrency = 0

End Function


'==============================================================================
' CURRENCY FORMAT
'==============================================================================

Private Function FormatAmount( _
    ByVal Amount As Currency) As String

    On Error GoTo ErrorHandler


    FormatAmount = _
        Format$(Amount, "#,##0.00")


    Exit Function


ErrorHandler:

    FormatAmount = _
        CStr(Amount)

End Function

