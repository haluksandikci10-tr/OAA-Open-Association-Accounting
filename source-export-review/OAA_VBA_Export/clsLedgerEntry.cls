VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLedgerEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA – Open Association Accounting
'
' Class : clsLedgerEntry
' Layer : Entity
' Purpose:
'   Ýþletme Defteri hareketini temsil eder.
'
' Sorumluluk:
'   Yalnýzca iþletme defteri hareketinin verisini taþýr.
'
' ÖNEMLÝ:
'   - Excel worksheet eriþimi yoktur.
'   - ListObject eriþimi yoktur.
'   - Repository çaðrýsý yoktur.
'   - Ýþ kuralý uygulanmaz.
'   - Kullanýcý arayüzü iþlemi yoktur.
'==============================================================================


'==============================================================================
' MODULE LEVEL FIELDS
'==============================================================================

Private mLedgerID As Long

Private mExpenseID As Long

Private mIncomeID As Long

Private mEntryDate As Date

Private mDocumentNo As String

Private mReceiptInvoice As String

Private mDescription As String

Private mIncomeType As String

Private mExpenseType As String

Private mIncomeAmount As Currency

Private mExpenseAmount As Currency

Private mCashAmount As Currency

Private mBankAmount As Currency

Private mBalanceAmount As Currency


'==============================================================================
' LedgerID
'==============================================================================

Public Property Get LedgerID() As Long

    LedgerID = mLedgerID

End Property


Public Property Let LedgerID(ByVal Value As Long)

    mLedgerID = Value

End Property


'==============================================================================
' ExpenseID - Gider kaydi ile iliski
'==============================================================================

Public Property Get ExpenseID() As Long

    ExpenseID = mExpenseID

End Property


Public Property Let ExpenseID(ByVal Value As Long)

    mExpenseID = Value

End Property


'==============================================================================
' IncomeID - Gelir kaydi ile iliski
'==============================================================================

Public Property Get IncomeID() As Long

    IncomeID = mIncomeID

End Property


Public Property Let IncomeID(ByVal Value As Long)

    mIncomeID = Value

End Property


'==============================================================================
' EntryDate
'==============================================================================

Public Property Get EntryDate() As Date

    EntryDate = mEntryDate

End Property


Public Property Let EntryDate(ByVal Value As Date)

    mEntryDate = Value

End Property


'==============================================================================
' DocumentNo
'==============================================================================

Public Property Get DocumentNo() As String

    DocumentNo = mDocumentNo

End Property


Public Property Let DocumentNo(ByVal Value As String)

    mDocumentNo = Trim$(Value)

End Property


'==============================================================================
' ReceiptInvoice
'==============================================================================

Public Property Get ReceiptInvoice() As String

    ReceiptInvoice = mReceiptInvoice

End Property


Public Property Let ReceiptInvoice(ByVal Value As String)

    mReceiptInvoice = Trim$(Value)

End Property


'==============================================================================
' Description
'==============================================================================

Public Property Get Description() As String

    Description = mDescription

End Property


Public Property Let Description(ByVal Value As String)

    mDescription = Trim$(Value)

End Property


'==============================================================================
' IncomeType
'==============================================================================

Public Property Get IncomeType() As String

    IncomeType = mIncomeType

End Property


Public Property Let IncomeType(ByVal Value As String)

    mIncomeType = Trim$(Value)

End Property


'==============================================================================
' ExpenseType
'==============================================================================

Public Property Get ExpenseType() As String

    ExpenseType = mExpenseType

End Property


Public Property Let ExpenseType(ByVal Value As String)

    mExpenseType = Trim$(Value)

End Property


'==============================================================================
' IncomeAmount
'==============================================================================

Public Property Get IncomeAmount() As Currency

    IncomeAmount = mIncomeAmount

End Property


Public Property Let IncomeAmount(ByVal Value As Currency)

    mIncomeAmount = Value

End Property


'==============================================================================
' ExpenseAmount
'==============================================================================

Public Property Get ExpenseAmount() As Currency

    ExpenseAmount = mExpenseAmount

End Property


Public Property Let ExpenseAmount(ByVal Value As Currency)

    mExpenseAmount = Value

End Property


'==============================================================================
' CashAmount
'==============================================================================

Public Property Get CashAmount() As Currency

    CashAmount = mCashAmount

End Property


Public Property Let CashAmount(ByVal Value As Currency)

    mCashAmount = Value

End Property


'==============================================================================
' BankAmount
'==============================================================================

Public Property Get BankAmount() As Currency

    BankAmount = mBankAmount

End Property


Public Property Let BankAmount(ByVal Value As Currency)

    mBankAmount = Value

End Property


'==============================================================================
' BalanceAmount
'==============================================================================

Public Property Get BalanceAmount() As Currency

    BalanceAmount = mBalanceAmount

End Property


Public Property Let BalanceAmount(ByVal Value As Currency)

    mBalanceAmount = Value

End Property

