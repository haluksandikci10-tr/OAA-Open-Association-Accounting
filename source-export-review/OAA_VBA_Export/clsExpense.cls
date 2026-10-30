VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Class : clsExpense
' Purpose: Gider veri nesnesi
'
' Entity sorumluluðu:
'   - Gider verisini taþýr.
'   - Excel / Worksheet eriþimi yapmaz.
'   - Repository çaðýrmaz.
'   - Ýþ kuralý uygulamaz.
'==============================================================================

Private mExpenseID As Long
Private mEntryDate As Date
Private mDocumentNo As String
Private mExpenseType As String
Private mDescription As String
Private mExpenseAmount As Currency
Private mCashAmount As Currency
Private mBankAmount As Currency


'==============================================================================
' EXPENSE ID
'==============================================================================

Public Property Get ExpenseID() As Long

    ExpenseID = mExpenseID

End Property


Public Property Let ExpenseID(ByVal Value As Long)

    mExpenseID = Value

End Property


'==============================================================================
' ENTRY DATE
'==============================================================================

Public Property Get EntryDate() As Date

    EntryDate = mEntryDate

End Property


Public Property Let EntryDate(ByVal Value As Date)

    mEntryDate = Value

End Property


'==============================================================================
' DOCUMENT NO
'==============================================================================

Public Property Get DocumentNo() As String

    DocumentNo = mDocumentNo

End Property


Public Property Let DocumentNo(ByVal Value As String)

    mDocumentNo = Value

End Property


'==============================================================================
' EXPENSE TYPE
'==============================================================================

Public Property Get ExpenseType() As String

    ExpenseType = mExpenseType

End Property


Public Property Let ExpenseType(ByVal Value As String)

    mExpenseType = Value

End Property


'==============================================================================
' DESCRIPTION
'==============================================================================

Public Property Get Description() As String

    Description = mDescription

End Property


Public Property Let Description(ByVal Value As String)

    mDescription = Value

End Property


'==============================================================================
' EXPENSE AMOUNT
'==============================================================================

Public Property Get ExpenseAmount() As Currency

    ExpenseAmount = mExpenseAmount

End Property


Public Property Let ExpenseAmount(ByVal Value As Currency)

    mExpenseAmount = Value

End Property


'==============================================================================
' CASH AMOUNT
'==============================================================================

Public Property Get CashAmount() As Currency

    CashAmount = mCashAmount

End Property


Public Property Let CashAmount(ByVal Value As Currency)

    mCashAmount = Value

End Property


'==============================================================================
' BANK AMOUNT
'==============================================================================

Public Property Get BankAmount() As Currency

    BankAmount = mBankAmount

End Property


Public Property Let BankAmount(ByVal Value As Currency)

    mBankAmount = Value

End Property
