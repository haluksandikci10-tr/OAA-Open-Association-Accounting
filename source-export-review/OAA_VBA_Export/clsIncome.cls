VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Entity  : clsIncome
' Amaç    : Gelir hareketini temsil eder.
'
' Entity yalnýzca veri taþýr.
'
' Ýþ kuralý içermez.
' Excel tablosuna doðrudan eriþmez.
' Repository veya Service çaðýrmaz.
'
' Alanlar:
'
'   IncomeID
'   EntryDate
'   DocumentNo
'   IncomeType
'   Description
'   IncomeAmount
'   CashAmount
'   BankAmount
'==============================================================================


'==============================================================================
' PRIVATE FIELDS
'==============================================================================

Private mIncomeID As Long

Private mEntryDate As Date

Private mDocumentNo As String

Private mIncomeType As String

Private mDescription As String

Private mIncomeAmount As Currency

Private mCashAmount As Currency

Private mBankAmount As Currency


'==============================================================================
' IncomeID
'==============================================================================

Public Property Get IncomeID() As Long

    IncomeID = mIncomeID

End Property


Public Property Let IncomeID( _
    ByVal Value As Long)

    mIncomeID = Value

End Property


'==============================================================================
' EntryDate
'==============================================================================

Public Property Get EntryDate() As Date

    EntryDate = mEntryDate

End Property


Public Property Let EntryDate( _
    ByVal Value As Date)

    mEntryDate = Value

End Property


'==============================================================================
' DocumentNo
'==============================================================================

Public Property Get DocumentNo() As String

    DocumentNo = mDocumentNo

End Property


Public Property Let DocumentNo( _
    ByVal Value As String)

    mDocumentNo = Value

End Property


'==============================================================================
' IncomeType
'==============================================================================

Public Property Get IncomeType() As String

    IncomeType = mIncomeType

End Property


Public Property Let IncomeType( _
    ByVal Value As String)

    mIncomeType = Value

End Property


'==============================================================================
' Description
'==============================================================================

Public Property Get Description() As String

    Description = mDescription

End Property


Public Property Let Description( _
    ByVal Value As String)

    mDescription = Value

End Property


'==============================================================================
' IncomeAmount
'==============================================================================

Public Property Get IncomeAmount() As Currency

    IncomeAmount = mIncomeAmount

End Property


Public Property Let IncomeAmount( _
    ByVal Value As Currency)

    mIncomeAmount = Value

End Property


'==============================================================================
' CashAmount
'==============================================================================

Public Property Get CashAmount() As Currency

    CashAmount = mCashAmount

End Property


Public Property Let CashAmount( _
    ByVal Value As Currency)

    mCashAmount = Value

End Property


'==============================================================================
' BankAmount
'==============================================================================

Public Property Get BankAmount() As Currency

    BankAmount = mBankAmount

End Property


Public Property Let BankAmount( _
    ByVal Value As Currency)

    mBankAmount = Value

End Property

