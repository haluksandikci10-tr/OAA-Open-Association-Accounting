VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Sýnýf   : clsReceipt
' Amaç    : Tahsilat / Makbuz Entity
'
' OAA-001 temel modeli:
'   Seri/No
'   Tarih
'   Üye
'   Tutar
'   Ödeme yöntemi
'
' Not:
'   Entity yalnýzca veri taþýr.
'   Excel / Worksheet eriþimi içermez.
'   Repository veya Service çaðýrmaz.
'==============================================================================

Private mReceiptID As Long
Private mSeries As String
Private mReceiptNumber As Long

Private mReceiptDate As Date
Private mMemberID As Long
Private mChargeID As Long
Private mAmount As Currency
Private mPaymentMethod As String
'==============================================================================
' CHARGE ID
'==============================================================================

Public Property Get ChargeID() As Long

    ChargeID = mChargeID

End Property


Public Property Let ChargeID(ByVal Value As Long)

    mChargeID = Value

End Property

'==============================================================================
' RECEIPT ID
'==============================================================================

Public Property Get ReceiptID() As Long

    ReceiptID = mReceiptID

End Property


Public Property Let ReceiptID(ByVal Value As Long)

    mReceiptID = Value

End Property


'==============================================================================
' SERÝ
'==============================================================================

Public Property Get Series() As String

    Series = mSeries

End Property


Public Property Let Series(ByVal Value As String)

    mSeries = Trim$(Value)

End Property


'==============================================================================
' MAKBUZ NUMARASI
'==============================================================================

Public Property Get ReceiptNumber() As Long

    ReceiptNumber = mReceiptNumber

End Property


Public Property Let ReceiptNumber(ByVal Value As Long)

    mReceiptNumber = Value

End Property


'==============================================================================
' TAHSÝLAT TARÝHÝ
'==============================================================================

Public Property Get ReceiptDate() As Date

    ReceiptDate = mReceiptDate

End Property


Public Property Let ReceiptDate(ByVal Value As Date)

    mReceiptDate = Value

End Property


'==============================================================================
' ÜYE ID
'==============================================================================

Public Property Get MemberID() As Long

    MemberID = mMemberID

End Property


Public Property Let MemberID(ByVal Value As Long)

    mMemberID = Value

End Property


'==============================================================================
' TUTAR
'==============================================================================

Public Property Get Amount() As Currency

    Amount = mAmount

End Property


Public Property Let Amount(ByVal Value As Currency)

    mAmount = Value

End Property


'==============================================================================
' ÖDEME YÖNTEMÝ
'==============================================================================

Public Property Get PaymentMethod() As String

    PaymentMethod = mPaymentMethod

End Property


Public Property Let PaymentMethod(ByVal Value As String)

    mPaymentMethod = Trim$(Value)

End Property
