VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMemberCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' clsMemberCharge
'
' WF-003 - Üye Borç Hareketi Entity
'
' Bu sýnýf yalnýzca üye borç hareketinin verisini temsil eder.
' Excel tablosuna doðrudan yazmaz.
'
' Veri eriþimi:
'     clsMemberChargeRepository
'
' Ýþ kurallarý:
'     clsMemberChargeService
'==============================================================================

Private mChargeID As Long
Private mMemberID As Long
Private mChargeType As String
Private mChargeDate As Date
Private mFiscalYear As Long
Private mDueMonth As Long
Private mAmount As Currency
Private mAdjustmentType As String
Private mSourceChargeID As Long
Private mDescription As String
Private mDecisionNo As String
Private mDecisionDate As Date
Private mEntryDate As Date
Private mStatus As String


'==============================================================================
' ChargeID
'==============================================================================

Public Property Get ChargeID() As Long
    ChargeID = mChargeID
End Property

Public Property Let ChargeID(ByVal Value As Long)
    mChargeID = Value
End Property


'==============================================================================
' MemberID
'==============================================================================

Public Property Get MemberID() As Long
    MemberID = mMemberID
End Property

Public Property Let MemberID(ByVal Value As Long)
    mMemberID = Value
End Property


'==============================================================================
' ChargeType
'==============================================================================

Public Property Get ChargeType() As String
    ChargeType = mChargeType
End Property

Public Property Let ChargeType(ByVal Value As String)
    mChargeType = Trim$(Value)
End Property


'==============================================================================
' ChargeDate
'==============================================================================

Public Property Get ChargeDate() As Date
    ChargeDate = mChargeDate
End Property

Public Property Let ChargeDate(ByVal Value As Date)
    mChargeDate = Value
End Property


'==============================================================================
' FiscalYear
'==============================================================================

Public Property Get FiscalYear() As Long
    FiscalYear = mFiscalYear
End Property

Public Property Let FiscalYear(ByVal Value As Long)
    mFiscalYear = Value
End Property


'==============================================================================
' DueMonth
'==============================================================================

Public Property Get DueMonth() As Long
    DueMonth = mDueMonth
End Property

Public Property Let DueMonth(ByVal Value As Long)
    mDueMonth = Value
End Property


'==============================================================================
' Amount
'==============================================================================

Public Property Get Amount() As Currency
    Amount = mAmount
End Property

Public Property Let Amount(ByVal Value As Currency)
    mAmount = Value
End Property


'==============================================================================
' AdjustmentType
'==============================================================================

Public Property Get AdjustmentType() As String
    AdjustmentType = mAdjustmentType
End Property

Public Property Let AdjustmentType(ByVal Value As String)
    mAdjustmentType = Trim$(Value)
End Property


'==============================================================================
' SourceChargeID
'==============================================================================

Public Property Get SourceChargeID() As Long
    SourceChargeID = mSourceChargeID
End Property

Public Property Let SourceChargeID(ByVal Value As Long)
    mSourceChargeID = Value
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
' DecisionNo
'==============================================================================

Public Property Get DecisionNo() As String
    DecisionNo = mDecisionNo
End Property

Public Property Let DecisionNo(ByVal Value As String)
    mDecisionNo = Trim$(Value)
End Property


'==============================================================================
' DecisionDate
'==============================================================================

Public Property Get DecisionDate() As Date
    DecisionDate = mDecisionDate
End Property

Public Property Let DecisionDate(ByVal Value As Date)
    mDecisionDate = Value
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
' Status
'==============================================================================

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Let Status(ByVal Value As String)
    mStatus = Trim$(Value)
End Property
