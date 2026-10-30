VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Sýnýf   : clsMember
' Amaç   : Üye Entity sýnýfý
'==============================================================================

Private mMemberID As Long
Private mFirstName As String
Private mLastName As String
Private mStatus As String

Private mJoinDate As Date
Private mLeaveDate As Date
Private mBirthDate As Date
Private mMarriageDate As Date

Private mMobile As String
Private mEmail As String
Private mAddress As String

Private mSpouseName As String
Private mNotes As String


'==============================================================================
' ÜYE NUMARASI
'==============================================================================

Public Property Get MemberID() As Long

    MemberID = mMemberID

End Property


Public Property Let MemberID(ByVal Value As Long)

    mMemberID = Value

End Property


'==============================================================================
' AD
'==============================================================================

Public Property Get FirstName() As String

    FirstName = mFirstName

End Property


Public Property Let FirstName(ByVal Value As String)

    mFirstName = Trim$(Value)

End Property


'==============================================================================
' SOYAD
'==============================================================================

Public Property Get LastName() As String

    LastName = mLastName

End Property


Public Property Let LastName(ByVal Value As String)

    mLastName = Trim$(Value)

End Property


'==============================================================================
' ÜYE DURUMU
'==============================================================================

Public Property Get Status() As String

    Status = mStatus

End Property


Public Property Let Status(ByVal Value As String)

    mStatus = Trim$(Value)

End Property


'==============================================================================
' ÜYELÝK TARÝHÝ
'==============================================================================

Public Property Get JoinDate() As Date

    JoinDate = mJoinDate

End Property


Public Property Let JoinDate(ByVal Value As Date)

    mJoinDate = Value

End Property


'==============================================================================
' AYRILIÞ TARÝHÝ
'==============================================================================

Public Property Get LeaveDate() As Date

    LeaveDate = mLeaveDate

End Property


Public Property Let LeaveDate(ByVal Value As Date)

    mLeaveDate = Value

End Property


'==============================================================================
' GSM
'==============================================================================

Public Property Get Mobile() As String

    Mobile = mMobile

End Property


Public Property Let Mobile(ByVal Value As String)

    mMobile = Trim$(Value)

End Property


'==============================================================================
' E-POSTA
'==============================================================================

Public Property Get Email() As String

    Email = mEmail

End Property


Public Property Let Email(ByVal Value As String)

    mEmail = LCase$(Trim$(Value))

End Property


'==============================================================================
' ADRES
'==============================================================================

Public Property Get Address() As String

    Address = mAddress

End Property


Public Property Let Address(ByVal Value As String)

    mAddress = Trim$(Value)

End Property


'==============================================================================
' DOÐUM TARÝHÝ
'==============================================================================

Public Property Get BirthDate() As Date

    BirthDate = mBirthDate

End Property


Public Property Let BirthDate(ByVal Value As Date)

    mBirthDate = Value

End Property


'==============================================================================
' EVLÝLÝK TARÝHÝ
'==============================================================================

Public Property Get MarriageDate() As Date

    MarriageDate = mMarriageDate

End Property


Public Property Let MarriageDate(ByVal Value As Date)

    mMarriageDate = Value

End Property


'==============================================================================
' EÞ ADI
'==============================================================================

Public Property Get SpouseName() As String

    SpouseName = mSpouseName

End Property


Public Property Let SpouseName(ByVal Value As String)

    mSpouseName = Trim$(Value)

End Property


'==============================================================================
' NOTLAR
'==============================================================================

Public Property Get Notes() As String

    Notes = mNotes

End Property


Public Property Let Notes(ByVal Value As String)

    mNotes = Trim$(Value)

End Property


'==============================================================================
' TAM AD
'==============================================================================

Public Property Get FullName() As String

    FullName = Trim$(mFirstName & " " & mLastName)

End Property
