VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA Framework
' Class : clsValidationResult
' Purpose : Validation Result Object
'==============================================================================
Private mSuccess As Boolean
Private mMessage As String
Private mFieldName As String
Private mErrorCode As Long
'------------------------------------------------------------------------------
' Validation baþarýlý mý?
'------------------------------------------------------------------------------
Public Property Get Success() As Boolean

    Success = mSuccess

End Property

Public Property Let Success(ByVal Value As Boolean)

    mSuccess = Value

End Property
'------------------------------------------------------------------------------
' Doðrulama mesajý.
'------------------------------------------------------------------------------
Public Property Get Message() As String

    Message = mMessage

End Property

Public Property Let Message(ByVal Value As String)

    mMessage = Trim$(Value)

End Property
'------------------------------------------------------------------------------
' Hatalý alan adý.
'------------------------------------------------------------------------------
Public Property Get FieldName() As String

    FieldName = mFieldName

End Property

Public Property Let FieldName(ByVal Value As String)

    mFieldName = Trim$(Value)

End Property
'------------------------------------------------------------------------------
' Hata kodu.
'------------------------------------------------------------------------------
Public Property Get ErrorCode() As Long

    ErrorCode = mErrorCode

End Property

Public Property Let ErrorCode(ByVal Value As Long)

    mErrorCode = Value

End Property
'------------------------------------------------------------------------------
' Validation sonucunu baþarýlý olarak ayarlar.
'------------------------------------------------------------------------------
Public Sub SetSuccess()

    mSuccess = True
    mMessage = vbNullString
    mFieldName = vbNullString
    mErrorCode = 0

End Sub
'------------------------------------------------------------------------------
' Validation sonucunu hata olarak ayarlar.
'------------------------------------------------------------------------------
Public Sub SetError( _
            ByVal ErrorCode As Long, _
            ByVal FieldName As String, _
            ByVal Message As String)

    mSuccess = False
    mErrorCode = ErrorCode
    mFieldName = Trim$(FieldName)
    mMessage = Trim$(Message)

End Sub

