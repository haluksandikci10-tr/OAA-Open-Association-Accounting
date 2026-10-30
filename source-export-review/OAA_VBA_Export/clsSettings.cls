VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Class : clsSettings
' Purpose : Settings Model
'==============================================================================

Private mSettingKey As String
Private mSettingValue As String
Private mDataType As String
Private mCategory As String
Private mIsSystem As Boolean
Private mDescription As String

'------------------------------------------------------------------------------
' Setting Key
'------------------------------------------------------------------------------
Public Property Get SettingKey() As String
    SettingKey = mSettingKey
End Property

Public Property Let SettingKey(ByVal Value As String)
    mSettingKey = Value
End Property

'------------------------------------------------------------------------------
' Setting Value
'------------------------------------------------------------------------------
Public Property Get SettingValue() As String
    SettingValue = mSettingValue
End Property

Public Property Let SettingValue(ByVal Value As String)
    mSettingValue = Value
End Property

'------------------------------------------------------------------------------
' Data Type
'------------------------------------------------------------------------------
Public Property Get DataType() As String
    DataType = mDataType
End Property

Public Property Let DataType(ByVal Value As String)
    mDataType = Value
End Property

'------------------------------------------------------------------------------
' Category
'------------------------------------------------------------------------------
Public Property Get Category() As String
    Category = mCategory
End Property

Public Property Let Category(ByVal Value As String)
    mCategory = Value
End Property

'------------------------------------------------------------------------------
' Is System
'------------------------------------------------------------------------------
Public Property Get IsSystem() As Boolean
    IsSystem = mIsSystem
End Property

Public Property Let IsSystem(ByVal Value As Boolean)
    mIsSystem = Value
End Property

'------------------------------------------------------------------------------
' Description
'------------------------------------------------------------------------------
Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal Value As String)
    mDescription = Value
End Property
