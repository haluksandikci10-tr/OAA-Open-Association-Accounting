Attribute VB_Name = "modUtilities"

Option Explicit

'==============================================================================
' OAA v2 Framework
' Utility Functions
'==============================================================================
'------------------------------------------------------------------------------
' Metin boþ mu?
'------------------------------------------------------------------------------
Public Function IsNullOrEmpty( _
                ByVal Value As String) As Boolean

    IsNullOrEmpty = (Trim$(Value) = vbNullString)

End Function

'------------------------------------------------------------------------------
' Null deðeri boþ metne çevirir.
'------------------------------------------------------------------------------
Public Function Nz( _
                ByVal Value As Variant, _
                Optional ByVal DefaultValue As String = "") As String

    If IsNull(Value) Then
        Nz = DefaultValue
    Else
        Nz = CStr(Value)
    End If

End Function
'------------------------------------------------------------------------------
' Güvenli Trim.
'------------------------------------------------------------------------------
Public Function SafeTrim( _
                ByVal Value As Variant) As String

    SafeTrim = Trim$(Nz(Value))

End Function

'------------------------------------------------------------------------------
' Hücre deðerini güvenli biçimde Long'a çevirir. Sayýsal deðilse 0 döner.
'------------------------------------------------------------------------------
Public Function NzLongValue( _
    ByVal Value As Variant) As Long

    On Error GoTo ErrorHandler

    If IsNumeric(Value) Then
        NzLongValue = CLng(Value)
    Else
        NzLongValue = 0
    End If

    Exit Function

ErrorHandler:

    NzLongValue = 0

End Function

'------------------------------------------------------------------------------
' Hücre deðerini güvenli biçimde Currency'ye çevirir. Sayýsal deðilse 0 döner.
'------------------------------------------------------------------------------
Public Function NzCurrencyValue( _
    ByVal Value As Variant) As Currency

    On Error GoTo ErrorHandler

    If IsNumeric(Value) Then
        NzCurrencyValue = CCur(Value)
    Else
        NzCurrencyValue = 0
    End If

    Exit Function

ErrorHandler:

    NzCurrencyValue = 0

End Function

'------------------------------------------------------------------------------
' Hücre deðerini güvenli biçimde metne çevirir. Hata/Null ise boþ metin döner.
'------------------------------------------------------------------------------
Public Function NzStringValue( _
    ByVal Value As Variant) As String

    On Error GoTo ErrorHandler

    If IsError(Value) Then
        NzStringValue = vbNullString
    ElseIf IsNull(Value) Then
        NzStringValue = vbNullString
    Else
        NzStringValue = CStr(Value)
    End If

    Exit Function

ErrorHandler:

    NzStringValue = vbNullString

End Function

'------------------------------------------------------------------------------
' Hücre deðerini güvenli biçimde Date'e çevirir. Tarih deðilse 0 döner.
'------------------------------------------------------------------------------
Public Function NzDateValue( _
    ByVal Value As Variant) As Date

    On Error GoTo ErrorHandler

    If IsDate(Value) Then
        NzDateValue = CDate(Value)
    Else
        NzDateValue = 0
    End If

    Exit Function

ErrorHandler:

    NzDateValue = 0

End Function


