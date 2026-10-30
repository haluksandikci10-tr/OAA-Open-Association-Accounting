Attribute VB_Name = "modValidation"
Option Explicit

'==============================================================================
' OAA v2 Framework
' Validation Functions
'==============================================================================
'------------------------------------------------------------------------------
' Zorunlu alan kontrolü.
'------------------------------------------------------------------------------
Public Function ValidateRequired( _
                ByVal Value As String) As Boolean

    ValidateRequired = (Trim$(Value) <> vbNullString)

End Function

'------------------------------------------------------------------------------
' Para birimi doðrulamasý.
'------------------------------------------------------------------------------
Public Function ValidateCurrency( _
                ByVal CurrencyCode As String) As Boolean

    CurrencyCode = UCase$(Trim$(CurrencyCode))

    Select Case CurrencyCode

        Case "TRY", "USD", "EUR"
            ValidateCurrency = True

        Case Else
            ValidateCurrency = False

    End Select

End Function
'------------------------------------------------------------------------------
' E-posta doðrulamasý.
'------------------------------------------------------------------------------
Public Function ValidateEmail( _
                ByVal Email As String) As Boolean

    Email = Trim$(Email)

    If Email = vbNullString Then
        ValidateEmail = False
        Exit Function
    End If

    ValidateEmail = _
        (InStr(Email, "@") > 1 And _
         InStrRev(Email, ".") > InStr(Email, "@"))

End Function
'------------------------------------------------------------------------------
' Ay doðrulamasý.
'------------------------------------------------------------------------------
Public Function ValidateMonth( _
                ByVal MonthNumber As Integer) As Boolean

    ValidateMonth = (MonthNumber >= 1 And MonthNumber <= 12)

End Function

