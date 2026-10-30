Attribute VB_Name = "modErrors"
Option Explicit

'==============================================================================
' OAA Framework
' Module : modErrors
' Purpose: Standard Error Codes
'==============================================================================

'==============================================================================
' VALIDATION ERRORS
'==============================================================================

Public Const ERR_REQUIRED_FIELD      As Long = 1001
Public Const ERR_INVALID_EMAIL       As Long = 1002
Public Const ERR_INVALID_DATE        As Long = 1003
Public Const ERR_INVALID_VALUE       As Long = 1004

'==============================================================================
' DATA ERRORS
'==============================================================================

Public Const ERR_DUPLICATE_RECORD    As Long = 2001
Public Const ERR_RECORD_NOT_FOUND    As Long = 2002
Public Const ERR_RECORD_IN_USE       As Long = 2003

'==============================================================================
' DATABASE ERRORS
'==============================================================================

Public Const ERR_TABLE_NOT_FOUND     As Long = 3001
Public Const ERR_COLUMN_NOT_FOUND    As Long = 3002

'==============================================================================
' SYSTEM ERRORS
'==============================================================================

Public Const ERR_UNKNOWN             As Long = 9000

'------------------------------------------------------------------------------
' Hata kodunun açýklamasýný döndürür.
'------------------------------------------------------------------------------
Public Function GetErrorDescription(ByVal ErrorCode As Long) As String

    Select Case ErrorCode

        Case ERR_REQUIRED_FIELD
            GetErrorDescription = "Zorunlu alan boþ."

        Case ERR_INVALID_EMAIL
            GetErrorDescription = "Geçersiz e-posta adresi."

        Case ERR_INVALID_DATE
            GetErrorDescription = "Geçersiz tarih."

        Case ERR_INVALID_VALUE
            GetErrorDescription = "Geçersiz deðer."

        Case ERR_DUPLICATE_RECORD
            GetErrorDescription = "Kayýt zaten mevcut."

        Case ERR_RECORD_NOT_FOUND
            GetErrorDescription = "Kayýt bulunamadý."

        Case ERR_RECORD_IN_USE
            GetErrorDescription = "Kayýt kullanýmda."

        Case ERR_TABLE_NOT_FOUND
            GetErrorDescription = "Tablo bulunamadý."

        Case ERR_COLUMN_NOT_FOUND
            GetErrorDescription = "Sütun bulunamadý."

        Case Else
            GetErrorDescription = "Bilinmeyen hata."

    End Select

End Function
