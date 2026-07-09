Attribute VB_Name = "modGlobals"
Option Explicit

'=========================================================
' OAA - Open Association Accounting
' Module : modGlobals
' Version: 1.0.0
'=========================================================

'---------------------------------------------------------
' Sayfa Adlarý
'---------------------------------------------------------
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_BOOK As String = "ÝþletmeDefteri"
Public Const SHEET_CASH As String = "Kasa"
Public Const SHEET_BANK As String = "Banka"
Public Const SHEET_MEMBERS As String = "Üyeler"
Public Const SHEET_DUES As String = "Aidatlar"
Public Const SHEET_INCOME As String = "Gelir"
Public Const SHEET_EXPENSE As String = "Gider"
Public Const SHEET_RECEIPTS As String = "Makbuzlar"
Public Const SHEET_REPORTS As String = "Raporlar"
Public Const SHEET_SETTINGS As String = "Ayarlar"
Public Const SHEET_LOG As String = "Log"
Public Const TBL_LOG As String = "tblLog"
Public Const LOG_INFO As String = "INFO"
Public Const LOG_WARNING As String = "WARNING"
Public Const LOG_ERROR As String = "ERROR"
Public Const LOG_DEBUG As String = "DEBUG"
'---------------------------------------------------------
' Tablo Adlarý
'---------------------------------------------------------
Public Const TBL_SETTINGS As String = "tblSettings"
Public Const TBL_MEMBERS As String = "tblMembers"

'---------------------------------------------------------
' Varsayýlan Deðerler
'---------------------------------------------------------
Public Const DEFAULT_RECEIPT_SERIES As String = "A"
Public Const DEFAULT_BANK As String = "Garanti BBVA"

'---------------------------------------------------------
' Para Birimi
'---------------------------------------------------------
Public Const CUR_TRY As String = "TRY"
'---------------------------------------------------------
' Üye Durumlarý
'---------------------------------------------------------
Public Const MEMBER_ACTIVE As String = "Aktif"
Public Const MEMBER_PASSIVE As String = "Pasif"
Public Const MEMBER_LEAVE As String = "Ýzinli"
'---------------------------------------------------------
' Program Bilgileri
'---------------------------------------------------------
Public Const APP_NAME As String = "Open Association Accounting"
Public Const APP_VERSION As String = "1.0.0"
'---------------------------------------------------------
' Tarih Formatlarý
'---------------------------------------------------------
Public Const DATE_FORMAT As String = "dd.mm.yyyy"
Public Const DATETIME_FORMAT As String = "dd.mm.yyyy hh:mm:ss"
'=========================================================
' GLOBAL OBJECTS
'=========================================================

Public gSettings As clsSettings

'=========================================================
' INITIALIZE
'=========================================================

Public Sub InitializeGlobals()

    Set gSettings = New clsSettings

End Sub

'=========================================================
' RELEASE
'=========================================================

Public Sub ReleaseGlobals()

    Set gSettings = Nothing

End Sub

