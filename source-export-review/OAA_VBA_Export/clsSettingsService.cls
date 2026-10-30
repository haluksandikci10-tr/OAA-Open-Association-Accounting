VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettingsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Class : clsSettingsService
' Purpose : Settings Service
'==============================================================================

Private mRepository As clsSettingsRepository

'------------------------------------------------------------------------------
' Sýnýf baþlatýlýrken Repository nesnesini oluþturur.
'------------------------------------------------------------------------------
Private Sub Class_Initialize()

    Set mRepository = New clsSettingsRepository

End Sub

'------------------------------------------------------------------------------
' Sýnýf sonlandýrýlýrken nesneleri serbest býrakýr.
'------------------------------------------------------------------------------
Private Sub Class_Terminate()

    Set mRepository = Nothing

End Sub
'------------------------------------------------------------------------------
' Tüm ayar kayýtlarýný döndürür.
'------------------------------------------------------------------------------
Public Function GetAllSettings() As Collection

    Set GetAllSettings = mRepository.GetAll

End Function
'------------------------------------------------------------------------------
' Belirtilen ayar anahtarýný döndürür.
'------------------------------------------------------------------------------
Public Function FindByKey(ByVal SettingKey As String) As clsSettings

    Set FindByKey = mRepository.GetByKey(SettingKey)

End Function

'------------------------------------------------------------------------------
' Belirtilen ayar anahtarýnýn mevcut olup olmadýðýný kontrol eder.
'------------------------------------------------------------------------------
Public Function Exists(ByVal SettingKey As String) As Boolean

    Exists = mRepository.Exists(SettingKey)

End Function

'------------------------------------------------------------------------------
' Mevcut bir ayar kaydýný günceller.
'------------------------------------------------------------------------------
Public Function Update(ByVal Setting As clsSettings) As Boolean

    Update = mRepository.Update(Setting)

End Function
'------------------------------------------------------------------------------
' Belirtilen anahtara ait ayar deðerini döndürür.
'------------------------------------------------------------------------------
Public Function GetValue(ByVal SettingKey As String, _
                         Optional ByVal DefaultValue As String = "") As String

    Dim Setting As clsSettings

    Set Setting = mRepository.GetByKey(SettingKey)

    If Setting Is Nothing Then
        GetValue = DefaultValue
    Else
        GetValue = Setting.SettingValue
    End If

End Function
'------------------------------------------------------------------------------
' Belirtilen anahtara ait ayar deðerini deðiþtirir.
' Yeni kayýt yoksa oluþturur.
'------------------------------------------------------------------------------
Public Function SetValue(ByVal SettingKey As String, _
                         ByVal SettingValue As String) As Boolean

    Dim Setting As clsSettings

    Set Setting = mRepository.GetByKey(SettingKey)

    If Setting Is Nothing Then

        Set Setting = New clsSettings

        Setting.SettingKey = SettingKey
        Setting.SettingValue = SettingValue
        Setting.DataType = "String"
        Setting.Category = "General"
        Setting.IsSystem = False
        Setting.Description = ""

        SetValue = mRepository.Add(Setting)

    Else

        Setting.SettingValue = SettingValue

        SetValue = mRepository.Update(Setting)

    End If

End Function

'------------------------------------------------------------------------------
' Varsayýlan sistem ayarlarýný oluþturur.
'------------------------------------------------------------------------------
Public Sub InitializeDefaults()

    If Not Exists("Currency") Then
        SetValue "Currency", "TRY"
    End If

    If Not Exists("FiscalYearStart") Then
        SetValue "FiscalYearStart", "1"
    End If

    If Not Exists("EmailProvider") Then
        SetValue "EmailProvider", "Outlook"
    End If

End Sub
