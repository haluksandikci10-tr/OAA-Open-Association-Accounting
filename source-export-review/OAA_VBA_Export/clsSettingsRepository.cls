VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettingsRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Class : clsSettingsRepository
' Purpose : Settings Repository
'==============================================================================

'------------------------------------------------------------------------------
' Settings tablosunu döndürür.
'------------------------------------------------------------------------------
Private Function GetSettingsTable() As ListObject

    Set GetSettingsTable = _
        ThisWorkbook.Worksheets(SHEET_SETTINGS) _
                    .ListObjects(TABLE_SETTINGS)

End Function

'------------------------------------------------------------------------------
' Verilen sütun adýnýn tablo içindeki sýra numarasýný döndürür.
'------------------------------------------------------------------------------
Private Function Col(ByVal ColumnName As String) As Long

    Dim lo As ListObject

    Set lo = GetSettingsTable()

    Col = lo.ListColumns(ColumnName).Index

End Function
'------------------------------------------------------------------------------
' Tüm ayar kayýtlarýný döndürür.
'------------------------------------------------------------------------------
Public Function GetAll() As Collection

    Dim Settings As New Collection

    Dim lo As ListObject
    Dim lr As ListRow

    Set lo = GetSettingsTable()

    For Each lr In lo.ListRows

        If Len(Trim$(CStr(lr.Range.Cells(1, Col(COL_SETTING_KEY)).Value))) > 0 Then
            Settings.Add CreateSettingFromRow(lr)
        End If

    Next lr

    Set GetAll = Settings

End Function

'------------------------------------------------------------------------------
' Bir clsSettings nesnesini tablo satýrýna yazar (Add ve Update ortak noktasý).
' Anahtar (SettingKey) sütunu yazýlmaz; Add tarafýndan ayrýca yazýlýr.
'------------------------------------------------------------------------------
Private Sub WriteSettingToRow(ByVal lr As ListRow, ByVal Setting As clsSettings)

    With lr.Range

        .Cells(1, Col(COL_SETTING_VALUE)).Value = Setting.SettingValue
        .Cells(1, Col(COL_SETTING_DATA_TYPE)).Value = Setting.DataType
        .Cells(1, Col(COL_SETTING_CATEGORY)).Value = Setting.Category
        .Cells(1, Col(COL_SETTING_IS_SYSTEM)).Value = Setting.IsSystem
        .Cells(1, Col(COL_SETTING_DESCRIPTION)).Value = Setting.Description

    End With

End Sub

'------------------------------------------------------------------------------
' Tablodaki bir satýrý clsSettings nesnesine dönüþtürür.
'------------------------------------------------------------------------------
Private Function CreateSettingFromRow(ByVal lr As ListRow) As clsSettings

    Dim Setting As clsSettings

    Set Setting = New clsSettings

    With Setting

        .SettingKey = CStr(lr.Range.Cells(1, Col(COL_SETTING_KEY)).Value)
        .SettingValue = CStr(lr.Range.Cells(1, Col(COL_SETTING_VALUE)).Value)
        .DataType = CStr(lr.Range.Cells(1, Col(COL_SETTING_DATA_TYPE)).Value)
        .Category = CStr(lr.Range.Cells(1, Col(COL_SETTING_CATEGORY)).Value)
        .IsSystem = CBool(lr.Range.Cells(1, Col(COL_SETTING_IS_SYSTEM)).Value)
        .Description = CStr(lr.Range.Cells(1, Col(COL_SETTING_DESCRIPTION)).Value)

    End With

    Set CreateSettingFromRow = Setting

End Function
'------------------------------------------------------------------------------
' Yeni bir ayar kaydý ekler.
'------------------------------------------------------------------------------
Public Function Add(ByVal Setting As clsSettings) As Boolean

    On Error GoTo ErrorHandler

    Dim lo As ListObject
    Dim lr As ListRow

    If Setting Is Nothing Then
        Err.Raise vbObjectError + 7301, _
                  "clsSettingsRepository.Add", _
                  "Setting nesnesi Nothing."
    End If

    Set lo = GetSettingsTable()

    Set lr = lo.ListRows.Add

    lr.Range.Cells(1, Col(COL_SETTING_KEY)).Value = Setting.SettingKey

    WriteSettingToRow lr, Setting

    Add = True

    Exit Function

ErrorHandler:

    Add = False

    Err.Raise Err.Number, _
              "clsSettingsRepository.Add", _
              Err.Description

End Function
'------------------------------------------------------------------------------
' Ayar anahtarýna göre kaydý döndürür.
'------------------------------------------------------------------------------
Public Function GetByKey(ByVal SettingKey As String) As clsSettings

    Dim lo As ListObject
    Dim lr As ListRow

    Set lo = GetSettingsTable()

    For Each lr In lo.ListRows

        If StrComp( _
            Trim$(CStr(lr.Range.Cells(1, Col(COL_SETTING_KEY)).Value)), _
            Trim$(SettingKey), _
            vbTextCompare) = 0 Then

            Set GetByKey = CreateSettingFromRow(lr)
            Exit Function

        End If

    Next lr

    Set GetByKey = Nothing

End Function

'------------------------------------------------------------------------------
' Belirtilen ayar anahtarýnýn mevcut olup olmadýðýný kontrol eder.
'------------------------------------------------------------------------------
Public Function Exists(ByVal SettingKey As String) As Boolean

    Exists = Not (GetByKey(SettingKey) Is Nothing)

End Function
'------------------------------------------------------------------------------
' Mevcut bir ayar kaydýný günceller.
'------------------------------------------------------------------------------
Public Function Update(ByVal Setting As clsSettings) As Boolean

    On Error GoTo ErrorHandler

    Dim lo As ListObject
    Dim lr As ListRow

    If Setting Is Nothing Then
        Err.Raise vbObjectError + 7302, _
                  "clsSettingsRepository.Update", _
                  "Setting nesnesi Nothing."
    End If

    Set lo = GetSettingsTable()

    For Each lr In lo.ListRows

        If StrComp( _
            Trim$(CStr(lr.Range.Cells(1, Col(COL_SETTING_KEY)).Value)), _
            Trim$(Setting.SettingKey), _
            vbTextCompare) = 0 Then

            WriteSettingToRow lr, Setting

            Update = True
            Exit Function

        End If

    Next lr

    ' Kayýt bulunamadý: teknik hata deðil, False döner.
    Update = False
    Exit Function

ErrorHandler:

    Update = False

    Err.Raise Err.Number, _
              "clsSettingsRepository.Update", _
              Err.Description

End Function
