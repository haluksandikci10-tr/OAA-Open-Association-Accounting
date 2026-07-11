Attribute VB_Name = "modInstaller"
Option Explicit

'=========================================================
' OAA - Open Association Accounting
' Module : modInstaller
' Version: 1.0.0
'=========================================================

'---------------------------------------------------------
' Ana Kurulum
'---------------------------------------------------------
Public Sub OAA_Install()

    On Error GoTo ErrHandler

    BeginUpdate

    InstallSheets

    InstallTables

    EndUpdate

    MsgBox "Kurulum baÅŸarÄ±yla tamamlandÄ±.", vbInformation

    Exit Sub

ErrHandler:

    EndUpdate

    ShowError Err

End Sub

'---------------------------------------------------------
' SayfalarÄ± OluÅŸtur
'---------------------------------------------------------
Private Sub InstallSheets()

    CreateSheet SHEET_DASHBOARD
    CreateSheet SHEET_BOOK
    CreateSheet SHEET_CASH
    CreateSheet SHEET_BANK
    CreateSheet SHEET_MEMBERS
    CreateSheet SHEET_DUES
    CreateSheet SHEET_INCOME
    CreateSheet SHEET_EXPENSE
    CreateSheet SHEET_RECEIPTS
    CreateSheet SHEET_REPORTS
    CreateSheet SHEET_SETTINGS
    CreateSheet SHEET_LOG

End Sub

'---------------------------------------------------------
' TablolarÄ± OluÅŸtur
'---------------------------------------------------------
Private Sub InstallTables()

    CreateSettingsTable

    CreateMembersTable

    CreateDuesTable

End Sub
'---------------------------------------------------------
' Ayarlar Tablosu
'---------------------------------------------------------
Private Sub CreateSettingsTable()

    Dim ws As Worksheet

    Set ws = GetWorksheet(SHEET_SETTINGS)

    If TableExists(SHEET_SETTINGS, TBL_SETTINGS) Then Exit Sub

    ws.Range("A1:B1").Value = Array("Key", "Value")

    ws.Range("A2").Value = "AssociationName"
    ws.Range("A3").Value = "ReceiptSeries"
    ws.Range("A4").Value = "ReceiptNumber"
    ws.Range("A5").Value = "Currency"
    ws.Range("A6").Value = "DefaultBank"

    ws.Range("B2").Value = ""
    ws.Range("B3").Value = DEFAULT_RECEIPT_SERIES
    ws.Range("B4").Value = 1
    ws.Range("B5").Value = CUR_TRY
    ws.Range("B6").Value = DEFAULT_BANK

    CreateExcelTable ws, TBL_SETTINGS, ws.Range("A1:B6")

End Sub

'---------------------------------------------------------
' Üyeler Tablosu
'---------------------------------------------------------
Private Sub CreateMembersTable()

    Dim ws As Worksheet

    Set ws = GetWorksheet(SHEET_MEMBERS)

    If TableExists(SHEET_MEMBERS, TBL_MEMBERS) Then Exit Sub

    ws.Range("A1:J1").Value = Array( _
        "ID", _
        "Ad Soyad", _
        "Durum", _
        "Telefon", _
        "E-Posta", _
        "Katýlým Tarihi", _
        "Aidat", _
        "Borç", _
        "Alacak", _
        "Not")

    CreateExcelTable ws, TBL_MEMBERS, ws.Range("A1:J2")

End Sub
'---------------------------------------------------------
' Log Tablosu
'---------------------------------------------------------
Private Sub CreateLogTable()

    Dim ws As Worksheet

    Set ws = GetWorksheet(SHEET_LOG)

    If TableExists(SHEET_LOG, TBL_LOG) Then Exit Sub

    ws.Range("A1:E1").Value = Array( _
        "Tarih", _
        "Seviye", _
        "Modul", _
        "Mesaj", _
        "Kullanici")

    CreateExcelTable ws, TBL_LOG, ws.Range("A1:E2")

End Sub

'---------------------------------------------------------
' Aidatlar Tablosu
'---------------------------------------------------------
Private Sub CreateDuesTable()
MsgBox "CreateDuesTable çalýþtý"
    Dim ws As Worksheet

    Set ws = GetWorksheet(SHEET_DUES)

    If TableExists(SHEET_DUES, TBL_DUES) Then Exit Sub

    ws.Range("A1:K1").Value = Array( _
        "ID", _
        "MemberID", _
        "Donem", _
        "TahakkukTarihi", _
        "VadeTarihi", _
        "Tutar", _
        "Odenen", _
        "Bakiye", _
        "Durum", _
        "MakbuzNo", _
        "Aciklama")

    CreateExcelTable ws, TBL_DUES, ws.Range("A1:K2")
    MsgBox "Tablo oluþturuldu: " & ws.ListObjects.Count
End Sub

