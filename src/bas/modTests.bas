Attribute VB_Name = "modTests"
Option Explicit

'=========================================================
' OAA - Open Association Accounting
' Module : modTests
' Version: 1.0.0
'=========================================================

'---------------------------------------------------------
' Tüm Testleri Çalýþtýr
'---------------------------------------------------------
Public Sub RunAllTests()

    Debug.Print String(60, "=")
    Debug.Print APP_NAME
    Debug.Print "Test Baþlangýcý : " & Now
    Debug.Print String(60, "=")

    TestGlobals
    TestHelpers
    TestInstaller
    TestMain

    Debug.Print String(60, "=")
    Debug.Print "TÜM TESTLER TAMAMLANDI"
    Debug.Print String(60, "=")

    MsgBox "Tüm testler baþarýyla tamamlandý.", vbInformation

End Sub

'---------------------------------------------------------
' Globals Testi
'---------------------------------------------------------
Private Sub TestGlobals()

    Debug.Print ""
    Debug.Print "[TEST] Globals"

    Debug.Assert APP_NAME <> ""
    Debug.Assert APP_VERSION <> ""

    Debug.Print "   OK"

End Sub

'---------------------------------------------------------
' Helpers Testi
'---------------------------------------------------------
Private Sub TestHelpers()

    Debug.Print ""
    Debug.Print "[TEST] Helpers"

    Debug.Assert SheetExists(SHEET_SETTINGS) = True
    Debug.Assert SheetExists(SHEET_MEMBERS) = True

    Debug.Print "   OK"

End Sub

'---------------------------------------------------------
' Installer Testi
'---------------------------------------------------------
Private Sub TestInstaller()

    Debug.Print ""
    Debug.Print "[TEST] Installer"

    Debug.Assert TableExists(SHEET_SETTINGS, TBL_SETTINGS)
    Debug.Assert TableExists(SHEET_MEMBERS, TBL_MEMBERS)

    Debug.Print "   OK"

End Sub

'---------------------------------------------------------
' Main Testi
'---------------------------------------------------------
Private Sub TestMain()

    Debug.Print ""
    Debug.Print "[TEST] Main"

    Debug.Assert OAA_IsReady = True

    Debug.Print "   OK"

End Sub
'---------------------------------------------------------
' Ayarlar Tablosu Testi
'---------------------------------------------------------
Private Sub TestSettings()

    Debug.Print ""
    Debug.Print "[TEST] Settings"

    Debug.Assert TableExists(SHEET_SETTINGS, TBL_SETTINGS)

    Debug.Print "   OK"

End Sub

'---------------------------------------------------------
' Üyeler Tablosu Testi
'---------------------------------------------------------
Private Sub TestMembers()

    Debug.Print ""
    Debug.Print "[TEST] Members"

    Debug.Assert TableExists(SHEET_MEMBERS, TBL_MEMBERS)

    Debug.Print "   OK"

End Sub

'---------------------------------------------------------
' Framework Testi
'---------------------------------------------------------
Public Sub TestFramework()

    TestGlobals
    TestHelpers
    TestInstaller
    TestMain
    TestSettings
    TestMembers

    MsgBox "Framework testleri baþarýyla tamamlandý.", vbInformation

End Sub
