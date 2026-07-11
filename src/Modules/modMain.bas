Attribute VB_Name = "modMain"
Option Explicit

'=========================================================
' OAA - Open Association Accounting
' Module : modMain
' Version: 1.0.0
'=========================================================

'---------------------------------------------------------
' Uygulamayý Baþlat
'---------------------------------------------------------
Public Sub OAA_Start()

    On Error GoTo ErrHandler

    BeginUpdate

    InitializeGlobals

    If Not SystemInstalled Then

        OAA_Install

    End If

    EndUpdate

    ShowInfo APP_NAME & " baþlatýldý."

    Exit Sub

ErrHandler:

    EndUpdate

    ShowError Err

End Sub

'---------------------------------------------------------
' Sistem Kurulu mu?
'---------------------------------------------------------
Public Function SystemInstalled() As Boolean

    If Not SheetExists(SHEET_SETTINGS) Then Exit Function

    If Not SheetExists(SHEET_MEMBERS) Then Exit Function

    If Not TableExists(SHEET_SETTINGS, TBL_SETTINGS) Then Exit Function

    If Not TableExists(SHEET_MEMBERS, TBL_MEMBERS) Then Exit Function

    SystemInstalled = True

End Function

'---------------------------------------------------------
' Sistemi Yeniden Baþlat
'---------------------------------------------------------
Public Sub RestartSystem()

    ReleaseGlobals

    InitializeGlobals

End Sub

'---------------------------------------------------------
' Sistem Bilgisi
'---------------------------------------------------------
Public Sub AboutOAA()

    Dim Msg As String

    Msg = APP_NAME & vbCrLf
    Msg = Msg & "Sürüm : " & APP_VERSION & vbCrLf
    Msg = Msg & "Excel : " & Application.Version

    MsgBox Msg, vbInformation, APP_NAME

End Sub
'---------------------------------------------------------
' Sistem Bilgileri
'---------------------------------------------------------
Public Function OAA_Version() As String

    OAA_Version = APP_VERSION

End Function

'---------------------------------------------------------
' Sistem Hazýr mý?
'---------------------------------------------------------
Public Function OAA_IsReady() As Boolean

    OAA_IsReady = SystemInstalled

End Function

'---------------------------------------------------------
' Sistemi Kapat
'---------------------------------------------------------
Public Sub OAA_Shutdown()

    On Error Resume Next

    ReleaseGlobals

    Application.StatusBar = False

End Sub

'---------------------------------------------------------
' Sistem Durumu
'---------------------------------------------------------
Public Sub OAA_Status()

    Dim s As String

    s = "Program : " & APP_NAME & vbCrLf
    s = s & "Sürüm : " & APP_VERSION & vbCrLf
    s = s & "Kurulu : " & IIf(SystemInstalled, "Evet", "Hayýr")

    MsgBox s, vbInformation, APP_NAME

End Sub
'---------------------------------------------------------
' Üye Yönetimini Aç
'---------------------------------------------------------
Public Sub OpenMembers()

    frmMembers.Show

End Sub
