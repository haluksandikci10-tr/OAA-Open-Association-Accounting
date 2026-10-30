VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMemberSearch 
   Caption         =   "Üye Arama"
   ClientHeight    =   8745.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "frmMemberSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMemberSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' FORM    : frmMemberSearch
' AMAÇ    : Üye arama ve üye listesini görüntüleme
'
' Liste sütunlarý:
'
'   1 - Üye No
'   2 - Ad Soyad
'   3 - GSM
'   4 - E-Posta
'   5 - Durum
'
' Telefon alaný kullanýlmamaktadýr.
'==============================================================================


'==============================================================================
' FORM DEÐÝÞKENLERÝ
'==============================================================================

Private mService As clsMemberService


'==============================================================================
' FORM AÇILIÞI
'==============================================================================

Private Sub UserForm_Initialize()

    On Error GoTo ErrorHandler

    '----------------------------------------------------------------------
    ' Servis oluþturulur.
    '----------------------------------------------------------------------

    Set mService = New clsMemberService


    '----------------------------------------------------------------------
    ' Liste özellikleri
    '----------------------------------------------------------------------

    ConfigureMemberList


    '----------------------------------------------------------------------
    ' Arama kutusunu temizle
    '----------------------------------------------------------------------

    txtSearch.Value = ""


    '----------------------------------------------------------------------
    ' Üyeleri yükle
    '----------------------------------------------------------------------

    LoadMembers


    Exit Sub


ErrorHandler:

    MsgBox _
        "Üye arama formu baþlatýlýrken hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbCritical, _
        "OAA - Hata"

End Sub


'==============================================================================
' LÝSTE KONTROLÜNÜ YAPILANDIR
'==============================================================================

Private Sub ConfigureMemberList()

    With lstMembers

        '------------------------------------------------------------------
        ' Beþ sütun:
        '
        ' 0 = Üye No
        ' 1 = Ad Soyad
        ' 2 = GSM
        ' 3 = E-Posta
        ' 4 = Durum
        '------------------------------------------------------------------

        .ColumnCount = 5

        '------------------------------------------------------------------
        ' ListBox kendi baþlýklarýný kullanmayacak.
        '
        ' Form üzerindeki baþlýk etiketleri kullanýlacaktýr.
        '------------------------------------------------------------------

        .ColumnHeads = False

        '------------------------------------------------------------------
        ' Sütun geniþlikleri
        '------------------------------------------------------------------

        .ColumnWidths = _
            "60 pt;150 pt;100 pt;170 pt;70 pt"

        .Clear

    End With

End Sub


'==============================================================================
' ÜYE LÝSTESÝNÝ YÜKLE
'==============================================================================

Private Sub LoadMembers()

    On Error GoTo ErrorHandler

    Dim Members As Collection
    Dim Member As clsMember

    Dim SearchText As String
    Dim FullName As String
    Dim GSM As String
    Dim Email As String
    Dim Status As String
    Dim MemberNo As String


    '----------------------------------------------------------------------
    ' Arama metnini al
    '----------------------------------------------------------------------

    SearchText = _
        LCase$(Trim$(CStr(txtSearch.Value)))


    '----------------------------------------------------------------------
    ' Servisten tüm üyeleri al
    '----------------------------------------------------------------------

    Set Members = mService.GetAllMembers


    '----------------------------------------------------------------------
    ' Listeyi temizle
    '----------------------------------------------------------------------

    lstMembers.Clear


    '----------------------------------------------------------------------
    ' Üyeler yoksa çýk
    '----------------------------------------------------------------------

    If Members Is Nothing Then Exit Sub


    '----------------------------------------------------------------------
    ' Üyeleri listeye ekle
    '----------------------------------------------------------------------

    For Each Member In Members

        '--------------------------------------------------------------
        ' Gösterilecek bilgiler hazýrlanýr.
        '--------------------------------------------------------------

        MemberNo = _
            Format$(Member.MemberID, "000000")

        FullName = _
            Trim$(Member.FirstName & " " & Member.LastName)

        GSM = _
            CStr(Member.Mobile)

        Email = _
            CStr(Member.Email)

        Status = _
            CStr(Member.Status)


        '--------------------------------------------------------------
        ' Arama filtresi
        '
        ' Arama kutusu boþsa bütün üyeler gösterilir.
        '
        ' Arama;
        '   Üye No
        '   Ad
        '   Soyad
        '   GSM
        '   E-Posta
        '   Durum
        '
        ' alanlarýnda yapýlýr.
        '--------------------------------------------------------------

        If SearchText = "" Then

            AddMemberToList _
                MemberNo, _
                FullName, _
                GSM, _
                Email, _
                Status

        Else

            If _
                InStr(1, LCase$(MemberNo), SearchText, vbTextCompare) > 0 _
                Or _
                InStr(1, LCase$(FullName), SearchText, vbTextCompare) > 0 _
                Or _
                InStr(1, LCase$(GSM), SearchText, vbTextCompare) > 0 _
                Or _
                InStr(1, LCase$(Email), SearchText, vbTextCompare) > 0 _
                Or _
                InStr(1, LCase$(Status), SearchText, vbTextCompare) > 0 Then


                AddMemberToList _
                    MemberNo, _
                    FullName, _
                    GSM, _
                    Email, _
                    Status

            End If

        End If

    Next Member


    Exit Sub


ErrorHandler:

    MsgBox _
        "Üye listesi yüklenirken hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbCritical, _
        "OAA - Hata"

End Sub


'==============================================================================
' ÜYEYÝ LÝSTEYE EKLE
'==============================================================================

Private Sub AddMemberToList( _
        ByVal MemberNo As String, _
        ByVal FullName As String, _
        ByVal GSM As String, _
        ByVal Email As String, _
        ByVal Status As String)


    Dim RowIndex As Long


    '----------------------------------------------------------------------
    ' Yeni satýr oluþtur
    '----------------------------------------------------------------------

    lstMembers.AddItem MemberNo


    RowIndex = _
        lstMembers.ListCount - 1


    '----------------------------------------------------------------------
    ' Ad Soyad
    '----------------------------------------------------------------------

    lstMembers.List(RowIndex, 1) = _
        FullName


    '----------------------------------------------------------------------
    ' GSM
    '
    ' Telefon alaný YOKTUR.
    '----------------------------------------------------------------------

    lstMembers.List(RowIndex, 2) = _
        CStr(GSM)


    '----------------------------------------------------------------------
    ' E-Posta
    '----------------------------------------------------------------------

    lstMembers.List(RowIndex, 3) = _
        Email


    '----------------------------------------------------------------------
    ' Durum
    '----------------------------------------------------------------------

    lstMembers.List(RowIndex, 4) = _
        Status

End Sub


'==============================================================================
' ARAMA KUTUSU DEÐÝÞTÝÐÝNDE
'==============================================================================

Private Sub txtSearch_Change()

    LoadMembers

End Sub


'==============================================================================
' LÝSTEDEN ÜYE SEÇÝLDÝÐÝNDE
'==============================================================================

Private Sub lstMembers_Click()

    '----------------------------------------------------------------------
    ' Þimdilik herhangi bir iþlem yapýlmýyor.
    '
    ' Üye açma iþlemi çift týklama veya AÇ düðmesi ile yapýlýr.
    '----------------------------------------------------------------------

End Sub


'==============================================================================
' ÜYEYÝ ÇÝFT TIKLAYARAK AÇ
'==============================================================================

Private Sub lstMembers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    OpenSelectedMember

End Sub


'==============================================================================
' AÇ DÜÐMESÝ
'==============================================================================

Private Sub cmdOpen_Click()

    OpenSelectedMember

End Sub


'==============================================================================
' SEÇÝLÝ ÜYEYÝ AÇ
'==============================================================================

Private Sub OpenSelectedMember()

    On Error GoTo ErrorHandler

    Dim MemberID As Long


    '----------------------------------------------------------------------
    ' Üye seçilmiþ mi?
    '----------------------------------------------------------------------

    If lstMembers.ListIndex < 0 Then

        MsgBox _
            "Lütfen açmak istediðiniz üyeyi seçiniz.", _
            vbExclamation, _
            "OAA - Üye Arama"

        Exit Sub

    End If


    '----------------------------------------------------------------------
    ' ListBox'ýn ilk sütunundan üye numarasýný al.
    '----------------------------------------------------------------------

    MemberID = _
        CLng(lstMembers.List(lstMembers.ListIndex, 0))


    '----------------------------------------------------------------------
    ' Üye numarasý geçerli mi?
    '----------------------------------------------------------------------

    If MemberID <= 0 Then

        MsgBox _
            "Geçerli bir üye numarasý bulunamadý.", _
            vbExclamation, _
            "OAA - Üye Arama"

        Exit Sub

    End If


    '----------------------------------------------------------------------
    ' Üye kartýný aç.
    '----------------------------------------------------------------------

    frmMemberCard.OpenMember MemberID


    Exit Sub


ErrorHandler:

    MsgBox _
        "Üye açýlýrken hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbCritical, _
        "OAA - Hata"

End Sub


'==============================================================================
' YENÝ ÜYE DÜÐMESÝ
'==============================================================================

Private Sub cmdNew_Click()

    On Error GoTo ErrorHandler


    '----------------------------------------------------------------------
    ' Yeni üye kartýný aç
    '----------------------------------------------------------------------

    frmMemberCard.Show


    '----------------------------------------------------------------------
    ' Form kapatýlmadýysa liste daha sonra yenilenecektir.
    '----------------------------------------------------------------------


    Exit Sub


ErrorHandler:

    MsgBox _
        "Yeni üye formu açýlýrken hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata: " & Err.Description, _
        vbCritical, _
        "OAA - Hata"

End Sub


'==============================================================================
' KAPAT DÜÐMESÝ
'==============================================================================

Private Sub cmdClose_Click()

    Unload Me

End Sub


'==============================================================================
' FORM KAPATILIRKEN
'==============================================================================

Private Sub UserForm_Terminate()

    Set mService = Nothing

End Sub

