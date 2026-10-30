VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMemberRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Sýnýf  : clsMemberRepository
' Amaç   : Üye verilerinin tblMembers tablosundan okunmasý ve yazýlmasý
'
' Ýletiþim bilgileri:
'   GSM
'   E-Posta
'   Adres
'
' Telefon ve WhatsApp kullanýlmamaktadýr.
'
' ÖNEMLÝ:
' GSM NUMARASI SAYI DEÐÝLDÝR.
' GSM alaný Excel'de METÝN olarak saklanýr.
'
' Örnek:
'   05323352982
'
' baþýndaki 0 kesinlikle korunmalýdýr.
'==============================================================================


'==============================================================================
' ÜYE TABLOSUNU GETÝR
'==============================================================================

Private Function GetMembersTable() As ListObject

    Set GetMembersTable = _
        ThisWorkbook.Worksheets(SHEET_MEMBERS) _
            .ListObjects(TABLE_MEMBERS)

End Function


'==============================================================================
' SÜTUN NUMARASINI GETÝR
'==============================================================================

Private Function Col(ByVal ColumnName As String) As Long

    Dim lo As ListObject

    Set lo = GetMembersTable()

    Col = lo.ListColumns(ColumnName).Index

End Function


'==============================================================================
' GSM ALANINI METÝN OLARAK HAZIRLA
'==============================================================================

Private Sub PrepareGSMCell(ByVal Cell As Range, _
                           ByVal GSM As String)

    '----------------------------------------------------------------------
    ' GSM alanýný kesin olarak METÝN biçimine getirir.
    '----------------------------------------------------------------------

    Cell.NumberFormat = "@"

    '----------------------------------------------------------------------
    ' GSM deðerini String olarak yazar.
    '----------------------------------------------------------------------

    Cell.Value = CStr(Trim$(GSM))

End Sub


'==============================================================================
' TABLO SATIRINDAN CLSMEMBER OLUÞTUR
'==============================================================================

Private Function CreateMemberFromRow(ByVal lr As ListRow) As clsMember

    Dim Member As clsMember

    Set Member = New clsMember

    With Member

        '------------------------------------------------------------------
        ' Temel bilgiler
        '------------------------------------------------------------------

        .MemberID = _
            CLng(lr.Range.Cells(1, Col(COL_MEMBER_ID)).Value)

        .FirstName = _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_FIRST_NAME)).Value)

        .LastName = _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_LAST_NAME)).Value)

        .Status = _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_STATUS)).Value)


        '------------------------------------------------------------------
        ' Tarihler
        '----------------------------------------------------------------------

        If Trim$( _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_JOIN_DATE)).Value)) <> "" Then

            .JoinDate = _
                CDate(lr.Range.Cells(1, Col(COL_MEMBER_JOIN_DATE)).Value)

        End If


        If Trim$( _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_LEAVE_DATE)).Value)) <> "" Then

            .LeaveDate = _
                CDate(lr.Range.Cells(1, Col(COL_MEMBER_LEAVE_DATE)).Value)

        End If


        If Trim$( _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_BIRTH_DATE)).Value)) <> "" Then

            .BirthDate = _
                CDate(lr.Range.Cells(1, Col(COL_MEMBER_BIRTH_DATE)).Value)

        End If


        If Trim$( _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_MARRIAGE_DATE)).Value)) <> "" Then

            .MarriageDate = _
                CDate(lr.Range.Cells(1, Col(COL_MEMBER_MARRIAGE_DATE)).Value)

        End If


        '------------------------------------------------------------------
        ' Ýletiþim bilgileri
        '----------------------------------------------------------------------

        ' Telefon KULLANILMIYOR.

        ' GSM
        .Mobile = _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_GSM)).Value)

        ' E-Posta
        .Email = _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_EMAIL)).Value)

        ' Adres
        .Address = _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_ADDRESS)).Value)


        '------------------------------------------------------------------
        ' Aile / diðer bilgiler
        '----------------------------------------------------------------------

        .SpouseName = _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_SPOUSE_NAME)).Value)

        .Notes = _
            CStr(lr.Range.Cells(1, Col(COL_MEMBER_NOTES)).Value)

    End With


    Set CreateMemberFromRow = Member

End Function


'==============================================================================
' ÜYE ALANLARINI SATIRA YAZ (Add ve Update ortak noktasý)
' MemberID yazýlmaz; Add tarafýndan ayrýca yazýlýr.
'==============================================================================

Private Sub WriteMemberToRow(ByVal lr As ListRow, ByVal Member As clsMember)

    With lr.Range

        .Cells(1, Col(COL_MEMBER_FIRST_NAME)).Value = Member.FirstName
        .Cells(1, Col(COL_MEMBER_LAST_NAME)).Value = Member.LastName
        .Cells(1, Col(COL_MEMBER_STATUS)).Value = Member.Status

        .Cells(1, Col(COL_MEMBER_JOIN_DATE)).Value = Member.JoinDate
        .Cells(1, Col(COL_MEMBER_LEAVE_DATE)).Value = Member.LeaveDate
        .Cells(1, Col(COL_MEMBER_BIRTH_DATE)).Value = Member.BirthDate
        .Cells(1, Col(COL_MEMBER_MARRIAGE_DATE)).Value = Member.MarriageDate

        ' GSM sayý deðildir; METÝN olarak yazýlýr (baþýndaki 0 korunur).
        PrepareGSMCell .Cells(1, Col(COL_MEMBER_GSM)), Member.Mobile

        .Cells(1, Col(COL_MEMBER_EMAIL)).Value = Member.Email
        .Cells(1, Col(COL_MEMBER_ADDRESS)).Value = Member.Address

        .Cells(1, Col(COL_MEMBER_SPOUSE_NAME)).Value = Member.SpouseName
        .Cells(1, Col(COL_MEMBER_NOTES)).Value = Member.Notes

    End With

End Sub


'==============================================================================
' YENÝ ÜYE EKLE
'==============================================================================

Public Function Add(ByVal Member As clsMember) As Boolean

    On Error GoTo ErrorHandler

    Dim lo As ListObject
    Dim lr As ListRow

    If Member Is Nothing Then
        Err.Raise vbObjectError + 7201, _
                  "clsMemberRepository.Add", _
                  "Member nesnesi Nothing."
    End If

    Set lo = GetMembersTable()

    Set lr = lo.ListRows.Add

    lr.Range.Cells(1, Col(COL_MEMBER_ID)).Value = Member.MemberID

    WriteMemberToRow lr, Member

    Add = True

    Exit Function

ErrorHandler:

    Add = False

    Err.Raise Err.Number, _
              "clsMemberRepository.Add", _
              Err.Description

End Function


'==============================================================================
' ÜYE NUMARASINA GÖRE BUL
'==============================================================================

Public Function GetByID(ByVal MemberID As Long) As clsMember

    Dim lo As ListObject
    Dim lr As ListRow


    Set lo = GetMembersTable()


    For Each lr In lo.ListRows

        If CLng( _
            lr.Range.Cells(1, Col(COL_MEMBER_ID)).Value) = MemberID Then

            Set GetByID = _
                CreateMemberFromRow(lr)

            Exit Function

        End If

    Next lr


    Set GetByID = Nothing

End Function


'==============================================================================
' ÜYE GÜNCELLE
'==============================================================================

Public Function Update(ByVal Member As clsMember) As Boolean

    On Error GoTo ErrorHandler

    Dim lo As ListObject
    Dim lr As ListRow

    If Member Is Nothing Then
        Err.Raise vbObjectError + 7202, _
                  "clsMemberRepository.Update", _
                  "Member nesnesi Nothing."
    End If

    Set lo = GetMembersTable()

    For Each lr In lo.ListRows

        If CLng(lr.Range.Cells(1, Col(COL_MEMBER_ID)).Value) = _
           Member.MemberID Then

            WriteMemberToRow lr, Member

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
              "clsMemberRepository.Update", _
              Err.Description

End Function


'==============================================================================
' ÜYEYÝ PASÝF / ÝSTÝFA DURUMUNA GETÝR
'==============================================================================

Public Function SetPassive( _
        ByVal MemberID As Long, _
        ByVal LeaveDate As Date) As Boolean

    On Error GoTo ErrorHandler

    Dim lo As ListObject
    Dim lr As ListRow


    Set lo = GetMembersTable()


    For Each lr In lo.ListRows

        If CLng( _
            lr.Range.Cells(1, Col(COL_MEMBER_ID)).Value) = _
            MemberID Then


            '--------------------------------------------------------------
            ' Üye silinmez.
            ' Sadece durumu deðiþtirilir.
            '--------------------------------------------------------------

            lr.Range.Cells(1, Col(COL_MEMBER_STATUS)).Value = _
                MEMBER_STATUS_RESIGNED


            '--------------------------------------------------------------
            ' Ayrýlýþ tarihi
            '--------------------------------------------------------------

            lr.Range.Cells(1, Col(COL_MEMBER_LEAVE_DATE)).Value = _
                LeaveDate


            SetPassive = True

            Exit Function

        End If

    Next lr


    SetPassive = False

    Exit Function


ErrorHandler:

    SetPassive = False

    Err.Raise Err.Number, _
              "clsMemberRepository.SetPassive", _
              Err.Description

End Function


'==============================================================================
' ÜYE VAR MI?
'==============================================================================

Public Function Exists(ByVal MemberID As Long) As Boolean

    Dim lo As ListObject
    Dim lr As ListRow


    Set lo = GetMembersTable()


    For Each lr In lo.ListRows

        If Val( _
            lr.Range.Cells(1, Col(COL_MEMBER_ID)).Value) = _
            MemberID Then

            Exists = True

            Exit Function

        End If

    Next lr


    Exists = False

End Function


'==============================================================================
' SONRAKÝ ÜYE NUMARASINI GETÝR
'==============================================================================

Public Function GetNextMemberID() As Long

    Dim lo As ListObject
    Dim lr As ListRow

    Dim MaxID As Long


    Set lo = GetMembersTable()


    MaxID = 0


    If lo.ListRows.Count = 0 Then

        GetNextMemberID = 1

        Exit Function

    End If


    For Each lr In lo.ListRows

        If Val( _
            lr.Range.Cells(1, Col(COL_MEMBER_ID)).Value) > _
            MaxID Then

            MaxID = _
                Val(lr.Range.Cells(1, Col(COL_MEMBER_ID)).Value)

        End If

    Next lr


    GetNextMemberID = MaxID + 1

End Function


'==============================================================================
' TÜM ÜYELERÝ GETÝR
'==============================================================================

Public Function GetAll() As Collection

    Dim Members As New Collection

    Dim lo As ListObject
    Dim lr As ListRow


    Set lo = GetMembersTable()


    For Each lr In lo.ListRows

        Members.Add _
            CreateMemberFromRow(lr)

    Next lr


    Set GetAll = Members

End Function


'==============================================================================
' E-POSTA KONTROLÜ
'==============================================================================

Public Function EmailExists( _
        ByVal Email As String, _
        Optional ByVal ExcludeMemberID As Long = 0) As Boolean

    Dim lo As ListObject
    Dim lr As ListRow


    Set lo = GetMembersTable()


    If Trim$(Email) = "" Then

        EmailExists = False

        Exit Function

    End If


    For Each lr In lo.ListRows

        If ExcludeMemberID > 0 And _
           NzLongValue(lr.Range.Cells(1, Col(COL_MEMBER_ID)).Value) = ExcludeMemberID Then

            GoTo NextEmailRow

        End If

        If LCase$( _
            Trim$( _
                CStr( _
                    lr.Range.Cells(1, Col(COL_MEMBER_EMAIL)).Value))) = _
            LCase$(Trim$(Email)) Then


            EmailExists = True

            Exit Function

        End If

NextEmailRow:

    Next lr


    EmailExists = False

End Function


'==============================================================================
' GSM KONTROLÜ
'==============================================================================

Public Function GSMExists( _
        ByVal GSM As String, _
        Optional ByVal ExcludeMemberID As Long = 0) As Boolean

    Dim lo As ListObject
    Dim lr As ListRow


    Set lo = GetMembersTable()


    If Trim$(GSM) = "" Then

        GSMExists = False

        Exit Function

    End If


    For Each lr In lo.ListRows

        If ExcludeMemberID > 0 And _
           NzLongValue(lr.Range.Cells(1, Col(COL_MEMBER_ID)).Value) = ExcludeMemberID Then

            GoTo NextGsmRow

        End If

        If Trim$( _
            CStr( _
                lr.Range.Cells(1, Col(COL_MEMBER_GSM)).Value)) = _
            Trim$(GSM) Then


            GSMExists = True

            Exit Function

        End If

NextGsmRow:

    Next lr


    GSMExists = False

End Function
