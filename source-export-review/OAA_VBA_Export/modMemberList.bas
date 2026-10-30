Attribute VB_Name = "modMemberList"
Option Explicit

'==============================================================================
' OAA v2
' Modül  : modMemberList
' Amaç   : Üye listesinin yenilenmesi ve görüntülenmesi
'==============================================================================

'------------------------------------------------------------------------------
' Üye listesini veritabanýndaki güncel üye kayýtlarýndan yeniden oluþturur.
'------------------------------------------------------------------------------
Public Sub RefreshMemberList()

    Dim ws As Worksheet
    Dim lo As ListObject

    Dim Members As Collection
    Dim Member As clsMember

    Dim rw As ListRow

    Dim Service As clsMemberService


    '----------------------------------------------------------------------
    ' Üye listesi çalýþma sayfasýný ve tabloyu al.
    '----------------------------------------------------------------------

    Set ws = ThisWorkbook.Worksheets("Üye Listesi")

    Set lo = ws.ListObjects("tblMemberList")


    '----------------------------------------------------------------------
    ' Üye servis sýnýfýný oluþtur.
    '----------------------------------------------------------------------

    Set Service = New clsMemberService


    '----------------------------------------------------------------------
    ' Repository/Service üzerinden tüm üyeleri al.
    '----------------------------------------------------------------------

    Set Members = Service.GetAllMembers


    '----------------------------------------------------------------------
    ' Mevcut listeyi temizle.
    '----------------------------------------------------------------------

    If Not lo.DataBodyRange Is Nothing Then

        lo.DataBodyRange.Delete

    End If


    '----------------------------------------------------------------------
    ' Üyeleri liste tablosuna ekle.
    '
    ' tblMemberList kolon sýrasý:
    '
    ' 1  MemberID
    ' 2  FirstName
    ' 3  LastName
    ' 4  Status
    ' 5  JoinDate
    ' 6  LeaveDate
    ' 7  BirthDate
    ' 8  MarriageDate
    ' 9  GSM
    ' 10 Email
    ' 11 Address
    ' 12 SpouseName
    ' 13 Notes
    '----------------------------------------------------------------------

    For Each Member In Members

        Set rw = lo.ListRows.Add

        With rw.Range

            .Cells(1, 1).Value = _
                Format$(Member.MemberID, "000000")

            .Cells(1, 2).Value = _
                Member.FirstName

            .Cells(1, 3).Value = _
                Member.LastName

            .Cells(1, 4).Value = _
                Member.Status

            .Cells(1, 5).Value = _
                Member.JoinDate

            .Cells(1, 6).Value = _
                Member.LeaveDate

            .Cells(1, 7).Value = _
                Member.BirthDate

            .Cells(1, 8).Value = _
                Member.MarriageDate

            ' GSM
            .Cells(1, 9).Value = _
                Member.Mobile

            ' E-posta
            .Cells(1, 10).Value = _
                Member.Email

            ' Adres
            .Cells(1, 11).Value = _
                Member.Address

            ' Eþ adý
            .Cells(1, 12).Value = _
                Member.SpouseName

            ' Notlar
            .Cells(1, 13).Value = _
                Member.Notes

        End With

    Next Member


End Sub

