Attribute VB_Name = "modDueTests"
Option Explicit

'=========================================================
' OAA - Due Repository Tests
'=========================================================

Public Sub Test_AddDue()

    Dim repo As clsRepository
    Dim DueID As Long

    Set repo = New clsRepository

    DueID = repo.AddDue( _
                1, _
                "2026-07", _
                Date, _
                Date + 30, _
                500)

    If DueID > 0 Then

        MsgBox "Test baþarýlý." & vbCrLf & _
               "Oluþturulan Aidat ID : " & DueID, _
               vbInformation

    Else

        MsgBox "Test baþarýsýz.", vbCritical

    End If

    Set repo = Nothing

End Sub

'=========================================================
' Test - Find Due
'=========================================================
Public Sub Test_FindDue()

    Dim repo As clsRepository
    Dim rw As ListRow

    Set repo = New clsRepository

    Set rw = repo.FindDue(1)

    If rw Is Nothing Then

        MsgBox "Aidat bulunamadý.", vbCritical

    Else

        MsgBox "Aidat bulundu." & vbCrLf & _
               "Üye ID : " & rw.Range.Cells(1, 2).Value & vbCrLf & _
               "Tutar : " & rw.Range.Cells(1, 6).Value, vbInformation

    End If

    Set repo = Nothing

End Sub
'=========================================================
' Test - Update Due
'=========================================================
Public Sub Test_UpdateDue()

    Dim repo As clsRepository
    Dim Ok As Boolean

    Set repo = New clsRepository

    Ok = repo.UpdateDue( _
            1, _
            Date, _
            Date + 30, _
            500, _
            250, _
            "Kýsmi Ödeme", _
            "A000001", _
            "Ýlk ödeme")

    If Ok Then

        MsgBox "Aidat güncellendi.", vbInformation

    Else

        MsgBox "Güncelleme baþarýsýz.", vbCritical

    End If

    Set repo = Nothing

End Sub
'=========================================================
' Test - Delete Due
'=========================================================
Public Sub Test_DeleteDue()

    Dim repo As clsRepository

    Set repo = New clsRepository

    If repo.DeleteDue(1) Then

        MsgBox "Aidat silindi.", vbInformation

    Else

        MsgBox "Silme baþarýsýz.", vbCritical

    End If

    Set repo = Nothing

End Sub
