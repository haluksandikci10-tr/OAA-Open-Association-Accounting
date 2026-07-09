VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMembers 
   Caption         =   "Üye Yönetimi"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   14160
   OleObjectBlob   =   "frmMembers.frx":0000
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRepository As clsRepository

Private Sub UserForm_Initialize()

    Set mRepository = New clsRepository

    LoadStatusList

    ClearForm

    RefreshMemberList

End Sub

Private Sub UserForm_Terminate()

    Set mRepository = Nothing

End Sub
'---------------------------------------------------------
' Durum listesini yükle
'---------------------------------------------------------
Private Sub LoadStatusList()

    With cmbStatus

        .Clear

        .AddItem MEMBER_ACTIVE
        .AddItem MEMBER_PASSIVE
        .AddItem MEMBER_LEAVE

        .ListIndex = 0

    End With

End Sub
'---------------------------------------------------------
' Formu Temizle
'---------------------------------------------------------
Private Sub ClearForm()

    txtMemberID.Value = ""

    txtFullName.Value = ""

    txtPhone.Value = ""

    txtEmail.Value = ""

    txtJoinDate.Value = Format(Date, "dd.mm.yyyy")

    txtDue.Value = Format(0, "#,##0.00")

    If cmbStatus.ListCount > 0 Then
        cmbStatus.ListIndex = 0
    End If

    txtFullName.SetFocus

End Sub
'=========================================================
' FORM HELPERS
'=========================================================

Private Sub RefreshMemberList()

    Dim tbl As ListObject
    Dim r As ListRow

    Set tbl = mRepository.Members

    lstMembers.Clear

    If tbl Is Nothing Then Exit Sub

   

    For Each r In tbl.ListRows

        With lstMembers
            If Trim$(CStr(r.Range.Cells(1, 1).Value)) = "" Then
    GoTo NextRow

End If
            .AddItem CStr(r.Range.Cells(1, 1).Value)
           
            .List(.ListCount - 1, 1) = CStr(r.Range.Cells(1, 2).Value)
            .List(.ListCount - 1, 2) = CStr(r.Range.Cells(1, 3).Value)
            .List(.ListCount - 1, 3) = CStr(r.Range.Cells(1, 4).Value)
            .List(.ListCount - 1, 4) = CStr(r.Range.Cells(1, 5).Value)
            .List(.ListCount - 1, 5) = CStr(r.Range.Cells(1, 6).Value)
        End With
        
NextRow:
    Next r

    UpdateCounters

End Sub

'---------------------------------------------------------

Private Sub UpdateCounters()

    lblTotalMembers.Caption = _
        "Toplam Üye : " & _
        mRepository.MemberCount

    lblActiveMembers.Caption = _
        "Aktif : " & _
        mRepository.ActiveMemberCount

    lblPassiveMembers.Caption = _
        "Pasif : " & _
        mRepository.PassiveMemberCount

End Sub

'---------------------------------------------------------

Private Function SelectedMemberID() As Long

    If lstMembers.ListIndex = -1 Then Exit Function

    SelectedMemberID = _
        CLng(lstMembers.List(lstMembers.ListIndex, 0))

End Function

'---------------------------------------------------------

Private Sub LoadSelectedMember()

    Dim rw As ListRow

    If lstMembers.ListIndex = -1 Then Exit Sub

    Set rw = _
        mRepository.FindMember(SelectedMemberID)

    If rw Is Nothing Then Exit Sub

    txtMemberID.Value = rw.Range.Cells(1, 1).Value
    txtFullName.Value = rw.Range.Cells(1, 2).Value
    cmbStatus.Value = rw.Range.Cells(1, 3).Value
    txtPhone.Value = rw.Range.Cells(1, 4).Value
    txtEmail.Value = rw.Range.Cells(1, 5).Value
    txtJoinDate.Value = rw.Range.Cells(1, 6).Value
    txtDue.Value = rw.Range.Cells(1, 7).Value

End Sub

'=========================================================
' BUTTON EVENTS
'=========================================================

Private Sub cmdNew_Click()

    ClearForm

End Sub

'---------------------------------------------------------

Private Sub cmdClear_Click()

    ClearForm

End Sub

'---------------------------------------------------------

Private Sub cmdClose_Click()

    Unload Me

End Sub

'---------------------------------------------------------

Private Sub cmdSave_Click()

    Dim ID As Long

    If Trim$(txtFullName.Value) = "" Then

        MsgBox "Ad Soyad boþ olamaz.", vbExclamation

        txtFullName.SetFocus

        Exit Sub

    End If

    ID = mRepository.AddMember( _
            txtFullName.Value, _
            cmbStatus.Value)

    If ID = 0 Then

        MsgBox "Kayýt yapýlamadý.", vbCritical

        Exit Sub

    End If

    SetMemberDetails ID

    RefreshMemberList

ClearForm

MsgBox "Üye baþarýyla kaydedildi.", vbInformation
End Sub

'---------------------------------------------------------

Private Sub cmdUpdate_Click()

    If txtMemberID.Value = "" Then Exit Sub

    If mRepository.UpdateMember( _
            CLng(txtMemberID.Value), _
            txtFullName.Value, _
            cmbStatus.Value, _
            txtPhone.Value, _
            txtEmail.Value) Then

       RefreshMemberList

ClearForm

MsgBox "Kayýt güncellendi.", vbInformation
    Else

        MsgBox "Güncelleme yapýlamadý.", vbCritical

    End If

End Sub

'---------------------------------------------------------

Private Sub cmdDelete_Click()

    If txtMemberID.Value = "" Then Exit Sub

    If MsgBox( _
        "Bu üyeyi silmek istiyor musunuz?", _
        vbYesNo + vbQuestion) = vbNo Then Exit Sub

    If mRepository.DeleteMember(CLng(txtMemberID.Value)) Then

        RefreshMemberList

        ClearForm

        MsgBox "Üye silindi.", vbInformation

    Else

        MsgBox "Silme iþlemi baþarýsýz.", vbCritical

    End If

End Sub

'=========================================================
' HELPERS
'=========================================================

Private Sub SetMemberDetails(ByVal MemberID As Long)

    Dim rw As ListRow

    Set rw = mRepository.FindMember(MemberID)

    If rw Is Nothing Then Exit Sub

    rw.Range.Cells(1, 4).Value = txtPhone.Value
    rw.Range.Cells(1, 5).Value = txtEmail.Value
    rw.Range.Cells(1, 6).Value = txtJoinDate.Value
    rw.Range.Cells(1, 7).Value = Val(txtDue.Value)

End Sub

Private Sub SelectMember(ByVal MemberID As Long)

    Dim i As Long
    Dim v As Variant

    If lstMembers.ListCount = 0 Then Exit Sub

    For i = 0 To lstMembers.ListCount - 1

        v = lstMembers.List(i, 0)

        If Len(Trim$(CStr(v))) > 0 Then

            If IsNumeric(v) Then

                If CLng(v) = MemberID Then

                    lstMembers.ListIndex = i
                    Exit Sub

                End If

            End If

        End If

    Next i

End Sub

'=========================================================
' LIST EVENTS
'=========================================================

Private Sub lstMembers_Click()

    LoadSelectedMember

End Sub

'---------------------------------------------------------

Private Sub lstMembers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    LoadSelectedMember

End Sub

'=========================================================
' SEARCH
'=========================================================

Private Sub txtSearch_Change()

    Dim i As Long
    Dim SearchText As String

    SearchText = LCase$(Trim$(txtSearch.Text))

    If SearchText = "" Then

        RefreshMemberList

        Exit Sub

    End If

    For i = lstMembers.ListCount - 1 To 0 Step -1

        If InStr(1, _
                 LCase$(lstMembers.List(i, 1)), _
                 SearchText, _
                 vbTextCompare) = 0 Then

            lstMembers.RemoveItem i

        End If

    Next i

End Sub

'=========================================================
' VALIDATION
'=========================================================

Private Function ValidateForm() As Boolean

    If Trim$(txtFullName.Text) = "" Then

        MsgBox "Ad Soyad boþ olamaz.", vbExclamation

        txtFullName.SetFocus

        Exit Function

    End If

    If cmbStatus.ListIndex = -1 Then

        MsgBox "Durum seçiniz.", vbExclamation

        cmbStatus.SetFocus

        Exit Function

    End If

    ValidateForm = True

End Function

'=========================================================
' KEY EVENTS
'=========================================================

Private Sub txtDue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii

        Case 48 To 57
        Case Asc(",")
        Case Asc(".")
        Case 8

        Case Else

            KeyAscii = 0

    End Select

End Sub

'---------------------------------------------------------

Private Sub txtPhone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii

        Case 48 To 57
        Case Asc("+")
        Case Asc("(")
        Case Asc(")")
        Case Asc("-")
        Case 8
        Case 32

        Case Else

            KeyAscii = 0

    End Select

End Sub

