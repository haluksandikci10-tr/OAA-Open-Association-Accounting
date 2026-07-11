VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDues 
   Caption         =   "Aidat Yönetimi"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14955
   OleObjectBlob   =   "frmDues.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRepository As clsRepository

'=========================================================
' FORM INITIALIZE
'=========================================================

Private Sub UserForm_Initialize()

    Set mRepository = New clsRepository

    LoadMemberList

    LoadPeriodList

    LoadStatusList

    ClearForm
    RefreshDueList
End Sub
' Üye Listesini Yükle
'---------------------------------------------------------
Private Sub LoadMemberList()

    Dim tbl As ListObject
    Dim rw As ListRow

    cmbMember.Clear

    Set tbl = mRepository.Members

    If tbl Is Nothing Then Exit Sub

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For Each rw In tbl.ListRows

        If Trim$(CStr(rw.Range.Cells(1, 2).Value)) <> "" Then

            cmbMember.AddItem rw.Range.Cells(1, 2).Value

            cmbMember.List(cmbMember.ListCount - 1, 1) = _
                rw.Range.Cells(1, 1).Value

        End If

    Next rw

End Sub
'---------------------------------------------------------
' Dönem Listesi
'---------------------------------------------------------
Private Sub LoadPeriodList()

    Dim Y As Integer
    Dim M As Integer
    Dim D As String

    cmbPeriod.Clear

    For Y = Year(Date) - 5 To Year(Date) + 1

        For M = 1 To 12

            D = Format(DateSerial(Y, M, 1), "yyyy-mm")

            cmbPeriod.AddItem D

        Next M

    Next Y

    cmbPeriod.Value = Format(Date, "yyyy-mm")

End Sub
'---------------------------------------------------------
' Durum Listesini Yükle
'---------------------------------------------------------
Private Sub LoadStatusList()

    With cmbStatus

        .Clear

        .AddItem "Bekliyor"
        .AddItem "Kýsmi Ödeme"
        .AddItem "Ödendi"

        .ListIndex = 0

    End With

End Sub
'---------------------------------------------------------
' Formu Temizle
'---------------------------------------------------------
Private Sub ClearForm()

    If cmbMember.ListCount > 0 Then
        cmbMember.ListIndex = -1
    End If

    cmbPeriod.Value = Format(Date, "yyyy-mm")

    txtDueDate.Value = Format(Date, "dd.mm.yyyy")

    txtDueLimit.Value = Format(Date + 30, "dd.mm.yyyy")

    txtAmount.Value = Format(0, "#,##0.00")

    txtPaid.Value = Format(0, "#,##0.00")

    txtBalance.Value = Format(0, "#,##0.00")

    If cmbStatus.ListCount > 0 Then
        cmbStatus.ListIndex = 0
    End If

    txtReceiptNo.Value = ""

    txtDescription.Value = ""

    cmbMember.SetFocus

End Sub
'---------------------------------------------------------
' Bakiyeyi Hesapla
'---------------------------------------------------------
Private Sub CalculateBalance()

    Dim Amount As Double
    Dim Paid As Double

    Amount = Val(Replace(txtAmount.Value, ",", "."))
    Paid = Val(Replace(txtPaid.Value, ",", "."))

    txtBalance.Value = Format(Amount - Paid, "#,##0.00")

End Sub
'---------------------------------------------------------
' Aidat Listesini Yenile
'---------------------------------------------------------
Private Sub RefreshDueList()

    Dim tbl As ListObject
    Dim rw As ListRow

    Set tbl = mRepository.Dues

    lstDues.Clear

    If tbl Is Nothing Then Exit Sub

    If tbl.DataBodyRange Is Nothing Then
        UpdateCounters
        Exit Sub
    End If

    lstDues.ColumnCount = 8
    lstDues.ColumnWidths = "35 pt;90 pt;55 pt;60 pt;60 pt;60 pt;70 pt;0 pt"

    For Each rw In tbl.ListRows

        With lstDues

            .AddItem CStr(rw.Range.Cells(1, 1).Value)      ' ID
            .List(.ListCount - 1, 1) = CStr(rw.Range.Cells(1, 2).Value) ' MemberID
            .List(.ListCount - 1, 2) = CStr(rw.Range.Cells(1, 3).Value) ' Dönem
            .List(.ListCount - 1, 3) = Format(rw.Range.Cells(1, 6).Value, "#,##0.00") ' Tutar
            .List(.ListCount - 1, 4) = Format(rw.Range.Cells(1, 7).Value, "#,##0.00") ' Ödenen
            .List(.ListCount - 1, 5) = Format(rw.Range.Cells(1, 8).Value, "#,##0.00") ' Bakiye
            .List(.ListCount - 1, 6) = CStr(rw.Range.Cells(1, 9).Value) ' Durum
            .List(.ListCount - 1, 7) = CStr(rw.Range.Cells(1, 10).Value) ' Makbuz

        End With

    Next rw

    UpdateCounters

End Sub
'---------------------------------------------------------
' Sayaçlarý Güncelle
'---------------------------------------------------------
Private Sub UpdateCounters()

    lblTotal.Caption = _
        "Toplam Aidat : " & mRepository.DueCount

    lblPending.Caption = _
        "Bekleyen : 0"

    lblPaidTotal.Caption = _
        "Tahsil Edilen : 0"

End Sub
Private Sub txtAmount_Change()

    CalculateBalance

End Sub
Private Sub UserForm_Terminate()

    Set mRepository = Nothing

End Sub
Private Sub txtPaid_Change()

    CalculateBalance

End Sub

Private Sub cmdAccrue_Click()

    Dim DueID As Long

    If Not ValidateDueForm Then Exit Sub

    DueID = mRepository.AddDue( _
                CLng(cmbMember.Column(1)), _
                cmbPeriod.Value, _
                CDate(txtDueDate.Value), _
                CDate(txtDueLimit.Value), _
                CDbl(txtAmount.Value))

    If DueID = 0 Then

    MsgBox "Aidat oluþturulamadý.", vbCritical

    Exit Sub

End If

RefreshDueList

ClearForm

MsgBox "Aidat baþarýyla oluþturuldu.", vbInformation

End Sub
'---------------------------------------------------------
' Form Doðrulama
'---------------------------------------------------------
Private Function ValidateDueForm() As Boolean

    If cmbMember.ListIndex = -1 Then

        MsgBox "Lütfen bir üye seçiniz.", vbExclamation

        cmbMember.SetFocus

        Exit Function

    End If

    If Trim$(cmbPeriod.Value) = "" Then

        MsgBox "Lütfen dönem seçiniz.", vbExclamation

        cmbPeriod.SetFocus

        Exit Function

    End If

    If Val(Replace(txtAmount.Value, ",", ".")) <= 0 Then

        MsgBox "Aidat tutarý sýfýrdan büyük olmalýdýr.", vbExclamation

        txtAmount.SetFocus

        Exit Function

    End If

    ValidateDueForm = True

End Function
