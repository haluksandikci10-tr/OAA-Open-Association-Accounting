VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' OAA - Open Association Accounting
' Class : clsRepository
' Version : 1.0.0
'=========================================================

Private mMembers As ListObject
Private mDues As ListObject
Private mSettings As ListObject
'---------------------------------------------------------
' Constructor
'---------------------------------------------------------
Private Sub Class_Initialize()

    If TableExists(SHEET_MEMBERS, TBL_MEMBERS) Then
        Set mMembers = GetTable(SHEET_MEMBERS, TBL_MEMBERS)
    End If

    If TableExists(SHEET_DUES, TBL_DUES) Then
        Set mDues = GetTable(SHEET_DUES, TBL_DUES)
    End If

    If TableExists(SHEET_SETTINGS, TBL_SETTINGS) Then
        Set mSettings = GetTable(SHEET_SETTINGS, TBL_SETTINGS)
    End If

End Sub
'---------------------------------------------------------
' Destructor
'---------------------------------------------------------
Private Sub Class_Terminate()

    Set mMembers = Nothing
    Set mSettings = Nothing

End Sub

'=========================================================
' PUBLIC
'=========================================================

Public Function Members() As ListObject

    Set Members = mMembers

End Function
'---------------------------------------------------------
' Dues Table
'---------------------------------------------------------
Public Function Dues() As ListObject

    Set Dues = mDues

End Function

Public Function Settings() As ListObject

    Set Settings = mSettings

End Function
'=========================================================
' MEMBER FUNCTIONS
'=========================================================

Public Function MemberCount() As Long

    Dim rw As ListRow

    MemberCount = 0

    If mMembers Is Nothing Then Exit Function

    If mMembers.DataBodyRange Is Nothing Then Exit Function

    For Each rw In mMembers.ListRows

        If Trim$(CStr(rw.Range.Cells(1, 2).Value)) <> "" Then
            MemberCount = MemberCount + 1
        End If

    Next rw

End Function
'---------------------------------------------------------
' Aidat Sayýsý
'---------------------------------------------------------
Public Function DueCount() As Long

    If mDues Is Nothing Then
        DueCount = 0
        Exit Function
    End If

    If mDues.DataBodyRange Is Nothing Then
        DueCount = 0
    Else
        DueCount = mDues.ListRows.Count
    End If

End Function
'---------------------------------------------------------
' Yeni Aidat Ekle
'---------------------------------------------------------
Public Function AddDue( _
        ByVal MemberID As Long, _
        ByVal Donem As String, _
        ByVal TahakkukTarihi As Date, _
        ByVal VadeTarihi As Date, _
        ByVal Tutar As Double) As Long

    Dim rw As ListRow

    If mDues Is Nothing Then Exit Function

    AddDue = GetNextID(mDues)
    Set rw = mDues.ListRows.Add

    

    With rw.Range

        .Cells(1, 1).Value = AddDue
        .Cells(1, 2).Value = MemberID
        .Cells(1, 3).Value = Donem
        .Cells(1, 4).Value = TahakkukTarihi
        .Cells(1, 5).Value = VadeTarihi
        .Cells(1, 6).Value = Tutar
        .Cells(1, 7).Value = 0
        .Cells(1, 8).Value = Tutar
        .Cells(1, 9).Value = "Bekliyor"
        .Cells(1, 10).Value = ""
        .Cells(1, 11).Value = ""

    End With

End Function
'---------------------------------------------------------
' Aidat Bul
'---------------------------------------------------------
Public Function FindDue(ByVal DueID As Long) As ListRow

    Dim rw As ListRow

    If mDues Is Nothing Then Exit Function

    If mDues.DataBodyRange Is Nothing Then Exit Function

    For Each rw In mDues.ListRows

        If CLng(rw.Range.Cells(1, 1).Value) = DueID Then

            Set FindDue = rw

            Exit Function

        End If

    Next rw

End Function
'---------------------------------------------------------
' Aidat Güncelle
'---------------------------------------------------------
Public Function UpdateDue( _
        ByVal DueID As Long, _
        ByVal TahakkukTarihi As Date, _
        ByVal VadeTarihi As Date, _
        ByVal Tutar As Double, _
        ByVal Odenen As Double, _
        ByVal Durum As String, _
        Optional ByVal MakbuzNo As String = "", _
        Optional ByVal Aciklama As String = "") As Boolean

    Dim rw As ListRow

    Set rw = FindDue(DueID)

    If rw Is Nothing Then Exit Function

    With rw.Range

        .Cells(1, 4).Value = TahakkukTarihi
        .Cells(1, 5).Value = VadeTarihi
        .Cells(1, 6).Value = Tutar
        .Cells(1, 7).Value = Odenen
        .Cells(1, 8).Value = Tutar - Odenen
        .Cells(1, 9).Value = Durum
        .Cells(1, 10).Value = MakbuzNo
        .Cells(1, 11).Value = Aciklama

    End With

    UpdateDue = True

End Function
'---------------------------------------------------------
' Aidat Sil
'---------------------------------------------------------
Public Function DeleteDue(ByVal DueID As Long) As Boolean

    Dim rw As ListRow

    Set rw = FindDue(DueID)

    If rw Is Nothing Then Exit Function

    rw.Delete

    DeleteDue = True

End Function
Public Function AddMember( _
        ByVal FullName As String, _
        Optional ByVal Status As String = MEMBER_ACTIVE) As Long

    Dim rw As ListRow

    If mMembers Is Nothing Then Exit Function

    Set rw = mMembers.ListRows.Add

    AddMember = GetNextID(mMembers)

    With rw.Range

        .Cells(1, 1).Value = AddMember
        .Cells(1, 2).Value = ProperText(FullName)
        .Cells(1, 3).Value = Status
        .Cells(1, 6).Value = Date
        .Cells(1, 7).Value = 0
        .Cells(1, 8).Value = 0
        .Cells(1, 9).Value = 0

    End With

End Function

'---------------------------------------------------------

Public Function FindMember(ByVal MemberID As Long) As ListRow

    Set FindMember = _
        FindTableRow(mMembers, 1, MemberID)

End Function

'---------------------------------------------------------

Public Function MemberExists(ByVal MemberID As Long) As Boolean

    MemberExists = _
        RecordExists(mMembers, 1, MemberID)

End Function

'---------------------------------------------------------

Public Function DeleteMember(ByVal MemberID As Long) As Boolean

    Dim rw As ListRow

    Set rw = FindMember(MemberID)

    If rw Is Nothing Then Exit Function

    rw.Delete

    DeleteMember = True

End Function

'=========================================================
' MEMBER UPDATE FUNCTIONS
'=========================================================

Public Function UpdateMember( _
        ByVal MemberID As Long, _
        ByVal FullName As String, _
        ByVal Status As String, _
        ByVal Phone As String, _
        ByVal Email As String) As Boolean

    Dim rw As ListRow

    Set rw = FindMember(MemberID)

    If rw Is Nothing Then Exit Function

    With rw.Range

        .Cells(1, 2).Value = ProperText(FullName)
        .Cells(1, 3).Value = Status
        .Cells(1, 4).Value = TrimSafe(Phone)
        .Cells(1, 5).Value = LCase(TrimSafe(Email))

    End With

    UpdateMember = True

End Function

'---------------------------------------------------------

Public Function GetMemberName(ByVal MemberID As Long) As String

    Dim rw As ListRow

    Set rw = FindMember(MemberID)

    If rw Is Nothing Then Exit Function

    GetMemberName = CStr(rw.Range.Cells(1, 2).Value)

End Function

'---------------------------------------------------------

Public Function GetMemberStatus(ByVal MemberID As Long) As String

    Dim rw As ListRow

    Set rw = FindMember(MemberID)

    If rw Is Nothing Then Exit Function

    GetMemberStatus = CStr(rw.Range.Cells(1, 3).Value)

End Function

'---------------------------------------------------------

Public Function SetMemberStatus( _
        ByVal MemberID As Long, _
        ByVal Status As String) As Boolean

    Dim rw As ListRow

    Set rw = FindMember(MemberID)

    If rw Is Nothing Then Exit Function

    rw.Range.Cells(1, 3).Value = Status

    SetMemberStatus = True

End Function

'---------------------------------------------------------

Public Function MemberPhone(ByVal MemberID As Long) As String

    Dim rw As ListRow

    Set rw = FindMember(MemberID)

    If rw Is Nothing Then Exit Function

    MemberPhone = CStr(rw.Range.Cells(1, 4).Value)

End Function

'---------------------------------------------------------

Public Function MemberEmail(ByVal MemberID As Long) As String

    Dim rw As ListRow

    Set rw = FindMember(MemberID)

    If rw Is Nothing Then Exit Function

    MemberEmail = CStr(rw.Range.Cells(1, 5).Value)

End Function

'=========================================================
' SETTINGS FUNCTIONS
'=========================================================

Public Function GetSetting(ByVal KeyName As String) As String

    Dim rw As ListRow

    If mSettings Is Nothing Then Exit Function

    For Each rw In mSettings.ListRows

        If StrComp(CStr(rw.Range.Cells(1, 1).Value), _
                   KeyName, _
                   vbTextCompare) = 0 Then

            GetSetting = CStr(rw.Range.Cells(1, 2).Value)
            Exit Function

        End If

    Next rw

End Function

'---------------------------------------------------------

Public Function SetSetting(ByVal KeyName As String, _
                           ByVal Value As String) As Boolean

    Dim rw As ListRow

    If mSettings Is Nothing Then Exit Function

    For Each rw In mSettings.ListRows

        If StrComp(CStr(rw.Range.Cells(1, 1).Value), _
                   KeyName, _
                   vbTextCompare) = 0 Then

            rw.Range.Cells(1, 2).Value = Value
            SetSetting = True
            Exit Function

        End If

    Next rw

End Function

'=========================================================
' GENERAL FUNCTIONS
'=========================================================

Public Function IsReady() As Boolean

    IsReady = _
        Not mMembers Is Nothing _
        And Not mSettings Is Nothing

End Function

'---------------------------------------------------------

Public Sub Refresh()

    Class_Terminate
    Class_Initialize

End Sub

'=========================================================
' MEMBER STATISTICS
'=========================================================

Public Function ActiveMemberCount() As Long

    Dim rw As ListRow

    If mMembers Is Nothing Then Exit Function

    For Each rw In mMembers.ListRows

        If CStr(rw.Range.Cells(1, 3).Value) = MEMBER_ACTIVE Then
            ActiveMemberCount = ActiveMemberCount + 1
        End If

    Next rw

End Function

'---------------------------------------------------------

Public Function PassiveMemberCount() As Long

    Dim rw As ListRow

    If mMembers Is Nothing Then Exit Function

    For Each rw In mMembers.ListRows

        If CStr(rw.Range.Cells(1, 3).Value) = MEMBER_PASSIVE Then
            PassiveMemberCount = PassiveMemberCount + 1
        End If

    Next rw

End Function

'---------------------------------------------------------

Public Function NextMemberID() As Long

    If mMembers Is Nothing Then

        NextMemberID = 1

    Else

        NextMemberID = GetNextID(mMembers)

    End If

End Function

'=========================================================
' TABLE INFORMATION
'=========================================================

Public Function HasMembers() As Boolean

    HasMembers = (MemberCount > 0)

End Function

'---------------------------------------------------------

Public Function HasSettings() As Boolean

    HasSettings = Not mSettings Is Nothing

End Function

'---------------------------------------------------------

Public Function RepositoryVersion() As String

    RepositoryVersion = APP_VERSION

End Function
'=========================================================
' SELF TEST
'=========================================================

Public Function RepositorySelfTest() As Boolean

    On Error GoTo ErrHandler

    If mMembers Is Nothing Then GoTo ErrHandler
    If mSettings Is Nothing Then GoTo ErrHandler

    RepositorySelfTest = True
    Exit Function

ErrHandler:

    RepositorySelfTest = False

End Function

'---------------------------------------------------------
' RESET
'---------------------------------------------------------

Public Sub ResetRepository()

    Refresh

End Sub

'---------------------------------------------------------
' RELEASE
'---------------------------------------------------------

Public Sub Dispose()

    Set mMembers = Nothing
    Set mSettings = Nothing

End Sub

'=========================================================
' END OF CLASS
'=========================================================
