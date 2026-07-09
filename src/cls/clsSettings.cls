VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' OAA - Open Association Accounting
' Class  : clsSettings
' Version: 1.0.0
'=========================================================

Private mTable As ListObject

'---------------------------------------------------------
' Constructor
'---------------------------------------------------------
Private Sub Class_Initialize()

    Dim ws As Worksheet

    Set ws = GetWorksheet(SHEET_SETTINGS)

    If ws Is Nothing Then Exit Sub

    If TableExists(SHEET_SETTINGS, TBL_SETTINGS) Then
        Set mTable = ws.ListObjects(TBL_SETTINGS)
    End If

End Sub

'---------------------------------------------------------
' Ayar Oku
'---------------------------------------------------------
Public Function GetValue(ByVal Key As String) As String

    Dim rw As ListRow

    If mTable Is Nothing Then Exit Function

    For Each rw In mTable.ListRows

        If StrComp(CStr(rw.Range.Cells(1, 1).Value), Key, vbTextCompare) = 0 Then

            GetValue = CStr(rw.Range.Cells(1, 2).Value)

            Exit Function

        End If

    Next rw

End Function

'---------------------------------------------------------
' Ayar Yaz
'---------------------------------------------------------
Public Sub SetValue(ByVal Key As String, ByVal Value As String)

    Dim rw As ListRow

    If mTable Is Nothing Then Exit Sub

    For Each rw In mTable.ListRows

        If StrComp(CStr(rw.Range.Cells(1, 1).Value), Key, vbTextCompare) = 0 Then

            rw.Range.Cells(1, 2).Value = Value

            Exit Sub

        End If

    Next rw

End Sub
