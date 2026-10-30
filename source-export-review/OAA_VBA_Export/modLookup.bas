Attribute VB_Name = "modLookup"
Option Explicit

'==============================================================================
' OAA v2
' Module : modLookup
' Purpose: Lookup Lists
'==============================================================================

Private Const LOOKUP_SHEET As String = "Settings"
Private Const LOOKUP_TABLE As String = "tblLookup"

Private Const COL_LOOKUP_TYPE As String = "LookupType"
Private Const COL_LOOKUP_CODE As String = "LookupCode"
Private Const COL_LOOKUP_VALUE As String = "LookupValue"
Private Const COL_SORT_ORDER As String = "SortOrder"
Private Const COL_SORT_ORDER2 As String = "SortOrder2"


'------------------------------------------------------------------------------
' Üye Durumlarý
'------------------------------------------------------------------------------
Public Function GetMemberStatusList() As Variant

    GetMemberStatusList = Array( _
        MEMBER_STATUS_ACTIVE, _
        MEMBER_STATUS_RESIGNED, _
        MEMBER_STATUS_LEAVE)

End Function


'------------------------------------------------------------------------------
' Genel Lookup Listesi
'
' Belirtilen LookupType deðerine ait LookupValue kayýtlarýný döndürür.
'
' Veri kaynaðý:
'   Settings / tblLookup
'
' Repository veya UserForm doðrudan worksheet'e eriþmez.
'------------------------------------------------------------------------------

Public Function GetLookupValues(ByVal LookupType As String) As Collection

    On Error GoTo ErrHandler

    Dim Result As New Collection
    Dim Table As ListObject
    Dim RowIndex As Long

    Dim TypeValue As String
    Dim LookupValue As String

    If Trim$(LookupType) = "" Then
        Set GetLookupValues = Result
        Exit Function
    End If

    Set Table = GetTable(LOOKUP_SHEET, LOOKUP_TABLE)

    If Table.DataBodyRange Is Nothing Then
        Set GetLookupValues = Result
        Exit Function
    End If

    For RowIndex = 1 To Table.ListRows.Count

        TypeValue = Trim$(CStr( _
            Table.DataBodyRange.Cells( _
                RowIndex, _
                GetColumnIndex(Table, COL_LOOKUP_TYPE) _
            ).Value))

        If StrComp(TypeValue, Trim$(LookupType), vbTextCompare) = 0 Then

            LookupValue = Trim$(CStr( _
                Table.DataBodyRange.Cells( _
                    RowIndex, _
                    GetColumnIndex(Table, COL_LOOKUP_VALUE) _
                ).Value))

            If LookupValue <> "" Then
                Result.Add LookupValue
            End If

        End If

    Next RowIndex

    Set GetLookupValues = Result
    Exit Function

ErrHandler:

    Set GetLookupValues = Nothing

End Function

