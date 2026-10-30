Attribute VB_Name = "modDatabase"
Option Explicit

'==============================================================================
' OAA v2 Framework
' Database Functions
'==============================================================================
'------------------------------------------------------------------------------
' Sayfadaki tabloyu döndürür.
'------------------------------------------------------------------------------
Public Function GetTable( _
            ByVal SheetName As String, _
            ByVal TableName As String) As ListObject

    Set GetTable = _
        ThisWorkbook.Worksheets(SheetName) _
                    .ListObjects(TableName)

End Function

'------------------------------------------------------------------------------
' Tablodaki sütun numarasýný döndürür.
'------------------------------------------------------------------------------
Public Function GetColumnIndex( _
            ByVal Table As ListObject, _
            ByVal ColumnName As String) As Long

    GetColumnIndex = Table.ListColumns(ColumnName).Index

End Function
'------------------------------------------------------------------------------
' Tablo mevcut mu?
'------------------------------------------------------------------------------
Public Function TableExists( _
            ByVal SheetName As String, _
            ByVal TableName As String) As Boolean

    On Error GoTo ErrHandler

    Dim lo As ListObject

    Set lo = ThisWorkbook.Worksheets(SheetName).ListObjects(TableName)

    TableExists = True
    Exit Function

ErrHandler:

    TableExists = False

End Function
'------------------------------------------------------------------------------
' Sütun mevcut mu?
'------------------------------------------------------------------------------
Public Function ColumnExists( _
            ByVal Table As ListObject, _
            ByVal ColumnName As String) As Boolean

    On Error GoTo ErrHandler

    Dim lc As ListColumn

    Set lc = Table.ListColumns(ColumnName)

    ColumnExists = True
    Exit Function

ErrHandler:

    ColumnExists = False

End Function

