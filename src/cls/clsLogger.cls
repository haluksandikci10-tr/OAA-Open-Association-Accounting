VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' OAA - Open Association Accounting
' Class : clsLogger
' Version : 1.0.0
'=========================================================

Private mLog As ListObject

'---------------------------------------------------------
' Constructor
'---------------------------------------------------------
Private Sub Class_Initialize()

    If TableExists(SHEET_LOG, TBL_LOG) Then
        Set mLog = GetTable(SHEET_LOG, TBL_LOG)
    End If

End Sub

'---------------------------------------------------------
' Destructor
'---------------------------------------------------------
Private Sub Class_Terminate()

    Set mLog = Nothing

End Sub

'=========================================================
' PRIVATE
'=========================================================

Private Sub WriteLog( _
    ByVal Level As String, _
    ByVal ModuleName As String, _
    ByVal Message As String)

    Dim rw As ListRow

    If mLog Is Nothing Then Exit Sub

    Set rw = mLog.ListRows.Add

    With rw.Range

        .Cells(1, 1).Value = Now
        .Cells(1, 2).Value = Level
        .Cells(1, 3).Value = ModuleName
        .Cells(1, 4).Value = Message
        .Cells(1, 5).Value = CurrentUser()

    End With

End Sub
'=========================================================
' PUBLIC LOG METHODS
'=========================================================

Public Sub Info(ByVal ModuleName As String, _
                ByVal Message As String)

    WriteLog LOG_INFO, ModuleName, Message

End Sub

'---------------------------------------------------------

Public Sub Warning(ByVal ModuleName As String, _
                   ByVal Message As String)

    WriteLog LOG_WARNING, ModuleName, Message

End Sub

'---------------------------------------------------------

Public Sub ErrorLog(ByVal ModuleName As String, _
                    ByVal Message As String)

    WriteLog LOG_ERROR, ModuleName, Message

End Sub

'---------------------------------------------------------

Public Sub DebugLog(ByVal ModuleName As String, _
                    ByVal Message As String)

    WriteLog LOG_DEBUG, ModuleName, Message

End Sub

'=========================================================
' LOG INFORMATION
'=========================================================

Public Function LogCount() As Long

    If mLog Is Nothing Then Exit Function

    If mLog.DataBodyRange Is Nothing Then
        LogCount = 0
    Else
        LogCount = mLog.ListRows.Count
    End If

End Function

'---------------------------------------------------------

Public Function IsReady() As Boolean

    IsReady = Not mLog Is Nothing

End Function
'=========================================================
' LOG MAINTENANCE
'=========================================================

Public Sub ClearLog()

    If mLog Is Nothing Then Exit Sub

    If mLog.DataBodyRange Is Nothing Then Exit Sub

    mLog.DataBodyRange.Delete

End Sub

'---------------------------------------------------------

Public Function LastMessage() As String

    If mLog Is Nothing Then Exit Function

    If mLog.DataBodyRange Is Nothing Then Exit Function

    LastMessage = _
        CStr(mLog.DataBodyRange.Cells( _
            mLog.ListRows.Count, 4).Value)

End Function

'---------------------------------------------------------

Public Function LoggerSelfTest() As Boolean

    On Error GoTo ErrHandler

    If mLog Is Nothing Then GoTo ErrHandler

    Info "Logger", "Logger Self Test"

    LoggerSelfTest = True

    Exit Function

ErrHandler:

    LoggerSelfTest = False

End Function

'---------------------------------------------------------

Public Sub Dispose()

    Set mLog = Nothing

End Sub

