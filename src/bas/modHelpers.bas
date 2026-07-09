Attribute VB_Name = "modHelpers"
Option Explicit

'=========================================================
' OAA - Open Association Accounting
' Module : modHelpers
' Author : Haluk Sand�k��
' Version: 1.0.0
'=========================================================

'=========================================================
' WORKSHEET FUNCTIONS
'=========================================================

Public Function SheetExists(ByVal SheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(SheetName)

    SheetExists = Not ws Is Nothing

    Set ws = Nothing

    On Error GoTo 0

End Function

'---------------------------------------------------------

Public Function GetWorksheet(ByVal SheetName As String) As Worksheet

    If SheetExists(SheetName) Then

        Set GetWorksheet = ThisWorkbook.Worksheets(SheetName)

    End If

End Function

'---------------------------------------------------------

Public Function CreateSheet(ByVal SheetName As String) As Worksheet

    If SheetExists(SheetName) Then

        Set CreateSheet = Worksheets(SheetName)

    Else

        Set CreateSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))

        CreateSheet.Name = SheetName

    End If

End Function

'---------------------------------------------------------

Public Sub DeleteSheet(ByVal SheetName As String)

    If Not SheetExists(SheetName) Then Exit Sub

    Application.DisplayAlerts = False

    Worksheets(SheetName).Delete

    Application.DisplayAlerts = True

End Sub

'---------------------------------------------------------

Public Sub ClearSheet(ByVal SheetName As String)

    If Not SheetExists(SheetName) Then Exit Sub

    Worksheets(SheetName).Cells.Clear

End Sub

'=========================================================
' TABLE FUNCTIONS
'=========================================================

Public Function TableExists( _
        ByVal SheetName As String, _
        ByVal TableName As String) As Boolean

    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error Resume Next

    Set ws = Worksheets(SheetName)

    If ws Is Nothing Then Exit Function

    Set tbl = ws.ListObjects(TableName)

    TableExists = Not tbl Is Nothing

    Set tbl = Nothing
    Set ws = Nothing

    On Error GoTo 0

End Function

'---------------------------------------------------------

Public Function GetTable( _
        ByVal SheetName As String, _
        ByVal TableName As String) As ListObject

    If TableExists(SheetName, TableName) Then

        Set GetTable = Worksheets(SheetName).ListObjects(TableName)

    End If

End Function

'---------------------------------------------------------

Public Function CreateExcelTable( _
        ByVal ws As Worksheet, _
        ByVal TableName As String, _
        ByVal SourceRange As Range) As ListObject

    Dim tbl As ListObject

    If TableExists(ws.Name, TableName) Then

        Set CreateExcelTable = ws.ListObjects(TableName)

        Exit Function

    End If

    Set tbl = ws.ListObjects.Add( _
                    SourceType:=xlSrcRange, _
                    Source:=SourceRange, _
                    XlListObjectHasHeaders:=xlYes)

    tbl.Name = TableName

    Set CreateExcelTable = tbl

End Function

'---------------------------------------------------------

Public Sub DeleteTable( _
        ByVal SheetName As String, _
        ByVal TableName As String)

    If Not TableExists(SheetName, TableName) Then Exit Sub

    Worksheets(SheetName).ListObjects(TableName).Delete

End Sub

'=========================================================
' RANGE FUNCTIONS
'=========================================================

Public Function LastRow(ByVal ws As Worksheet) As Long

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

End Function

'---------------------------------------------------------

Public Function LastColumn(ByVal ws As Worksheet) As Long

    LastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

End Function

'---------------------------------------------------------

Public Function DataRowCount(ByVal tbl As ListObject) As Long

    If tbl.DataBodyRange Is Nothing Then

        DataRowCount = 0

    Else

        DataRowCount = tbl.DataBodyRange.Rows.Count

    End If

End Function
'=========================================================
' VALIDATION FUNCTIONS
'=========================================================

Public Function Nz( _
        ByVal Value As Variant, _
        Optional ByVal DefaultValue As Variant = "") As Variant

    If IsError(Value) Then

        Nz = DefaultValue

    ElseIf IsNull(Value) Then

        Nz = DefaultValue

    ElseIf IsEmpty(Value) Then

        Nz = DefaultValue

    ElseIf Trim(CStr(Value)) = "" Then

        Nz = DefaultValue

    Else

        Nz = Value

    End If

End Function

'---------------------------------------------------------

Public Function IsNullOrEmpty(ByVal Value As Variant) As Boolean

    If IsNull(Value) Then

        IsNullOrEmpty = True

    ElseIf IsEmpty(Value) Then

        IsNullOrEmpty = True

    ElseIf Trim(CStr(Value)) = "" Then

        IsNullOrEmpty = True

    Else

        IsNullOrEmpty = False

    End If

End Function

'---------------------------------------------------------

Public Function IsWorksheetEmpty(ByVal ws As Worksheet) As Boolean

    IsWorksheetEmpty = _
        (WorksheetFunction.CountA(ws.Cells) = 0)

End Function

'=========================================================
' NUMBER FUNCTIONS
'=========================================================

Public Function GetNextID(ByVal tbl As ListObject) As Long

    On Error GoTo NewTable

    If tbl.DataBodyRange Is Nothing Then

        GetNextID = 1

    Else

        GetNextID = _
            Application.WorksheetFunction.Max( _
            tbl.ListColumns(1).DataBodyRange) + 1

    End If

    Exit Function

NewTable:

    GetNextID = 1

End Function

'---------------------------------------------------------

Public Function SafeLong(ByVal Value As Variant) As Long

    If IsNumeric(Value) Then

        SafeLong = CLng(Value)

    Else

        SafeLong = 0

    End If

End Function

'---------------------------------------------------------

Public Function SafeDouble(ByVal Value As Variant) As Double

    If IsNumeric(Value) Then

        SafeDouble = CDbl(Value)

    Else

        SafeDouble = 0

    End If

End Function

'=========================================================
' STRING FUNCTIONS
'=========================================================

Public Function TrimSafe(ByVal Value As Variant) As String

    TrimSafe = Trim(CStr(Nz(Value, "")))

End Function

'---------------------------------------------------------

Public Function ProperText(ByVal Text As String) As String

    ProperText = _
        Application.WorksheetFunction.Proper( _
        LCase(Trim(Text)))

End Function

'---------------------------------------------------------

Public Function UpperText(ByVal Text As String) As String

    UpperText = UCase(Trim(Text))

End Function

'---------------------------------------------------------

Public Function LowerText(ByVal Text As String) As String

    LowerText = LCase(Trim(Text))

End Function

'=========================================================
' UNIQUE ID
'=========================================================

Public Function GenerateGUID() As String

    GenerateGUID = _
        Mid(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)

End Function

'---------------------------------------------------------

Public Function Timestamp() As String

    Timestamp = _
        Format(Now, "yyyymmddhhnnss")

End Function
'=========================================================
' DATE FUNCTIONS
'=========================================================

Public Function TodayDate() As Date

    TodayDate = Date

End Function

'---------------------------------------------------------

Public Function CurrentDateTime() As Date

    CurrentDateTime = Now

End Function

'---------------------------------------------------------

Public Function FormatDateTR(ByVal Value As Variant) As String

    If IsDate(Value) Then

        FormatDateTR = Format(CDate(Value), DATE_FORMAT)

    Else

        FormatDateTR = ""

    End If

End Function

'---------------------------------------------------------

Public Function FormatDateTimeTR(ByVal Value As Variant) As String

    If IsDate(Value) Then

        FormatDateTimeTR = Format(CDate(Value), DATETIME_FORMAT)

    Else

        FormatDateTimeTR = ""

    End If

End Function

'---------------------------------------------------------

Public Function FirstDayOfMonth(ByVal Value As Date) As Date

    FirstDayOfMonth = DateSerial(Year(Value), Month(Value), 1)

End Function

'---------------------------------------------------------

Public Function LastDayOfMonth(ByVal Value As Date) As Date

    LastDayOfMonth = DateSerial(Year(Value), Month(Value) + 1, 0)

End Function

'---------------------------------------------------------

Public Function MonthNameTR(ByVal MonthNo As Integer) As String

    Select Case MonthNo

        Case 1: MonthNameTR = "Ocak"
        Case 2: MonthNameTR = "Şubat"
        Case 3: MonthNameTR = "Mart"
        Case 4: MonthNameTR = "Nisan"
        Case 5: MonthNameTR = "Mayıs"
        Case 6: MonthNameTR = "Haziran"
        Case 7: MonthNameTR = "Temmuz"
        Case 8: MonthNameTR = "Ağustos"
        Case 9: MonthNameTR = "Eylül"
        Case 10: MonthNameTR = "Ekim"
        Case 11: MonthNameTR = "Kasım"
        Case 12: MonthNameTR = "Aralık"

        Case Else

            MonthNameTR = ""

    End Select

End Function

'=========================================================
' CURRENCY FUNCTIONS
'=========================================================

Public Function FormatTRY(ByVal Value As Variant) As String

    FormatTRY = Format(SafeDouble(Value), "#,##0.00 ₺")

End Function

'---------------------------------------------------------

Public Function FormatUSD(ByVal Value As Variant) As String

    FormatUSD = Format(SafeDouble(Value), "#,##0.00 $")

End Function

'---------------------------------------------------------

Public Function FormatEUR(ByVal Value As Variant) As String

    FormatEUR = Format(SafeDouble(Value), "#,##0.00 €")

End Function

'=========================================================
' FORMAT FUNCTIONS
'=========================================================

Public Function PhoneFormat(ByVal Phone As String) As String

    Phone = Replace(Phone, " ", "")
    Phone = Replace(Phone, "-", "")
    Phone = Replace(Phone, "(", "")
    Phone = Replace(Phone, ")", "")

    PhoneFormat = Phone

End Function

'---------------------------------------------------------

Public Function EmailFormat(ByVal Email As String) As String

    EmailFormat = LCase(Trim(Email))

End Function

'---------------------------------------------------------

Public Function ReceiptNo(ByVal Series As String, _
                          ByVal Number As Long) As String

    ReceiptNo = _
        Series & "-" & Format(Number, "000000")

End Function

'---------------------------------------------------------

Public Function PadLeft(ByVal Text As String, _
                        ByVal TotalLength As Long, _
                        Optional ByVal PadChar As String = "0") As String

    If Len(Text) >= TotalLength Then

        PadLeft = Text

    Else

        PadLeft = String(TotalLength - Len(Text), PadChar) & Text

    End If

End Function
'=========================================================
' PERFORMANCE FUNCTIONS
'=========================================================

Public Sub BeginUpdate()

    With Application

        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = "İşlem yapılıyor..."

    End With

End Sub

'---------------------------------------------------------

Public Sub EndUpdate()

    With Application

        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = False

    End With

End Sub

'=========================================================
' WORKSHEET FUNCTIONS
'=========================================================

Public Sub AutoFitWorksheet(ByVal ws As Worksheet)

    ws.Cells.EntireColumn.AutoFit

End Sub

'---------------------------------------------------------

Public Sub FreezeTopRow(ByVal ws As Worksheet)

    ws.Activate

    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True

End Sub

'---------------------------------------------------------

Public Sub UnFreezePanes()

    ActiveWindow.FreezePanes = False

End Sub

'---------------------------------------------------------

Public Sub ProtectWorksheet(ByVal ws As Worksheet, _
                            Optional ByVal Password As String = "")

    ws.Protect Password:=Password, _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True

End Sub

'---------------------------------------------------------

Public Sub UnProtectWorksheet(ByVal ws As Worksheet, _
                              Optional ByVal Password As String = "")

    ws.Unprotect Password

End Sub

'---------------------------------------------------------

Public Sub ClearFilters(ByVal tbl As ListObject)

    On Error Resume Next

    If tbl.AutoFilter.FilterMode Then

        tbl.AutoFilter.ShowAllData

    End If

    On Error GoTo 0

End Sub

'=========================================================
' STATUS BAR
'=========================================================

Public Sub WriteStatusBar(ByVal Message As String)

    Application.StatusBar = Message

End Sub

'---------------------------------------------------------

Public Sub ClearStatusBar()

    Application.StatusBar = False

End Sub

'=========================================================
' EXCEL HELPERS
'=========================================================

Public Sub SelectFirstCell(ByVal ws As Worksheet)

    ws.Activate

    ws.Range("A1").Select

End Sub

'---------------------------------------------------------

Public Sub ActivateSheet(ByVal SheetName As String)

    If SheetExists(SheetName) Then

        Worksheets(SheetName).Activate

    End If

End Sub

'---------------------------------------------------------

Public Function WorkbookPath() As String

    WorkbookPath = ThisWorkbook.Path

End Function

'---------------------------------------------------------

Public Function WorkbookName() As String

    WorkbookName = ThisWorkbook.Name

End Function

'---------------------------------------------------------

Public Function WorkbookFullName() As String

    WorkbookFullName = ThisWorkbook.FullName

End Function
'=========================================================
' FILE FUNCTIONS
'=========================================================

Public Function FileExists(ByVal FileName As String) As Boolean

    FileExists = (Dir(FileName, vbNormal) <> "")

End Function

'---------------------------------------------------------

Public Function FolderExists(ByVal FolderName As String) As Boolean

    On Error Resume Next

    FolderExists = _
        (GetAttr(FolderName) And vbDirectory) = vbDirectory

    On Error GoTo 0

End Function

'---------------------------------------------------------

Public Sub CreateFolder(ByVal FolderName As String)

    If FolderExists(FolderName) Then Exit Sub

    MkDir FolderName

End Sub

'---------------------------------------------------------

Public Function GetFileName(ByVal FullName As String) As String

    GetFileName = Dir(FullName)

End Function

'---------------------------------------------------------

Public Function GetFolderName(ByVal FullName As String) As String

    Dim p As Long

    p = InStrRev(FullName, "\")

    If p > 0 Then

        GetFolderName = Left$(FullName, p - 1)

    Else

        GetFolderName = ""

    End If

End Function

'=========================================================
' BACKUP FUNCTIONS
'=========================================================

Public Function BackupFolder() As String

    BackupFolder = _
        WorkbookPath & "\Backup"

End Function

'---------------------------------------------------------

Public Function EnsureBackupFolder() As String

    If Not FolderExists(BackupFolder) Then

        CreateFolder BackupFolder

    End If

    EnsureBackupFolder = BackupFolder

End Function

'---------------------------------------------------------

Public Function BackupFileName() As String

    BackupFileName = _
        EnsureBackupFolder & "\" & _
        "OAA_" & _
        Format(Now, "yyyymmdd_hhnnss") & _
        ".xlsm"

End Function

'---------------------------------------------------------

Public Function BackupWorkbook() As Boolean

    On Error GoTo ErrHandler

    ThisWorkbook.SaveCopyAs BackupFileName

    BackupWorkbook = True

    Exit Function

ErrHandler:

    BackupWorkbook = False

End Function

'=========================================================
' ENVIRONMENT
'=========================================================

Public Function CurrentUser() As String

    CurrentUser = Environ$("USERNAME")

End Function

'---------------------------------------------------------

Public Function ComputerName() As String

    ComputerName = Environ$("COMPUTERNAME")

End Function

'---------------------------------------------------------

Public Function WindowsFolder() As String

    WindowsFolder = Environ$("WINDIR")

End Function

'---------------------------------------------------------

Public Function TempFolder() As String

    TempFolder = Environ$("TEMP")

End Function

'---------------------------------------------------------

Public Function DocumentsFolder() As String

    DocumentsFolder = _
        CreateObject("WScript.Shell") _
        .SpecialFolders("MyDocuments")

End Function
'=========================================================
' ERROR FUNCTIONS
'=========================================================

Public Function ErrorText(ByVal ErrObject As ErrObject) As String

    ErrorText = _
        "Error " & ErrObject.Number & _
        " : " & ErrObject.Description

End Function

'---------------------------------------------------------

Public Sub ShowError(ByVal ErrObject As ErrObject)

    MsgBox ErrorText(ErrObject), _
           vbCritical, _
           APP_NAME

End Sub

'---------------------------------------------------------

Public Sub ShowInfo(ByVal Message As String)

    MsgBox Message, vbInformation, APP_NAME

End Sub

'---------------------------------------------------------

Public Sub ShowWarning(ByVal Message As String)

    MsgBox Message, vbExclamation, APP_NAME

End Sub

'---------------------------------------------------------

Public Function Confirm(ByVal Message As String) As Boolean

    Confirm = _
        (MsgBox(Message, _
                vbYesNo + vbQuestion, _
                APP_NAME) = vbYes)

End Function

'=========================================================
' VALUE FUNCTIONS
'=========================================================

Public Function SafeDate(ByVal Value As Variant) As Date

    If IsDate(Value) Then

        SafeDate = CDate(Value)

    Else

        SafeDate = 0

    End If

End Function

'---------------------------------------------------------

Public Function SafeBoolean(ByVal Value As Variant) As Boolean

    On Error Resume Next

    SafeBoolean = CBool(Value)

    On Error GoTo 0

End Function

'---------------------------------------------------------

Public Function Between( _
        ByVal Value As Double, _
        ByVal MinValue As Double, _
        ByVal MaxValue As Double) As Boolean

    Between = _
        (Value >= MinValue And Value <= MaxValue)

End Function

'=========================================================
' TABLE FUNCTIONS
'=========================================================

Public Function AddTableRow(ByVal tbl As ListObject) As ListRow

    Set AddTableRow = tbl.ListRows.Add

End Function

'---------------------------------------------------------

Public Sub ClearTable(ByVal tbl As ListObject)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    tbl.DataBodyRange.Delete

End Sub

'---------------------------------------------------------

Public Function TableRowCount(ByVal tbl As ListObject) As Long

    If tbl.DataBodyRange Is Nothing Then

        TableRowCount = 0

    Else

        TableRowCount = tbl.ListRows.Count

    End If

End Function

'=========================================================
' APPLICATION
'=========================================================

Public Function ExcelVersion() As String

    ExcelVersion = Application.Version

End Function

'---------------------------------------------------------

Public Function CurrentWorkbook() As Workbook

    Set CurrentWorkbook = ThisWorkbook

End Function

'---------------------------------------------------------

Public Function CurrentWorksheet() As Worksheet

    Set CurrentWorksheet = ActiveSheet

End Function
'=========================================================
' LOOKUP FUNCTIONS
'=========================================================

Public Function FindTableRow( _
        ByVal tbl As ListObject, _
        ByVal ColumnIndex As Long, _
        ByVal SearchValue As Variant) As ListRow

    Dim rw As ListRow

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For Each rw In tbl.ListRows

        If CStr(rw.Range.Cells(1, ColumnIndex).Value) = _
           CStr(SearchValue) Then

            Set FindTableRow = rw
            Exit Function

        End If

    Next rw

End Function

'---------------------------------------------------------

Public Function RecordExists( _
        ByVal tbl As ListObject, _
        ByVal ColumnIndex As Long, _
        ByVal SearchValue As Variant) As Boolean

    RecordExists = _
        Not FindTableRow(tbl, ColumnIndex, SearchValue) Is Nothing

End Function

'=========================================================
' SETTINGS FUNCTIONS
'=========================================================

Public Function SettingValue(ByVal KeyName As String) As String

    Dim tbl As ListObject
    Dim rw As ListRow

    If Not TableExists(SHEET_SETTINGS, TBL_SETTINGS) Then Exit Function

    Set tbl = GetTable(SHEET_SETTINGS, TBL_SETTINGS)

    For Each rw In tbl.ListRows

        If StrComp(CStr(rw.Range.Cells(1, 1).Value), _
                   KeyName, _
                   vbTextCompare) = 0 Then

            SettingValue = CStr(rw.Range.Cells(1, 2).Value)
            Exit Function

        End If

    Next rw

End Function

'=========================================================
' MEMBER FUNCTIONS
'=========================================================

Public Function MemberTable() As ListObject

    Set MemberTable = _
        GetTable(SHEET_MEMBERS, TBL_MEMBERS)

End Function

'---------------------------------------------------------

Public Function SettingsTable() As ListObject

    Set SettingsTable = _
        GetTable(SHEET_SETTINGS, TBL_SETTINGS)

End Function

'=========================================================
' SYSTEM INFORMATION
'=========================================================

Public Function AppTitle() As String

    AppTitle = _
        APP_NAME & _
        "  v" & _
        APP_VERSION

End Function

'---------------------------------------------------------

Public Function CurrentYear() As Integer

    CurrentYear = Year(Date)

End Function

'---------------------------------------------------------

Public Function CurrentMonth() As Integer

    CurrentMonth = Month(Date)

End Function

'---------------------------------------------------------

Public Function CurrentDay() As Integer

    CurrentDay = Day(Date)

End Function

'=========================================================
' GENERAL UTILITIES
'=========================================================

Public Function BoolText(ByVal Value As Boolean) As String

    If Value Then

        BoolText = "Evet"

    Else

        BoolText = "Hayır"

    End If

End Function

'---------------------------------------------------------

Public Function EmptyIfZero(ByVal Value As Double) As Variant

    If Value = 0 Then

        EmptyIfZero = ""

    Else

        EmptyIfZero = Value

    End If

End Function

'---------------------------------------------------------

Public Function ZeroIfEmpty(ByVal Value As Variant) As Double

    If IsNullOrEmpty(Value) Then

        ZeroIfEmpty = 0

    Else

        ZeroIfEmpty = SafeDouble(Value)

    End If

End Function

'---------------------------------------------------------

Public Function SafeText(ByVal Value As Variant) As String

    SafeText = CStr(Nz(Value, ""))

End Function
'=========================================================
' VERSION INFORMATION
'=========================================================

Public Function HelpersVersion() As String

    HelpersVersion = "1.0.0"

End Function

'---------------------------------------------------------

Public Function HelpersAuthor() As String

    HelpersAuthor = "Haluk Sandıkçı"

End Function

'---------------------------------------------------------

Public Function HelpersDescription() As String

    HelpersDescription = _
        "Open Association Accounting Helper Library"

End Function

'=========================================================
' SELF TEST
'=========================================================

Public Function HelpersSelfTest() As Boolean

    On Error GoTo ErrHandler

    If Not SheetExists(SHEET_SETTINGS) Then GoTo ErrHandler
    If Not SheetExists(SHEET_MEMBERS) Then GoTo ErrHandler

    If Not TableExists(SHEET_SETTINGS, TBL_SETTINGS) Then GoTo ErrHandler
    If Not TableExists(SHEET_MEMBERS, TBL_MEMBERS) Then GoTo ErrHandler

    HelpersSelfTest = True

    Exit Function

ErrHandler:

    HelpersSelfTest = False

End Function

'=========================================================
' END OF MODULE
'=========================================================
