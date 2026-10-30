VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssociationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Class : clsAssociationService
' Purpose : Association Business Layer
'==============================================================================

Private Repository As clsAssociationRepository

Private Sub Class_Initialize()

    Set Repository = New clsAssociationRepository

End Sub

'------------------------------------------------------------------------------
' Association kaydý var mý?
'------------------------------------------------------------------------------
Public Function Exists() As Boolean

    Exists = Repository.Exists()

End Function
'------------------------------------------------------------------------------
' Association bilgilerini döndürür.
'------------------------------------------------------------------------------
Public Function GetAssociation() As clsAssociation

    Set GetAssociation = Repository.GetAssociation()

End Function
'------------------------------------------------------------------------------
' Association bilgilerini kaydeder.
'------------------------------------------------------------------------------
Public Function Save( _
            ByVal Association As clsAssociation) As Boolean

    On Error GoTo ErrorHandler

    If Not Validate(Association) Then
        Save = False
        Exit Function
    End If

    If Repository.Exists Then
        Save = Repository.Update(Association)
    Else
        Save = Repository.Add(Association)
    End If

    Exit Function

ErrorHandler:

    ' Repository teknik hata fýrlattý: form açýk hata iletisi almasýn,
    ' Save baþarýsýz (False) döner.
    Save = False

End Function

'------------------------------------------------------------------------------
' Association bilgilerini doðrular.
'------------------------------------------------------------------------------
Public Function Validate( _
            ByVal Association As clsAssociation) As Boolean

    Dim Result As clsValidationResult

    Set Result = New clsValidationResult

    Result.SetSuccess

    '----------------------------------------------------------
    ' Dernek Adý
    '----------------------------------------------------------
    If Not ValidateRequired(Association.AssociationName) Then

        Result.SetError _
            ERR_REQUIRED_FIELD, _
            COL_ASSOCIATION_NAME, _
            "Dernek adý zorunludur."

    End If

    '----------------------------------------------------------
    ' Kýsa Ad
    '----------------------------------------------------------
    If Result.Success Then

        If Not ValidateRequired(Association.ShortName) Then

            Result.SetError _
                ERR_REQUIRED_FIELD, _
                COL_SHORT_NAME, _
                "Kýsa ad zorunludur."

        End If

    End If

    '----------------------------------------------------------
    ' Para Birimi
    '----------------------------------------------------------
    If Result.Success Then

        If Not ValidateRequired(Association.DefaultCurrency) Then

            Result.SetError _
                ERR_REQUIRED_FIELD, _
                COL_DEFAULT_CURRENCY, _
                "Varsayýlan para birimi zorunludur."

        End If

    End If

    If Result.Success Then

        If Not ValidateCurrency(Association.DefaultCurrency) Then

            Result.SetError _
                ERR_INVALID_VALUE, _
                COL_DEFAULT_CURRENCY, _
                "Geçersiz para birimi."

        End If

    End If

    '----------------------------------------------------------
    ' Mali Yýl
    '----------------------------------------------------------
    If Result.Success Then

        If Not ValidateMonth(Association.FiscalYearStart) Then

            Result.SetError _
                ERR_INVALID_VALUE, _
                COL_FISCAL_YEAR_START, _
                "Geçersiz mali yýl baþlangýç ayý."

        End If

    End If

    Validate = Result.Success

End Function

