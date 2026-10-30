VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' OAA v2
' Class : ApplicationContext
' Purpose : Application Lifecycle Context
'==============================================================================

Private mIsInitialized As Boolean

Private Sub Class_Initialize()

    mIsInitialized = False

End Sub
Public Property Get IsInitialized() As Boolean

    IsInitialized = mIsInitialized

End Property

Public Property Let IsInitialized(ByVal Value As Boolean)

    mIsInitialized = Value

End Property
