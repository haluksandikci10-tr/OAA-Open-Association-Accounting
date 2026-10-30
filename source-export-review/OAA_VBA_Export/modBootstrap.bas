Attribute VB_Name = "modBootstrap"
Option Explicit

'==============================================================================
' OAA v2 - Open Association Accounting
'------------------------------------------------------------------------------
' Module Name : modBootstrap
' Layer       : Bootstrap
' Purpose     : Application Bootstrap and Lifecycle Management
' Author      : Haluk Sandýkçý
' Copyright   : (c) 2026 OAA Project
' Version     : 2.0.0-alpha.1
'------------------------------------------------------------------------------
' Change History
'------------------------------------------------------------------------------
' 2026-07-19  HS  Initial creation.
'==============================================================================

'------------------------------------------------------------------------------
' Initializes the OAA application framework.
'------------------------------------------------------------------------------
Public Sub InitializeApplication()

    ' TODO:
    ' Create ApplicationContext
    ' Initialize ServiceLocator
    ' Load Settings
    ' Initialize Logger

End Sub

'------------------------------------------------------------------------------
' Shuts down the OAA application framework.
'------------------------------------------------------------------------------
Public Sub ShutdownApplication()

    ' TODO:
    ' Flush Logger
    ' Save Settings
    ' Release Services
    ' Release ApplicationContext

End Sub

'------------------------------------------------------------------------------
' Restarts the OAA application framework.
'------------------------------------------------------------------------------
Public Sub RestartApplication()

    ShutdownApplication
    InitializeApplication

End Sub
