Attribute VB_Name = "modConstants"

Option Explicit

'==============================================================================
' OAA v2
' Module : modConstants
' Purpose: Global Application Constants
'==============================================================================

'--------------------------------------------------------------------------
' Application
'--------------------------------------------------------------------------
Public Const APP_NAME As String = "Open Association Accounting"

Public Const APP_VERSION As String = "2.0.0-alpha"

Public Const APP_AUTHOR As String = "Haluk Sandýkçý"

Public Const APP_COMPANY As String = "OAA Open Source Project"
'--------------------------------------------------------------------------
' Workbook
'--------------------------------------------------------------------------
Public Const DEFAULT_SHEET As String = "Dashboard"

Public Const SETTINGS_SHEET As String = "Settings"

Public Const LOG_SHEET As String = "Log"
'--------------------------------------------------------------------------
' Database
'--------------------------------------------------------------------------
Public Const DATABASE_VERSION As String = "2.0"
Public Const SHEET_MEMBERS As String = "Members"
Public Const TABLE_MEMBERS As String = "tblMembers"
Public Const SHEET_MEMBERLIST As String = "Üye Listesi"
Public Const TABLE_MEMBERLIST As String = "tblMemberList"
'==============================================================================
' Member Status
'==============================================================================

Public Const MEMBER_STATUS_ACTIVE As String = "Aktif"
Public Const MEMBER_STATUS_RESIGNED As String = "Ýstifa"
Public Const MEMBER_STATUS_LEAVE As String = "Ýzinli"
'==============================================================================
' SETTINGS SHEETS
'==============================================================================

Public Const SHEET_SETTINGS As String = "Settings"

'==============================================================================
' SETTINGS TABLES
'==============================================================================

Public Const TABLE_SETTINGS As String = "tblSettings"
Public Const TABLE_PARAMETERS As String = "tblParameters"
Public Const TABLE_LOOKUP As String = "tblLookup"
Public Const TABLE_CONFIGURATION As String = "tblConfiguration"
'==============================================================================
' ASSOCIATION SHEET
'==============================================================================

Public Const SHEET_ASSOCIATION As String = "Association"

'==============================================================================
' ASSOCIATION TABLE
'==============================================================================

Public Const TABLE_ASSOCIATION As String = "tblAssociation"
'==============================================================================
' ASSOCIATION TABLE COLUMNS
'==============================================================================

Public Const COL_ASSOCIATION_ID As String = "AssociationID"
Public Const COL_ASSOCIATION_CODE As String = "AssociationCode"
Public Const COL_ASSOCIATION_NAME As String = "AssociationName"
Public Const COL_SHORT_NAME As String = "ShortName"
Public Const COL_LOGO_FILENAME As String = "LogoFileName"

Public Const COL_ADDRESS As String = "Address"
Public Const COL_DISTRICT As String = "District"
Public Const COL_CITY As String = "City"
Public Const COL_POSTAL_CODE As String = "PostalCode"
Public Const COL_COUNTRY As String = "Country"

Public Const COL_PHONE As String = "Phone"
Public Const COL_MOBILE As String = "Mobile"
Public Const COL_EMAIL As String = "Email"
Public Const COL_WEBSITE As String = "Website"

Public Const COL_TAX_OFFICE As String = "TaxOffice"
Public Const COL_TAX_NUMBER As String = "TaxNumber"
Public Const COL_CLUB_NUMBER As String = "ClubNumber"

Public Const COL_DEFAULT_CURRENCY As String = "DefaultCurrency"
Public Const COL_FISCAL_YEAR_START As String = "FiscalYearStart"

Public Const COL_CHARTER_DATE As String = "CharterDate"
Public Const COL_FOUNDATION_DATE As String = "FoundationDate"

Public Const COL_CREATED_DATE As String = "CreatedDate"
Public Const COL_MODIFIED_DATE As String = "ModifiedDate"

Public Const COL_ACTIVE As String = "Active"
Public Const COL_NOTES As String = "Notes"


'==============================================================================
' Charge Type / Adjustment Type / Charge Status
'==============================================================================

Public Const CHARGE_TYPE_OFFICIAL As String = "RESMI"
Public Const CHARGE_TYPE_MEETING As String = "TOPLANTI"
Public Const CHARGE_TYPE_OTHER As String = "SAIR"

Public Const ADJUSTMENT_TYPE_NORMAL As String = "NORMAL"
Public Const ADJUSTMENT_TYPE_INCREASE As String = "ARTIS"

Public Const CHARGE_STATUS_ACTIVE As String = "AKTIF"


'==============================================================================
' Sheet / Table Names (Member Charges, Income, Receipts, Ledger)
'==============================================================================

Public Const SHEET_MEMBER_CHARGES As String = "ÜyeBorçlarý"
Public Const TABLE_MEMBER_CHARGES As String = "TBL_MEMBER_CHARGES"

Public Const SHEET_INCOME As String = "Gelir"
Public Const TABLE_INCOME As String = "TBL_INCOME"

Public Const SHEET_RECEIPTS As String = "Receipts"
Public Const TABLE_RECEIPTS As String = "TBL_RECEIPTS"

Public Const SHEET_LEDGER As String = "Ýþletme Defteri"
Public Const TABLE_LEDGER As String = "TBL_LEDGER"


'==============================================================================
' MEMBERS TABLE COLUMNS
'==============================================================================

Public Const COL_MEMBER_ID            As String = "MemberID"
Public Const COL_MEMBER_FIRST_NAME    As String = "FirstName"
Public Const COL_MEMBER_LAST_NAME     As String = "LastName"
Public Const COL_MEMBER_STATUS        As String = "Status"
Public Const COL_MEMBER_JOIN_DATE     As String = "JoinDate"
Public Const COL_MEMBER_LEAVE_DATE    As String = "LeaveDate"
Public Const COL_MEMBER_BIRTH_DATE    As String = "BirthDate"
Public Const COL_MEMBER_MARRIAGE_DATE As String = "MarriageDate"
Public Const COL_MEMBER_GSM           As String = "GSM"
Public Const COL_MEMBER_EMAIL         As String = "Email"
Public Const COL_MEMBER_ADDRESS       As String = "Address"
Public Const COL_MEMBER_SPOUSE_NAME   As String = "SpouseName"
Public Const COL_MEMBER_NOTES         As String = "Notes"


'==============================================================================
' MEMBER CHARGES TABLE COLUMNS
'==============================================================================

Public Const COL_CHARGE_ID               As String = "ChargeID"
Public Const COL_CHARGE_MEMBER_ID        As String = "MemberID"
Public Const COL_CHARGE_TYPE             As String = "ChargeType"
Public Const COL_CHARGE_DATE             As String = "ChargeDate"
Public Const COL_CHARGE_FISCAL_YEAR      As String = "FiscalYear"
Public Const COL_CHARGE_DUE_MONTH        As String = "DueMonth"
Public Const COL_CHARGE_AMOUNT           As String = "Amount"
Public Const COL_CHARGE_ADJUSTMENT_TYPE  As String = "AdjustmentType"
Public Const COL_CHARGE_SOURCE_CHARGE_ID As String = "SourceChargeID"
Public Const COL_CHARGE_DESCRIPTION      As String = "Description"
Public Const COL_CHARGE_DECISION_NO      As String = "DecisionNo"
Public Const COL_CHARGE_DECISION_DATE    As String = "DecisionDate"
Public Const COL_CHARGE_ENTRY_DATE       As String = "EntryDate"
Public Const COL_CHARGE_STATUS           As String = "Status"


'==============================================================================
' INCOME TABLE COLUMNS
'==============================================================================

Public Const COL_INCOME_ID          As String = "IncomeID"
Public Const COL_INCOME_ENTRY_DATE  As String = "EntryDate"
Public Const COL_INCOME_DOCUMENT_NO As String = "DocumentNo"
Public Const COL_INCOME_TYPE        As String = "IncomeType"
Public Const COL_INCOME_DESCRIPTION As String = "Description"
Public Const COL_INCOME_AMOUNT      As String = "IncomeAmount"
Public Const COL_INCOME_CASH_AMOUNT As String = "CashAmount"
Public Const COL_INCOME_BANK_AMOUNT As String = "BankAmount"


'==============================================================================
' RECEIPTS TABLE COLUMNS
'==============================================================================

Public Const COL_RECEIPT_ID             As String = "ReceiptID"
Public Const COL_RECEIPT_SERIES         As String = "Series"
Public Const COL_RECEIPT_NUMBER         As String = "ReceiptNumber"
Public Const COL_RECEIPT_DATE           As String = "ReceiptDate"
Public Const COL_RECEIPT_MEMBER_ID      As String = "MemberID"
Public Const COL_RECEIPT_CHARGE_ID      As String = "ChargeID"
Public Const COL_RECEIPT_AMOUNT         As String = "Amount"
Public Const COL_RECEIPT_PAYMENT_METHOD As String = "PaymentMethod"


'==============================================================================
' LEDGER TABLE COLUMNS
'==============================================================================

Public Const COL_LEDGER_ID              As String = "LedgerID"
Public Const COL_LEDGER_EXPENSE_ID        As String = "ExpenseID"
Public Const COL_LEDGER_INCOME_ID         As String = "IncomeID"
Public Const COL_LEDGER_ENTRY_DATE      As String = "EntryDate"
Public Const COL_LEDGER_DOCUMENT_NO     As String = "DocumentNo"
Public Const COL_LEDGER_RECEIPT_INVOICE As String = "ReceiptInvoice"
Public Const COL_LEDGER_DESCRIPTION     As String = "Description"
Public Const COL_LEDGER_INCOME_TYPE     As String = "IncomeType"
Public Const COL_LEDGER_EXPENSE_TYPE    As String = "ExpenseType"
Public Const COL_LEDGER_INCOME_AMOUNT   As String = "IncomeAmount"
Public Const COL_LEDGER_EXPENSE_AMOUNT  As String = "ExpenseAmount"
Public Const COL_LEDGER_CASH_AMOUNT     As String = "CashAmount"
Public Const COL_LEDGER_BANK_AMOUNT     As String = "BankAmount"
Public Const COL_LEDGER_BALANCE_AMOUNT  As String = "BalanceAmount"


'==============================================================================
' SETTINGS TABLE COLUMNS
'==============================================================================

Public Const COL_SETTING_KEY         As String = "SettingKey"
Public Const COL_SETTING_VALUE       As String = "SettingValue"
Public Const COL_SETTING_DATA_TYPE   As String = "DataType"
Public Const COL_SETTING_CATEGORY    As String = "Category"
Public Const COL_SETTING_IS_SYSTEM   As String = "IsSystem"
Public Const COL_SETTING_DESCRIPTION As String = "Description"


'==============================================================================
' EXPENSE (GÝDER) TABLE
'==============================================================================

Public Const SHEET_EXPENSE As String = "Gider"
Public Const TABLE_EXPENSE As String = "TBL_EXPENSE"

Public Const COL_EXPENSE_ID             As String = "ExpenseID"
Public Const COL_EXPENSE_ENTRY_DATE     As String = "EntryDate"
Public Const COL_EXPENSE_DOCUMENT_NO    As String = "DocumentNo"
Public Const COL_EXPENSE_TYPE           As String = "ExpenseType"
Public Const COL_EXPENSE_DESCRIPTION    As String = "Description"
Public Const COL_EXPENSE_AMOUNT         As String = "ExpenseAmount"
Public Const COL_EXPENSE_CASH_AMOUNT    As String = "CashAmount"
Public Const COL_EXPENSE_BANK_AMOUNT    As String = "BankAmount"

