Attribute VB_Name = "modConfig"
Option Explicit

'=============================================================================
' Module      : modConfig
' Project     : DUDS-PFMS (Project Financial Management System)
' Description : Application-wide constants and configuration parameters.
'               All magic numbers/strings are centralized here.
'=============================================================================

'------------------------------------------------------------------------------
' Application Constants
'------------------------------------------------------------------------------
Public Const APP_NAME          As String = "DUDS-PFMS"
Public Const APP_VERSION       As String = "1.0.0"
Public Const APP_AUTHOR        As String = "DUDS Team"
Public Const APP_COMPANY       As String = "DUDS"
Public Const APP_FULL_NAME     As String = "DUDS Project Financial Management System"
Public Const APP_WORKBOOK_NAME As String = "DUDS-PFMS.xlsm"

'------------------------------------------------------------------------------
' Sheet Names (System / Hidden)
'------------------------------------------------------------------------------
Public Const SHT_CONFIG        As String = "_Config"
Public Const SHT_USERS         As String = "_Users"
Public Const SHT_PROJECTS      As String = "_Projects"
Public Const SHT_TRANSACTIONS  As String = "_Transactions"
Public Const SHT_BUDGET        As String = "_Budget"
Public Const SHT_LOG           As String = "_Log"
Public Const SHT_AUDIT         As String = "_Audit"
Public Const SHT_REPORTS       As String = "_Reports"
Public Const SHT_CASHIN        As String = "_CashIn"
Public Const SHT_CASHOUT       As String = "_CashOut"
Public Const SHT_JOURNAL       As String = "_Journal"
Public Const SHT_ACCOUNT       As String = "_Account"
Public Const SHT_JOURNALHEADER As String = "_JournalHeader"
Public Const SHT_JOURNALLINE   As String = "_JournalLine"
Public Const SHT_LEDGER        As String = "_Ledger"
Public Const SHT_WORKITEM      As String = "_WorkItem"
Public Const SHT_BUDGETHEADER  As String = "_BudgetHeader"
Public Const SHT_BUDGETLINE    As String = "_BudgetLine"
Public Const SHT_PROGRESS      As String = "_Progress"
Public Const SHT_APPROVAL      As String = "_Approval"
Public Const SHT_VENDOR        As String = "_Vendor"
Public Const SHT_PURCHASEREQ   As String = "_PurchaseRequest"
Public Const SHT_PURCHASEORDER As String = "_PurchaseOrder"
Public Const SHT_MATERIAL      As String = "_Material"
Public Const SHT_MATTRANS      As String = "_MaterialTransaction"
Public Const SHT_EMPLOYEE      As String = "_Employee"
Public Const SHT_ATTENDANCE    As String = "_Attendance"
Public Const SHT_PAYROLL       As String = "_Payroll"

'------------------------------------------------------------------------------
' Sheet Indexes (for use with Workbooks/Sheets)
'------------------------------------------------------------------------------
Public Const IDX_CONFIG        As Long = 1
Public Const IDX_USERS         As Long = 2
Public Const IDX_PROJECTS      As Long = 3
Public Const IDX_TRANSACTIONS  As Long = 4
Public Const IDX_BUDGET        As Long = 5
Public Const IDX_LOG           As Long = 6
Public Const IDX_AUDIT         As Long = 7
Public Const IDX_REPORTS       As Long = 8
Public Const IDX_CASHIN        As Long = 9
Public Const IDX_CASHOUT       As Long = 10
Public Const IDX_JOURNAL       As Long = 11
Public Const IDX_ACCOUNT       As Long = 12
Public Const IDX_JOURNALHEADER As Long = 13
Public Const IDX_JOURNALLINE   As Long = 14
Public Const IDX_LEDGER        As Long = 15
Public Const IDX_WORKITEM      As Long = 16
Public Const IDX_BUDGETHEADER  As Long = 17
Public Const IDX_BUDGETLINE    As Long = 18
Public Const IDX_PROGRESS      As Long = 19
Public Const IDX_APPROVAL      As Long = 20
Public Const IDX_VENDOR        As Long = 21
Public Const IDX_PURCHASEREQ   As Long = 22
Public Const IDX_PURCHASEORDER As Long = 23
Public Const IDX_MATERIAL      As Long = 24
Public Const IDX_MATTRANS      As Long = 25
Public Const IDX_EMPLOYEE      As Long = 26
Public Const IDX_ATTENDANCE    As Long = 27
Public Const IDX_PAYROLL       As Long = 28

'------------------------------------------------------------------------------
' Column Indices (0-based for Range.Offset use)
'------------------------------------------------------------------------------
' _Users table columns
Public Const COL_USER_ID       As Long = 0
Public Const COL_USER_NAME     As Long = 1
Public Const COL_USER_PASSWORD As Long = 2
Public Const COL_USER_ROLE     As Long = 3
Public Const COL_USER_EMAIL    As Long = 4
Public Const COL_USER_ACTIVE   As Long = 5
Public Const COL_USER_CREATED  As Long = 6
Public Const COL_USER_LASTLOGIN As Long = 7

' _Projects table columns
Public Const COL_PROJ_ID           As Long = 0
Public Const COL_PROJ_CODE         As Long = 1
Public Const COL_PROJ_NAME         As Long = 2
Public Const COL_PROJ_CLIENT       As Long = 3
Public Const COL_PROJ_LOCATION     As Long = 4
Public Const COL_PROJ_START_DATE   As Long = 5
Public Const COL_PROJ_END_DATE     As Long = 6
Public Const COL_PROJ_CONTRACT_VAL As Long = 7
Public Const COL_PROJ_BUDGET       As Long = 8
Public Const COL_PROJ_STATUS       As Long = 9
Public Const COL_PROJ_PROGRESS     As Long = 10
Public Const COL_PROJ_NOTES        As Long = 11
Public Const COL_PROJ_CREATED_BY   As Long = 12
Public Const COL_PROJ_CREATED_ON   As Long = 13

' _Transactions table columns
Public Const COL_TXN_ID            As Long = 0
Public Const COL_TXN_DATE          As Long = 1
Public Const COL_TXN_PROJ_CODE     As Long = 2
Public Const COL_TXN_TYPE          As Long = 3
Public Const COL_TXN_DESCRIPTION   As Long = 4
Public Const COL_TXN_AMOUNT        As Long = 5
Public Const COL_TXN_VAT           As Long = 6
Public Const COL_TXN_TOTAL         As Long = 7
Public Const COL_TXN_APPROVED_BY   As Long = 8
Public Const COL_TXN_CREATED_BY    As Long = 9
Public Const COL_TXN_CREATED_ON    As Long = 10
Public Const COL_TXN_STATUS        As Long = 11

' _CashIn table columns
Public Const COL_CASHIN_ID          As Long = 0
Public Const COL_CASHIN_RECEIPT_NO  As Long = 1
Public Const COL_CASHIN_DATE        As Long = 2
Public Const COL_CASHIN_PROJECT     As Long = 3
Public Const COL_CASHIN_ACCOUNT     As Long = 4
Public Const COL_CASHIN_DESC        As Long = 5
Public Const COL_CASHIN_AMOUNT      As Long = 6
Public Const COL_CASHIN_CREATED_BY  As Long = 7
Public Const COL_CASHIN_CREATED_ON  As Long = 8

' _CashOut table columns
Public Const COL_CASHOUT_ID          As Long = 0
Public Const COL_CASHOUT_VOUCHER_NO  As Long = 1
Public Const COL_CASHOUT_DATE        As Long = 2
Public Const COL_CASHOUT_PROJECT     As Long = 3
Public Const COL_CASHOUT_ACCOUNT     As Long = 4
Public Const COL_CASHOUT_VENDOR      As Long = 5
Public Const COL_CASHOUT_DESC        As Long = 6
Public Const COL_CASHOUT_AMOUNT      As Long = 7
Public Const COL_CASHOUT_CREATED_BY  As Long = 8
Public Const COL_CASHOUT_CREATED_ON  As Long = 9

' _Journal table columns
Public Const COL_JOURNAL_ID         As Long = 0
Public Const COL_JOURNAL_DATE       As Long = 1
Public Const COL_JOURNAL_SOURCE     As Long = 2
Public Const COL_JOURNAL_REF_NO     As Long = 3
Public Const COL_JOURNAL_PROJECT    As Long = 4
Public Const COL_JOURNAL_ACCOUNT    As Long = 5
Public Const COL_JOURNAL_DESC       As Long = 6
Public Const COL_JOURNAL_DEBIT      As Long = 7
Public Const COL_JOURNAL_CREDIT     As Long = 8
Public Const COL_JOURNAL_CREATED_BY As Long = 9
Public Const COL_JOURNAL_CREATED_ON As Long = 10

' _Account table columns
Public Const COL_ACCT_ID          As Long = 0
Public Const COL_ACCT_CODE        As Long = 1
Public Const COL_ACCT_NAME        As Long = 2
Public Const COL_ACCT_TYPE        As Long = 3
Public Const COL_ACCT_CATEGORY    As Long = 4
Public Const COL_ACCT_NORMAL_BAL  As Long = 5
Public Const COL_ACCT_OPEN_BAL    As Long = 6
Public Const COL_ACCT_CURRENT_BAL As Long = 7
Public Const COL_ACCT_ACTIVE      As Long = 8
Public Const COL_ACCT_CREATED_BY  As Long = 9
Public Const COL_ACCT_CREATED_ON  As Long = 10

' _JournalHeader table columns
Public Const COL_JH_ID        As Long = 0
Public Const COL_JH_JOURNALNO As Long = 1
Public Const COL_JH_DATE      As Long = 2
Public Const COL_JH_SOURCE    As Long = 3
Public Const COL_JH_REF       As Long = 4
Public Const COL_JH_DESC      As Long = 5
Public Const COL_JH_POSTED    As Long = 6
Public Const COL_JH_LOCKED    As Long = 7
Public Const COL_JH_CREATED_BY As Long = 8
Public Const COL_JH_CREATED_ON As Long = 9

' _JournalLine table columns
Public Const COL_JL_ID          As Long = 0
Public Const COL_JL_HEADER_ID   As Long = 1
Public Const COL_JL_ACCOUNT_CODE As Long = 2
Public Const COL_JL_DESC        As Long = 3
Public Const COL_JL_DEBIT       As Long = 4
Public Const COL_JL_CREDIT      As Long = 5
Public Const COL_JL_PROJECT     As Long = 6
Public Const COL_JL_CREATED_BY  As Long = 7
Public Const COL_JL_CREATED_ON  As Long = 8

' _Ledger table columns
Public Const COL_LEDGER_ID          As Long = 0
Public Const COL_LEDGER_DATE        As Long = 1
Public Const COL_LEDGER_ACCOUNT_CODE As Long = 2
Public Const COL_LEDGER_ACCOUNT_NAME As Long = 3
Public Const COL_LEDGER_DESC        As Long = 4
Public Const COL_LEDGER_REF         As Long = 5
Public Const COL_LEDGER_DEBIT       As Long = 6
Public Const COL_LEDGER_CREDIT      As Long = 7
Public Const COL_LEDGER_BALANCE     As Long = 8
Public Const COL_LEDGER_FISCAL_YEAR As Long = 9
Public Const COL_LEDGER_CREATED_BY  As Long = 10
Public Const COL_LEDGER_CREATED_ON  As Long = 11

' _Budget table columns
Public Const COL_BUDGET_ID       As Long = 0
Public Const COL_BUDGET_PROJ     As Long = 1
Public Const COL_BUDGET_FY       As Long = 2
Public Const COL_BUDGET_CAT      As Long = 3
Public Const COL_BUDGET_PLANNED  As Long = 4
Public Const COL_BUDGET_ACTUAL   As Long = 5
Public Const COL_BUDGET_VARIANCE As Long = 6
Public Const COL_BUDGET_NOTES    As Long = 7

' _WorkItem table columns
Public Const COL_WI_ID           As Long = 0
Public Const COL_WI_PROJECTID    As Long = 1
Public Const COL_WI_ITEMCODE     As Long = 2
Public Const COL_WI_ITEMNAME     As Long = 3
Public Const COL_WI_UNIT         As Long = 4
Public Const COL_WI_VOLUME       As Long = 5
Public Const COL_WI_UNITPRICE    As Long = 6
Public Const COL_WI_TOTALBUDGET  As Long = 7
Public Const COL_WI_CATEGORY     As Long = 8
Public Const COL_WI_CREATED_BY   As Long = 9
Public Const COL_WI_CREATED_ON   As Long = 10

' _BudgetHeader table columns
Public Const COL_BH_ID           As Long = 0
Public Const COL_BH_PROJECTID    As Long = 1
Public Const COL_BH_BUDGETNO     As Long = 2
Public Const COL_BH_FISCALYEAR   As Long = 3
Public Const COL_BH_DESCRIPTION  As Long = 4
Public Const COL_BH_TOTALAMOUNT  As Long = 5
Public Const COL_BH_STATUS       As Long = 6
Public Const COL_BH_APPROVEDBY   As Long = 7
Public Const COL_BH_APPROVEDDATE As Long = 8
Public Const COL_BH_LOCKED       As Long = 9
Public Const COL_BH_CREATED_BY   As Long = 10
Public Const COL_BH_CREATED_ON   As Long = 11

' _BudgetLine table columns
Public Const COL_BL_ID           As Long = 0
Public Const COL_BL_HEADERID     As Long = 1
Public Const COL_BL_WORKITEMID   As Long = 2
Public Const COL_BL_VOLUME       As Long = 3
Public Const COL_BL_UNITPRICE    As Long = 4
Public Const COL_BL_TOTAL        As Long = 5
Public Const COL_BL_CATEGORY     As Long = 6
Public Const COL_BL_CREATED_BY   As Long = 7
Public Const COL_BL_CREATED_ON   As Long = 8

' _Progress table columns
Public Const COL_PRG_ID          As Long = 0
Public Const COL_PRG_PROJECTID   As Long = 1
Public Const COL_PRG_WORKITEMID  As Long = 2
Public Const COL_PRG_DATE        As Long = 3
Public Const COL_PRG_ACTUALVOL   As Long = 4
Public Const COL_PRG_BUDGETVOL   As Long = 5
Public Const COL_PRG_PHYSICALPCT As Long = 6
Public Const COL_PRG_FINANCIALPCT As Long = 7
Public Const COL_PRG_NOTES       As Long = 8
Public Const COL_PRG_CREATED_BY  As Long = 9
Public Const COL_PRG_CREATED_ON  As Long = 10

' _Approval table columns
Public Const COL_APR_ID          As Long = 0
Public Const COL_APR_RECORDTYPE  As Long = 1
Public Const COL_APR_RECORDID    As Long = 2
Public Const COL_APR_LEVEL       As Long = 3
Public Const COL_APR_APPROVER    As Long = 4
Public Const COL_APR_STATUS      As Long = 5
Public Const COL_APR_COMMENTS    As Long = 6
Public Const COL_APR_ACTIONDATE  As Long = 7
Public Const COL_APR_CREATED_BY  As Long = 8
Public Const COL_APR_CREATED_ON  As Long = 9

' Table names (ListObject names)
Public Const TBL_ACCOUNT        As String = "tblAccount"
Public Const TBL_JOURNALHEADER  As String = "tblJournalHeader"
Public Const TBL_JOURNALLINE    As String = "tblJournalLine"
Public Const TBL_LEDGER         As String = "tblLedger"
Public Const TBL_WORKITEM       As String = "tblWorkItem"
Public Const TBL_BUDGETHEADER   As String = "tblBudgetHeader"
Public Const TBL_BUDGETLINE     As String = "tblBudgetLine"
Public Const TBL_PROGRESS       As String = "tblProgress"
Public Const TBL_APPROVAL       As String = "tblApproval"
Public Const TBL_VENDOR         As String = "tblVendor"
Public Const TBL_PURCHASEREQ    As String = "tblPurchaseRequest"
Public Const TBL_PURCHASEORDER  As String = "tblPurchaseOrder"
Public Const TBL_MATERIAL       As String = "tblMaterial"
Public Const TBL_MATTRANS       As String = "tblMaterialTransaction"
Public Const TBL_EMPLOYEE       As String = "tblEmployee"
Public Const TBL_ATTENDANCE     As String = "tblAttendance"
Public Const TBL_PAYROLL        As String = "tblPayroll"

' _Vendor table columns
Public Const COL_VEN_ID          As Long = 0
Public Const COL_VEN_CODE        As Long = 1
Public Const COL_VEN_NAME        As Long = 2
Public Const COL_VEN_ADDRESS     As Long = 3
Public Const COL_VEN_CONTACT     As Long = 4
Public Const COL_VEN_PHONE       As Long = 5
Public Const COL_VEN_EMAIL       As Long = 6
Public Const COL_VEN_NPWP        As Long = 7
Public Const COL_VEN_BANKACCT    As Long = 8
Public Const COL_VEN_CREATED_BY  As Long = 9
Public Const COL_VEN_CREATED_ON  As Long = 10

' _PurchaseRequest table columns
Public Const COL_PR_ID           As Long = 0
Public Const COL_PR_NUMBER       As Long = 1
Public Const COL_PR_PROJECT      As Long = 2
Public Const COL_PR_REQDATE      As Long = 3
Public Const COL_PR_REQUESTEDBY  As Long = 4
Public Const COL_PR_STATUS       As Long = 5
Public Const COL_PR_NOTES        As Long = 6
Public Const COL_PR_CREATED_BY   As Long = 7
Public Const COL_PR_CREATED_ON   As Long = 8

' _PurchaseOrder table columns
Public Const COL_PO_ID           As Long = 0
Public Const COL_PO_NUMBER       As Long = 1
Public Const COL_PO_VENDOR       As Long = 2
Public Const COL_PO_PROJECT      As Long = 3
Public Const COL_PO_DELIVERY     As Long = 4
Public Const COL_PO_PAYMENTTERMS As Long = 5
Public Const COL_PO_TOTAL        As Long = 6
Public Const COL_PO_STATUS       As Long = 7
Public Const COL_PO_APPROVEDBY   As Long = 8
Public Const COL_PO_CREATED_BY   As Long = 9
Public Const COL_PO_CREATED_ON   As Long = 10

' _Material table columns
Public Const COL_MAT_ID          As Long = 0
Public Const COL_MAT_CODE        As Long = 1
Public Const COL_MAT_NAME        As Long = 2
Public Const COL_MAT_UNIT        As Long = 3
Public Const COL_MAT_BUDGETQTY   As Long = 4
Public Const COL_MAT_ACTUALQTY   As Long = 5
Public Const COL_MAT_REMAINING   As Long = 6
Public Const COL_MAT_UNITPRICE   As Long = 7
Public Const COL_MAT_TOTALCOST   As Long = 8
Public Const COL_MAT_PROJECT     As Long = 9
Public Const COL_MAT_CREATED_BY  As Long = 10
Public Const COL_MAT_CREATED_ON  As Long = 11

' _MaterialTransaction table columns
Public Const COL_MT_ID           As Long = 0
Public Const COL_MT_MATERIALID   As Long = 1
Public Const COL_MT_TYPE         As Long = 2
Public Const COL_MT_QTY          As Long = 3
Public Const COL_MT_DATE         As Long = 4
Public Const COL_MT_REFNO        As Long = 5
Public Const COL_MT_NOTES        As Long = 6
Public Const COL_MT_CREATED_BY   As Long = 7
Public Const COL_MT_CREATED_ON   As Long = 8

' _Employee table columns
Public Const COL_EMP_ID          As Long = 0
Public Const COL_EMP_NIK         As Long = 1
Public Const COL_EMP_NAME        As Long = 2
Public Const COL_EMP_ADDRESS     As Long = 3
Public Const COL_EMP_PHONE       As Long = 4
Public Const COL_EMP_POSITION    As Long = 5
Public Const COL_EMP_DAILYRATE   As Long = 6
Public Const COL_EMP_MONTHLYSAL  As Long = 7
Public Const COL_EMP_BANKACCT    As Long = 8
Public Const COL_EMP_ACTIVE      As Long = 9
Public Const COL_EMP_CREATED_BY  As Long = 10
Public Const COL_EMP_CREATED_ON  As Long = 11

' _Attendance table columns
Public Const COL_ATT_ID          As Long = 0
Public Const COL_ATT_EMPLOYEEID  As Long = 1
Public Const COL_ATT_DATE        As Long = 2
Public Const COL_ATT_CHECKIN     As Long = 3
Public Const COL_ATT_CHECKOUT    As Long = 4
Public Const COL_ATT_OVERTIME    As Long = 5
Public Const COL_ATT_STATUS      As Long = 6
Public Const COL_ATT_NOTES       As Long = 7
Public Const COL_ATT_CREATED_BY  As Long = 8
Public Const COL_ATT_CREATED_ON  As Long = 9

' _Payroll table columns
Public Const COL_PAY_ID          As Long = 0
Public Const COL_PAY_EMPLOYEEID  As Long = 1
Public Const COL_PAY_PERIOD      As Long = 2
Public Const COL_PAY_BASICSAL    As Long = 3
Public Const COL_PAY_OVERTIME    As Long = 4
Public Const COL_PAY_LOAN        As Long = 5
Public Const COL_PAY_DEDUCTION   As Long = 6
Public Const COL_PAY_NETPAY      As Long = 7
Public Const COL_PAY_STATUS      As Long = 8
Public Const COL_PAY_POSTED      As Long = 9
Public Const COL_PAY_CREATED_BY  As Long = 10
Public Const COL_PAY_CREATED_ON  As Long = 11

' Material transaction types
Public Const MT_RECEIVE As String = "Receive"
Public Const MT_ISSUE   As String = "Issue"
Public Const MT_RETURN  As String = "Return"

' Purchase statuses
Public Const PO_DRAFT     As String = "Draft"
Public Const PO_SUBMITTED As String = "Submitted"
Public Const PO_APPROVED  As String = "Approved"
Public Const PO_DELIVERED As String = "Delivered"
Public Const PO_COMPLETED As String = "Completed"

' Attendance statuses
Public Const ATT_PRESENT As String = "Present"
Public Const ATT_ABSENT  As String = "Absent"
Public Const ATT_LEAVE   As String = "Leave"
Public Const ATT_SICK    As String = "Sick"

' Approval workflow statuses
Public Const APR_DRAFT      As String = "Draft"
Public Const APR_SUBMITTED  As String = "Submitted"
Public Const APR_APPROVED   As String = "Approved"
Public Const APR_REJECTED   As String = "Rejected"
Public Const APR_POSTED     As String = "Posted"

' Approval levels
Public Const APR_LEVEL_PM        As String = "Project Manager"
Public Const APR_LEVEL_ACCTG     As String = "Accounting"
Public Const APR_LEVEL_DIRECTOR  As String = "Director"
Public Const APR_LEVEL_OWNER     As String = "Owner"

' Notification thresholds
Public Const NOTIF_BUDGET_PCT    As Double = 0.9
Public Const NOTIF_CASH_LOW      As Double = 1000000
Public Const NOTIF_DEADLINE_DAYS As Long = 7

' Account type constants
Public Const ACCT_TYPE_ASSET     As String = "Asset"
Public Const ACCT_TYPE_LIABILITY As String = "Liability"
Public Const ACCT_TYPE_EQUITY    As String = "Equity"
Public Const ACCT_TYPE_INCOME    As String = "Income"
Public Const ACCT_TYPE_EXPENSE   As String = "Expense"

' Normal balance indicators
Public Const BAL_DEBIT  As String = "Debit"
Public Const BAL_CREDIT As String = "Credit"

' Account categories
Public Const ACCT_CAT_CURRENT_ASSET   As String = "Current Asset"
Public Const ACCT_CAT_FIXED_ASSET     As String = "Fixed Asset"
Public Const ACCT_CAT_CURRENT_LIAB    As String = "Current Liability"
Public Const ACCT_CAT_LONG_TERM_LIAB  As String = "Long Term Liability"
Public Const ACCT_CAT_OWNERS_EQUITY   As String = "Owner's Equity"
Public Const ACCT_CAT_OPERATING_INCOME As String = "Operating Income"
Public Const ACCT_CAT_OTHER_INCOME    As String = "Other Income"
Public Const ACCT_CAT_OPERATING_EXP   As String = "Operating Expense"
Public Const ACCT_CAT_OTHER_EXPENSE   As String = "Other Expense"

' Journal status
Public Const JH_POSTED As String = "Posted"
Public Const JH_DRAFT  As String = "Draft"
Public Const JH_REVERSED As String = "Reversed"

' Fiscal year constants
Public Const FY_START_MONTH As Long = 1
Public Const FY_END_MONTH   As Long = 12

'------------------------------------------------------------------------------
' Financial Constants
'------------------------------------------------------------------------------
Public Const FIN_VAT_RATE         As Double = 0.11  ' 11% VAT
Public Const FIN_CURRENCY_SYMBOL  As String = "Rp"
Public Const FIN_CURRENCY_CODE    As String = "IDR"
Public Const FIN_DECIMAL_PLACES   As Long = 2
Public Const FIN_TAX_EXEMPT_LIMIT As Double = 0#

'------------------------------------------------------------------------------
' Status Constants
'------------------------------------------------------------------------------
Public Const STATUS_ACTIVE     As String = "Active"
Public Const STATUS_INACTIVE   As String = "Inactive"
Public Const STATUS_PENDING    As String = "Pending"
Public Const STATUS_APPROVED   As String = "Approved"
Public Const STATUS_REJECTED   As String = "Rejected"
Public Const STATUS_COMPLETED  As String = "Completed"
Public Const STATUS_CANCELLED  As String = "Cancelled"
Public Const STATUS_DRAFT      As String = "Draft"

'------------------------------------------------------------------------------
' Role Constants
'------------------------------------------------------------------------------
Public Const ROLE_ADMIN       As String = "Admin"
Public Const ROLE_MANAGER     As String = "Manager"
Public Const ROLE_FINANCE     As String = "Finance"
Public Const ROLE_USER        As String = "User"
Public Const ROLE_VIEWER      As String = "Viewer"

'------------------------------------------------------------------------------
' Transaction Types
'------------------------------------------------------------------------------
Public Const TXN_INCOME       As String = "Income"
Public Const TXN_EXPENSE      As String = "Expense"
Public Const TXN_TRANSFER     As String = "Transfer"
Public Const TXN_ADJUSTMENT   As String = "Adjustment"

'------------------------------------------------------------------------------
' Message Constants
'------------------------------------------------------------------------------
Public Const MSG_WELCOME        As String = "Welcome to DUDS Project Financial Management System"
Public Const MSG_LOGIN_SUCCESS  As String = "Login successful."
Public Const MSG_LOGIN_FAIL     As String = "Invalid username or password."
Public Const MSG_SESSION_EXP    As String = "Your session has expired. Please login again."
Public Const MSG_ACCESS_DENIED  As String = "Access denied. Insufficient privileges."
Public Const MSG_SAVE_SUCCESS   As String = "Record saved successfully."
Public Const MSG_SAVE_FAIL      As String = "Failed to save record."
Public Const MSG_DELETE_CONFIRM As String = "Are you sure you want to delete this record?"
Public Const MSG_RECORD_NOT_FOUND As String = "Record not found."
Public Const MSG_INVALID_INPUT  As String = "Invalid input. Please check your data."
Public Const MSG_DB_INIT        As String = "Database initialized successfully."
Public Const MSG_DB_INIT_FAIL   As String = "Database initialization failed."
Public Const MSG_SHUTDOWN       As String = "Shutting down DUDS-PFMS."

'------------------------------------------------------------------------------
' Path / File Constants
'------------------------------------------------------------------------------
Public Const PATH_CONFIG_SECTION As String = "Config"
Public Const PATH_LOG_FILE       As String = "duds_pfms_log.txt"
Public Const PATH_BACKUP_FOLDER  As String = "Backup"

'------------------------------------------------------------------------------
' Security Constants
'------------------------------------------------------------------------------
Public Const SEC_SALT            As String = "DUDS2024PFMS"
Public Const SEC_SESSION_TIMEOUT As Long = 30  ' minutes
Public Const SEC_MIN_PASSWORD_LEN As Long = 6
Public Const SEC_MAX_LOGIN_ATTEMPTS As Long = 5
Public Const SEC_LOCKOUT_MINUTES  As Long = 15

'------------------------------------------------------------------------------
' System Limits
'------------------------------------------------------------------------------
Public Const SYS_MAX_RECORDS     As Long = 1048576
Public Const SYS_MAX_PROJECTS    As Long = 10000
Public Const SYS_MAX_USERS       As Long = 500
Public Const SYS_PAGE_SIZE       As Long = 100

'------------------------------------------------------------------------------
' Color Constants (RGB Long values for UI)
'------------------------------------------------------------------------------
Public Const CLR_HEADER         As Long = 12632256  ' Gray
Public Const CLR_SUCCESS        As Long = 5287936   ' Green
Public Const CLR_ERROR          As Long = 255       ' Red
Public Const CLR_WARNING        As Long = 65535     ' Yellow
Public Const CLR_INFO           As Long = 15773696  ' Light Blue

'------------------------------------------------------------------------------
' Named Range Names (for programmatic access)
'------------------------------------------------------------------------------
Public Const RNAME_DB_VERSION   As String = "DB_Version"
Public Const RNAME_CURRENT_USER As String = "CurrentUser"
Public Const RNAME_APP_STATE    As String = "AppState"

