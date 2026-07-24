Attribute VB_Name = "modDatabase"
Option Explicit

'=============================================================================
' Module      : modDatabase
' Project     : DUDS-PFMS (Project Financial Management System)
' Description : Database initialization, schema management, and CRUD wrapper
'               procedures. Uses Excel worksheets as the database engine.
'=============================================================================

'------------------------------------------------------------------------------
' Database Initialization
'------------------------------------------------------------------------------

Public Function InitializeDatabase() As Boolean
    '-----------------------------------------------------------------------
    ' Main entry point for database initialization. Creates all system
    ' worksheets, populates the Config sheet with schema version info,
    ' and sets up named ranges. Returns True on success.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.InitializeDatabase"
    
    On Error GoTo InitializeDatabase_Err
    
    ' Suppress screen flickering and alerts during setup
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Create each system sheet
    If Not CreateSystemSheet(SHT_CONFIG, CreateConfigHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_USERS, CreateUsersHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_PROJECTS, CreateProjectsHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_TRANSACTIONS, CreateTransactionsHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_BUDGET, CreateBudgetHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_LOG, CreateLogHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_AUDIT, CreateAuditHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_REPORTS, CreateReportsHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_CASHIN, CreateCashInHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_CASHOUT, CreateCashOutHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_JOURNAL, CreateJournalHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_ACCOUNT, CreateAccountHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_JOURNALHEADER, CreateJournalHeaderHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_JOURNALLINE, CreateJournalLineHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_LEDGER, CreateLedgerHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_WORKITEM, CreateWorkItemHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_BUDGETHEADER, CreateBudgetHeaderHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_BUDGETLINE, CreateBudgetLineHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_PROGRESS, CreateProgressHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_APPROVAL, CreateApprovalHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_VENDOR, CreateVendorHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_PURCHASEREQ, CreatePurchaseRequestHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_PURCHASEORDER, CreatePurchaseOrderHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_MATERIAL, CreateMaterialHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_MATTRANS, CreateMaterialTransactionHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_EMPLOYEE, CreateEmployeeHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_ATTENDANCE, CreateAttendanceHeader()) Then GoTo InitializeDatabase_Fail
    If Not CreateSystemSheet(SHT_PAYROLL, CreatePayrollHeader()) Then GoTo InitializeDatabase_Fail
    
    ' Populate configuration
    If Not SetupConfigSheet() Then GoTo InitializeDatabase_Fail
    
    ' Setup named ranges
    SetupNamedRanges
    
    ' Everything succeeded
    InitializeDatabase = True
    LogInfo FnName, MSG_DB_INIT, "Startup"
    GoTo InitializeDatabase_Exit

InitializeDatabase_Fail:
    InitializeDatabase = False
    LogInfo FnName, MSG_DB_INIT_FAIL, "Startup"
    MsgBox MSG_DB_INIT_FAIL, vbCritical, APP_NAME
    GoTo InitializeDatabase_Exit

InitializeDatabase_Err:
    HandleError FnName, Err.Number, Err.Description
    InitializeDatabase = False
    GoTo InitializeDatabase_Exit

InitializeDatabase_Exit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Function

Private Function CreateSystemSheet(ByVal SheetName As String, _
                                   ByRef HeaderData() As String) As Boolean
    '-----------------------------------------------------------------------
    ' Creates a worksheet with the given name (if it doesn't exist) and
    ' writes the header row. The sheet is hidden (xlSheetVeryHidden) after
    ' creation except for _Config and _Log which remain visible for now.
    ' Returns True if the sheet was created or already exists.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.CreateSystemSheet"
    
    Dim ws As Worksheet
    Dim i As Long
    
    On Error GoTo CreateSystemSheet_Err
    
    ' Check if sheet already exists
    If SheetExists(SheetName) Then
        Set ws = ThisWorkbook.Worksheets(SheetName)
        ' Ensure header row is correct
        For i = LBound(HeaderData) To UBound(HeaderData)
            ws.Cells(1, i + 1).Value = HeaderData(i)
        Next i
        CreateSystemSheet = True
        Exit Function
    End If
    
    ' Create new worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SheetName
    
    ' Write headers
    For i = LBound(HeaderData) To UBound(HeaderData)
        ws.Cells(1, i + 1).Value = HeaderData(i)
    Next i
    
    ' Format header row
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(HeaderData) + 1))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = CLR_HEADER
        .HorizontalAlignment = xlCenter
    End With
    
    ' Hide system sheets (except the first visible ones that users need)
    Select Case SheetName
        Case SHT_CONFIG, SHT_LOG, SHT_REPORTS:
            ' Keep visible for debugging/reporting
            ws.Visible = xlSheetVisible
        Case Else:
            ws.Visible = xlSheetVeryHidden
    End Select
    
    CreateSystemSheet = True
    Exit Function

CreateSystemSheet_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    CreateSystemSheet = False
End Function

Private Function SetupConfigSheet() As Boolean
    '-----------------------------------------------------------------------
    ' Populates the _Config sheet with initial configuration key-value pairs.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.SetupConfigSheet"
    
    Dim wsConfig As Worksheet
    Dim lRow As Long
    
    On Error GoTo SetupConfigSheet_Err
    
    Set wsConfig = ThisWorkbook.Worksheets(SHT_CONFIG)
    
    ' Write config data starting from row 2 (row 1 is headers)
    lRow = 2
    
    wsConfig.Cells(lRow, 1).Value = "DB_Version"
    wsConfig.Cells(lRow, 2).Value = APP_VERSION
    wsConfig.Cells(lRow, 3).Value = "Database schema version"
    lRow = lRow + 1
    
    wsConfig.Cells(lRow, 1).Value = "App_Name"
    wsConfig.Cells(lRow, 2).Value = APP_NAME
    wsConfig.Cells(lRow, 3).Value = "Application name"
    lRow = lRow + 1
    
    wsConfig.Cells(lRow, 1).Value = "DB_Created"
    wsConfig.Cells(lRow, 2).Value = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    wsConfig.Cells(lRow, 3).Value = "Database creation timestamp"
    lRow = lRow + 1
    
    wsConfig.Cells(lRow, 1).Value = "DB_Last_Opened"
    wsConfig.Cells(lRow, 2).Value = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    wsConfig.Cells(lRow, 3).Value = "Last database open timestamp"
    lRow = lRow + 1
    
    wsConfig.Cells(lRow, 1).Value = "VAT_Rate"
    wsConfig.Cells(lRow, 2).Value = FIN_VAT_RATE
    wsConfig.Cells(lRow, 3).Value = "Value Added Tax rate"
    lRow = lRow + 1
    
    wsConfig.Cells(lRow, 1).Value = "Currency"
    wsConfig.Cells(lRow, 2).Value = FIN_CURRENCY_CODE
    wsConfig.Cells(lRow, 3).Value = "Default currency code"
    lRow = lRow + 1
    
    wsConfig.Cells(lRow, 1).Value = "Session_Timeout"
    wsConfig.Cells(lRow, 2).Value = SEC_SESSION_TIMEOUT
    wsConfig.Cells(lRow, 3).Value = "Session timeout in minutes"
    lRow = lRow + 1
    
    wsConfig.Cells(lRow, 1).Value = "Max_Login_Attempts"
    wsConfig.Cells(lRow, 2).Value = SEC_MAX_LOGIN_ATTEMPTS
    wsConfig.Cells(lRow, 3).Value = "Maximum failed login attempts before lockout"
    
    ' Auto-fit columns
    wsConfig.Columns("A:C").AutoFit
    
    SetupConfigSheet = True
    Exit Function

SetupConfigSheet_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    SetupConfigSheet = False
End Function

Private Sub SetupNamedRanges()
    '-----------------------------------------------------------------------
    ' Creates workbook-scoped named ranges pointing to important cells/ranges.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.SetupNamedRanges"
    
    Dim wsConfig As Worksheet
    
    On Error Resume Next
    
    Set wsConfig = ThisWorkbook.Worksheets(SHT_CONFIG)
    
    ' DB_Version -> _Config!B2
    ThisWorkbook.Names.Add Name:=RNAME_DB_VERSION, _
        RefersTo:=wsConfig.Range("B2")
    
    ' CurrentUser -> _Config!B? (will be updated on login)
    ThisWorkbook.Names.Add Name:=RNAME_CURRENT_USER, _
        RefersTo:=wsConfig.Range("B3")
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Schema / Header Definitions
'------------------------------------------------------------------------------

Private Function CreateConfigHeader() As String()
    '-----------------------------------------------------------------------
    ' Returns the header array for the _Config sheet.
    '-----------------------------------------------------------------------
    Dim arr(0 To 2) As String
    arr(0) = "Key"
    arr(1) = "Value"
    arr(2) = "Description"
    CreateConfigHeader = arr
End Function

Private Function CreateUsersHeader() As String()
    '-----------------------------------------------------------------------
    ' Returns the header array for the _Users sheet.
    '-----------------------------------------------------------------------
    Dim arr(0 To 7) As String
    arr(0) = "UserID"
    arr(1) = "UserName"
    arr(2) = "PasswordHash"
    arr(3) = "Role"
    arr(4) = "Email"
    arr(5) = "Active"
    arr(6) = "CreatedOn"
    arr(7) = "LastLogin"
    CreateUsersHeader = arr
End Function

Private Function CreateProjectsHeader() As String()
    '-----------------------------------------------------------------------
    ' Returns the header array for the _Projects sheet.
    '-----------------------------------------------------------------------
    Dim arr(0 To 13) As String
    arr(0) = "ProjectID"
    arr(1) = "ProjectCode"
    arr(2) = "ProjectName"
    arr(3) = "ClientName"
    arr(4) = "Location"
    arr(5) = "StartDate"
    arr(6) = "EndDate"
    arr(7) = "ContractValue"
    arr(8) = "BudgetValue"
    arr(9) = "Status"
    arr(10) = "Progress"
    arr(11) = "Notes"
    arr(12) = "CreatedBy"
    arr(13) = "CreatedOn"
    CreateProjectsHeader = arr
End Function

Private Function CreateTransactionsHeader() As String()
    '-----------------------------------------------------------------------
    ' Returns the header array for the _Transactions sheet.
    '-----------------------------------------------------------------------
    Dim arr(0 To 11) As String
    arr(0) = "TransactionID"
    arr(1) = "Date"
    arr(2) = "ProjectCode"
    arr(3) = "Type"
    arr(4) = "Description"
    arr(5) = "Amount"
    arr(6) = "VAT"
    arr(7) = "Total"
    arr(8) = "ApprovedBy"
    arr(9) = "CreatedBy"
    arr(10) = "CreatedOn"
    arr(11) = "Status"
    CreateTransactionsHeader = arr
End Function

Private Function CreateBudgetHeader() As String()
    '-----------------------------------------------------------------------
    ' Returns the header array for the _Budget sheet.
    '-----------------------------------------------------------------------
    Dim arr(0 To 7) As String
    arr(0) = "BudgetID"
    arr(1) = "ProjectCode"
    arr(2) = "FiscalYear"
    arr(3) = "Category"
    arr(4) = "PlannedAmount"
    arr(5) = "ActualAmount"
    arr(6) = "Variance"
    arr(7) = "Notes"
    CreateBudgetHeader = arr
End Function

Private Function CreateLogHeader() As String()
    '-----------------------------------------------------------------------
    ' Returns the header array for the _Log sheet.
    '-----------------------------------------------------------------------
    Dim arr(0 To 5) As String
    arr(0) = "Timestamp"
    arr(1) = "Source"
    arr(2) = "Type"
    arr(3) = "Description"
    arr(4) = "Category"
    arr(5) = "User"
    CreateLogHeader = arr
End Function

Private Function CreateAuditHeader() As String()
    '-----------------------------------------------------------------------
    ' Returns the header array for the _Audit sheet.
    '-----------------------------------------------------------------------
    Dim arr(0 To 6) As String
    arr(0) = "AuditID"
    arr(1) = "Timestamp"
    arr(2) = "User"
    arr(3) = "Action"
    arr(4) = "TableName"
    arr(5) = "RecordID"
    arr(6) = "OldValue"
    arr(7) = "NewValue"
    CreateAuditHeader = arr
End Function

Private Function CreateReportsHeader() As String()
    '-----------------------------------------------------------------------
    ' Returns the header array for the _Reports sheet (summary area).
    '-----------------------------------------------------------------------
    Dim arr(0 To 4) As String
    arr(0) = "ReportID"
    arr(1) = "ReportName"
    arr(2) = "GeneratedOn"
    arr(3) = "GeneratedBy"
    arr(4) = "Summary"
    CreateReportsHeader = arr
End Function

Private Function CreateCashInHeader() As String()
    Dim arr(0 To 8) As String
    arr(0) = "CashInID"
    arr(1) = "ReceiptNo"
    arr(2) = "Date"
    arr(3) = "Project"
    arr(4) = "Account"
    arr(5) = "Description"
    arr(6) = "Amount"
    arr(7) = "CreatedBy"
    arr(8) = "CreatedOn"
    CreateCashInHeader = arr
End Function

Private Function CreateCashOutHeader() As String()
    Dim arr(0 To 9) As String
    arr(0) = "CashOutID"
    arr(1) = "VoucherNo"
    arr(2) = "Date"
    arr(3) = "Project"
    arr(4) = "Account"
    arr(5) = "Vendor"
    arr(6) = "Description"
    arr(7) = "Amount"
    arr(8) = "CreatedBy"
    arr(9) = "CreatedOn"
    CreateCashOutHeader = arr
End Function

Private Function CreateJournalHeader() As String()
    Dim arr(0 To 10) As String
    arr(0) = "JournalID"
    arr(1) = "Date"
    arr(2) = "Source"
    arr(3) = "RefNo"
    arr(4) = "Project"
    arr(5) = "Account"
    arr(6) = "Description"
    arr(7) = "Debit"
    arr(8) = "Credit"
    arr(9) = "CreatedBy"
    arr(10) = "CreatedOn"
    CreateJournalHeader = arr
End Function

Private Function CreateAccountHeader() As String()
    Dim arr(0 To 10) As String
    arr(0) = "AccountID"
    arr(1) = "AccountCode"
    arr(2) = "AccountName"
    arr(3) = "AccountType"
    arr(4) = "Category"
    arr(5) = "NormalBalance"
    arr(6) = "OpeningBalance"
    arr(7) = "CurrentBalance"
    arr(8) = "Active"
    arr(9) = "CreatedBy"
    arr(10) = "CreatedOn"
    CreateAccountHeader = arr
End Function

Private Function CreateJournalHeaderHeader() As String()
    Dim arr(0 To 9) As String
    arr(0) = "HeaderID"
    arr(1) = "JournalNo"
    arr(2) = "Date"
    arr(3) = "Source"
    arr(4) = "Reference"
    arr(5) = "Description"
    arr(6) = "Posted"
    arr(7) = "Locked"
    arr(8) = "CreatedBy"
    arr(9) = "CreatedOn"
    CreateJournalHeaderHeader = arr
End Function

Private Function CreateJournalLineHeader() As String()
    Dim arr(0 To 8) As String
    arr(0) = "LineID"
    arr(1) = "HeaderID"
    arr(2) = "AccountCode"
    arr(3) = "Description"
    arr(4) = "Debit"
    arr(5) = "Credit"
    arr(6) = "Project"
    arr(7) = "CreatedBy"
    arr(8) = "CreatedOn"
    CreateJournalLineHeader = arr
End Function

Private Function CreateLedgerHeader() As String()
    Dim arr(0 To 11) As String
    arr(0) = "LedgerID"
    arr(1) = "Date"
    arr(2) = "AccountCode"
    arr(3) = "AccountName"
    arr(4) = "Description"
    arr(5) = "Reference"
    arr(6) = "Debit"
    arr(7) = "Credit"
    arr(8) = "Balance"
    arr(9) = "FiscalYear"
    arr(10) = "CreatedBy"
    arr(11) = "CreatedOn"
    CreateLedgerHeader = arr
End Function

Private Function CreateWorkItemHeader() As String()
    Dim arr(0 To 10) As String
    arr(0) = "WorkItemID"
    arr(1) = "ProjectID"
    arr(2) = "ItemCode"
    arr(3) = "ItemName"
    arr(4) = "Unit"
    arr(5) = "Volume"
    arr(6) = "UnitPrice"
    arr(7) = "TotalBudget"
    arr(8) = "Category"
    arr(9) = "CreatedBy"
    arr(10) = "CreatedOn"
    CreateWorkItemHeader = arr
End Function

Private Function CreateBudgetHeaderHeader() As String()
    Dim arr(0 To 11) As String
    arr(0) = "BudgetHeaderID"
    arr(1) = "ProjectID"
    arr(2) = "BudgetNo"
    arr(3) = "FiscalYear"
    arr(4) = "Description"
    arr(5) = "TotalAmount"
    arr(6) = "Status"
    arr(7) = "ApprovedBy"
    arr(8) = "ApprovedDate"
    arr(9) = "Locked"
    arr(10) = "CreatedBy"
    arr(11) = "CreatedOn"
    CreateBudgetHeaderHeader = arr
End Function

Private Function CreateBudgetLineHeader() As String()
    Dim arr(0 To 8) As String
    arr(0) = "BudgetLineID"
    arr(1) = "HeaderID"
    arr(2) = "WorkItemID"
    arr(3) = "Volume"
    arr(4) = "UnitPrice"
    arr(5) = "Total"
    arr(6) = "Category"
    arr(7) = "CreatedBy"
    arr(8) = "CreatedOn"
    CreateBudgetLineHeader = arr
End Function

Private Function CreateProgressHeader() As String()
    Dim arr(0 To 10) As String
    arr(0) = "ProgressID"
    arr(1) = "ProjectID"
    arr(2) = "WorkItemID"
    arr(3) = "Date"
    arr(4) = "ActualVolume"
    arr(5) = "BudgetVolume"
    arr(6) = "PhysicalPct"
    arr(7) = "FinancialPct"
    arr(8) = "Notes"
    arr(9) = "CreatedBy"
    arr(10) = "CreatedOn"
    CreateProgressHeader = arr
End Function

Private Function CreateApprovalHeader() As String()
    Dim arr(0 To 9) As String
    arr(0) = "ApprovalID"
    arr(1) = "RecordType"
    arr(2) = "RecordID"
    arr(3) = "Level"
    arr(4) = "Approver"
    arr(5) = "Status"
    arr(6) = "Comments"
    arr(7) = "ActionDate"
    arr(8) = "CreatedBy"
    arr(9) = "CreatedOn"
    CreateApprovalHeader = arr
End Function

Private Function CreateVendorHeader() As String()
    Dim arr(0 To 10) As String
    arr(0) = "VendorID"
    arr(1) = "VendorCode"
    arr(2) = "VendorName"
    arr(3) = "Address"
    arr(4) = "ContactPerson"
    arr(5) = "Phone"
    arr(6) = "Email"
    arr(7) = "NPWP"
    arr(8) = "BankAccount"
    arr(9) = "CreatedBy"
    arr(10) = "CreatedOn"
    CreateVendorHeader = arr
End Function

Private Function CreatePurchaseRequestHeader() As String()
    Dim arr(0 To 8) As String
    arr(0) = "PRID"
    arr(1) = "PRNumber"
    arr(2) = "Project"
    arr(3) = "RequestDate"
    arr(4) = "RequestedBy"
    arr(5) = "Status"
    arr(6) = "Notes"
    arr(7) = "CreatedBy"
    arr(8) = "CreatedOn"
    CreatePurchaseRequestHeader = arr
End Function

Private Function CreatePurchaseOrderHeader() As String()
    Dim arr(0 To 10) As String
    arr(0) = "POID"
    arr(1) = "PONumber"
    arr(2) = "Vendor"
    arr(3) = "Project"
    arr(4) = "DeliveryDate"
    arr(5) = "PaymentTerms"
    arr(6) = "Total"
    arr(7) = "Status"
    arr(8) = "ApprovedBy"
    arr(9) = "CreatedBy"
    arr(10) = "CreatedOn"
    CreatePurchaseOrderHeader = arr
End Function

Private Function CreateMaterialHeader() As String()
    Dim arr(0 To 11) As String
    arr(0) = "MaterialID"
    arr(1) = "MaterialCode"
    arr(2) = "MaterialName"
    arr(3) = "Unit"
    arr(4) = "BudgetQty"
    arr(5) = "ActualQty"
    arr(6) = "RemainingQty"
    arr(7) = "UnitPrice"
    arr(8) = "TotalCost"
    arr(9) = "Project"
    arr(10) = "CreatedBy"
    arr(11) = "CreatedOn"
    CreateMaterialHeader = arr
End Function

Private Function CreateMaterialTransactionHeader() As String()
    Dim arr(0 To 8) As String
    arr(0) = "TransID"
    arr(1) = "MaterialID"
    arr(2) = "TransType"
    arr(3) = "Quantity"
    arr(4) = "Date"
    arr(5) = "RefNo"
    arr(6) = "Notes"
    arr(7) = "CreatedBy"
    arr(8) = "CreatedOn"
    CreateMaterialTransactionHeader = arr
End Function

Private Function CreateEmployeeHeader() As String()
    Dim arr(0 To 11) As String
    arr(0) = "EmployeeID"
    arr(1) = "NIK"
    arr(2) = "EmployeeName"
    arr(3) = "Address"
    arr(4) = "Phone"
    arr(5) = "Position"
    arr(6) = "DailyRate"
    arr(7) = "MonthlySalary"
    arr(8) = "BankAccount"
    arr(9) = "Active"
    arr(10) = "CreatedBy"
    arr(11) = "CreatedOn"
    CreateEmployeeHeader = arr
End Function

Private Function CreateAttendanceHeader() As String()
    Dim arr(0 To 9) As String
    arr(0) = "AttendanceID"
    arr(1) = "EmployeeID"
    arr(2) = "Date"
    arr(3) = "CheckIn"
    arr(4) = "CheckOut"
    arr(5) = "Overtime"
    arr(6) = "Status"
    arr(7) = "Notes"
    arr(8) = "CreatedBy"
    arr(9) = "CreatedOn"
    CreateAttendanceHeader = arr
End Function

Private Function CreatePayrollHeader() As String()
    Dim arr(0 To 11) As String
    arr(0) = "PayrollID"
    arr(1) = "EmployeeID"
    arr(2) = "Period"
    arr(3) = "BasicSalary"
    arr(4) = "Overtime"
    arr(5) = "Loan"
    arr(6) = "Deduction"
    arr(7) = "NetPay"
    arr(8) = "Status"
    arr(9) = "Posted"
    arr(10) = "CreatedBy"
    arr(11) = "CreatedOn"
    CreatePayrollHeader = arr
End Function

'------------------------------------------------------------------------------
' CRUD Operations
'------------------------------------------------------------------------------

Public Function InsertRecord(ByVal SheetName As String, _
                             ByRef Values() As Variant) As Long
    '-----------------------------------------------------------------------
    ' Inserts a new record into the specified system sheet.
    ' Values() is a 0-based array of column values (excluding the ID column
    ' which is auto-generated).
    ' Returns the row number where the record was inserted, or 0 on failure.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.InsertRecord"
    
    Dim ws As Worksheet
    Dim lRow As Long
    Dim i As Long
    
    On Error GoTo InsertRecord_Err
    
    Set ws = ThisWorkbook.Worksheets(SheetName)
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Write each value to the corresponding column
    For i = LBound(Values) To UBound(Values)
        ws.Cells(lRow, i + 1).Value = Values(i)
    Next i
    
    ' Audit trail
    Call WriteAuditEntry SheetName, lRow, "INSERT", "", JoinValues(Values)
    
    InsertRecord = lRow
    Exit Function

InsertRecord_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    InsertRecord = 0
End Function

Public Function UpdateRecord(ByVal SheetName As String, _
                             ByVal RowNum As Long, _
                             ByRef Values() As Variant) As Boolean
    '-----------------------------------------------------------------------
    ' Updates an existing record (identified by RowNum) with new values.
    ' Returns True on success.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.UpdateRecord"
    
    Dim ws As Worksheet
    Dim i As Long
    Dim sOldValues As String
    
    On Error GoTo UpdateRecord_Err
    
    Set ws = ThisWorkbook.Worksheets(SheetName)
    
    ' Capture old values for audit
    sOldValues = ReadRowValues(ws, RowNum)
    
    ' Write new values
    For i = LBound(Values) To UBound(Values)
        ws.Cells(RowNum, i + 1).Value = Values(i)
    Next i
    
    ' Audit trail
    Call WriteAuditEntry SheetName, RowNum, "UPDATE", sOldValues, JoinValues(Values)
    
    UpdateRecord = True
    Exit Function

UpdateRecord_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    UpdateRecord = False
End Function

Public Function DeleteRecord(ByVal SheetName As String, _
                             ByVal RowNum As Long) As Boolean
    '-----------------------------------------------------------------------
    ' Deletes a record from the specified system sheet by clearing its row.
    ' Returns True on success.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.DeleteRecord"
    
    Dim ws As Worksheet
    Dim sOldValues As String
    Dim lColCount As Long
    
    On Error GoTo DeleteRecord_Err
    
    Set ws = ThisWorkbook.Worksheets(SheetName)
    
    ' Capture old values for audit
    sOldValues = ReadRowValues(ws, RowNum)
    
    ' Clear the row (delete contents only, not the row itself)
    lColCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ws.Range(ws.Cells(RowNum, 1), ws.Cells(RowNum, lColCount)).ClearContents
    
    ' Audit trail
    Call WriteAuditEntry SheetName, RowNum, "DELETE", sOldValues, ""
    
    DeleteRecord = True
    Exit Function

DeleteRecord_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    DeleteRecord = False
End Function

Public Function GetRecord(ByVal SheetName As String, _
                          ByVal RowNum As Long) As Variant()
    '-----------------------------------------------------------------------
    ' Reads an entire record (row) and returns it as a 0-based variant array.
    ' Returns an empty array if the row has no data.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.GetRecord"
    
    Dim ws As Worksheet
    Dim lColCount As Long
    Dim i As Long
    Dim arrResult() As Variant
    
    On Error GoTo GetRecord_Err
    
    Set ws = ThisWorkbook.Worksheets(SheetName)
    lColCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ReDim arrResult(0 To lColCount - 1)
    
    For i = 1 To lColCount
        arrResult(i - 1) = ws.Cells(RowNum, i).Value
    Next i
    
    GetRecord = arrResult
    Exit Function

GetRecord_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    ReDim arrResult(0)
    arrResult(0) = Empty
    GetRecord = arrResult
End Function

Public Function FindRecord(ByVal SheetName As String, _
                           ByVal ColumnIndex As Long, _
                           ByVal SearchValue As Variant, _
                           Optional ByVal StartRow As Long = 2) As Long
    '-----------------------------------------------------------------------
    ' Finds a record by searching for SearchValue in the specified column.
    ' ColumnIndex is 0-based. Returns the row number if found, or 0 if not.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.FindRecord"
    
    Dim ws As Worksheet
    Dim rngFound As Range
    Dim lCol As Long
    
    On Error GoTo FindRecord_Err
    
    Set ws = ThisWorkbook.Worksheets(SheetName)
    lCol = ColumnIndex + 1  ' Convert to 1-based for Range
    
    ' Use Find method for efficient searching
    Set rngFound = ws.Columns(lCol).Find(What:=SearchValue, _
                                         After:=ws.Cells(StartRow, lCol), _
                                         LookIn:=xlValues, _
                                         LookAt:=xlWhole)
    
    If Not rngFound Is Nothing Then
        FindRecord = rngFound.Row
    Else
        FindRecord = 0
    End If
    
    Exit Function

FindRecord_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    FindRecord = 0
End Function

Public Function GetRecordCount(ByVal SheetName As String) As Long
    '-----------------------------------------------------------------------
    ' Returns the number of data rows (excluding header) in the given sheet.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.GetRecordCount"
    
    Dim ws As Worksheet
    
    On Error GoTo GetRecordCount_Err
    
    Set ws = ThisWorkbook.Worksheets(SheetName)
    GetRecordCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
    If GetRecordCount < 0 Then GetRecordCount = 0
    
    Exit Function

GetRecordCount_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    GetRecordCount = 0
End Function

'------------------------------------------------------------------------------
' Query Helpers
'------------------------------------------------------------------------------

Public Function GetRecordsByColumn(ByVal SheetName As String, _
                                   ByVal ColumnIndex As Long, _
                                   ByVal SearchValue As Variant) As Collection
    '-----------------------------------------------------------------------
    ' Returns a Collection of row numbers where ColumnIndex matches
    ' SearchValue. Useful for filtering.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.GetRecordsByColumn"
    
    Dim ws As Worksheet
    Dim lLastRow As Long
    Dim lCol As Long
    Dim i As Long
    Dim colResults As Collection
    
    On Error GoTo GetRecordsByColumn_Err
    
    Set ws = ThisWorkbook.Worksheets(SheetName)
    lCol = ColumnIndex + 1
    lLastRow = ws.Cells(ws.Rows.Count, lCol).End(xlUp).Row
    Set colResults = New Collection
    
    For i = 2 To lLastRow
        If ws.Cells(i, lCol).Value = SearchValue Then
            colResults.Add i
        End If
    Next i
    
    Set GetRecordsByColumn = colResults
    Exit Function

GetRecordsByColumn_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    Set GetRecordsByColumn = Nothing
End Function

Public Function GetAllRecords(ByVal SheetName As String) As Variant()
    '-----------------------------------------------------------------------
    ' Returns all data rows (excluding header) as a 2D variant array.
    ' Each row is a 0-based array of values. Returns empty array if no data.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.GetAllRecords"
    
    Dim ws As Worksheet
    Dim lLastRow As Long
    Dim lLastCol As Long
    Dim rngData As Range
    Dim arrData As Variant
    Dim arrResult() As Variant
    Dim i As Long, j As Long
    
    On Error GoTo GetAllRecords_Err
    
    Set ws = ThisWorkbook.Worksheets(SheetName)
    lLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lLastRow < 2 Then
        GetAllRecords = Array()
        Exit Function
    End If
    
    Set rngData = ws.Range(ws.Cells(2, 1), ws.Cells(lLastRow, lLastCol))
    arrData = rngData.Value
    
    ' Convert 2D variant array to array of rows
    ReDim arrResult(1 To UBound(arrData, 1) - 1)
    For i = 1 To UBound(arrData, 1)
        ReDim rowArr(0 To UBound(arrData, 2) - 1)
        For j = 1 To UBound(arrData, 2)
            rowArr(j - 1) = arrData(i, j)
        Next j
        arrResult(i) = rowArr
    Next i
    
    GetAllRecords = arrResult
    Exit Function

GetAllRecords_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    GetAllRecords = Array()
End Function

'------------------------------------------------------------------------------
' Audit Trail
'------------------------------------------------------------------------------

Private Sub WriteAuditEntry(ByVal TableName As String, _
                            ByVal RecordID As Long, _
                            ByVal Action As String, _
                            ByVal OldValue As String, _
                            ByVal NewValue As String)
    '-----------------------------------------------------------------------
    ' Writes an entry to the _Audit sheet for change tracking.
    ' Called automatically by InsertRecord, UpdateRecord, DeleteRecord.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.WriteAuditEntry"
    
    Dim wsAudit As Worksheet
    Dim lRow As Long
    
    On Error Resume Next
    
    Set wsAudit = ThisWorkbook.Worksheets(SHT_AUDIT)
    If wsAudit Is Nothing Then Exit Sub
    
    lRow = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Generate AuditID
    wsAudit.Cells(lRow, 1).Value = GenerateGUID()
    wsAudit.Cells(lRow, 2).Value = Now()
    wsAudit.Cells(lRow, 3).Value = GetCurrentUserName()
    wsAudit.Cells(lRow, 4).Value = Action
    wsAudit.Cells(lRow, 5).Value = TableName
    wsAudit.Cells(lRow, 6).Value = RecordID
    wsAudit.Cells(lRow, 7).Value = Left$(OldValue, 255)
    wsAudit.Cells(lRow, 8).Value = Left$(NewValue, 255)
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Internal Helpers
'------------------------------------------------------------------------------

Private Function JoinValues(ByRef Values() As Variant) As String
    '-----------------------------------------------------------------------
    ' Joins a variant array into a pipe-delimited string for audit logging.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.JoinValues"
    
    Dim i As Long
    Dim sResult As String
    
    sResult = ""
    For i = LBound(Values) To UBound(Values)
        sResult = sResult & CStr(Values(i)) & "|"
    Next i
    
    If Len(sResult) > 0 Then
        sResult = Left$(sResult, Len(sResult) - 1)
    End If
    
    JoinValues = sResult
End Function

Private Function ReadRowValues(ByRef ws As Worksheet, _
                               ByVal RowNum As Long) As String
    '-----------------------------------------------------------------------
    ' Reads all values from a row and returns them as a pipe-delimited string.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.ReadRowValues"
    
    Dim lColCount As Long
    Dim i As Long
    Dim sResult As String
    
    On Error Resume Next
    
    lColCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    sResult = ""
    
    For i = 1 To lColCount
        sResult = sResult & CStr(ws.Cells(RowNum, i).Value) & "|"
    Next i
    
    If Len(sResult) > 0 Then
        sResult = Left$(sResult, Len(sResult) - 1)
    End If
    
    ReadRowValues = sResult
    On Error GoTo 0
End Function

Private Function GetCurrentUserName() As String
    '-----------------------------------------------------------------------
    ' Returns the current logged-in user name from the global variable.
    ' Falls back to Windows username if not logged in.
    '-----------------------------------------------------------------------
    Const FnName As String = "modDatabase.GetCurrentUserName"
    
    If Len(gCurrentUser) > 0 Then
        GetCurrentUserName = gCurrentUser
    Else
        GetCurrentUserName = "SYSTEM"
    End If
End Function

