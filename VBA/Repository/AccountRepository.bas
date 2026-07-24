Attribute VB_Name = "AccountRepository"
Option Explicit

Private Const TBL_ACCOUNT As String = "tblAccount"

Public Function CreateAccount(ByVal AccountID As String, ByVal AccountCode As String, ByVal AccountName As String, ByVal AccountType As String, ByVal Category As String, ByVal NormalBalance As String, ByVal OpeningBalance As Double) As Boolean
    Dim ws As Worksheet, lo As ListObject, lRow As Long
    On Error GoTo CreateAccount_Err
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set lo = ws.ListObjects(TBL_ACCOUNT)
    lRow = lo.ListRows.AddAlwaysNewRow.Position
    With lo.DataBodyRange
        .Cells(lRow, COL_ACCT_ID + 1).Value = AccountID
        .Cells(lRow, COL_ACCT_CODE + 1).Value = AccountCode
        .Cells(lRow, COL_ACCT_NAME + 1).Value = AccountName
        .Cells(lRow, COL_ACCT_TYPE + 1).Value = AccountType
        .Cells(lRow, COL_ACCT_CATEGORY + 1).Value = Category
        .Cells(lRow, COL_ACCT_NORMAL_BAL + 1).Value = NormalBalance
        .Cells(lRow, COL_ACCT_OPEN_BAL + 1).Value = OpeningBalance
        .Cells(lRow, COL_ACCT_CURRENT_BAL + 1).Value = OpeningBalance
        .Cells(lRow, COL_ACCT_ACTIVE + 1).Value = "Yes"
        .Cells(lRow, COL_ACCT_CREATED_BY + 1).Value = gCurrentUser
        .Cells(lRow, COL_ACCT_CREATED_ON + 1).Value = Now()
    End With
    CreateAccount = True: Exit Function
CreateAccount_Err: LogError "AccountRepository.CreateAccount", Err.Description: CreateAccount = False
End Function

Public Function UpdateAccountBalance(ByVal AccountCode As String, ByVal Amount As Double, ByVal IsDebit As Boolean) As Boolean
    Dim ws As Worksheet, lo As ListObject, i As Long, dBal As Double
    On Error GoTo UpdateAccountBalance_Err
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set lo = ws.ListObjects(TBL_ACCOUNT)
    If lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value = AccountCode Then
            dBal = SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_ACCT_CURRENT_BAL + 1).Value)
            If IsDebit Then dBal = dBal + Amount Else dBal = dBal - Amount
            lo.DataBodyRange.Cells(i, COL_ACCT_CURRENT_BAL + 1).Value = Round(dBal, FIN_DECIMAL_PLACES)
            UpdateAccountBalance = True: Exit Function
        End If
    Next i
    UpdateAccountBalance = False: Exit Function
UpdateAccountBalance_Err: LogError "AccountRepository.UpdateAccountBalance", Err.Description: UpdateAccountBalance = False
End Function

Public Function GetAccountBalance(ByVal AccountCode As String) As Double
    Dim ws As Worksheet, lo As ListObject, i As Long
    On Error GoTo GetAccountBalance_Err
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set lo = ws.ListObjects(TBL_ACCOUNT)
    If lo.DataBodyRange Is Nothing Then GetAccountBalance = 0: Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value = AccountCode Then
            GetAccountBalance = SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_ACCT_CURRENT_BAL + 1).Value): Exit Function
        End If
    Next i
    GetAccountBalance = 0
GetAccountBalance_Err: GetAccountBalance = 0
End Function

Public Function GetAccountCodeByName(ByVal AccountName As String) As String
    Dim ws As Worksheet, lo As ListObject, i As Long
    On Error GoTo GetAccountCodeByName_Err
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set lo = ws.ListObjects(TBL_ACCOUNT)
    If lo.DataBodyRange Is Nothing Then GetAccountCodeByName = "": Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value = AccountName Then
            GetAccountCodeByName = SafeConvertToString(lo.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value): Exit Function
        End If
    Next i
    GetAccountCodeByName = ""
GetAccountCodeByName_Err: GetAccountCodeByName = ""
End Function

Public Function LoadAccounts() As Variant
    Dim ws As Worksheet, lo As ListObject
    On Error GoTo LoadAccounts_Err
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set lo = ws.ListObjects(TBL_ACCOUNT)
    If lo.DataBodyRange Is Nothing Then LoadAccounts = Array(): Exit Function
    LoadAccounts = lo.DataBodyRange.Value: Exit Function
LoadAccounts_Err: LogError "AccountRepository.LoadAccounts", Err.Description: LoadAccounts = Array()
End Function

Public Function GetAccountsByType(ByVal AccountType As String) As Variant
    Dim ws As Worksheet, lo As ListObject, i As Long, j As Long, n As Long
    Dim arrData As Variant, arrResult As Variant
    On Error GoTo GetAccountsByType_Err
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set lo = ws.ListObjects(TBL_ACCOUNT)
    If lo.DataBodyRange Is Nothing Then GetAccountsByType = Array(): Exit Function
    arrData = lo.DataBodyRange.Value
    ReDim arrResult(1 To UBound(arrData, 1), 1 To UBound(arrData, 2))
    n = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, COL_ACCT_TYPE + 1) = AccountType Then
            n = n + 1
            For j = 1 To UBound(arrData, 2): arrResult(n, j) = arrData(i, j): Next j
        End If
    Next i
    If n = 0 Then GetAccountsByType = Array(): Exit Function
    ReDim Preserve arrResult(1 To n, 1 To UBound(arrData, 2))
    GetAccountsByType = arrResult
GetAccountsByType_Err: If Err.Number <> 0 Then LogError "AccountRepository.GetAccountsByType", Err.Description
End Function

Public Function GenerateAccountCode(ByVal AccountType As String) As String
    Dim sPrefix As String, ws As Worksheet, lo As ListObject, lMax As Long, i As Long, lCode As Long
    On Error GoTo GenerateAccountCode_Err
    Select Case AccountType
        Case ACCT_TYPE_ASSET: sPrefix = "1"
        Case ACCT_TYPE_LIABILITY: sPrefix = "2"
        Case ACCT_TYPE_EQUITY: sPrefix = "3"
        Case ACCT_TYPE_INCOME: sPrefix = "4"
        Case ACCT_TYPE_EXPENSE: sPrefix = "5"
        Case Else: sPrefix = "9"
    End Select
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set lo = ws.ListObjects(TBL_ACCOUNT)
    lMax = 0
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.DataBodyRange.Rows.Count
            lCode = SafeConvertToLong(lo.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value)
            If lCode > lMax Then lMax = lCode
        Next i
    End If
    GenerateAccountCode = sPrefix & Format$(lMax + 1, "0000")
    Exit Function
GenerateAccountCode_Err: LogError "AccountRepository.GenerateAccountCode", Err.Description: GenerateAccountCode = sPrefix & "0001"
End Function

Public Function CreateDefaultAccounts() As Boolean
    Dim bResult As Boolean: bResult = True
    bResult = bResult And CreateAccount(GenerateGUID(), "1001", "Cash", ACCT_TYPE_ASSET, ACCT_CAT_CURRENT_ASSET, BAL_DEBIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "1002", "Bank", ACCT_TYPE_ASSET, ACCT_CAT_CURRENT_ASSET, BAL_DEBIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "1003", "Accounts Receivable", ACCT_TYPE_ASSET, ACCT_CAT_CURRENT_ASSET, BAL_DEBIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "2001", "Accounts Payable", ACCT_TYPE_LIABILITY, ACCT_CAT_CURRENT_LIAB, BAL_CREDIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "2002", "Accrued Expenses", ACCT_TYPE_LIABILITY, ACCT_CAT_CURRENT_LIAB, BAL_CREDIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "3001", "Retained Earnings", ACCT_TYPE_EQUITY, ACCT_CAT_OWNERS_EQUITY, BAL_CREDIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "4001", "Project Revenue", ACCT_TYPE_INCOME, ACCT_CAT_OPERATING_INCOME, BAL_CREDIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "4002", "Other Income", ACCT_TYPE_INCOME, ACCT_CAT_OTHER_INCOME, BAL_CREDIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "5001", "Material Cost", ACCT_TYPE_EXPENSE, ACCT_CAT_OPERATING_EXP, BAL_DEBIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "5002", "Labor Cost", ACCT_TYPE_EXPENSE, ACCT_CAT_OPERATING_EXP, BAL_DEBIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "5003", "Operational Expense", ACCT_TYPE_EXPENSE, ACCT_CAT_OPERATING_EXP, BAL_DEBIT, 0)
    bResult = bResult And CreateAccount(GenerateGUID(), "5004", "Other Expense", ACCT_TYPE_EXPENSE, ACCT_CAT_OTHER_EXPENSE, BAL_DEBIT, 0)
    CreateDefaultAccounts = bResult
End Function

