Attribute VB_Name = "LedgerService"
Option Explicit

Private Const TBL_LEDGER As String = "tblLedger"

Public Function PostToLedger(ByVal HeaderID As String) As Boolean
    Dim wsLine As Worksheet, loLine As ListObject, wsLedger As Worksheet, loLedger As ListObject
    Dim wsJH As Worksheet, loJH As ListObject
    Dim i As Long, lRow As Long, dBal As Double, dPrevBal As Double
    Dim sAcctCode As String, sAcctName As String, sDesc As String
    Dim dDebit As Double, dCredit As Double, dDate As Date
    Dim sRef As String, sAcctType As String, sNormalBal As String

    On Error GoTo PostToLedger_Err

    Set wsLine = ThisWorkbook.Worksheets(SHT_JOURNALLINE): Set loLine = wsLine.ListObjects(TBL_JOURNALLINE)
    Set wsLedger = ThisWorkbook.Worksheets(SHT_LEDGER): Set loLedger = wsLedger.ListObjects(TBL_LEDGER)
    Set wsJH = ThisWorkbook.Worksheets(SHT_JOURNALHEADER): Set loJH = wsJH.ListObjects(TBL_JOURNALHEADER)

    If loLine.DataBodyRange Is Nothing Or loJH.DataBodyRange Is Nothing Then PostToLedger = False: Exit Function

    ' Get journal date and reference from header
    For i = 1 To loJH.DataBodyRange.Rows.Count
        If loJH.DataBodyRange.Cells(i, COL_JH_ID + 1).Value = HeaderID Then
            dDate = CDate(loJH.DataBodyRange.Cells(i, COL_JH_DATE + 1).Value)
            sRef = SafeConvertToString(loJH.DataBodyRange.Cells(i, COL_JH_JOURNALNO + 1).Value)
            Exit For
        End If
    Next i

    For i = 1 To loLine.DataBodyRange.Rows.Count
        If loLine.DataBodyRange.Cells(i, COL_JL_HEADER_ID + 1).Value = HeaderID Then
            sAcctCode = SafeConvertToString(loLine.DataBodyRange.Cells(i, COL_JL_ACCOUNT_CODE + 1).Value)
            sDesc = SafeConvertToString(loLine.DataBodyRange.Cells(i, COL_JL_DESC + 1).Value)
            dDebit = SafeConvertToDouble(loLine.DataBodyRange.Cells(i, COL_JL_DEBIT + 1).Value)
            dCredit = SafeConvertToDouble(loLine.DataBodyRange.Cells(i, COL_JL_CREDIT + 1).Value)
            sAcctName = GetAccountName(sAcctCode)
            sAcctType = GetAccountType(sAcctCode)
            sNormalBal = GetNormalBalance(sAcctCode)

            ' Calculate running balance
            dPrevBal = GetAccountBalance(sAcctCode)
            If sNormalBal = BAL_DEBIT Then
                dBal = Round(dPrevBal + dDebit - dCredit, FIN_DECIMAL_PLACES)
            Else
                dBal = Round(dPrevBal + dCredit - dDebit, FIN_DECIMAL_PLACES)
            End If

            lRow = loLedger.ListRows.AddAlwaysNewRow.Position
            With loLedger.DataBodyRange
                .Cells(lRow, COL_LEDGER_ID + 1).Value = GenerateGUID()
                .Cells(lRow, COL_LEDGER_DATE + 1).Value = dDate
                .Cells(lRow, COL_LEDGER_ACCOUNT_CODE + 1).Value = sAcctCode
                .Cells(lRow, COL_LEDGER_ACCOUNT_NAME + 1).Value = sAcctName
                .Cells(lRow, COL_LEDGER_DESC + 1).Value = sDesc
                .Cells(lRow, COL_LEDGER_REF + 1).Value = sRef
                .Cells(lRow, COL_LEDGER_DEBIT + 1).Value = dDebit
                .Cells(lRow, COL_LEDGER_CREDIT + 1).Value = dCredit
                .Cells(lRow, COL_LEDGER_BALANCE + 1).Value = dBal
                .Cells(lRow, COL_LEDGER_FISCAL_YEAR + 1).Value = FiscalYear(dDate)
                .Cells(lRow, COL_LEDGER_CREATED_BY + 1).Value = gCurrentUser
                .Cells(lRow, COL_LEDGER_CREATED_ON + 1).Value = Now()
            End With

            ' Update account current balance
            UpdateAccountBalance sAcctCode, dDebit, True
            UpdateAccountBalance sAcctCode, dCredit, False
        End If
    Next i

    PostToLedger = True
    Exit Function
PostToLedger_Err: LogError "LedgerService.PostToLedger", Err.Description: PostToLedger = False
End Function

Public Function ReverseLedgerEntries(ByVal OriginalHeaderID As String) As Boolean
    Dim wsLedger As Worksheet, loLedger As ListObject, i As Long
    On Error GoTo ReverseLedgerEntries_Err
    Set wsLedger = ThisWorkbook.Worksheets(SHT_LEDGER): Set loLedger = wsLedger.ListObjects(TBL_LEDGER)
    If loLedger.DataBodyRange Is Nothing Then ReverseLedgerEntries = True: Exit Function
    ' Find all ledger entries referencing the reversed journal and clear them
    i = 1
    Do While i <= loLedger.DataBodyRange.Rows.Count
        If loLedger.DataBodyRange.Cells(i, COL_LEDGER_REF + 1).Value = OriginalHeaderID Then
            loLedger.DataBodyRange.Rows(i).Delete
        Else
            i = i + 1
        End If
    Loop
    ReverseLedgerEntries = True
ReverseLedgerEntries_Err: LogError "LedgerService.ReverseLedgerEntries", Err.Description: ReverseLedgerEntries = False
End Function

Private Function GetAccountName(ByVal AccountCode As String) As String
    Dim ws As Worksheet, lo As ListObject, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set lo = ws.ListObjects(TBL_ACCOUNT)
    If lo.DataBodyRange Is Nothing Then GetAccountName = "": Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value = AccountCode Then
            GetAccountName = SafeConvertToString(lo.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value): Exit Function
        End If
    Next i
    GetAccountName = ""
End Function

Private Function GetAccountType(ByVal AccountCode As String) As String
    Dim ws As Worksheet, lo As ListObject, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set lo = ws.ListObjects(TBL_ACCOUNT)
    If lo.DataBodyRange Is Nothing Then GetAccountType = "": Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value = AccountCode Then
            GetAccountType = SafeConvertToString(lo.DataBodyRange.Cells(i, COL_ACCT_TYPE + 1).Value): Exit Function
        End If
    Next i
    GetAccountType = ""
End Function

Private Function GetNormalBalance(ByVal AccountCode As String) As String
    Dim ws As Worksheet, lo As ListObject, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set lo = ws.ListObjects(TBL_ACCOUNT)
    If lo.DataBodyRange Is Nothing Then GetNormalBalance = BAL_DEBIT: Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value = AccountCode Then
            GetNormalBalance = SafeConvertToString(lo.DataBodyRange.Cells(i, COL_ACCT_NORMAL_BAL + 1).Value): Exit Function
        End If
    Next i
    GetNormalBalance = BAL_DEBIT
End Function

Public Function GetLedgerEntries(ByVal AccountCode As String, ByVal FiscalYear As Long) As Variant
    Dim ws As Worksheet, lo As ListObject, i As Long, n As Long, j As Long
    Dim arrResult As Variant
    On Error GoTo GetLedgerEntries_Err
    Set ws = ThisWorkbook.Worksheets(SHT_LEDGER): Set lo = ws.ListObjects(TBL_LEDGER)
    If lo.DataBodyRange Is Nothing Then GetLedgerEntries = Array(): Exit Function
    ReDim arrResult(1 To lo.DataBodyRange.Rows.Count, 1 To lo.DataBodyRange.Columns.Count)
    n = 0
    For i = 1 To lo.DataBodyRange.Rows.Count
        If SafeConvertToString(lo.DataBodyRange.Cells(i, COL_LEDGER_ACCOUNT_CODE + 1).Value) = AccountCode Then
            If FiscalYear = 0 Or SafeConvertToLong(lo.DataBodyRange.Cells(i, COL_LEDGER_FISCAL_YEAR + 1).Value) = FiscalYear Then
                n = n + 1
                For j = 1 To lo.DataBodyRange.Columns.Count
                    arrResult(n, j) = lo.DataBodyRange.Cells(i, j).Value
                Next j
            End If
        End If
    Next i
    If n = 0 Then GetLedgerEntries = Array(): Exit Function
    ReDim Preserve arrResult(1 To n, 1 To lo.DataBodyRange.Columns.Count)
    GetLedgerEntries = arrResult
GetLedgerEntries_Err: If Err.Number <> 0 Then LogError "LedgerService.GetLedgerEntries", Err.Description
End Function

