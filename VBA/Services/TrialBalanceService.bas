Attribute VB_Name = "TrialBalanceService"
Option Explicit

Public Function GenerateTrialBalance(ByVal FiscalYear As Long) As Variant
    Dim wsAcct As Worksheet, loAcct As ListObject, wsLedger As Worksheet, loLedger As ListObject
    Dim i As Long, n As Long, j As Long
    Dim dTotalDebit As Double, dTotalCredit As Double
    Dim sAcctCode As String, sAcctName As String
    Dim dDebit As Double, dCredit As Double
    Dim arrResult As Variant
    Dim dOpeningBal As Double, dNormalBal As String

    On Error GoTo GenerateTrialBalance_Err

    Set wsAcct = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set loAcct = wsAcct.ListObjects(TBL_ACCOUNT)
    Set wsLedger = ThisWorkbook.Worksheets(SHT_LEDGER): Set loLedger = wsLedger.ListObjects(TBL_LEDGER)

    If loAcct.DataBodyRange Is Nothing Then GenerateTrialBalance = Array(): Exit Function

    ReDim arrResult(1 To loAcct.DataBodyRange.Rows.Count + 3, 1 To 6)

    ' Header row
    arrResult(1, 1) = "Account Code": arrResult(1, 2) = "Account Name"
    arrResult(1, 3) = "Type": arrResult(1, 4) = "Opening Balance"
    arrResult(1, 5) = "Debit": arrResult(1, 6) = "Credit"

    n = 2
    dTotalDebit = 0: dTotalCredit = 0

    For i = 1 To loAcct.DataBodyRange.Rows.Count
        sAcctCode = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value)
        sAcctName = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value)
        dNormalBal = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NORMAL_BAL + 1).Value)
        dOpeningBal = SafeConvertToDouble(loAcct.DataBodyRange.Cells(i, COL_ACCT_OPEN_BAL + 1).Value)

        dDebit = 0: dCredit = 0

        ' Sum ledger entries for this account in the fiscal year
        If Not loLedger.DataBodyRange Is Nothing Then
            Dim k As Long
            For k = 1 To loLedger.DataBodyRange.Rows.Count
                If SafeConvertToString(loLedger.DataBodyRange.Cells(k, COL_LEDGER_ACCOUNT_CODE + 1).Value) = sAcctCode Then
                    If FiscalYear = 0 Or SafeConvertToLong(loLedger.DataBodyRange.Cells(k, COL_LEDGER_FISCAL_YEAR + 1).Value) = FiscalYear Then
                        dDebit = dDebit + SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_DEBIT + 1).Value)
                        dCredit = dCredit + SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_CREDIT + 1).Value)
                    End If
                End If
            Next k
        End If

        ' Determine debit/credit columns based on normal balance
        If dNormalBal = BAL_DEBIT Then
            dDebit = dDebit + dOpeningBal
        Else
            dCredit = dCredit + dOpeningBal
        End If

        arrResult(n, 1) = sAcctCode: arrResult(n, 2) = sAcctName
        arrResult(n, 3) = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_TYPE + 1).Value)
        arrResult(n, 4) = dOpeningBal
        arrResult(n, 5) = Round(dDebit, FIN_DECIMAL_PLACES)
        arrResult(n, 6) = Round(dCredit, FIN_DECIMAL_PLACES)

        dTotalDebit = dTotalDebit + dDebit
        dTotalCredit = dTotalCredit + dCredit

        n = n + 1
    Next i

    ' Totals row
    arrResult(n, 1) = "": arrResult(n, 2) = "TOTAL": arrResult(n, 3) = ""
    arrResult(n, 4) = "": arrResult(n, 5) = Round(dTotalDebit, FIN_DECIMAL_PLACES)
    arrResult(n, 6) = Round(dTotalCredit, FIN_DECIMAL_PLACES)

    n = n + 1
    arrResult(n, 1) = "": arrResult(n, 2) = "Difference": arrResult(n, 3) = ""
    arrResult(n, 4) = "": arrResult(n, 5) = ""
    arrResult(n, 6) = Round(dTotalDebit - dTotalCredit, FIN_DECIMAL_PLACES)

    ReDim Preserve arrResult(1 To n, 1 To 6)
    GenerateTrialBalance = arrResult
    Exit Function

GenerateTrialBalance_Err:
    LogError "TrialBalanceService.GenerateTrialBalance", Err.Description
    GenerateTrialBalance = Array()
End Function

Public Function ExportTrialBalanceToSheet(ByVal FiscalYear As Long, ByVal TargetSheetName As String) As Boolean
    Dim arrData As Variant, ws As Worksheet, i As Long, j As Long
    On Error GoTo ExportTrialBalanceToSheet_Err
    arrData = GenerateTrialBalance(FiscalYear)
    If Not IsArray(arrData) Then ExportTrialBalanceToSheet = False: Exit Function
    If UBound(arrData, 1) = 0 Then ExportTrialBalanceToSheet = False: Exit Function
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TargetSheetName)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = TargetSheetName
    ws.Cells.Clear
    For i = 1 To UBound(arrData, 1)
        For j = 1 To UBound(arrData, 2)
            ws.Cells(i, j).Value = arrData(i, j)
        Next j
    Next i
    ws.Range("A1:F1").Font.Bold = True
    ws.Columns("A:F").AutoFit
    ExportTrialBalanceToSheet = True
    Exit Function
ExportTrialBalanceToSheet_Err: LogError "TrialBalanceService.ExportTrialBalanceToSheet", Err.Description: ExportTrialBalanceToSheet = False
End Function

