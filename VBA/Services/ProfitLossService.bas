Attribute VB_Name = "ProfitLossService"
Option Explicit

Public Function GenerateProfitLoss(ByVal FiscalYear As Long) As Variant
    Dim wsAcct As Worksheet, loAcct As ListObject, wsLedger As Worksheet, loLedger As ListObject
    Dim i As Long, n As Long, k As Long
    Dim dIncomeTotal As Double, dExpenseTotal As Double
    Dim arrResult As Variant
    Dim sAcctCode As String, sAcctName As String, sAcctType As String, sCategory As String
    Dim dBal As Double, dNormalBal As String

    On Error GoTo GenerateProfitLoss_Err

    Set wsAcct = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set loAcct = wsAcct.ListObjects(TBL_ACCOUNT)
    Set wsLedger = ThisWorkbook.Worksheets(SHT_LEDGER): Set loLedger = wsLedger.ListObjects(TBL_LEDGER)

    If loAcct.DataBodyRange Is Nothing Then GenerateProfitLoss = Array(): Exit Function

    ReDim arrResult(1 To loAcct.DataBodyRange.Rows.Count + 5, 1 To 4)
    n = 1: arrResult(n, 1) = "INCOME": arrResult(n, 2) = "": arrResult(n, 3) = "": arrResult(n, 4) = "": n = 2
    dIncomeTotal = 0: dExpenseTotal = 0

    For i = 1 To loAcct.DataBodyRange.Rows.Count
        sAcctType = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_TYPE + 1).Value)
        If sAcctType = ACCT_TYPE_INCOME Then
            sAcctCode = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value)
            sAcctName = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value)
            sCategory = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CATEGORY + 1).Value)
            dNormalBal = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NORMAL_BAL + 1).Value)
            dBal = 0
            If Not loLedger.DataBodyRange Is Nothing Then
                For k = 1 To loLedger.DataBodyRange.Rows.Count
                    If SafeConvertToString(loLedger.DataBodyRange.Cells(k, COL_LEDGER_ACCOUNT_CODE + 1).Value) = sAcctCode Then
                        If FiscalYear = 0 Or SafeConvertToLong(loLedger.DataBodyRange.Cells(k, COL_LEDGER_FISCAL_YEAR + 1).Value) = FiscalYear Then
                            dBal = dBal + SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_CREDIT + 1).Value)
                            dBal = dBal - SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_DEBIT + 1).Value)
                        End If
                    End If
                Next k
            End If
            If dBal > 0 Then
                arrResult(n, 1) = "  " & sAcctName: arrResult(n, 2) = sCategory
                arrResult(n, 3) = Round(dBal, FIN_DECIMAL_PLACES): arrResult(n, 4) = ""
                dIncomeTotal = dIncomeTotal + dBal: n = n + 1
            End If
        End If
    Next i

    arrResult(n, 1) = "Total Income": arrResult(n, 2) = "": arrResult(n, 3) = ""
    arrResult(n, 4) = Round(dIncomeTotal, FIN_DECIMAL_PLACES): n = n + 1

    arrResult(n, 1) = "": arrResult(n, 2) = "": arrResult(n, 3) = "": arrResult(n, 4) = "": n = n + 1
    arrResult(n, 1) = "EXPENSES": arrResult(n, 2) = "": arrResult(n, 3) = "": arrResult(n, 4) = "": n = n + 1

    For i = 1 To loAcct.DataBodyRange.Rows.Count
        sAcctType = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_TYPE + 1).Value)
        If sAcctType = ACCT_TYPE_EXPENSE Then
            sAcctCode = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value)
            sAcctName = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value)
            sCategory = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CATEGORY + 1).Value)
            dBal = 0
            If Not loLedger.DataBodyRange Is Nothing Then
                For k = 1 To loLedger.DataBodyRange.Rows.Count
                    If SafeConvertToString(loLedger.DataBodyRange.Cells(k, COL_LEDGER_ACCOUNT_CODE + 1).Value) = sAcctCode Then
                        If FiscalYear = 0 Or SafeConvertToLong(loLedger.DataBodyRange.Cells(k, COL_LEDGER_FISCAL_YEAR + 1).Value) = FiscalYear Then
                            dBal = dBal + SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_DEBIT + 1).Value)
                            dBal = dBal - SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_CREDIT + 1).Value)
                        End If
                    End If
                Next k
            End If
            If dBal > 0 Then
                arrResult(n, 1) = "  " & sAcctName: arrResult(n, 2) = sCategory
                arrResult(n, 3) = Round(dBal, FIN_DECIMAL_PLACES): arrResult(n, 4) = ""
                dExpenseTotal = dExpenseTotal + dBal: n = n + 1
            End If
        End If
    Next i

    arrResult(n, 1) = "Total Expenses": arrResult(n, 2) = "": arrResult(n, 3) = ""
    arrResult(n, 4) = Round(dExpenseTotal, FIN_DECIMAL_PLACES): n = n + 1
    arrResult(n, 1) = "": arrResult(n, 2) = "": arrResult(n, 3) = "": arrResult(n, 4) = "": n = n + 1
    arrResult(n, 1) = "NET PROFIT / (LOSS)": arrResult(n, 2) = "": arrResult(n, 3) = ""
    arrResult(n, 4) = Round(dIncomeTotal - dExpenseTotal, FIN_DECIMAL_PLACES)

    ReDim Preserve arrResult(1 To n, 1 To 4)
    GenerateProfitLoss = arrResult
    Exit Function

GenerateProfitLoss_Err:
    LogError "ProfitLossService.GenerateProfitLoss", Err.Description
    GenerateProfitLoss = Array()
End Function

Public Function ExportProfitLossToSheet(ByVal FiscalYear As Long, ByVal TargetSheetName As String) As Boolean
    Dim arrData As Variant, ws As Worksheet, i As Long, j As Long
    On Error GoTo ExportProfitLossToSheet_Err
    arrData = GenerateProfitLoss(FiscalYear)
    If Not IsArray(arrData) Then ExportProfitLossToSheet = False: Exit Function
    If UBound(arrData, 1) = 0 Then ExportProfitLossToSheet = False: Exit Function
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
    ws.Range("A1:D1").Font.Bold = True
    ws.Columns("A:D").AutoFit
    ExportProfitLossToSheet = True
    Exit Function
ExportProfitLossToSheet_Err: LogError "ProfitLossService.ExportProfitLossToSheet", Err.Description: ExportProfitLossToSheet = False
End Function
