Attribute VB_Name = "BalanceSheetService"
Option Explicit

Public Function GenerateBalanceSheet(ByVal FiscalYear As Long) As Variant
    Dim wsAcct As Worksheet, loAcct As ListObject, wsLedger As Worksheet, loLedger As ListObject
    Dim i As Long, n As Long, k As Long
    Dim dTotalAsset As Double, dTotalLiability As Double, dTotalEquity As Double
    Dim dNetProfit As Double
    Dim arrResult As Variant
    Dim sAcctCode As String, sAcctName As String, sAcctType As String
    Dim dBal As Double, dNormalBal As String, dOpeningBal As Double

    On Error GoTo GenerateBalanceSheet_Err

    Set wsAcct = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set loAcct = wsAcct.ListObjects(TBL_ACCOUNT)
    Set wsLedger = ThisWorkbook.Worksheets(SHT_LEDGER): Set loLedger = wsLedger.ListObjects(TBL_LEDGER)

    If loAcct.DataBodyRange Is Nothing Then GenerateBalanceSheet = Array(): Exit Function

    ReDim arrResult(1 To loAcct.DataBodyRange.Rows.Count + 10, 1 To 3)
    n = 1: arrResult(n, 1) = "ASSETS": arrResult(n, 2) = "": arrResult(n, 3) = "": n = 2
    dTotalAsset = 0: dTotalLiability = 0: dTotalEquity = 0

    For i = 1 To loAcct.DataBodyRange.Rows.Count
        sAcctType = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_TYPE + 1).Value)
        If sAcctType = ACCT_TYPE_ASSET Then
            sAcctCode = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value)
            sAcctName = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value)
            dNormalBal = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NORMAL_BAL + 1).Value)
            dOpeningBal = SafeConvertToDouble(loAcct.DataBodyRange.Cells(i, COL_ACCT_OPEN_BAL + 1).Value)
            dBal = dOpeningBal
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
            If dBal <> 0 Then
                arrResult(n, 1) = "  " & sAcctName: arrResult(n, 2) = Round(dBal, FIN_DECIMAL_PLACES): arrResult(n, 3) = ""
                dTotalAsset = dTotalAsset + dBal: n = n + 1
            End If
        End If
    Next i
    arrResult(n, 1) = "Total Assets": arrResult(n, 2) = "": arrResult(n, 3) = Round(dTotalAsset, FIN_DECIMAL_PLACES)
    n = n + 1: arrResult(n, 1) = "": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1

    ' Liabilities
    arrResult(n, 1) = "LIABILITIES": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    For i = 1 To loAcct.DataBodyRange.Rows.Count
        sAcctType = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_TYPE + 1).Value)
        If sAcctType = ACCT_TYPE_LIABILITY Then
            sAcctCode = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value)
            sAcctName = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value)
            dOpeningBal = SafeConvertToDouble(loAcct.DataBodyRange.Cells(i, COL_ACCT_OPEN_BAL + 1).Value)
            dBal = dOpeningBal
            If Not loLedger.DataBodyRange Is Nothing Then
                For k = 1 To loLedger.DataBodyRange.Rows.Count
                    If SafeConvertToString(loLedger.DataBodyRange.Cells(k, COL_LEDGER_ACCOUNT_CODE + 1).Value) = sAcctCode Then
                        If FiscalYear = 0 Or SafeConvertToLong(loLedger.DataBodyRange.Cells(k, COL_LEDGER_FISCAL_YEAR + 1).Value) = FiscalYear Then
                            dBal = dBal - SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_DEBIT + 1).Value)
                            dBal = dBal + SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_CREDIT + 1).Value)
                        End If
                    End If
                Next k
            End If
            If dBal <> 0 Then
                arrResult(n, 1) = "  " & sAcctName: arrResult(n, 2) = Round(dBal, FIN_DECIMAL_PLACES): arrResult(n, 3) = ""
                dTotalLiability = dTotalLiability + dBal: n = n + 1
            End If
        End If
    Next i
    arrResult(n, 1) = "Total Liabilities": arrResult(n, 2) = "": arrResult(n, 3) = Round(dTotalLiability, FIN_DECIMAL_PLACES)
    n = n + 1: arrResult(n, 1) = "": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1

    ' Equity
    arrResult(n, 1) = "EQUITY": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    For i = 1 To loAcct.DataBodyRange.Rows.Count
        sAcctType = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_TYPE + 1).Value)
        If sAcctType = ACCT_TYPE_EQUITY Then
            sAcctCode = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value)
            sAcctName = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value)
            dOpeningBal = SafeConvertToDouble(loAcct.DataBodyRange.Cells(i, COL_ACCT_OPEN_BAL + 1).Value)
            dBal = dOpeningBal
            If Not loLedger.DataBodyRange Is Nothing Then
                For k = 1 To loLedger.DataBodyRange.Rows.Count
                    If SafeConvertToString(loLedger.DataBodyRange.Cells(k, COL_LEDGER_ACCOUNT_CODE + 1).Value) = sAcctCode Then
                        If FiscalYear = 0 Or SafeConvertToLong(loLedger.DataBodyRange.Cells(k, COL_LEDGER_FISCAL_YEAR + 1).Value) = FiscalYear Then
                            dBal = dBal - SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_DEBIT + 1).Value)
                            dBal = dBal + SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_CREDIT + 1).Value)
                        End If
                    End If
                Next k
            End If
            arrResult(n, 1) = "  " & sAcctName: arrResult(n, 2) = Round(dBal, FIN_DECIMAL_PLACES): arrResult(n, 3) = ""
            dTotalEquity = dTotalEquity + dBal: n = n + 1
        End If
    Next i

    ' Add Net Profit/Loss to Equity
    Dim arrPL As Variant
    arrPL = GenerateProfitLoss(FiscalYear)
    If IsArray(arrPL) Then
        If UBound(arrPL, 1) > 0 Then
            dNetProfit = SafeConvertToDouble(arrPL(UBound(arrPL, 1), 4))
            arrResult(n, 1) = "  Current Year Profit/(Loss)": arrResult(n, 2) = Round(dNetProfit, FIN_DECIMAL_PLACES): arrResult(n, 3) = ""
            dTotalEquity = dTotalEquity + dNetProfit: n = n + 1
        End If
    End If

    arrResult(n, 1) = "Total Equity": arrResult(n, 2) = "": arrResult(n, 3) = Round(dTotalEquity, FIN_DECIMAL_PLACES)
    n = n + 1: arrResult(n, 1) = "": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    arrResult(n, 1) = "TOTAL LIABILITIES & EQUITY": arrResult(n, 2) = ""
    arrResult(n, 3) = Round(dTotalLiability + dTotalEquity, FIN_DECIMAL_PLACES)

    ReDim Preserve arrResult(1 To n, 1 To 3)
    GenerateBalanceSheet = arrResult
    Exit Function

GenerateBalanceSheet_Err:
    LogError "BalanceSheetService.GenerateBalanceSheet", Err.Description
    GenerateBalanceSheet = Array()
End Function

Public Function ExportBalanceSheetToSheet(ByVal FiscalYear As Long, ByVal TargetSheetName As String) As Boolean
    Dim arrData As Variant, ws As Worksheet, i As Long, j As Long
    On Error GoTo ExportBalanceSheetToSheet_Err
    arrData = GenerateBalanceSheet(FiscalYear)
    If Not IsArray(arrData) Then ExportBalanceSheetToSheet = False: Exit Function
    If UBound(arrData, 1) = 0 Then ExportBalanceSheetToSheet = False: Exit Function
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
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
    ExportBalanceSheetToSheet = True
    Exit Function
ExportBalanceSheetToSheet_Err: LogError "BalanceSheetService.ExportBalanceSheetToSheet", Err.Description: ExportBalanceSheetToSheet = False
End Function
