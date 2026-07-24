Attribute VB_Name = "CashFlowService"
Option Explicit

Public Function GenerateCashFlow(ByVal FiscalYear As Long) As Variant
    Dim wsAcct As Worksheet, loAcct As ListObject, wsLedger As Worksheet, loLedger As ListObject
    Dim i As Long, n As Long, k As Long
    Dim dOperatingIn As Double, dOperatingOut As Double
    Dim dInvestingIn As Double, dInvestingOut As Double
    Dim dFinancingIn As Double, dFinancingOut As Double
    Dim arrResult As Variant
    Dim sAcctCode As String, sAcctName As String, sAcctType As String
    Dim dDebit As Double, dCredit As Double, dNet As Double
    Dim dBegCash As Double, dEndCash As Double

    On Error GoTo GenerateCashFlow_Err

    Set wsAcct = ThisWorkbook.Worksheets(SHT_ACCOUNT): Set loAcct = wsAcct.ListObjects(TBL_ACCOUNT)
    Set wsLedger = ThisWorkbook.Worksheets(SHT_LEDGER): Set loLedger = wsLedger.ListObjects(TBL_LEDGER)

    If loAcct.DataBodyRange Is Nothing Then GenerateCashFlow = Array(): Exit Function

    ' Find Cash account code
    Dim sCashCode As String
    sCashCode = ""
    For i = 1 To loAcct.DataBodyRange.Rows.Count
        If SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value) = "Cash" Or _
           SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value) = "Bank" Then
            sCashCode = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value)
            Exit For
        End If
    Next i
    If sCashCode = "" Then
        sCashCode = SafeConvertToString(loAcct.DataBodyRange.Cells(1, COL_ACCT_CODE + 1).Value)
    End If

    ' Get beginning cash balance (opening balance)
    For i = 1 To loAcct.DataBodyRange.Rows.Count
        If SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value) = sCashCode Then
            dBegCash = SafeConvertToDouble(loAcct.DataBodyRange.Cells(i, COL_ACCT_OPEN_BAL + 1).Value)
            Exit For
        End If
    Next i

    ReDim arrResult(1 To 25, 1 To 3)
    n = 1

    ' Header
    arrResult(n, 1) = "CASH FLOW STATEMENT": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    If FiscalYear > 0 Then
        arrResult(n, 1) = "Fiscal Year: " & FiscalYear: arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    End If
    arrResult(n, 1) = "": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1

    ' Operating Activities
    arrResult(n, 1) = "CASH FLOW FROM OPERATING ACTIVITIES": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    dOperatingIn = 0: dOperatingOut = 0

    For i = 1 To loAcct.DataBodyRange.Rows.Count
        sAcctType = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_TYPE + 1).Value)
        If sAcctType = ACCT_TYPE_INCOME Or sAcctType = ACCT_TYPE_EXPENSE Then
            sAcctCode = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_CODE + 1).Value)
            sAcctName = SafeConvertToString(loAcct.DataBodyRange.Cells(i, COL_ACCT_NAME + 1).Value)
            dDebit = 0: dCredit = 0
            If Not loLedger.DataBodyRange Is Nothing Then
                For k = 1 To loLedger.DataBodyRange.Rows.Count
                    If SafeConvertToString(loLedger.DataBodyRange.Cells(k, COL_LEDGER_ACCOUNT_CODE + 1).Value) = sAcctCode Then
                        If FiscalYear = 0 Or SafeConvertToLong(loLedger.DataBodyRange.Cells(k, COL_LEDGER_FISCAL_YEAR + 1).Value) = FiscalYear Then
                            dDebit = dDebit + SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_DEBIT + 1).Value)
                            dCredit = dCredit + SafeConvertToDouble(loLedger.DataBodyRange.Cells(k, COL_LEDGER_CREDIT + 1).Value)
                        End If
                    End If
                Next k
            End If
            dNet = dCredit - dDebit
            If dNet <> 0 Then
                arrResult(n, 1) = "  " & sAcctName
                If dNet > 0 Then
                    arrResult(n, 2) = Round(dNet, FIN_DECIMAL_PLACES): dOperatingIn = dOperatingIn + dNet
                Else
                    arrResult(n, 3) = Round(Abs(dNet), FIN_DECIMAL_PLACES): dOperatingOut = dOperatingOut + Abs(dNet)
                End If
                n = n + 1
            End If
        End If
    Next i

    arrResult(n, 1) = "  Net Cash from Operations": arrResult(n, 2) = ""
    arrResult(n, 3) = Round(dOperatingIn - dOperatingOut, FIN_DECIMAL_PLACES): n = n + 1
    arrResult(n, 1) = "": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1

    ' Investing Activities
    arrResult(n, 1) = "CASH FLOW FROM INVESTING ACTIVITIES": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    arrResult(n, 1) = "  (Not yet implemented)": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    arrResult(n, 1) = "": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1

    ' Financing Activities
    arrResult(n, 1) = "CASH FLOW FROM FINANCING ACTIVITIES": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    arrResult(n, 1) = "  (Not yet implemented)": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    arrResult(n, 1) = "": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1

    ' Summary
    dEndCash = dBegCash + (dOperatingIn - dOperatingOut)
    arrResult(n, 1) = "SUMMARY": arrResult(n, 2) = "": arrResult(n, 3) = "": n = n + 1
    arrResult(n, 1) = "  Beginning Cash Balance": arrResult(n, 2) = Round(dBegCash, FIN_DECIMAL_PLACES)
    arrResult(n, 3) = "": n = n + 1
    arrResult(n, 1) = "  Net Cash from Operations": arrResult(n, 2) = ""
    arrResult(n, 3) = Round(dOperatingIn - dOperatingOut, FIN_DECIMAL_PLACES): n = n + 1
    arrResult(n, 1) = "  Ending Cash Balance": arrResult(n, 2) = ""
    arrResult(n, 3) = Round(dEndCash, FIN_DECIMAL_PLACES): n = n + 1

    ReDim Preserve arrResult(1 To n, 1 To 3)
    GenerateCashFlow = arrResult
    Exit Function

GenerateCashFlow_Err:
    LogError "CashFlowService.GenerateCashFlow", Err.Description
    GenerateCashFlow = Array()
End Function

Public Function ExportCashFlowToSheet(ByVal FiscalYear As Long, ByVal TargetSheetName As String) As Boolean
    Dim arrData As Variant, ws As Worksheet, i As Long, j As Long
    On Error GoTo ExportCashFlowToSheet_Err
    arrData = GenerateCashFlow(FiscalYear)
    If Not IsArray(arrData) Then ExportCashFlowToSheet = False: Exit Function
    If UBound(arrData, 1) = 0 Then ExportCashFlowToSheet = False: Exit Function
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
    ExportCashFlowToSheet = True
    Exit Function
ExportCashFlowToSheet_Err: LogError "CashFlowService.ExportCashFlowToSheet", Err.Description: ExportCashFlowToSheet = False
End Function
