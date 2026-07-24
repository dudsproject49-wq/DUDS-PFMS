Attribute VB_Name = "DashboardService"
Option Explicit

Private Const DSH_MOD As String = "DashboardService"

Public Sub RefreshDashboard()
    Dim wsDashboard As Worksheet
    On Error Resume Next
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    If wsDashboard Is Nothing Then
        Set wsDashboard = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsDashboard.Name = "Dashboard"
    End If
    On Error GoTo 0
    
    wsDashboard.Cells.Clear
    wsDashboard.Cells(1, 1).Value = "DUDS-PFMS Executive Dashboard"
    wsDashboard.Range("A1:C1").Font.Bold = True
    wsDashboard.Range("A1:C1").Font.Size = 14
    
    Dim lRow As Long: lRow = 3
    
    ' Revenue
    wsDashboard.Cells(lRow, 1).Value = "Total Revenue:"
    wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(GetTotalRevenue())
    lRow = lRow + 1
    
    ' Expense
    wsDashboard.Cells(lRow, 1).Value = "Total Expense:"
    wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(GetTotalExpense())
    lRow = lRow + 1
    
    ' Profit
    wsDashboard.Cells(lRow, 1).Value = "Net Profit:"
    wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(GetNetProfit())
    If GetNetProfit() >= 0 Then wsDashboard.Cells(lRow, 2).Font.Color = RGB(0, 128, 0) Else wsDashboard.Cells(lRow, 2).Font.Color = RGB(255, 0, 0)
    lRow = lRow + 1
    
    ' Cash Position
    wsDashboard.Cells(lRow, 1).Value = "Cash Position:"
    wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(GetCashPosition())
    lRow = lRow + 2
    
    ' Budget vs Actual
    wsDashboard.Cells(lRow, 1).Value = "Budget vs Actual:"
    Dim dBudget As Double, dActual As Double
    dBudget = GetTotalBudget()
    dActual = GetTotalActual()
    wsDashboard.Cells(lRow, 2).Value = "Budget: " & FormatCurrencyIDR(dBudget)
    wsDashboard.Cells(lRow + 1, 2).Value = "Actual: " & FormatCurrencyIDR(dActual)
    If dBudget > 0 Then wsDashboard.Cells(lRow + 2, 2).Value = "Utilization: " & Format$(dActual / dBudget * 100, "0.0") & "%"
    lRow = lRow + 4
    
    ' Project Progress
    wsDashboard.Cells(lRow, 1).Value = "Project Progress:"
    wsDashboard.Cells(lRow, 2).Value = Format$(GetOverallProjectProgress(), "0.0") & "%"
    lRow = lRow + 2
    
    ' Cost Breakdown
    wsDashboard.Cells(lRow, 1).Value = "Material Cost:"
    wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(GetMaterialCost())
    lRow = lRow + 1
    
    wsDashboard.Cells(lRow, 1).Value = "Labor Cost:"
    wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(GetLaborCost())
    lRow = lRow + 1
    
    wsDashboard.Cells(lRow, 1).Value = "Equipment Cost:"
    wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(GetEquipmentCost())
    lRow = lRow + 2
    
    ' Top 10 Expenses
    wsDashboard.Cells(lRow, 1).Value = "Top 10 Expenses:"
    wsDashboard.Cells(lRow, 1).Font.Bold = True
    lRow = lRow + 1
    Dim arrExp As Variant: arrExp = GetTopExpenses(10)
    If IsArray(arrExp) Then
        Dim i As Long
        For i = LBound(arrExp) To UBound(arrExp)
            If Not IsEmpty(arrExp(i)) Then
                wsDashboard.Cells(lRow, 1).Value = arrExp(i)(0)
                wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(CDbl(arrExp(i)(1)))
                lRow = lRow + 1
            End If
        Next i
    End If
    lRow = lRow + 1
    
    ' Monthly Cash Flow
    wsDashboard.Cells(lRow, 1).Value = "Monthly Cash Flow:"
    wsDashboard.Cells(lRow, 1).Font.Bold = True
    lRow = lRow + 1
    Dim dCashIn As Double, dCashOut As Double
    dCashIn = GetMonthlyCashIn(Month(Date), Year(Date))
    dCashOut = GetMonthlyCashOut(Month(Date), Year(Date))
    wsDashboard.Cells(lRow, 1).Value = "Cash In:": wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(dCashIn)
    lRow = lRow + 1
    wsDashboard.Cells(lRow, 1).Value = "Cash Out:": wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(dCashOut)
    lRow = lRow + 1
    wsDashboard.Cells(lRow, 1).Value = "Net Cash Flow:": wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(dCashIn - dCashOut)
    lRow = lRow + 2
    
    ' Outstanding
    wsDashboard.Cells(lRow, 1).Value = "Outstanding Receivable:"
    wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(GetOutstandingReceivable())
    lRow = lRow + 1
    
    wsDashboard.Cells(lRow, 1).Value = "Outstanding Payable:"
    wsDashboard.Cells(lRow, 2).Value = FormatCurrencyIDR(GetOutstandingPayable())
    lRow = lRow + 2
    
    ' Project Status
    wsDashboard.Cells(lRow, 1).Value = "Active Projects:"
    wsDashboard.Cells(lRow, 2).Value = GetActiveProjectCount()
    lRow = lRow + 1
    
    wsDashboard.Cells(lRow, 1).Value = "Completed Projects:"
    wsDashboard.Cells(lRow, 2).Value = GetCompletedProjectCount()
    lRow = lRow + 1
    
    wsDashboard.Columns("A:B").AutoFit
    wsDashboard.Range("A:A").Font.Bold = True
    LogInfo DSH_MOD, "Dashboard refreshed", "Dashboard"
End Sub

Private Function GetTotalRevenue() As Double
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_CASHIN)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast: dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_CASHIN_AMOUNT + 1).Value): Next lRow
    On Error GoTo 0
    GetTotalRevenue = dTotal
End Function

Private Function GetTotalExpense() As Double
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_CASHOUT)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast: dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_CASHOUT_AMOUNT + 1).Value): Next lRow
    On Error GoTo 0
    GetTotalExpense = dTotal
End Function

Private Function GetNetProfit() As Double
    GetNetProfit = GetTotalRevenue() - GetTotalExpense()
End Function

Public Function GetCashPosition() As Double
    GetCashPosition = GetTotalRevenue() - GetTotalExpense()
End Function

Private Function GetTotalBudget() As Double
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_BUDGET)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast: dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_BUDGET_PLANNED + 1).Value): Next lRow
    On Error GoTo 0
    GetTotalBudget = dTotal
End Function

Private Function GetTotalActual() As Double
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_BUDGET)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast: dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_BUDGET_ACTUAL + 1).Value): Next lRow
    On Error GoTo 0
    GetTotalActual = dTotal
End Function

Private Function GetOverallProjectProgress() As Double
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double, lCount As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_PROJ_PROGRESS + 1).Value)
        lCount = lCount + 1
    Next lRow
    On Error GoTo 0
    If lCount = 0 Then GetOverallProjectProgress = 0 Else GetOverallProjectProgress = dTotal / lCount
End Function

Private Function GetMaterialCost() As Double
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_MATERIAL)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast: dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_MAT_TOTALCOST + 1).Value): Next lRow
    On Error GoTo 0
    GetMaterialCost = dTotal
End Function

Private Function GetLaborCost() As Double
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_PAYROLL)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast: dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_PAY_NETPAY + 1).Value): Next lRow
    On Error GoTo 0
    GetLaborCost = dTotal
End Function

Private Function GetEquipmentCost() As Double
    ' Equipment cost tracked via transactions with "Equipment" category
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_TRANSACTIONS)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        If InStr(1, SafeConvertToString(ws.Cells(lRow, COL_TXN_DESCRIPTION + 1).Value), "Equipment", vbTextCompare) > 0 Then
            dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_TXN_AMOUNT + 1).Value)
        End If
    Next lRow
    On Error GoTo 0
    GetEquipmentCost = dTotal
End Function

Private Function GetTopExpenses(ByVal lCount As Long) As Variant()
    Dim ws As Worksheet, lLast As Long, lRow As Long, n As Long
    Dim arrTemp() As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_CASHOUT)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    n = 0: ReDim arrTemp(1 To lLast, 0 To 1)
    For lRow = 2 To lLast
        If SafeConvertToDouble(ws.Cells(lRow, COL_CASHOUT_AMOUNT + 1).Value) > 0 Then
            n = n + 1
            arrTemp(n, 0) = SafeConvertToString(ws.Cells(lRow, COL_CASHOUT_DESC + 1).Value)
            arrTemp(n, 1) = SafeConvertToDouble(ws.Cells(lRow, COL_CASHOUT_AMOUNT + 1).Value)
        End If
    Next lRow
    On Error GoTo 0
    If n = 0 Then GetTopExpenses = Array(): Exit Function
    ' Simple bubble sort descending
    Dim i As Long, j As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If CDbl(arrTemp(j, 1)) > CDbl(arrTemp(i, 1)) Then
                Dim sTemp As String, dTemp As Double
                sTemp = arrTemp(i, 0): dTemp = CDbl(arrTemp(i, 1))
                arrTemp(i, 0) = arrTemp(j, 0): arrTemp(i, 1) = CDbl(arrTemp(j, 1))
                arrTemp(j, 0) = sTemp: arrTemp(j, 1) = dTemp
            End If
        Next j
    Next i
    Dim lTop As Long: lTop = IIf(n < lCount, n, lCount)
    Dim arrResult() As Variant
    ReDim arrResult(1 To lTop, 0 To 1)
    For i = 1 To lTop
        arrResult(i, 0) = arrTemp(i, 0): arrResult(i, 1) = CDbl(arrTemp(i, 1))
    Next i
    GetTopExpenses = arrResult
End Function

Private Function GetMonthlyCashIn(ByVal lMonth As Long, ByVal lYear As Long) As Double
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_CASHIN)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        Dim dDate As Date: dDate = CDate(ws.Cells(lRow, COL_CASHIN_DATE + 1).Value)
        If Month(dDate) = lMonth And Year(dDate) = lYear Then
            dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_CASHIN_AMOUNT + 1).Value)
        End If
    Next lRow
    On Error GoTo 0
    GetMonthlyCashIn = dTotal
End Function

Private Function GetMonthlyCashOut(ByVal lMonth As Long, ByVal lYear As Long) As Double
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_CASHOUT)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        Dim dDate As Date: dDate = CDate(ws.Cells(lRow, COL_CASHOUT_DATE + 1).Value)
        If Month(dDate) = lMonth And Year(dDate) = lYear Then
            dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_CASHOUT_AMOUNT + 1).Value)
        End If
    Next lRow
    On Error GoTo 0
    GetMonthlyCashOut = dTotal
End Function

Private Function GetOutstandingReceivable() As Double
    ' Receivable tracked via CashIn with status "Pending"
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_CASHIN)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        If SafeConvertToString(ws.Cells(lRow, COL_CASHIN_CREATED_BY + 1).Value) = "PENDING" Then
            dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_CASHIN_AMOUNT + 1).Value)
        End If
    Next lRow
    On Error GoTo 0
    GetOutstandingReceivable = dTotal
End Function

Private Function GetOutstandingPayable() As Double
    ' Payable tracked via CashOut with status "Pending"
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_CASHOUT)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        If SafeConvertToString(ws.Cells(lRow, COL_CASHOUT_CREATED_BY + 1).Value) = "PENDING" Then
            dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_CASHOUT_AMOUNT + 1).Value)
        End If
    Next lRow
    On Error GoTo 0
    GetOutstandingPayable = dTotal
End Function

Private Function GetActiveProjectCount() As Long
    Dim ws As Worksheet, lLast As Long, lRow As Long, lCount As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        If SafeConvertToString(ws.Cells(lRow, COL_PROJ_STATUS + 1).Value) = STATUS_ACTIVE Then
            lCount = lCount + 1
        End If
    Next lRow
    On Error GoTo 0
    GetActiveProjectCount = lCount
End Function

Private Function GetCompletedProjectCount() As Long
    Dim ws As Worksheet, lLast As Long, lRow As Long, lCount As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        If SafeConvertToString(ws.Cells(lRow, COL_PROJ_STATUS + 1).Value) = STATUS_COMPLETED Then
            lCount = lCount + 1
        End If
    Next lRow
    On Error GoTo 0
    GetCompletedProjectCount = lCount
End Function
