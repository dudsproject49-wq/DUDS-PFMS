Attribute VB_Name = "PayrollRepository"
Option Explicit

Private Const PAY_MOD As String = "PayrollRepository"

Public Function CalculatePayroll(ByVal sEmpID As String, ByVal sPeriod As String, _
    ByVal dBasicSal As Double, ByVal dOvertime As Double, ByVal dLoan As Double, _
    ByVal dDeduction As Double) As Double
    CalculatePayroll = dBasicSal + dOvertime - dLoan - dDeduction
End Function

Public Function SavePayroll(ByVal lRow As Long, ByVal sEmpID As String, ByVal sPeriod As String, _
    ByVal dBasicSal As Double, ByVal dOvertime As Double, ByVal dLoan As Double, _
    ByVal dDeduction As Double, ByVal dNetPay As Double, ByVal sStatus As String, _
    ByVal sPosted As String) As Boolean
    Dim v(0 To 11) As Variant
    If lRow <= 0 Then
        v(0) = GenerateGUID(): v(1) = sEmpID: v(2) = sPeriod: v(3) = dBasicSal
        v(4) = dOvertime: v(5) = dLoan: v(6) = dDeduction: v(7) = dNetPay
        v(8) = sStatus: v(9) = sPosted: v(10) = gCurrentUser: v(11) = Now()
        lRow = InsertRecord(SHT_PAYROLL, v)
        SavePayroll = (lRow > 0)
    Else
        v(1) = sEmpID: v(2) = sPeriod: v(3) = dBasicSal: v(4) = dOvertime
        v(5) = dLoan: v(6) = dDeduction: v(7) = dNetPay: v(8) = sStatus: v(9) = sPosted
        SavePayroll = UpdateRecord(SHT_PAYROLL, lRow, v)
    End If
End Function

Public Function PostPayrollJournal(ByVal sPeriod As String) As Boolean
    Dim ws As Worksheet, lLast As Long, lRow As Long
    Dim dTotalNetPay As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_PAYROLL)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        If SafeConvertToString(ws.Cells(lRow, COL_PAY_PERIOD + 1).Value) = sPeriod Then
            If SafeConvertToString(ws.Cells(lRow, COL_PAY_POSTED + 1).Value) <> "Yes" Then
                dTotalNetPay = dTotalNetPay + SafeConvertToDouble(ws.Cells(lRow, COL_PAY_NETPAY + 1).Value)
                ws.Cells(lRow, COL_PAY_POSTED + 1).Value = "Yes"
            End If
        End If
    Next lRow
    
    If dTotalNetPay > 0 Then
        Dim v1(0 To 10) As Variant
        v1(0) = GenerateGUID(): v1(1) = Now(): v1(2) = "Payroll"
        v1(3) = "PR-" & sPeriod: v1(4) = "": v1(5) = "5-2000"
        v1(6) = "Payroll expense - " & sPeriod: v1(7) = dTotalNetPay: v1(8) = 0
        v1(9) = gCurrentUser: v1(10) = Now()
        InsertRecord SHT_JOURNAL, v1
        
        Dim v2(0 To 10) As Variant
        v2(0) = GenerateGUID(): v2(1) = Now(): v2(2) = "Payroll"
        v2(3) = "PR-" & sPeriod: v2(4) = "": v2(5) = "2-1000"
        v2(6) = "Salary payable - " & sPeriod: v2(7) = 0: v2(8) = dTotalNetPay
        v2(9) = gCurrentUser: v2(10) = Now()
        InsertRecord SHT_JOURNAL, v2
        
        LogInfo PAY_MOD, "Payroll posted for period " & sPeriod & " Total: " & FormatCurrencyIDR(dTotalNetPay)
    End If
    
    PostPayrollJournal = True
    On Error GoTo 0
End Function

Public Function GetPayrollByPeriod(ByVal sPeriod As String) As Variant()
    Dim col As Collection: Set col = GetRecordsByColumn(SHT_PAYROLL, COL_PAY_PERIOD, sPeriod)
    If col Is Nothing Then GetPayrollByPeriod = Array(): Exit Function
    If col.Count = 0 Then GetPayrollByPeriod = Array(): Exit Function
    Dim vRec As Variant, i As Long, n As Long: n = 0
    ReDim arrResult(1 To col.Count)
    For i = 1 To col.Count
        vRec = GetRecord(SHT_PAYROLL, col(i))
        If Not IsEmpty(vRec(0)) Then n = n + 1: arrResult(n) = vRec
    Next i
    If n = 0 Then GetPayrollByPeriod = Array() Else GetPayrollByPeriod = arrResult
End Function

Public Function ProcessDailyWage(ByVal sEmpID As String, ByVal sPeriod As String, _
    ByVal dDays As Double, ByVal dDailyRate As Double) As Double
    ProcessDailyWage = dDays * dDailyRate
End Function

Public Function ProcessMonthlySalary(ByVal sEmpID As String, ByVal sPeriod As String, _
    ByVal dMonthlySal As Double, ByVal dOvertime As Double, _
    ByVal dLoan As Double, ByVal dDeduction As Double) As Double
    ProcessMonthlySalary = CalculatePayroll(sEmpID, sPeriod, dMonthlySal, dOvertime, dLoan, dDeduction)
End Function
