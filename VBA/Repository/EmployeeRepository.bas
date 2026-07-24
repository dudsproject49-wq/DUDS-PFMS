Attribute VB_Name = "EmployeeRepository"
Option Explicit

Private Const EMP_MOD As String = "EmployeeRepository"

Public Function GetNextNIK() As String
    Dim lCount As Long: lCount = GetRecordCount(SHT_EMPLOYEE)
    GetNextNIK = "EMP" & Format$(lCount + 1, "0000")
End Function

Public Function SaveEmployee(ByVal lRow As Long, ByVal sNIK As String, ByVal sName As String, _
    ByVal sAddress As String, ByVal sPhone As String, ByVal sPosition As String, _
    ByVal dDailyRate As Double, ByVal dMonthlySal As Double, ByVal sBank As String, _
    ByVal sActive As String) As Boolean
    Dim v(0 To 11) As Variant
    If lRow <= 0 Then
        v(0) = GenerateGUID(): v(1) = sNIK: v(2) = sName: v(3) = sAddress
        v(4) = sPhone: v(5) = sPosition: v(6) = dDailyRate: v(7) = dMonthlySal
        v(8) = sBank: v(9) = sActive: v(10) = gCurrentUser: v(11) = Now()
        lRow = InsertRecord(SHT_EMPLOYEE, v)
        SaveEmployee = (lRow > 0)
    Else
        v(1) = sNIK: v(2) = sName: v(3) = sAddress: v(4) = sPhone
        v(5) = sPosition: v(6) = dDailyRate: v(7) = dMonthlySal: v(8) = sBank
        v(9) = sActive
        SaveEmployee = UpdateRecord(SHT_EMPLOYEE, lRow, v)
    End If
End Function

Public Function DeleteEmployee(ByVal lRow As Long) As Boolean
    DeleteEmployee = DeleteRecord(SHT_EMPLOYEE, lRow)
End Function

Public Function GetEmployeeList() As Variant()
    GetEmployeeList = GetAllRecords(SHT_EMPLOYEE)
End Function

Public Function FindEmployeeByNIK(ByVal sNIK As String) As Long
    FindEmployeeByNIK = FindRecord(SHT_EMPLOYEE, COL_EMP_NIK, sNIK)
End Function

Public Function GetEmployeeDropdown() As Collection
    Dim col As New Collection, ws As Worksheet, lRow As Long, lLast As Long
    Set ws = ThisWorkbook.Worksheets(SHT_EMPLOYEE)
    lLast = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For lRow = 2 To lLast
        If Len(SafeConvertToString(ws.Cells(lRow, 2).Value)) > 0 Then
            col.Add SafeConvertToString(ws.Cells(lRow, 2).Value) & " - " & SafeConvertToString(ws.Cells(lRow, 3).Value)
        End If
    Next lRow
    Set GetEmployeeDropdown = col
End Function

Public Function GetEmployeeDailyRate(ByVal sNIK As String) As Double
    Dim lRow As Long: lRow = FindEmployeeByNIK(sNIK)
    If lRow <= 0 Then GetEmployeeDailyRate = 0: Exit Function
    Dim vRec As Variant: vRec = GetRecord(SHT_EMPLOYEE, lRow)
    If IsEmpty(vRec(0)) Then GetEmployeeDailyRate = 0 Else GetEmployeeDailyRate = SafeConvertToDouble(vRec(COL_EMP_DAILYRATE + 1))
End Function

Public Function GetEmployeeMonthlySalary(ByVal sNIK As String) As Double
    Dim lRow As Long: lRow = FindEmployeeByNIK(sNIK)
    If lRow <= 0 Then GetEmployeeMonthlySalary = 0: Exit Function
    Dim vRec As Variant: vRec = GetRecord(SHT_EMPLOYEE, lRow)
    If IsEmpty(vRec(0)) Then GetEmployeeMonthlySalary = 0 Else GetEmployeeMonthlySalary = SafeConvertToDouble(vRec(COL_EMP_MONTHLYSAL + 1))
End Function
