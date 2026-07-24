Attribute VB_Name = "AttendanceRepository"
Option Explicit

Private Const ATT_MOD As String = "AttendanceRepository"

Public Function SaveAttendance(ByVal lRow As Long, ByVal sEmpID As String, ByVal dDate As Date, _
    ByVal dCheckIn As Date, ByVal dCheckOut As Date, ByVal dOvertime As Double, _
    ByVal sStatus As String, ByVal sNotes As String) As Boolean
    Dim v(0 To 9) As Variant
    If lRow <= 0 Then
        v(0) = GenerateGUID(): v(1) = sEmpID: v(2) = dDate: v(3) = dCheckIn
        v(4) = dCheckOut: v(5) = dOvertime: v(6) = sStatus: v(7) = sNotes
        v(8) = gCurrentUser: v(9) = Now()
        lRow = InsertRecord(SHT_ATTENDANCE, v)
        SaveAttendance = (lRow > 0)
    Else
        v(1) = sEmpID: v(2) = dDate: v(3) = dCheckIn: v(4) = dCheckOut
        v(5) = dOvertime: v(6) = sStatus: v(7) = sNotes
        SaveAttendance = UpdateRecord(SHT_ATTENDANCE, lRow, v)
    End If
End Function

Public Function GetAttendanceByPeriod(ByVal sEmpID As String, ByVal dStart As Date, ByVal dEnd As Date) As Variant()
    Dim ws As Worksheet, lLast As Long, lRow As Long, n As Long
    Dim arrResult() As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_ATTENDANCE)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    n = 0: ReDim arrResult(1 To lLast)
    For lRow = 2 To lLast
        If SafeConvertToString(ws.Cells(lRow, COL_ATT_EMPLOYEEID + 1).Value) = sEmpID Then
            Dim dAttDate As Date: dAttDate = CDate(ws.Cells(lRow, COL_ATT_DATE + 1).Value)
            If dAttDate >= dStart And dAttDate <= dEnd Then
                n = n + 1: arrResult(n) = GetRecord(SHT_ATTENDANCE, lRow)
            End If
        End If
    Next lRow
    If n = 0 Then GetAttendanceByPeriod = Array() Else GetAttendanceByPeriod = arrResult
    On Error GoTo 0
End Function

Public Function GetAttendanceSummary(ByVal sEmpID As String, ByVal dStart As Date, ByVal dEnd As Date) As Double
    Dim arrAtt As Variant: arrAtt = GetAttendanceByPeriod(sEmpID, dStart, dEnd)
    If Not IsArray(arrAtt) Then GetAttendanceSummary = 0: Exit Function
    If UBound(arrAtt) < LBound(arrAtt) Then GetAttendanceSummary = 0: Exit Function
    Dim dTotalDays As Double, i As Long
    For i = LBound(arrAtt) To UBound(arrAtt)
        If Not IsEmpty(arrAtt(i)) Then
            dTotalDays = dTotalDays + 1
        End If
    Next i
    GetAttendanceSummary = dTotalDays
End Function

Public Function GetAttendanceOvertime(ByVal sEmpID As String, ByVal dStart As Date, ByVal dEnd As Date) As Double
    Dim arrAtt As Variant: arrAtt = GetAttendanceByPeriod(sEmpID, dStart, dEnd)
    If Not IsArray(arrAtt) Then GetAttendanceOvertime = 0: Exit Function
    If UBound(arrAtt) < LBound(arrAtt) Then GetAttendanceOvertime = 0: Exit Function
    Dim dTotalOT As Double, i As Long
    For i = LBound(arrAtt) To UBound(arrAtt)
        If Not IsEmpty(arrAtt(i)) Then
            dTotalOT = dTotalOT + SafeConvertToDouble(arrAtt(i)(COL_ATT_OVERTIME + 1))
        End If
    Next i
    GetAttendanceOvertime = dTotalOT
End Function
