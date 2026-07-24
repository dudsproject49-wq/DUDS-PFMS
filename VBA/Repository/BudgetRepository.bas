Attribute VB_Name = "BudgetRepository"
Option Explicit

Public Function CreateBudget(ByVal ProjectID As String, ByVal BudgetNo As String, _
                             ByVal FiscalYear As Long, ByVal Description As String) As Long
    Dim arr(0 To 11) As Variant
    arr(0) = GenerateGUID()
    arr(1) = ProjectID
    arr(2) = BudgetNo
    arr(3) = FiscalYear
    arr(4) = Description
    arr(5) = 0
    arr(6) = APR_DRAFT
    arr(7) = ""
    arr(8) = ""
    arr(9) = "No"
    arr(10) = gCurrentUser
    arr(11) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    CreateBudget = InsertRecord(SHT_BUDGETHEADER, arr)
End Function

Public Function UpdateBudget(ByVal RowNum As Long, ByVal Description As String, _
                             ByVal TotalAmount As Double, ByVal Status As String) As Boolean
    Dim arr(0 To 11) As Variant
    Dim vRec As Variant
    vRec = GetRecord(SHT_BUDGETHEADER, RowNum)
    If IsEmpty(vRec(0)) Then UpdateBudget = False: Exit Function
    If vRec(COL_BH_LOCKED + 1) = "Yes" Then UpdateBudget = False: Exit Function
    arr(0) = vRec(0): arr(1) = vRec(1): arr(2) = vRec(2): arr(3) = vRec(3)
    arr(4) = Description: arr(5) = TotalAmount: arr(6) = Status
    arr(7) = vRec(7): arr(8) = vRec(8)
    arr(9) = vRec(9): arr(10) = vRec(10): arr(11) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    UpdateBudget = UpdateRecord(SHT_BUDGETHEADER, RowNum, arr)
End Function

Public Function DeleteBudget(ByVal RowNum As Long) As Boolean
    Dim vRec As Variant
    vRec = GetRecord(SHT_BUDGETHEADER, RowNum)
    If Not IsEmpty(vRec(0)) Then
        If vRec(COL_BH_LOCKED + 1) = "Yes" Then DeleteBudget = False: Exit Function
    End If
    Dim ws As Worksheet, lo As ListObject, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_BUDGETLINE): Set lo = ws.ListObjects(TBL_BUDGETLINE)
    If Not lo.DataBodyRange Is Nothing Then
        For i = lo.DataBodyRange.Rows.Count To 1 Step -1
            If SafeConvertToString(lo.DataBodyRange.Cells(i, COL_BL_HEADERID + 1).Value) = vRec(0) Then
                DeleteRecord SHT_BUDGETLINE, lo.DataBodyRange.Cells(i, 1).Row
            End If
        Next i
    End If
    On Error GoTo 0
    DeleteBudget = DeleteRecord(SHT_BUDGETHEADER, RowNum)
End Function

Public Function GetBudget(ByVal RowNum As Long) As Variant()
    GetBudget = GetRecord(SHT_BUDGETHEADER, RowNum)
End Function

Public Function CalculateBudget(ByVal HeaderID As String) As Double
    Dim ws As Worksheet, lo As ListObject, i As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_BUDGETLINE): Set lo = ws.ListObjects(TBL_BUDGETLINE)
    If lo.DataBodyRange Is Nothing Then CalculateBudget = 0: Exit Function
    dTotal = 0
    For i = 1 To lo.DataBodyRange.Rows.Count
        If SafeConvertToString(lo.DataBodyRange.Cells(i, COL_BL_HEADERID + 1).Value) = HeaderID Then
            dTotal = dTotal + SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_BL_TOTAL + 1).Value)
        End If
    Next i
    CalculateBudget = Round(dTotal, FIN_DECIMAL_PLACES)
End Function

Public Function AddBudgetLine(ByVal HeaderID As String, ByVal WorkItemID As String, _
                              ByVal Volume As Double, ByVal UnitPrice As Double, _
                              ByVal Category As String) As Long
    Dim arr(0 To 8) As Variant
    arr(0) = GenerateGUID()
    arr(1) = HeaderID
    arr(2) = WorkItemID
    arr(3) = Volume
    arr(4) = UnitPrice
    arr(5) = Round(Volume * UnitPrice, FIN_DECIMAL_PLACES)
    arr(6) = Category
    arr(7) = gCurrentUser
    arr(8) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    AddBudgetLine = InsertRecord(SHT_BUDGETLINE, arr)
End Function

Public Function UpdateBudgetLine(ByVal RowNum As Long, ByVal Volume As Double, _
                                 ByVal UnitPrice As Double, ByVal Category As String) As Boolean
    Dim arr(0 To 8) As Variant
    Dim vRec As Variant
    vRec = GetRecord(SHT_BUDGETLINE, RowNum)
    If IsEmpty(vRec(0)) Then UpdateBudgetLine = False: Exit Function
    arr(0) = vRec(0): arr(1) = vRec(1): arr(2) = vRec(2)
    arr(3) = Volume: arr(4) = UnitPrice
    arr(5) = Round(Volume * UnitPrice, FIN_DECIMAL_PLACES)
    arr(6) = Category: arr(7) = vRec(7): arr(8) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    UpdateBudgetLine = UpdateRecord(SHT_BUDGETLINE, RowNum, arr)
End Function

Public Function DeleteBudgetLine(ByVal RowNum As Long) As Boolean
    DeleteBudgetLine = DeleteRecord(SHT_BUDGETLINE, RowNum)
End Function

Public Function GetBudgetLinesByHeader(ByVal HeaderID As String) As Variant()
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_BUDGETLINE): Set lo = ws.ListObjects(TBL_BUDGETLINE)
    If lo.DataBodyRange Is Nothing Then GetBudgetLinesByHeader = Array(): Exit Function
    Dim arrData As Variant, arrResult() As Variant, i As Long, n As Long
    arrData = lo.DataBodyRange.Value: n = 0
    ReDim arrResult(1 To UBound(arrData, 1))
    For i = 1 To UBound(arrData, 1)
        If SafeConvertToString(arrData(i, COL_BL_HEADERID + 1)) = HeaderID Then
            n = n + 1: arrResult(n) = i
        End If
    Next i
    If n = 0 Then GetBudgetLinesByHeader = Array(): Exit Function
    ReDim Preserve arrResult(1 To n)
    GetBudgetLinesByHeader = arrResult
End Function

Public Function LoadBudgetsByProject(ByVal ProjectID As String) As Variant()
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_BUDGETHEADER): Set lo = ws.ListObjects(TBL_BUDGETHEADER)
    If lo.DataBodyRange Is Nothing Then LoadBudgetsByProject = Array(): Exit Function
    Dim arrData As Variant, arrResult() As Variant, i As Long, n As Long
    arrData = lo.DataBodyRange.Value: n = 0
    ReDim arrResult(1 To UBound(arrData, 1))
    For i = 1 To UBound(arrData, 1)
        If SafeConvertToString(arrData(i, COL_BH_PROJECTID + 1)) = ProjectID Then
            n = n + 1: arrResult(n) = i
        End If
    Next i
    If n = 0 Then LoadBudgetsByProject = Array(): Exit Function
    ReDim Preserve arrResult(1 To n)
    LoadBudgetsByProject = arrResult
End Function

Public Sub LockBudget(ByVal RowNum As Long)
    Dim vRec As Variant: vRec = GetRecord(SHT_BUDGETHEADER, RowNum)
    If IsEmpty(vRec(0)) Then Exit Sub
    Dim arr(0 To 11) As Variant
    arr(0) = vRec(0): arr(1) = vRec(1): arr(2) = vRec(2): arr(3) = vRec(3)
    arr(4) = vRec(4): arr(5) = vRec(5): arr(6) = vRec(6)
    arr(7) = vRec(7): arr(8) = vRec(8)
    arr(9) = "Yes": arr(10) = vRec(10): arr(11) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    UpdateRecord SHT_BUDGETHEADER, RowNum, arr
End Sub

