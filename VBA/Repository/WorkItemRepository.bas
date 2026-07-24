Attribute VB_Name = "WorkItemRepository"
Option Explicit

Public Function CreateWorkItem(ByVal ProjectID As String, ByVal ItemCode As String, _
                               ByVal ItemName As String, ByVal Unit As String, _
                               ByVal Volume As Double, ByVal UnitPrice As Double, _
                               ByVal Category As String) As Long
    Dim arr(0 To 10) As Variant
    Dim dTotal As Double
    dTotal = Round(Volume * UnitPrice, FIN_DECIMAL_PLACES)
    arr(0) = GenerateGUID()
    arr(1) = ProjectID
    arr(2) = ItemCode
    arr(3) = ItemName
    arr(4) = Unit
    arr(5) = Volume
    arr(6) = UnitPrice
    arr(7) = dTotal
    arr(8) = Category
    arr(9) = gCurrentUser
    arr(10) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    CreateWorkItem = InsertRecord(SHT_WORKITEM, arr)
End Function

Public Function UpdateWorkItem(ByVal RowNum As Long, ByVal ItemCode As String, _
                               ByVal ItemName As String, ByVal Unit As String, _
                               ByVal Volume As Double, ByVal UnitPrice As Double, _
                               ByVal Category As String) As Boolean
    Dim arr(0 To 10) As Variant
    Dim vRec As Variant
    vRec = GetRecord(SHT_WORKITEM, RowNum)
    If IsEmpty(vRec(0)) Then UpdateWorkItem = False: Exit Function
    arr(0) = vRec(0)
    arr(1) = vRec(1)
    arr(2) = ItemCode
    arr(3) = ItemName
    arr(4) = Unit
    arr(5) = Volume
    arr(6) = UnitPrice
    arr(7) = Round(Volume * UnitPrice, FIN_DECIMAL_PLACES)
    arr(8) = Category
    arr(9) = vRec(9)
    arr(10) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    UpdateWorkItem = UpdateRecord(SHT_WORKITEM, RowNum, arr)
End Function

Public Function DeleteWorkItem(ByVal RowNum As Long) As Boolean
    DeleteWorkItem = DeleteRecord(SHT_WORKITEM, RowNum)
End Function

Public Function GetWorkItem(ByVal RowNum As Long) As Variant()
    GetWorkItem = GetRecord(SHT_WORKITEM, RowNum)
End Function

Public Function FindWorkItemByCode(ByVal ProjectID As String, ByVal ItemCode As String) As Long
    Dim ws As Worksheet, i As Long, lRow As Long
    Set ws = ThisWorkbook.Worksheets(SHT_WORKITEM)
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lRow
        If ws.Cells(i, COL_WI_ITEMCODE + 1).Value = ItemCode And ws.Cells(i, COL_WI_PROJECTID + 1).Value = ProjectID Then
            FindWorkItemByCode = i: Exit Function
        End If
    Next i
    FindWorkItemByCode = 0
End Function

Public Function LoadWorkItemsByProject(ByVal ProjectID As String) As Variant()
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_WORKITEM): Set lo = ws.ListObjects(TBL_WORKITEM)
    If lo.DataBodyRange Is Nothing Then LoadWorkItemsByProject = Array(): Exit Function
    Dim arrData As Variant, arrResult() As Variant, i As Long, n As Long
    arrData = lo.DataBodyRange.Value
    n = 0
    ReDim arrResult(1 To UBound(arrData, 1))
    For i = 1 To UBound(arrData, 1)
        If SafeConvertToString(arrData(i, COL_WI_PROJECTID + 1)) = ProjectID Then
            n = n + 1: arrResult(n) = i
        End If
    Next i
    If n = 0 Then LoadWorkItemsByProject = Array(): Exit Function
    ReDim Preserve arrResult(1 To n)
    LoadWorkItemsByProject = arrResult
End Function

Public Function GetTotalBudgetByProject(ByVal ProjectID As String) As Double
    Dim ws As Worksheet, lo As ListObject, i As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_WORKITEM): Set lo = ws.ListObjects(TBL_WORKITEM)
    If lo.DataBodyRange Is Nothing Then GetTotalBudgetByProject = 0: Exit Function
    dTotal = 0
    For i = 1 To lo.DataBodyRange.Rows.Count
        If SafeConvertToString(lo.DataBodyRange.Cells(i, COL_WI_PROJECTID + 1).Value) = ProjectID Then
            dTotal = dTotal + SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_WI_TOTALBUDGET + 1).Value)
        End If
    Next i
    GetTotalBudgetByProject = Round(dTotal, FIN_DECIMAL_PLACES)
End Function

Public Function LoadWorkItems() As Variant()
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_WORKITEM): Set lo = ws.ListObjects(TBL_WORKITEM)
    If lo.DataBodyRange Is Nothing Then LoadWorkItems = Array(): Exit Function
    LoadWorkItems = lo.DataBodyRange.Value
End Function

