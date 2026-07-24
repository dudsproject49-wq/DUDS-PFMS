Attribute VB_Name = "ProgressRepository"
Option Explicit

Public Function InputProgress(ByVal ProjectID As String, ByVal WorkItemID As String, _
                              ByVal ProgressDate As Date, ByVal ActualVolume As Double, _
                              ByVal BudgetVolume As Double, ByVal Notes As String) As Long
    Dim arr(0 To 10) As Variant
    Dim dPhysical As Double, dFinancial As Double
    If BudgetVolume > 0 Then
        dPhysical = Round((ActualVolume / BudgetVolume) * 100, 2)
    Else
        dPhysical = 0
    End If
    dFinancial = dPhysical
    arr(0) = GenerateGUID()
    arr(1) = ProjectID
    arr(2) = WorkItemID
    arr(3) = Format$(ProgressDate, "yyyy-mm-dd")
    arr(4) = ActualVolume
    arr(5) = BudgetVolume
    arr(6) = dPhysical
    arr(7) = dFinancial
    arr(8) = Notes
    arr(9) = gCurrentUser
    arr(10) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    InputProgress = InsertRecord(SHT_PROGRESS, arr)
End Function

Public Function UpdateProgress(ByVal RowNum As Long, ByVal ActualVolume As Double, _
                               ByVal BudgetVolume As Double, ByVal Notes As String) As Boolean
    Dim arr(0 To 10) As Variant
    Dim vRec As Variant
    vRec = GetRecord(SHT_PROGRESS, RowNum)
    If IsEmpty(vRec(0)) Then UpdateProgress = False: Exit Function
    Dim dPhysical As Double, dFinancial As Double
    If BudgetVolume > 0 Then
        dPhysical = Round((ActualVolume / BudgetVolume) * 100, 2)
    Else
        dPhysical = 0
    End If
    dFinancial = dPhysical
    arr(0) = vRec(0): arr(1) = vRec(1): arr(2) = vRec(2): arr(3) = vRec(3)
    arr(4) = ActualVolume: arr(5) = BudgetVolume
    arr(6) = dPhysical: arr(7) = dFinancial: arr(8) = Notes
    arr(9) = vRec(9): arr(10) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    UpdateProgress = UpdateRecord(SHT_PROGRESS, RowNum, arr)
End Function

Public Function DeleteProgress(ByVal RowNum As Long) As Boolean
    DeleteProgress = DeleteRecord(SHT_PROGRESS, RowNum)
End Function

Public Function CalculatePhysicalProgress(ByVal ProjectID As String, Optional ByVal WorkItemID As String = "") As Double
    Dim ws As Worksheet, lo As ListObject, i As Long
    Dim dActualTotal As Double, dBudgetTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_PROGRESS): Set lo = ws.ListObjects(TBL_PROGRESS)
    If lo.DataBodyRange Is Nothing Then CalculatePhysicalProgress = 0: Exit Function
    dActualTotal = 0: dBudgetTotal = 0
    For i = 1 To lo.DataBodyRange.Rows.Count
        If SafeConvertToString(lo.DataBodyRange.Cells(i, COL_PRG_PROJECTID + 1).Value) = ProjectID Then
            If Len(WorkItemID) = 0 Or SafeConvertToString(lo.DataBodyRange.Cells(i, COL_PRG_WORKITEMID + 1).Value) = WorkItemID Then
                dActualTotal = dActualTotal + SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_PRG_ACTUALVOL + 1).Value)
                dBudgetTotal = dBudgetTotal + SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_PRG_BUDGETVOL + 1).Value)
            End If
        End If
    Next i
    If dBudgetTotal > 0 Then
        CalculatePhysicalProgress = Round((dActualTotal / dBudgetTotal) * 100, 2)
    Else
        CalculatePhysicalProgress = 0
    End If
End Function

Public Function CalculateFinancialProgress(ByVal ProjectID As String, Optional ByVal WorkItemID As String = "") As Double
    Dim ws As Worksheet, lo As ListObject, i As Long
    Dim dActualCost As Double, dTotalBudget As Double
    Dim wsWI As Worksheet, loWI As ListObject, j As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_PROGRESS): Set lo = ws.ListObjects(TBL_PROGRESS)
    Set wsWI = ThisWorkbook.Worksheets(SHT_WORKITEM): Set loWI = wsWI.ListObjects(TBL_WORKITEM)
    If lo.DataBodyRange Is Nothing Or loWI.DataBodyRange Is Nothing Then CalculateFinancialProgress = 0: Exit Function
    dActualCost = 0: dTotalBudget = 0
    For i = 1 To lo.DataBodyRange.Rows.Count
        If SafeConvertToString(lo.DataBodyRange.Cells(i, COL_PRG_PROJECTID + 1).Value) = ProjectID Then
            Dim sWIID As String
            sWIID = SafeConvertToString(lo.DataBodyRange.Cells(i, COL_PRG_WORKITEMID + 1).Value)
            If Len(WorkItemID) = 0 Or sWIID = WorkItemID Then
                Dim dVolRatio As Double
                Dim dBudgetVol As Double, dActualVol As Double
                dBudgetVol = SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_PRG_BUDGETVOL + 1).Value)
                dActualVol = SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_PRG_ACTUALVOL + 1).Value)
                If dBudgetVol > 0 Then
                    dVolRatio = dActualVol / dBudgetVol
                Else
                    dVolRatio = 0
                End If
                For j = 1 To loWI.DataBodyRange.Rows.Count
                    If SafeConvertToString(loWI.DataBodyRange.Cells(j, COL_WI_ID + 1).Value) = sWIID Then
                        dActualCost = dActualCost + (dVolRatio * SafeConvertToDouble(loWI.DataBodyRange.Cells(j, COL_WI_TOTALBUDGET + 1).Value))
                        dTotalBudget = dTotalBudget + SafeConvertToDouble(loWI.DataBodyRange.Cells(j, COL_WI_TOTALBUDGET + 1).Value)
                    End If
                Next j
            End If
        End If
    Next i
    If dTotalBudget > 0 Then
        CalculateFinancialProgress = Round((dActualCost / dTotalBudget) * 100, 2)
    Else
        CalculateFinancialProgress = 0
    End If
End Function

Public Function GetProgressByWorkItem(ByVal ProjectID As String, ByVal WorkItemID As String) As Variant()
    Dim ws As Worksheet, lo As ListObject, i As Long, n As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_PROGRESS): Set lo = ws.ListObjects(TBL_PROGRESS)
    If lo.DataBodyRange Is Nothing Then GetProgressByWorkItem = Array(): Exit Function
    Dim arrData As Variant, arrResult() As Variant
    arrData = lo.DataBodyRange.Value: n = 0
    ReDim arrResult(1 To UBound(arrData, 1))
    For i = 1 To UBound(arrData, 1)
        If SafeConvertToString(arrData(i, COL_PRG_PROJECTID + 1)) = ProjectID And _
           SafeConvertToString(arrData(i, COL_PRG_WORKITEMID + 1)) = WorkItemID Then
            n = n + 1: arrResult(n) = i
        End If
    Next i
    If n = 0 Then GetProgressByWorkItem = Array(): Exit Function
    ReDim Preserve arrResult(1 To n)
    GetProgressByWorkItem = arrResult
End Function

Public Function LoadProgressByProject(ByVal ProjectID As String) As Variant()
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_PROGRESS): Set lo = ws.ListObjects(TBL_PROGRESS)
    If lo.DataBodyRange Is Nothing Then LoadProgressByProject = Array(): Exit Function
    Dim arrData As Variant, arrResult() As Variant, i As Long, n As Long
    arrData = lo.DataBodyRange.Value: n = 0
    ReDim arrResult(1 To UBound(arrData, 1))
    For i = 1 To UBound(arrData, 1)
        If SafeConvertToString(arrData(i, COL_PRG_PROJECTID + 1)) = ProjectID Then
            n = n + 1: arrResult(n) = i
        End If
    Next i
    If n = 0 Then LoadProgressByProject = Array(): Exit Function
    ReDim Preserve arrResult(1 To n)
    LoadProgressByProject = arrResult
End Function

