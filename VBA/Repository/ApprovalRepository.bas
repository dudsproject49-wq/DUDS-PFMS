Attribute VB_Name = "ApprovalRepository"
Option Explicit

Public Function SubmitForApproval(ByVal RecordType As String, ByVal RecordID As String) As Boolean
    Dim arrLevels As Variant
    Dim i As Long, lResult As Long
    arrLevels = Array(APR_LEVEL_PM, APR_LEVEL_ACCTG, APR_LEVEL_DIRECTOR, APR_LEVEL_OWNER)
    For i = LBound(arrLevels) To UBound(arrLevels)
        Dim arr(0 To 9) As Variant
        arr(0) = GenerateGUID(): arr(1) = RecordType: arr(2) = RecordID
        arr(3) = arrLevels(i): arr(4) = "": arr(5) = "Pending"
        arr(6) = "": arr(7) = "": arr(8) = gCurrentUser
        arr(9) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
        lResult = InsertRecord(SHT_APPROVAL, arr)
        If lResult = 0 Then SubmitForApproval = False: Exit Function
    Next i
    SubmitForApproval = True
End Function

Public Function ApproveRecord(ByVal ApprovalRow As Long, ByVal Comments As String) As Boolean
    Dim vRec As Variant: vRec = GetRecord(SHT_APPROVAL, ApprovalRow)
    If IsEmpty(vRec(0)) Then ApproveRecord = False: Exit Function
    Dim arr(0 To 9) As Variant
    arr(0) = vRec(0): arr(1) = vRec(1): arr(2) = vRec(2): arr(3) = vRec(3)
    arr(4) = gCurrentUser: arr(5) = APR_APPROVED: arr(6) = Comments
    arr(7) = Format$(Now, "yyyy-mm-dd hh:mm:ss"): arr(8) = vRec(8)
    arr(9) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    ApproveRecord = UpdateRecord(SHT_APPROVAL, ApprovalRow, arr)
End Function

Public Function RejectRecord(ByVal ApprovalRow As Long, ByVal Comments As String) As Boolean
    Dim vRec As Variant: vRec = GetRecord(SHT_APPROVAL, ApprovalRow)
    If IsEmpty(vRec(0)) Then RejectRecord = False: Exit Function
    Dim arr(0 To 9) As Variant
    arr(0) = vRec(0): arr(1) = vRec(1): arr(2) = vRec(2): arr(3) = vRec(3)
    arr(4) = gCurrentUser: arr(5) = APR_REJECTED: arr(6) = Comments
    arr(7) = Format$(Now, "yyyy-mm-dd hh:mm:ss"): arr(8) = vRec(8)
    arr(9) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    RejectRecord = UpdateRecord(SHT_APPROVAL, ApprovalRow, arr)
End Function

Public Function IsFullyApproved(ByVal RecordType As String, ByVal RecordID As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, i As Long, lApproved As Long, lTotal As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_APPROVAL): Set lo = ws.ListObjects(TBL_APPROVAL)
    If lo.DataBodyRange Is Nothing Then IsFullyApproved = False: Exit Function
    lApproved = 0: lTotal = 0
    For i = 1 To lo.DataBodyRange.Rows.Count
        If SafeConvertToString(lo.DataBodyRange.Cells(i, COL_APR_RECORDTYPE + 1).Value) = RecordType And _
           SafeConvertToString(lo.DataBodyRange.Cells(i, COL_APR_RECORDID + 1).Value) = RecordID Then
            lTotal = lTotal + 1
            If SafeConvertToString(lo.DataBodyRange.Cells(i, COL_APR_STATUS + 1).Value) = APR_APPROVED Then
                lApproved = lApproved + 1
            End If
        End If
    Next i
    IsFullyApproved = (lTotal > 0 And lApproved = lTotal)
End Function

Public Function IsRejected(ByVal RecordType As String, ByVal RecordID As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_APPROVAL): Set lo = ws.ListObjects(TBL_APPROVAL)
    If lo.DataBodyRange Is Nothing Then IsRejected = False: Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If SafeConvertToString(lo.DataBodyRange.Cells(i, COL_APR_RECORDTYPE + 1).Value) = RecordType And _
           SafeConvertToString(lo.DataBodyRange.Cells(i, COL_APR_RECORDID + 1).Value) = RecordID And _
           SafeConvertToString(lo.DataBodyRange.Cells(i, COL_APR_STATUS + 1).Value) = APR_REJECTED Then
            IsRejected = True: Exit Function
        End If
    Next i
    IsRejected = False
End Function

Public Function GetPendingApprovals(ByVal ApproverRole As String) As Variant()
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_APPROVAL): Set lo = ws.ListObjects(TBL_APPROVAL)
    If lo.DataBodyRange Is Nothing Then GetPendingApprovals = Array(): Exit Function
    Dim arrData As Variant, arrResult() As Variant, i As Long, n As Long
    arrData = lo.DataBodyRange.Value: n = 0
    ReDim arrResult(1 To UBound(arrData, 1))
    For i = 1 To UBound(arrData, 1)
        If SafeConvertToString(arrData(i, COL_APR_STATUS + 1)) = "Pending" Then
            Dim sLevel As String: sLevel = SafeConvertToString(arrData(i, COL_APR_LEVEL + 1))
            If sLevel = ApproverRole Or ApproverRole = ROLE_ADMIN Then
                n = n + 1: arrResult(n) = lo.DataBodyRange.Cells(i, 1).Row
            End If
        End If
    Next i
    If n = 0 Then GetPendingApprovals = Array(): Exit Function
    ReDim Preserve arrResult(1 To n)
    GetPendingApprovals = arrResult
End Function

Public Sub LockApprovedRecords(ByVal RecordType As String, ByVal RecordID As String)
    Select Case RecordType
        Case "Budget":
            Dim lRow As Long: lRow = FindRecord(SHT_BUDGETHEADER, COL_BH_ID, RecordID)
            If lRow > 0 Then LockBudget lRow
        Case "Project":
            lRow = FindRecord(SHT_PROJECTS, COL_PROJ_ID, RecordID)
            If lRow > 0 Then
                Dim arr(0 To 13) As Variant: Dim vRec As Variant
                vRec = GetRecord(SHT_PROJECTS, lRow)
                If Not IsEmpty(vRec(0)) Then
                    arr(0) = vRec(0): arr(1) = vRec(1): arr(2) = vRec(2)
                    arr(3) = vRec(3): arr(4) = vRec(4): arr(5) = vRec(5)
                    arr(6) = vRec(6): arr(7) = vRec(7): arr(8) = vRec(8)
                    arr(9) = STATUS_APPROVED: arr(10) = vRec(10): arr(11) = vRec(11)
                    arr(12) = vRec(12): arr(13) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
                    UpdateRecord SHT_PROJECTS, lRow, arr
                End If
            End If
    End Select
End Sub

Public Sub UpdateRecordStatus(ByVal RecordType As String, ByVal RecordID As String, ByVal NewStatus As String)
    Select Case RecordType
        Case "Budget":
            Dim lRow As Long: lRow = FindRecord(SHT_BUDGETHEADER, COL_BH_ID, RecordID)
            If lRow > 0 Then
                Dim vRec As Variant: vRec = GetRecord(SHT_BUDGETHEADER, lRow)
                If Not IsEmpty(vRec(0)) Then
                    Dim arr(0 To 11) As Variant
                    arr(0) = vRec(0): arr(1) = vRec(1): arr(2) = vRec(2): arr(3) = vRec(3)
                    arr(4) = vRec(4): arr(5) = vRec(5): arr(6) = NewStatus
                    arr(7) = vRec(7): arr(8) = vRec(8): arr(9) = vRec(9)
                    arr(10) = vRec(10): arr(11) = Format$(Now, "yyyy-mm-dd hh:mm:ss")
                    UpdateRecord SHT_BUDGETHEADER, lRow, arr
                End If
            End If
    End Select
End Sub

