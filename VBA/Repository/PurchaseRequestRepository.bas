Attribute VB_Name = "PurchaseRequestRepository"
Option Explicit

Private Const PR_MOD As String = "PurchaseRequestRepository"

Public Function GetNextPRNumber() As String
    Dim lCount As Long: lCount = GetRecordCount(SHT_PURCHASEREQ)
    GetNextPRNumber = "PR" & Format$(Now, "yy") & Format$(lCount + 1, "0000")
End Function

Public Function SavePurchaseRequest(ByVal lRow As Long, ByVal sNumber As String, _
    ByVal sProject As String, ByVal dDate As Date, ByVal sRequestedBy As String, _
    ByVal sStatus As String, ByVal sNotes As String) As Boolean
    Dim v(0 To 8) As Variant
    If lRow <= 0 Then
        v(0) = GenerateGUID(): v(1) = sNumber: v(2) = sProject
        v(3) = dDate: v(4) = sRequestedBy: v(5) = sStatus
        v(6) = sNotes: v(7) = gCurrentUser: v(8) = Now()
        lRow = InsertRecord(SHT_PURCHASEREQ, v)
        SavePurchaseRequest = (lRow > 0)
    Else
        v(1) = sNumber: v(2) = sProject: v(3) = dDate: v(4) = sRequestedBy
        v(5) = sStatus: v(6) = sNotes
        SavePurchaseRequest = UpdateRecord(SHT_PURCHASEREQ, lRow, v)
    End If
End Function

Public Function ApprovePurchaseRequest(ByVal lRow As Long) As Boolean
    Dim v(0 To 5) As Variant
    Dim vRec As Variant: vRec = GetRecord(SHT_PURCHASEREQ, lRow)
    If IsEmpty(vRec(0)) Then ApprovePurchaseRequest = False: Exit Function
    v(1) = SafeConvertToString(vRec(COL_PR_NUMBER + 1))
    v(2) = SafeConvertToString(vRec(COL_PR_PROJECT + 1))
    v(3) = vRec(COL_PR_REQDATE + 1)
    v(4) = SafeConvertToString(vRec(COL_PR_REQUESTEDBY + 1))
    v(5) = STATUS_APPROVED
    v(6) = SafeConvertToString(vRec(COL_PR_NOTES + 1))
    ApprovePurchaseRequest = UpdateRecord(SHT_PURCHASEREQ, lRow, v)
End Function

Public Function RejectPurchaseRequest(ByVal lRow As Long) As Boolean
    Dim vRec As Variant: vRec = GetRecord(SHT_PURCHASEREQ, lRow)
    If IsEmpty(vRec(0)) Then RejectPurchaseRequest = False: Exit Function
    Dim v(0 To 6) As Variant
    v(1) = SafeConvertToString(vRec(COL_PR_NUMBER + 1))
    v(2) = SafeConvertToString(vRec(COL_PR_PROJECT + 1))
    v(3) = vRec(COL_PR_REQDATE + 1)
    v(4) = SafeConvertToString(vRec(COL_PR_REQUESTEDBY + 1))
    v(5) = STATUS_REJECTED
    v(6) = SafeConvertToString(vRec(COL_PR_NOTES + 1))
    RejectPurchaseRequest = UpdateRecord(SHT_PURCHASEREQ, lRow, v)
End Function

Public Function GetPurchaseRequestList() As Variant()
    GetPurchaseRequestList = GetAllRecords(SHT_PURCHASEREQ)
End Function

Public Function FindPRByNumber(ByVal sNumber As String) As Long
    FindPRByNumber = FindRecord(SHT_PURCHASEREQ, COL_PR_NUMBER, sNumber)
End Function
