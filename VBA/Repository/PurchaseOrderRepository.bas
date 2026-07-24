Attribute VB_Name = "PurchaseOrderRepository"
Option Explicit

Private Const PO_MOD As String = "PurchaseOrderRepository"

Public Function GetNextPONumber() As String
    Dim lCount As Long: lCount = GetRecordCount(SHT_PURCHASEORDER)
    GetNextPONumber = "PO" & Format$(Now, "yy") & Format$(lCount + 1, "0000")
End Function

Public Function SavePurchaseOrder(ByVal lRow As Long, ByVal sNumber As String, _
    ByVal sVendor As String, ByVal sProject As String, ByVal dDelivery As Date, _
    ByVal sPaymentTerms As String, ByVal dTotal As Double, ByVal sStatus As String, _
    ByVal sApprovedBy As String) As Boolean
    Dim v(0 To 10) As Variant
    If lRow <= 0 Then
        v(0) = GenerateGUID(): v(1) = sNumber: v(2) = sVendor
        v(3) = sProject: v(4) = dDelivery: v(5) = sPaymentTerms
        v(6) = dTotal: v(7) = sStatus: v(8) = sApprovedBy
        v(9) = gCurrentUser: v(10) = Now()
        lRow = InsertRecord(SHT_PURCHASEORDER, v)
        SavePurchaseOrder = (lRow > 0)
    Else
        v(1) = sNumber: v(2) = sVendor: v(3) = sProject
        v(4) = dDelivery: v(5) = sPaymentTerms: v(6) = dTotal
        v(7) = sStatus: v(8) = sApprovedBy
        SavePurchaseOrder = UpdateRecord(SHT_PURCHASEORDER, lRow, v)
    End If
End Function

Public Function ApprovePurchaseOrder(ByVal lRow As Long) As Boolean
    Dim vRec As Variant: vRec = GetRecord(SHT_PURCHASEORDER, lRow)
    If IsEmpty(vRec(0)) Then ApprovePurchaseOrder = False: Exit Function
    Dim v(0 To 8) As Variant
    v(1) = SafeConvertToString(vRec(COL_PO_NUMBER + 1))
    v(2) = SafeConvertToString(vRec(COL_PO_VENDOR + 1))
    v(3) = SafeConvertToString(vRec(COL_PO_PROJECT + 1))
    v(4) = vRec(COL_PO_DELIVERY + 1)
    v(5) = SafeConvertToString(vRec(COL_PO_PAYMENTTERMS + 1))
    v(6) = SafeConvertToDouble(vRec(COL_PO_TOTAL + 1))
    v(7) = PO_APPROVED
    v(8) = gCurrentUser
    ApprovePurchaseOrder = UpdateRecord(SHT_PURCHASEORDER, lRow, v)
End Function

Public Function ReceivePurchaseOrder(ByVal lRow As Long) As Boolean
    Dim vRec As Variant: vRec = GetRecord(SHT_PURCHASEORDER, lRow)
    If IsEmpty(vRec(0)) Then ReceivePurchaseOrder = False: Exit Function
    Dim v(0 To 8) As Variant
    v(1) = SafeConvertToString(vRec(COL_PO_NUMBER + 1))
    v(2) = SafeConvertToString(vRec(COL_PO_VENDOR + 1))
    v(3) = SafeConvertToString(vRec(COL_PO_PROJECT + 1))
    v(4) = vRec(COL_PO_DELIVERY + 1)
    v(5) = SafeConvertToString(vRec(COL_PO_PAYMENTTERMS + 1))
    v(6) = SafeConvertToDouble(vRec(COL_PO_TOTAL + 1))
    v(7) = PO_DELIVERED
    v(8) = SafeConvertToString(vRec(COL_PO_APPROVEDBY + 1))
    ReceivePurchaseOrder = UpdateRecord(SHT_PURCHASEORDER, lRow, v)
End Function

Public Function GetPurchaseOrderList() As Variant()
    GetPurchaseOrderList = GetAllRecords(SHT_PURCHASEORDER)
End Function

Public Function FindPOByNumber(ByVal sNumber As String) As Long
    FindPOByNumber = FindRecord(SHT_PURCHASEORDER, COL_PO_NUMBER, sNumber)
End Function
