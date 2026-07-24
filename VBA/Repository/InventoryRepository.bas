Attribute VB_Name = "InventoryRepository"
Option Explicit

Private Const INV_MOD As String = "InventoryRepository"

Public Function GetInventorySummary() As String
    Dim ws As Worksheet, lLast As Long, lRow As Long
    Dim dTotalQty As Double, dTotalCost As Double, lCount As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_MATERIAL)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        If Len(SafeConvertToString(ws.Cells(lRow, 2).Value)) > 0 Then
            lCount = lCount + 1
            dTotalQty = dTotalQty + SafeConvertToDouble(ws.Cells(lRow, COL_MAT_REMAINING + 1).Value)
            dTotalCost = dTotalCost + SafeConvertToDouble(ws.Cells(lRow, COL_MAT_TOTALCOST + 1).Value)
        End If
    Next lRow
    GetInventorySummary = "Items:" & lCount & " Qty:" & dTotalQty & " Value:" & FormatCurrencyIDR(dTotalCost)
    On Error GoTo 0
End Function

Public Function GetLowStockMaterials(ByVal dThreshold As Double) As Variant()
    Dim ws As Worksheet, lLast As Long, lRow As Long, n As Long
    Dim arrResult() As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_MATERIAL)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    n = 0: ReDim arrResult(1 To lLast)
    For lRow = 2 To lLast
        If SafeConvertToDouble(ws.Cells(lRow, COL_MAT_REMAINING + 1).Value) <= dThreshold Then
            n = n + 1: arrResult(n) = GetRecord(SHT_MATERIAL, lRow)
        End If
    Next lRow
    If n = 0 Then GetLowStockMaterials = Array() Else GetLowStockMaterials = arrResult
    On Error GoTo 0
End Function

Public Function GetMaterialBalance(ByVal sCode As String) As Double
    Dim lRow As Long: lRow = FindMaterialByCode(sCode)
    If lRow <= 0 Then GetMaterialBalance = 0: Exit Function
    Dim vRec As Variant: vRec = GetRecord(SHT_MATERIAL, lRow)
    If IsEmpty(vRec(0)) Then GetMaterialBalance = 0 Else GetMaterialBalance = SafeConvertToDouble(vRec(COL_MAT_REMAINING + 1))
End Function

Public Function GetTotalInventoryValue() As Double
    Dim ws As Worksheet, lLast As Long, lRow As Long, dTotal As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_MATERIAL)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For lRow = 2 To lLast
        dTotal = dTotal + SafeConvertToDouble(ws.Cells(lRow, COL_MAT_TOTALCOST + 1).Value)
    Next lRow
    GetTotalInventoryValue = dTotal
    On Error GoTo 0
End Function
