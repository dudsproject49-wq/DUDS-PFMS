Attribute VB_Name = "MaterialRepository"
Option Explicit

Private Const MAT_MOD As String = "MaterialRepository"

Public Function GetNextMaterialCode(ByVal sProject As String) As String
    Dim lCount As Long: lCount = GetRecordCount(SHT_MATERIAL)
    GetNextMaterialCode = "MAT" & Format$(lCount + 1, "0000")
End Function

Public Function SaveMaterial(ByVal lRow As Long, ByVal sCode As String, ByVal sName As String, _
    ByVal sUnit As String, ByVal dBudgetQty As Double, ByVal dActualQty As Double, _
    ByVal dRemaining As Double, ByVal dUnitPrice As Double, ByVal dTotalCost As Double, _
    ByVal sProject As String) As Boolean
    Dim v(0 To 11) As Variant
    If lRow <= 0 Then
        v(0) = GenerateGUID(): v(1) = sCode: v(2) = sName: v(3) = sUnit
        v(4) = dBudgetQty: v(5) = dActualQty: v(6) = dRemaining
        v(7) = dUnitPrice: v(8) = dTotalCost: v(9) = sProject
        v(10) = gCurrentUser: v(11) = Now()
        lRow = InsertRecord(SHT_MATERIAL, v)
        SaveMaterial = (lRow > 0)
    Else
        v(1) = sCode: v(2) = sName: v(3) = sUnit: v(4) = dBudgetQty
        v(5) = dActualQty: v(6) = dRemaining: v(7) = dUnitPrice
        v(8) = dTotalCost: v(9) = sProject
        SaveMaterial = UpdateRecord(SHT_MATERIAL, lRow, v)
    End If
End Function

Public Function ReceiveMaterial(ByVal lMaterialRow As Long, ByVal dQty As Double, _
    ByVal sRefNo As String, ByVal sNotes As String) As Boolean
    Dim vRec As Variant: vRec = GetRecord(SHT_MATERIAL, lMaterialRow)
    If IsEmpty(vRec(0)) Then ReceiveMaterial = False: Exit Function
    
    Dim dNewActual As Double, dNewRemaining As Double
    dNewActual = SafeConvertToDouble(vRec(COL_MAT_ACTUALQTY + 1)) + dQty
    dNewRemaining = SafeConvertToDouble(vRec(COL_MAT_REMAINING + 1)) + dQty
    Dim dTotalCost As Double: dTotalCost = dNewActual * SafeConvertToDouble(vRec(COL_MAT_UNITPRICE + 1))
    
    Dim vMat(0 To 9) As Variant
    vMat(1) = SafeConvertToString(vRec(COL_MAT_CODE + 1))
    vMat(2) = SafeConvertToString(vRec(COL_MAT_NAME + 1))
    vMat(3) = SafeConvertToString(vRec(COL_MAT_UNIT + 1))
    vMat(4) = SafeConvertToDouble(vRec(COL_MAT_BUDGETQTY + 1))
    vMat(5) = dNewActual
    vMat(6) = dNewRemaining
    vMat(7) = SafeConvertToDouble(vRec(COL_MAT_UNITPRICE + 1))
    vMat(8) = dTotalCost
    vMat(9) = SafeConvertToString(vRec(COL_MAT_PROJECT + 1))
    If Not UpdateRecord(SHT_MATERIAL, lMaterialRow, vMat) Then
        ReceiveMaterial = False: Exit Function
    End If
    
    Dim vTrans(0 To 8) As Variant
    vTrans(0) = GenerateGUID()
    vTrans(1) = SafeConvertToString(vRec(COL_MAT_CODE + 1))
    vTrans(2) = MT_RECEIVE
    vTrans(3) = dQty
    vTrans(4) = Now()
    vTrans(5) = sRefNo
    vTrans(6) = sNotes
    vTrans(7) = gCurrentUser
    vTrans(8) = Now()
    InsertRecord SHT_MATTRANS, vTrans
    
    LogInfo MAT_MOD & ".ReceiveMaterial", "Received " & dQty & " of " & SafeConvertToString(vRec(COL_MAT_NAME + 1)) & " ref:" & sRefNo
    ReceiveMaterial = True
End Function

Public Function IssueMaterial(ByVal lMaterialRow As Long, ByVal dQty As Double, _
    ByVal sRefNo As String, ByVal sNotes As String) As Boolean
    Dim vRec As Variant: vRec = GetRecord(SHT_MATERIAL, lMaterialRow)
    If IsEmpty(vRec(0)) Then IssueMaterial = False: Exit Function
    
    Dim dRemaining As Double: dRemaining = SafeConvertToDouble(vRec(COL_MAT_REMAINING + 1))
    If dQty > dRemaining Then
        MsgBox "Insufficient material quantity. Available: " & dRemaining, vbExclamation
        IssueMaterial = False: Exit Function
    End If
    
    Dim dNewActual As Double, dNewRemaining As Double
    dNewActual = SafeConvertToDouble(vRec(COL_MAT_ACTUALQTY + 1)) + dQty
    dNewRemaining = dRemaining - dQty
    
    Dim vMat(0 To 9) As Variant
    vMat(1) = SafeConvertToString(vRec(COL_MAT_CODE + 1))
    vMat(2) = SafeConvertToString(vRec(COL_MAT_NAME + 1))
    vMat(3) = SafeConvertToString(vRec(COL_MAT_UNIT + 1))
    vMat(4) = SafeConvertToDouble(vRec(COL_MAT_BUDGETQTY + 1))
    vMat(5) = dNewActual
    vMat(6) = dNewRemaining
    vMat(7) = SafeConvertToDouble(vRec(COL_MAT_UNITPRICE + 1))
    vMat(8) = dNewActual * SafeConvertToDouble(vRec(COL_MAT_UNITPRICE + 1))
    vMat(9) = SafeConvertToString(vRec(COL_MAT_PROJECT + 1))
    If Not UpdateRecord(SHT_MATERIAL, lMaterialRow, vMat) Then
        IssueMaterial = False: Exit Function
    End If
    
    Dim vTrans(0 To 8) As Variant
    vTrans(0) = GenerateGUID()
    vTrans(1) = SafeConvertToString(vRec(COL_MAT_CODE + 1))
    vTrans(2) = MT_ISSUE
    vTrans(3) = dQty
    vTrans(4) = Now()
    vTrans(5) = sRefNo
    vTrans(6) = sNotes
    vTrans(7) = gCurrentUser
    vTrans(8) = Now()
    InsertRecord SHT_MATTRANS, vTrans
    
    Dim dCost As Double: dCost = dQty * SafeConvertToDouble(vRec(COL_MAT_UNITPRICE + 1))
    Dim sProject As String: sProject = SafeConvertToString(vRec(COL_MAT_PROJECT + 1))
    UpdateProjectCost sProject, dCost
    PostMaterialJournal sProject, "Material Issue - " & SafeConvertToString(vRec(COL_MAT_NAME + 1)), dCost, sRefNo
    
    LogInfo MAT_MOD & ".IssueMaterial", "Issued " & dQty & " of " & SafeConvertToString(vRec(COL_MAT_NAME + 1)) & " ref:" & sRefNo
    IssueMaterial = True
End Function

Public Function ReturnMaterial(ByVal lMaterialRow As Long, ByVal dQty As Double, _
    ByVal sRefNo As String, ByVal sNotes As String) As Boolean
    ReceiveMaterial lMaterialRow, dQty, sRefNo, "RETURN: " & sNotes
    Dim vRec As Variant: vRec = GetRecord(SHT_MATERIAL, lMaterialRow)
    Dim dCost As Double: dCost = dQty * SafeConvertToDouble(vRec(COL_MAT_UNITPRICE + 1))
    Dim sProject As String: sProject = SafeConvertToString(vRec(COL_MAT_PROJECT + 1))
    UpdateProjectCost sProject, -dCost
    PostMaterialJournal sProject, "Material Return - " & SafeConvertToString(vRec(COL_MAT_NAME + 1)), -dCost, sRefNo
    LogInfo MAT_MOD & ".ReturnMaterial", "Returned " & dQty & " of " & SafeConvertToString(vRec(COL_MAT_NAME + 1)) & " ref:" & sRefNo
    ReturnMaterial = True
End Function

Public Function GetMaterialList() As Variant()
    GetMaterialList = GetAllRecords(SHT_MATERIAL)
End Function

Public Function FindMaterialByCode(ByVal sCode As String) As Long
    FindMaterialByCode = FindRecord(SHT_MATERIAL, COL_MAT_CODE, sCode)
End Function

Public Function GetMaterialTransactions(ByVal sMaterialCode As String) As Variant()
    Dim col As Collection: Set col = GetRecordsByColumn(SHT_MATTRANS, COL_MT_MATERIALID, sMaterialCode)
    If col Is Nothing Then GetMaterialTransactions = Array(): Exit Function
    If col.Count = 0 Then GetMaterialTransactions = Array(): Exit Function
    Dim vRec As Variant, i As Long, n As Long: n = 0
    ReDim arrResult(1 To col.Count)
    For i = 1 To col.Count
        vRec = GetRecord(SHT_MATTRANS, col(i))
        If Not IsEmpty(vRec(0)) Then n = n + 1: arrResult(n) = vRec
    Next i
    If n = 0 Then GetMaterialTransactions = Array() Else GetMaterialTransactions = arrResult
End Function

Private Sub UpdateProjectCost(ByVal sProject As String, ByVal dAmount As Double)
    Dim lRow As Long: lRow = FindRecord(SHT_PROJECTS, COL_PROJ_CODE, sProject)
    If lRow <= 0 Then Exit Sub
    Dim vRec As Variant: vRec = GetRecord(SHT_PROJECTS, lRow)
    If IsEmpty(vRec(0)) Then Exit Sub
    Dim v(0 To 13) As Variant
    v(1) = SafeConvertToString(vRec(COL_PROJ_CODE + 1))
    v(2) = SafeConvertToString(vRec(COL_PROJ_NAME + 1))
    v(3) = SafeConvertToString(vRec(COL_PROJ_CLIENT + 1))
    v(4) = SafeConvertToString(vRec(COL_PROJ_LOCATION + 1))
    v(5) = vRec(COL_PROJ_START_DATE + 1)
    v(6) = vRec(COL_PROJ_END_DATE + 1)
    v(7) = SafeConvertToDouble(vRec(COL_PROJ_CONTRACT_VAL + 1))
    v(8) = SafeConvertToDouble(vRec(COL_PROJ_BUDGET + 1))
    v(9) = SafeConvertToString(vRec(COL_PROJ_STATUS + 1))
    v(10) = SafeConvertToDouble(vRec(COL_PROJ_PROGRESS + 1))
    v(11) = SafeConvertToString(vRec(COL_PROJ_NOTES + 1))
    UpdateRecord SHT_PROJECTS, lRow, v
End Sub

Private Sub PostMaterialJournal(ByVal sProject As String, ByVal sDesc As String, _
    ByVal dAmount As Double, ByVal sRefNo As String)
    Dim v(0 To 10) As Variant
    v(0) = GenerateGUID(): v(1) = Now(): v(2) = "Material"
    v(3) = sRefNo: v(4) = sProject: v(5) = "5-1000"
    v(6) = sDesc
    If dAmount > 0 Then v(7) = dAmount Else v(7) = 0
    If dAmount < 0 Then v(8) = -dAmount Else v(8) = 0
    v(9) = gCurrentUser: v(10) = Now()
    InsertRecord SHT_JOURNAL, v
End Function
