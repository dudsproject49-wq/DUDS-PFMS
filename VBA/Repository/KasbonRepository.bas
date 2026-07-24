Attribute VB_Name = "KasbonRepository"
Option Explicit

Private Const FnName As String = "KasbonRepository"

Public Function SaveKasbon(ByVal ksNumber As String, ByVal ksDate As Date, _
                           ByVal ksEmployee As String, ByVal ksProject As String, _
                           ByVal ksDescription As String, ByVal ksAmount As Double, _
                           ByVal ksApprovedBy As String, ByVal ksStatus As String, _
                           ByVal ksCreatedBy As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    Dim lRow As Long
    On Error GoTo SaveKasbon_Err
    Set ws = ThisWorkbook.Worksheets(SHT_KASBON)
    Set lo = ws.ListObjects("tblKasbon")
    If lo.DataBodyRange Is Nothing Then
        lRow = 2
    Else
        lRow = lo.DataBodyRange.Rows.Count + 2
    End If
    ws.Cells(lRow, 1).Value = lRow - 1
    ws.Cells(lRow, 2).Value = ksNumber
    ws.Cells(lRow, 3).Value = ksDate
    ws.Cells(lRow, 4).Value = ksEmployee
    ws.Cells(lRow, 5).Value = ksProject
    ws.Cells(lRow, 6).Value = ksDescription
    ws.Cells(lRow, 7).Value = ksAmount
    ws.Cells(lRow, 8).Value = ksApprovedBy
    ws.Cells(lRow, 9).Value = ksStatus
    ws.Cells(lRow, 10).Value = ksCreatedBy
    ws.Cells(lRow, 11).Value = Now
    SaveKasbon = True
    Exit Function
SaveKasbon_Err:
    LogError "KasbonRepository.SaveKasbon", Err.Description
    SaveKasbon = False
End Function

Public Function UpdateKasbon(ByVal ksId As Long, ByVal ksNumber As String, _
                             ByVal ksDate As Date, ByVal ksEmployee As String, _
                             ByVal ksProject As String, ByVal ksDescription As String, _
                             ByVal ksAmount As Double, ByVal ksApprovedBy As String, _
                             ByVal ksStatus As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    Dim i As Long
    On Error GoTo UpdateKasbon_Err
    Set ws = ThisWorkbook.Worksheets(SHT_KASBON)
    Set lo = ws.ListObjects("tblKasbon")
    If lo.DataBodyRange Is Nothing Then UpdateKasbon = False: Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If CLng(lo.DataBodyRange.Cells(i, 1).Value) = ksId Then
            lo.DataBodyRange.Cells(i, 2).Value = ksNumber
            lo.DataBodyRange.Cells(i, 3).Value = ksDate
            lo.DataBodyRange.Cells(i, 4).Value = ksEmployee
            lo.DataBodyRange.Cells(i, 5).Value = ksProject
            lo.DataBodyRange.Cells(i, 6).Value = ksDescription
            lo.DataBodyRange.Cells(i, 7).Value = ksAmount
            lo.DataBodyRange.Cells(i, 8).Value = ksApprovedBy
            lo.DataBodyRange.Cells(i, 9).Value = ksStatus
            UpdateKasbon = True
            Exit Function
        End If
    Next i
    UpdateKasbon = False
    Exit Function
UpdateKasbon_Err:
    LogError "KasbonRepository.UpdateKasbon", Err.Description
    UpdateKasbon = False
End Function

Public Function GetKasbonById(ByVal ksId As Long) As Variant
    Dim ws As Worksheet, lo As ListObject
    Dim i As Long
    On Error GoTo GetKasbonById_Err
    Set ws = ThisWorkbook.Worksheets(SHT_KASBON)
    Set lo = ws.ListObjects("tblKasbon")
    If lo.DataBodyRange Is Nothing Then GetKasbonById = Null: Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If CLng(lo.DataBodyRange.Cells(i, 1).Value) = ksId Then
            GetKasbonById = lo.DataBodyRange.Rows(i).Value
            Exit Function
        End If
    Next i
    GetKasbonById = Null
    Exit Function
GetKasbonById_Err:
    LogError "KasbonRepository.GetKasbonById", Err.Description
    GetKasbonById = Null
End Function

Public Function GetAllKasbon() As Variant
    Dim ws As Worksheet, lo As ListObject
    On Error GoTo GetAllKasbon_Err
    Set ws = ThisWorkbook.Worksheets(SHT_KASBON)
    Set lo = ws.ListObjects("tblKasbon")
    If lo.DataBodyRange Is Nothing Then
        GetAllKasbon = Array()
    Else
        GetAllKasbon = lo.DataBodyRange.Value
    End If
    Exit Function
GetAllKasbon_Err:
    LogError "KasbonRepository.GetAllKasbon", Err.Description
    GetAllKasbon = Array()
End Function

Public Function DeleteKasbon(ByVal ksId As Long) As Boolean
    Dim ws As Worksheet, lo As ListObject
    Dim i As Long
    On Error GoTo DeleteKasbon_Err
    Set ws = ThisWorkbook.Worksheets(SHT_KASBON)
    Set lo = ws.ListObjects("tblKasbon")
    If lo.DataBodyRange Is Nothing Then DeleteKasbon = False: Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If CLng(lo.DataBodyRange.Cells(i, 1).Value) = ksId Then
            lo.DataBodyRange.Rows(i).Delete
            DeleteKasbon = True
            Exit Function
        End If
    Next i
    DeleteKasbon = False
    Exit Function
DeleteKasbon_Err:
    LogError "KasbonRepository.DeleteKasbon", Err.Description
    DeleteKasbon = False
End Function

Public Function GetKasbonByEmployee(ByVal ksEmployee As String) As Variant
    Dim ws As Worksheet, lo As ListObject
    Dim arrResult() As Variant, i As Long, n As Long
    On Error GoTo GetKasbonByEmployee_Err
    Set ws = ThisWorkbook.Worksheets(SHT_KASBON)
    Set lo = ws.ListObjects("tblKasbon")
    If lo.DataBodyRange Is Nothing Then GetKasbonByEmployee = Array(): Exit Function
    n = 0
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, 4).Value = ksEmployee Then
            n = n + 1
            ReDim Preserve arrResult(1 To n)
            arrResult(n) = lo.DataBodyRange.Rows(i).Value
        End If
    Next i
    If n = 0 Then GetKasbonByEmployee = Array() Else GetKasbonByEmployee = arrResult
    Exit Function
GetKasbonByEmployee_Err:
    LogError "KasbonRepository.GetKasbonByEmployee", Err.Description
    GetKasbonByEmployee = Array()
End Function

Public Function GetKasbonByProject(ByVal ksProject As String) As Variant
    Dim ws As Worksheet, lo As ListObject
    Dim arrResult() As Variant, i As Long, n As Long
    On Error GoTo GetKasbonByProject_Err
    Set ws = ThisWorkbook.Worksheets(SHT_KASBON)
    Set lo = ws.ListObjects("tblKasbon")
    If lo.DataBodyRange Is Nothing Then GetKasbonByProject = Array(): Exit Function
    n = 0
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, 5).Value = ksProject Then
            n = n + 1
            ReDim Preserve arrResult(1 To n)
            arrResult(n) = lo.DataBodyRange.Rows(i).Value
        End If
    Next i
    If n = 0 Then GetKasbonByProject = Array() Else GetKasbonByProject = arrResult
    Exit Function
GetKasbonByProject_Err:
    LogError "KasbonRepository.GetKasbonByProject", Err.Description
    GetKasbonByProject = Array()
End Function

Public Function GetKasbonTotalByEmployee(ByVal ksEmployee As String) As Double
    Dim ws As Worksheet, lo As ListObject
    Dim dTotal As Double, i As Long
    On Error GoTo GetKasbonTotalByEmployee_Err
    Set ws = ThisWorkbook.Worksheets(SHT_KASBON)
    Set lo = ws.ListObjects("tblKasbon")
    If lo.DataBodyRange Is Nothing Then GetKasbonTotalByEmployee = 0: Exit Function
    dTotal = 0
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_KSB_EMPLOYEE + 1).Value = ksEmployee Then
            dTotal = dTotal + SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_KSB_AMOUNT + 1).Value)
        End If
    Next i
    GetKasbonTotalByEmployee = dTotal
    Exit Function
GetKasbonTotalByEmployee_Err:
    LogError "KasbonRepository.GetKasbonTotalByEmployee", Err.Description
    GetKasbonTotalByEmployee = 0
End Function

Public Function GenerateKasbonNumber() As String
    Dim ws As Worksheet, lo As ListObject
    Dim lCount As Long
    On Error GoTo GenerateKasbonNumber_Err
    Set ws = ThisWorkbook.Worksheets(SHT_KASBON)
    Set lo = ws.ListObjects("tblKasbon")
    If lo.DataBodyRange Is Nothing Then
        lCount = 0
    Else
        lCount = lo.DataBodyRange.Rows.Count
    End If
    GenerateKasbonNumber = "KSB-" & Format$(Now, "yyMMdd") & "-" & Format$(lCount + 1, "0000")
    Exit Function
GenerateKasbonNumber_Err:
    LogError "KasbonRepository.GenerateKasbonNumber", Err.Description
    GenerateKasbonNumber = "KSB-" & Format$(Now, "yyMMdd") & "-0001"
End Function
