Attribute VB_Name = "ReportService"
Option Explicit

Public Sub GenerateCashInReport(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal ProjectFilter As String, _
                                ByVal TargetSheetName As String)
    Dim ws As Worksheet, wsSrc As Worksheet, lo As ListObject
    Dim arrData As Variant, i As Long, n As Long, j As Long
    Dim dTotal As Double

    On Error GoTo GenerateCashInReport_Err
    Application.ScreenUpdating = False

    Set wsSrc = ThisWorkbook.Worksheets(SHT_CASHIN): Set lo = wsSrc.ListObjects("tblCashIn")
    If lo.DataBodyRange Is Nothing Then GoTo NoData

    arrData = lo.DataBodyRange.Value
    n = 2

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TargetSheetName)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = TargetSheetName
    On Error GoTo GenerateCashInReport_Err

    ws.Cells.Clear
    ws.Cells(1, 1).Value = "CASH IN REPORT"
    ws.Cells(2, 1).Value = "Period: " & Format$(DateFrom, "dd/mm/yyyy") & " - " & Format$(DateTo, "dd/mm/yyyy")
    If Len(ProjectFilter) > 0 Then ws.Cells(2, 2).Value = "Project: " & ProjectFilter

    ' Headers
    ws.Cells(4, 1).Value = "No": ws.Cells(4, 2).Value = "Receipt No"
    ws.Cells(4, 3).Value = "Date": ws.Cells(4, 4).Value = "Project"
    ws.Cells(4, 5).Value = "Account": ws.Cells(4, 6).Value = "Description"
    ws.Cells(4, 7).Value = "Amount"
    ws.Range("A4:G4").Font.Bold = True: ws.Range("A4:G4").Interior.Color = CLR_HEADER

    dTotal = 0
    For i = 1 To UBound(arrData, 1)
        If CDate(arrData(i, COL_CASHIN_DATE + 1)) >= DateFrom And CDate(arrData(i, COL_CASHIN_DATE + 1)) <= DateTo Then
            If Len(ProjectFilter) = 0 Or arrData(i, COL_CASHIN_PROJECT + 1) = ProjectFilter Then
                ws.Cells(n + 3, 1).Value = n - 1
                ws.Cells(n + 3, 2).Value = arrData(i, COL_CASHIN_RECEIPT_NO + 1)
                ws.Cells(n + 3, 3).Value = Format$(arrData(i, COL_CASHIN_DATE + 1), "dd/mm/yyyy")
                ws.Cells(n + 3, 4).Value = arrData(i, COL_CASHIN_PROJECT + 1)
                ws.Cells(n + 3, 5).Value = arrData(i, COL_CASHIN_ACCOUNT + 1)
                ws.Cells(n + 3, 6).Value = arrData(i, COL_CASHIN_DESC + 1)
                ws.Cells(n + 3, 7).Value = arrData(i, COL_CASHIN_AMOUNT + 1)
                ws.Cells(n + 3, 7).NumberFormat = "#,##0.00"
                dTotal = dTotal + SafeConvertToDouble(arrData(i, COL_CASHIN_AMOUNT + 1))
                n = n + 1
            End If
        End If
    Next i

    ws.Cells(n + 3, 1).Value = "": ws.Cells(n + 3, 6).Value = "TOTAL"
    ws.Cells(n + 3, 6).Font.Bold = True
    ws.Cells(n + 3, 7).Value = dTotal: ws.Cells(n + 3, 7).NumberFormat = "#,##0.00"
    ws.Cells(n + 3, 7).Font.Bold = True
    ws.Columns("A:G").AutoFit
    Application.ScreenUpdating = True
    Exit Sub

NoData:
    MsgBox "No Cash In data found.", vbInformation, APP_NAME
    Application.ScreenUpdating = True
    Exit Sub
GenerateCashInReport_Err:
    LogError "ReportService.GenerateCashInReport", Err.Description
    Application.ScreenUpdating = True
End Sub

Public Sub GenerateCashOutReport(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal ProjectFilter As String, _
                                 ByVal TargetSheetName As String)
    Dim ws As Worksheet, wsSrc As Worksheet, lo As ListObject
    Dim arrData As Variant, i As Long, n As Long
    Dim dTotal As Double

    On Error GoTo GenerateCashOutReport_Err
    Application.ScreenUpdating = False

    Set wsSrc = ThisWorkbook.Worksheets(SHT_CASHOUT): Set lo = wsSrc.ListObjects("tblCashOut")
    If lo.DataBodyRange Is Nothing Then GoTo NoData2

    arrData = lo.DataBodyRange.Value
    n = 2

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TargetSheetName)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = TargetSheetName
    On Error GoTo GenerateCashOutReport_Err

    ws.Cells.Clear
    ws.Cells(1, 1).Value = "CASH OUT REPORT"
    ws.Cells(2, 1).Value = "Period: " & Format$(DateFrom, "dd/mm/yyyy") & " - " & Format$(DateTo, "dd/mm/yyyy")
    If Len(ProjectFilter) > 0 Then ws.Cells(2, 2).Value = "Project: " & ProjectFilter

    ws.Cells(4, 1).Value = "No": ws.Cells(4, 2).Value = "Voucher No"
    ws.Cells(4, 3).Value = "Date": ws.Cells(4, 4).Value = "Project"
    ws.Cells(4, 5).Value = "Account": ws.Cells(4, 6).Value = "Vendor"
    ws.Cells(4, 7).Value = "Description": ws.Cells(4, 8).Value = "Amount"
    ws.Range("A4:H4").Font.Bold = True: ws.Range("A4:H4").Interior.Color = CLR_HEADER

    dTotal = 0
    For i = 1 To UBound(arrData, 1)
        If CDate(arrData(i, COL_CASHOUT_DATE + 1)) >= DateFrom And CDate(arrData(i, COL_CASHOUT_DATE + 1)) <= DateTo Then
            If Len(ProjectFilter) = 0 Or arrData(i, COL_CASHOUT_PROJECT + 1) = ProjectFilter Then
                ws.Cells(n + 3, 1).Value = n - 1
                ws.Cells(n + 3, 2).Value = arrData(i, COL_CASHOUT_VOUCHER_NO + 1)
                ws.Cells(n + 3, 3).Value = Format$(arrData(i, COL_CASHOUT_DATE + 1), "dd/mm/yyyy")
                ws.Cells(n + 3, 4).Value = arrData(i, COL_CASHOUT_PROJECT + 1)
                ws.Cells(n + 3, 5).Value = arrData(i, COL_CASHOUT_ACCOUNT + 1)
                ws.Cells(n + 3, 6).Value = arrData(i, COL_CASHOUT_VENDOR + 1)
                ws.Cells(n + 3, 7).Value = arrData(i, COL_CASHOUT_DESC + 1)
                ws.Cells(n + 3, 8).Value = arrData(i, COL_CASHOUT_AMOUNT + 1)
                ws.Cells(n + 3, 8).NumberFormat = "#,##0.00"
                dTotal = dTotal + SafeConvertToDouble(arrData(i, COL_CASHOUT_AMOUNT + 1))
                n = n + 1
            End If
        End If
    Next i

    ws.Cells(n + 3, 1).Value = "": ws.Cells(n + 3, 7).Value = "TOTAL"
    ws.Cells(n + 3, 7).Font.Bold = True
    ws.Cells(n + 3, 8).Value = dTotal: ws.Cells(n + 3, 8).NumberFormat = "#,##0.00"
    ws.Cells(n + 3, 8).Font.Bold = True
    ws.Columns("A:H").AutoFit
    Application.ScreenUpdating = True
    Exit Sub

NoData2:
    MsgBox "No Cash Out data found.", vbInformation, APP_NAME
    Application.ScreenUpdating = True
    Exit Sub
GenerateCashOutReport_Err:
    LogError "ReportService.GenerateCashOutReport", Err.Description
    Application.ScreenUpdating = True
End Sub

Public Sub GenerateProjectCostReport(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal ProjectFilter As String, _
                                     ByVal TargetSheetName As String)
    Dim ws As Worksheet, wsSrc As Worksheet, lo As ListObject
    Dim arrData As Variant, i As Long, n As Long
    Dim dTotal As Double

    On Error GoTo GenerateProjectCostReport_Err
    Application.ScreenUpdating = False

    Set wsSrc = ThisWorkbook.Worksheets(SHT_TRANSACTIONS): Set lo = wsSrc.ListObjects("tblTransaction")
    If lo.DataBodyRange Is Nothing Then GoTo NoData3

    arrData = lo.DataBodyRange.Value
    n = 2

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TargetSheetName)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = TargetSheetName
    On Error GoTo GenerateProjectCostReport_Err

    ws.Cells.Clear
    ws.Cells(1, 1).Value = "PROJECT COST REPORT"
    ws.Cells(2, 1).Value = "Period: " & Format$(DateFrom, "dd/mm/yyyy") & " - " & Format$(DateTo, "dd/mm/yyyy")
    If Len(ProjectFilter) > 0 Then ws.Cells(2, 2).Value = "Project: " & ProjectFilter

    ws.Cells(4, 1).Value = "No": ws.Cells(4, 2).Value = "Date"
    ws.Cells(4, 3).Value = "Project": ws.Cells(4, 4).Value = "Type"
    ws.Cells(4, 5).Value = "Description": ws.Cells(4, 6).Value = "Amount"
    ws.Cells(4, 7).Value = "VAT": ws.Cells(4, 8).Value = "Total"
    ws.Range("A4:H4").Font.Bold = True: ws.Range("A4:H4").Interior.Color = CLR_HEADER

    dTotal = 0
    For i = 1 To UBound(arrData, 1)
        If CDate(arrData(i, COL_TXN_DATE + 1)) >= DateFrom And CDate(arrData(i, COL_TXN_DATE + 1)) <= DateTo Then
            If Len(ProjectFilter) = 0 Or arrData(i, COL_TXN_PROJ_CODE + 1) = ProjectFilter Then
                ws.Cells(n + 3, 1).Value = n - 1
                ws.Cells(n + 3, 2).Value = Format$(arrData(i, COL_TXN_DATE + 1), "dd/mm/yyyy")
                ws.Cells(n + 3, 3).Value = arrData(i, COL_TXN_PROJ_CODE + 1)
                ws.Cells(n + 3, 4).Value = arrData(i, COL_TXN_TYPE + 1)
                ws.Cells(n + 3, 5).Value = arrData(i, COL_TXN_DESCRIPTION + 1)
                ws.Cells(n + 3, 6).Value = arrData(i, COL_TXN_AMOUNT + 1)
                ws.Cells(n + 3, 7).Value = arrData(i, COL_TXN_VAT + 1)
                ws.Cells(n + 3, 8).Value = arrData(i, COL_TXN_TOTAL + 1)
                ws.Range(ws.Cells(n + 3, 6), ws.Cells(n + 3, 8)).NumberFormat = "#,##0.00"
                dTotal = dTotal + SafeConvertToDouble(arrData(i, COL_TXN_TOTAL + 1))
                n = n + 1
            End If
        End If
    Next i

    ws.Cells(n + 3, 1).Value = "": ws.Cells(n + 3, 7).Value = "TOTAL"
    ws.Cells(n + 3, 7).Font.Bold = True
    ws.Cells(n + 3, 8).Value = dTotal: ws.Cells(n + 3, 8).NumberFormat = "#,##0.00"
    ws.Cells(n + 3, 8).Font.Bold = True
    ws.Columns("A:H").AutoFit
    Application.ScreenUpdating = True
    Exit Sub

NoData3:
    MsgBox "No transaction data found.", vbInformation, APP_NAME
    Application.ScreenUpdating = True
    Exit Sub
GenerateProjectCostReport_Err:
    LogError "ReportService.GenerateProjectCostReport", Err.Description
    Application.ScreenUpdating = True
End Sub

Public Sub GenerateBudgetVsActual(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal ProjectFilter As String, _
                                  ByVal TargetSheetName As String)
    Dim ws As Worksheet, wsProj As Worksheet, wsBud As Worksheet
    Dim loProj As ListObject, loBud As ListObject
    Dim arrProj As Variant, i As Long, n As Long, j As Long
    Dim dPlanned As Double, dActual As Double, dVariance As Double, dPct As Double

    On Error GoTo GenerateBudgetVsActual_Err
    Application.ScreenUpdating = False

    Set wsProj = ThisWorkbook.Worksheets(SHT_PROJECTS): Set loProj = wsProj.ListObjects("tblProject")
    Set wsBud = ThisWorkbook.Worksheets(SHT_BUDGET): Set loBud = wsBud.ListObjects("tblBudget")
    If loProj.DataBodyRange Is Nothing Then GoTo NoData4

    arrProj = loProj.DataBodyRange.Value
    n = 2

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TargetSheetName)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = TargetSheetName
    On Error GoTo GenerateBudgetVsActual_Err

    ws.Cells.Clear
    ws.Cells(1, 1).Value = "BUDGET VS ACTUAL REPORT"
    ws.Cells(2, 1).Value = "Period: " & Format$(DateFrom, "dd/mm/yyyy") & " - " & Format$(DateTo, "dd/mm/yyyy")
    If Len(ProjectFilter) > 0 Then ws.Cells(2, 2).Value = "Project: " & ProjectFilter

    ws.Cells(4, 1).Value = "No": ws.Cells(4, 2).Value = "Project Code"
    ws.Cells(4, 3).Value = "Project Name": ws.Cells(4, 4).Value = "Budget"
    ws.Cells(4, 5).Value = "Actual": ws.Cells(4, 6).Value = "Variance"
    ws.Cells(4, 7).Value = "% Used"
    ws.Range("A4:G4").Font.Bold = True: ws.Range("A4:G4").Interior.Color = CLR_HEADER

    For i = 1 To UBound(arrProj, 1)
        If Len(ProjectFilter) = 0 Or arrProj(i, COL_PROJ_CODE + 1) = ProjectFilter Then
            dPlanned = SafeConvertToDouble(arrProj(i, COL_PROJ_BUDGET + 1))
            dActual = 0
            If Not loBud.DataBodyRange Is Nothing Then
                For j = 1 To loBud.DataBodyRange.Rows.Count
                    If loBud.DataBodyRange.Cells(j, COL_BUDGET_PROJ + 1).Value = arrProj(i, COL_PROJ_CODE + 1) Then
                        dActual = dActual + SafeConvertToDouble(loBud.DataBodyRange.Cells(j, COL_BUDGET_ACTUAL + 1).Value)
                    End If
                Next j
            End If
            dVariance = dPlanned - dActual
            If dPlanned > 0 Then dPct = (dActual / dPlanned) * 100 Else dPct = 0

            ws.Cells(n + 3, 1).Value = n - 1
            ws.Cells(n + 3, 2).Value = arrProj(i, COL_PROJ_CODE + 1)
            ws.Cells(n + 3, 3).Value = arrProj(i, COL_PROJ_NAME + 1)
            ws.Cells(n + 3, 4).Value = dPlanned
            ws.Cells(n + 3, 5).Value = dActual
            ws.Cells(n + 3, 6).Value = dVariance
            ws.Cells(n + 3, 7).Value = Round(dPct, 1) & "%"
            ws.Range(ws.Cells(n + 3, 4), ws.Cells(n + 3, 6)).NumberFormat = "#,##0.00"
            n = n + 1
        End If
    Next i

    ws.Columns("A:G").AutoFit
    Application.ScreenUpdating = True
    Exit Sub

NoData4:
    MsgBox "No project data found.", vbInformation, APP_NAME
    Application.ScreenUpdating = True
    Exit Sub
GenerateBudgetVsActual_Err:
    LogError "ReportService.GenerateBudgetVsActual", Err.Description
    Application.ScreenUpdating = True
End Sub

Public Sub GenerateGLReport(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal AccountCode As String, _
                            ByVal TargetSheetName As String)
    Dim ws As Worksheet, wsLedger As Worksheet, loLedger As ListObject
    Dim arrData As Variant, i As Long, n As Long

    On Error GoTo GenerateGLReport_Err
    Application.ScreenUpdating = False

    Set wsLedger = ThisWorkbook.Worksheets(SHT_LEDGER): Set loLedger = wsLedger.ListObjects(TBL_LEDGER)
    If loLedger.DataBodyRange Is Nothing Then GoTo NoData5

    arrData = loLedger.DataBodyRange.Value
    n = 2

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TargetSheetName)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = TargetSheetName
    On Error GoTo GenerateGLReport_Err

    ws.Cells.Clear
    ws.Cells(1, 1).Value = "GENERAL LEDGER REPORT"
    ws.Cells(2, 1).Value = "Period: " & Format$(DateFrom, "dd/mm/yyyy") & " - " & Format$(DateTo, "dd/mm/yyyy")
    If Len(AccountCode) > 0 Then ws.Cells(2, 2).Value = "Account: " & AccountCode

    ws.Cells(4, 1).Value = "No": ws.Cells(4, 2).Value = "Date"
    ws.Cells(4, 3).Value = "Account Code": ws.Cells(4, 4).Value = "Account Name"
    ws.Cells(4, 5).Value = "Description": ws.Cells(4, 6).Value = "Reference"
    ws.Cells(4, 7).Value = "Debit": ws.Cells(4, 8).Value = "Credit"
    ws.Cells(4, 9).Value = "Balance"
    ws.Range("A4:I4").Font.Bold = True: ws.Range("A4:I4").Interior.Color = CLR_HEADER

    For i = 1 To UBound(arrData, 1)
        If CDate(arrData(i, COL_LEDGER_DATE + 1)) >= DateFrom And CDate(arrData(i, COL_LEDGER_DATE + 1)) <= DateTo Then
            If Len(AccountCode) = 0 Or arrData(i, COL_LEDGER_ACCOUNT_CODE + 1) = AccountCode Then
                ws.Cells(n + 3, 1).Value = n - 1
                ws.Cells(n + 3, 2).Value = Format$(arrData(i, COL_LEDGER_DATE + 1), "dd/mm/yyyy")
                ws.Cells(n + 3, 3).Value = arrData(i, COL_LEDGER_ACCOUNT_CODE + 1)
                ws.Cells(n + 3, 4).Value = arrData(i, COL_LEDGER_ACCOUNT_NAME + 1)
                ws.Cells(n + 3, 5).Value = arrData(i, COL_LEDGER_DESC + 1)
                ws.Cells(n + 3, 6).Value = arrData(i, COL_LEDGER_REF + 1)
                ws.Cells(n + 3, 7).Value = arrData(i, COL_LEDGER_DEBIT + 1)
                ws.Cells(n + 3, 8).Value = arrData(i, COL_LEDGER_CREDIT + 1)
                ws.Cells(n + 3, 9).Value = arrData(i, COL_LEDGER_BALANCE + 1)
                ws.Range(ws.Cells(n + 3, 7), ws.Cells(n + 3, 9)).NumberFormat = "#,##0.00"
                n = n + 1
            End If
        End If
    Next i

    If n = 2 Then
        ws.Cells(6, 1).Value = "No records found for the selected period."
    End If
    ws.Columns("A:I").AutoFit
    Application.ScreenUpdating = True
    Exit Sub

NoData5:
    MsgBox "No ledger data found.", vbInformation, APP_NAME
    Application.ScreenUpdating = True
    Exit Sub
GenerateGLReport_Err:
    LogError "ReportService.GenerateGLReport", Err.Description
    Application.ScreenUpdating = True
End Sub

Public Sub ExportReportToPDF(ByVal SourceSheetName As String)
    Dim ws As Worksheet
    On Error GoTo ExportReportToPDF_Err
    Set ws = ThisWorkbook.Worksheets(SourceSheetName)
    If ws Is Nothing Then MsgBox "Sheet not found: " & SourceSheetName, vbExclamation: Exit Sub
    Dim sFileName As String
    sFileName = ThisWorkbook.Path & "\" & SourceSheetName & "_" & Format$(Now, "yyyymmdd_hhmmss") & ".pdf"
    ws.ExportAsFixedFormat xlTypePDF, sFileName, OpenAfterPublish:=True
    LogInfo "ReportService.ExportReportToPDF", "Exported: " & sFileName
ExportReportToPDF_Err:
End Sub

Public Sub PrintReport(ByVal SourceSheetName As String)
    Dim ws As Worksheet
    On Error GoTo PrintReport_Err
    Set ws = ThisWorkbook.Worksheets(SourceSheetName)
    If ws Is Nothing Then MsgBox "Sheet not found: " & SourceSheetName, vbExclamation: Exit Sub
    ws.PrintOut
PrintReport_Err:
End Sub

Public Sub ExportReportToExcel(ByVal SourceSheetName As String)
    Dim ws As Worksheet
    Dim wbNew As Workbook
    On Error GoTo ExportReportToExcel_Err
    Set ws = ThisWorkbook.Worksheets(SourceSheetName)
    If ws Is Nothing Then MsgBox "Sheet not found: " & SourceSheetName, vbExclamation: Exit Sub
    ws.Copy
    Set wbNew = ActiveWorkbook
    Dim sFileName As String
    sFileName = ThisWorkbook.Path & "\" & SourceSheetName & "_" & Format$(Now, "yyyymmdd_hhmmss") & ".xlsx"
    wbNew.SaveAs sFileName, xlOpenXMLWorkbook
    wbNew.Close
    LogInfo "ReportService.ExportReportToExcel", "Exported: " & sFileName
ExportReportToExcel_Err:
End Sub
