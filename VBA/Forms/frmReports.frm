Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mReportType As String

Private Sub UserForm_Initialize()
    Me.Caption = "DUDS-PFMS - Reports"
    Me.StartUpPosition = 0
    Me.Left = (Application.Width - Me.Width) / 2
    Me.Top = (Application.Height - Me.Height) / 2
    mReportType = ""
    dtpFrom.Value = DateSerial(Year(Now), Month(Now), 1)
    dtpTo.Value = Now
    LoadProjectCombo
    LoadAccountCombo
    cmbReportType.Clear
    cmbReportType.AddItem "Cash In Report"
    cmbReportType.AddItem "Cash Out Report"
    cmbReportType.AddItem "Project Cost Report"
    cmbReportType.AddItem "Budget vs Actual"
    cmbReportType.AddItem "General Ledger"
    cmbReportType.AddItem "Trial Balance"
    cmbReportType.AddItem "Profit & Loss"
    cmbReportType.AddItem "Balance Sheet"
    cmbReportType.AddItem "Cash Flow"
    If cmbReportType.ListCount > 0 Then cmbReportType.ListIndex = 0
    cmbReportType.SetFocus
End Sub

Private Sub LoadProjectCombo()
    Dim arrData As Variant, i As Long
    On Error Resume Next
    cboProject.Clear
    cboProject.AddItem "All Projects"
    arrData = LoadProjects()
    If IsArray(arrData) Then
        For i = LBound(arrData, 1) To UBound(arrData, 1)
            cboProject.AddItem SafeConvertToString(arrData(i, 2))
        Next i
    End If
    cboProject.ListIndex = 0
    On Error GoTo 0
End Sub

Private Sub LoadAccountCombo()
    Dim arrData As Variant, i As Long
    On Error Resume Next
    cboAccount.Clear
    cboAccount.AddItem "All Accounts"
    arrData = LoadAccounts()
    If IsArray(arrData) Then
        If UBound(arrData, 1) > 0 Then
            For i = LBound(arrData, 1) To UBound(arrData, 1)
                cboAccount.AddItem SafeConvertToString(arrData(i, 2))
            Next i
        End If
    End If
    cboAccount.ListIndex = 0
    On Error GoTo 0
End Sub

Private Sub cmbReportType_Change()
    mReportType = cmbReportType.Text
    ' Show/hide account filter based on report type
    Select Case mReportType
        Case "General Ledger":
            cboAccount.Visible = True: lblAccount.Visible = True
        Case Else:
            cboAccount.Visible = False: lblAccount.Visible = False
    End Select
End Sub

Private Sub cmdGenerate_Click()
    Dim sTarget As String, sProject As String, sAccount As String
    Dim dFrom As Date, dTo As Date

    If Len(Trim$(cmbReportType.Text)) = 0 Then
        MsgBox "Please select a report type.", vbExclamation, APP_NAME: Exit Sub
    End If

    mReportType = cmbReportType.Text
    dFrom = CDate(dtpFrom.Value): dTo = CDate(dtpTo.Value)
    If cboProject.ListIndex > 0 Then sProject = cboProject.Text Else sProject = ""
    If cboAccount.ListIndex > 0 Then sAccount = cboAccount.Text Else sAccount = ""

    sTarget = Replace(mReportType, " ", "_") & "_" & Format$(Now, "yyyymmdd")

    Application.ScreenUpdating = False

    Select Case mReportType
        Case "Cash In Report":
            GenerateCashInReport dFrom, dTo, sProject, sTarget
        Case "Cash Out Report":
            GenerateCashOutReport dFrom, dTo, sProject, sTarget
        Case "Project Cost Report":
            GenerateProjectCostReport dFrom, dTo, sProject, sTarget
        Case "Budget vs Actual":
            GenerateBudgetVsActual dFrom, dTo, sProject, sTarget
        Case "General Ledger":
            Dim sAcctCode As String
            sAcctCode = GetAccountCodeByName(sAccount)
            GenerateGLReport dFrom, dTo, sAcctCode, sTarget
        Case "Trial Balance":
            Dim lFY As Long
            lFY = Year(dFrom)
            ExportTrialBalanceToSheet lFY, sTarget
        Case "Profit & Loss":
            lFY = Year(dFrom)
            ExportProfitLossToSheet lFY, sTarget
        Case "Balance Sheet":
            lFY = Year(dFrom)
            ExportBalanceSheetToSheet lFY, sTarget
        Case "Cash Flow":
            lFY = Year(dFrom)
            ExportCashFlowToSheet lFY, sTarget
    End Select

    Application.ScreenUpdating = True
    MsgBox "Report generated: " & sTarget, vbInformation, APP_NAME
End Sub

Private Sub cmdExportPDF_Click()
    Dim sTarget As String
    If Len(Trim$(cmbReportType.Text)) = 0 Then Exit Sub
    sTarget = Replace(cmbReportType.Text, " ", "_") & "_" & Format$(Now, "yyyymmdd")
    On Error Resume Next
    If SheetExists(sTarget) Then
        ExportReportToPDF sTarget
    Else
        MsgBox "Please generate the report first.", vbExclamation
    End If
    On Error GoTo 0
End Sub

Private Sub cmdPrint_Click()
    Dim sTarget As String
    If Len(Trim$(cmbReportType.Text)) = 0 Then Exit Sub
    sTarget = Replace(cmbReportType.Text, " ", "_") & "_" & Format$(Now, "yyyymmdd")
    On Error Resume Next
    If SheetExists(sTarget) Then
        PrintReport sTarget
    Else
        MsgBox "Please generate the report first.", vbExclamation
    End If
    On Error GoTo 0
End Sub

Private Sub cmdExportExcel_Click()
    Dim sTarget As String
    If Len(Trim$(cmbReportType.Text)) = 0 Then Exit Sub
    sTarget = Replace(cmbReportType.Text, " ", "_") & "_" & Format$(Now, "yyyymmdd")
    On Error Resume Next
    If SheetExists(sTarget) Then
        ExportReportToExcel sTarget
    Else
        MsgBox "Please generate the report first.", vbExclamation
    End If
    On Error GoTo 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
