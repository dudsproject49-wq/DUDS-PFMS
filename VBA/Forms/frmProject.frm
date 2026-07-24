Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Project Management Form"
Option Explicit

Private mEditMode As Boolean
Private mCurrentProjectID As String
Private mProjectData As Variant
Private mRowIndex As Long

Private Sub UserForm_Initialize()
    Me.Caption = "DUDS-PFMS - Project Management"
    Me.StartUpPosition = 0
    Me.Left = (Application.Width - Me.Width) / 2
    Me.Top = (Application.Height - Me.Height) / 2
    Me.BackColor = RGB(240, 240, 240)
    FormMode False
    ClearFields
    PopulateStatusCombo
    PopulateProgressCombo
    LoadProjectList
End Sub

Private Sub FormMode(ByVal EditMode As Boolean)
    Dim ctrl As Control
    mEditMode = EditMode
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Then
            ctrl.Enabled = EditMode
        End If
    Next ctrl
    txtProjectCode.Enabled = EditMode
    cmdSave.Enabled = EditMode
    cmdCancel.Enabled = EditMode
    cmdAdd.Enabled = Not EditMode
    cmdEdit.Enabled = Not EditMode
    cmdDelete.Enabled = Not EditMode
    cmdSearch.Enabled = Not EditMode
End Sub

Private Sub ClearFields()
    txtProjectID.Value = ""
    txtProjectCode.Value = GenerateProjectCode()
    txtProjectName.Value = ""
    txtClientName.Value = ""
    txtLocation.Value = ""
    txtStartDate.Value = ""
    txtEndDate.Value = ""
    txtContractValue.Value = ""
    txtBudgetValue.Value = ""
    txtNotes.Value = ""
    cboStatus.Value = STATUS_DRAFT
    cboProgress.Value = "0"
    mCurrentProjectID = ""
    mRowIndex = 0
End Sub

Private Sub PopulateStatusCombo()
    cboStatus.Clear
    cboStatus.AddItem STATUS_DRAFT
    cboStatus.AddItem STATUS_ACTIVE
    cboStatus.AddItem STATUS_PENDING
    cboStatus.AddItem STATUS_APPROVED
    cboStatus.AddItem STATUS_COMPLETED
    cboStatus.AddItem STATUS_CANCELLED
    cboStatus.Value = STATUS_DRAFT
End Sub

Private Sub PopulateProgressCombo()
    Dim i As Long
    cboProgress.Clear
    For i = 0 To 100 Step 10
        cboProgress.AddItem CStr(i)
    Next i
    cboProgress.Value = "0"
End Sub

Private Sub LoadProjectList(Optional ByVal Filter As String = "")
    Dim i As Long
    Dim j As Long
    Dim sDisplay As String

    lstProjects.Clear
    mProjectData = LoadProjects()

    If IsArray(mProjectData) Then
        If UBound(mProjectData, 1) = 0 And IsEmpty(mProjectData(1, 1)) Then Exit Sub
    Else
        Exit Sub
    End If

    For i = LBound(mProjectData, 1) To UBound(mProjectData, 1)
        sDisplay = SafeConvertToString(mProjectData(i, 2)) & " | " & _
                   SafeConvertToString(mProjectData(i, 3))
        If Len(Filter) > 0 Then
            If InStr(1, SafeConvertToString(mProjectData(i, 2)), Filter, vbTextCompare) > 0 Or _
               InStr(1, SafeConvertToString(mProjectData(i, 3)), Filter, vbTextCompare) > 0 Then
                lstProjects.AddItem sDisplay
            End If
        Else
            lstProjects.AddItem sDisplay
        End If
    Next i
End Sub

Private Sub PopulateForm(ByVal RowIndex As Long)
    If Not IsArray(mProjectData) Then Exit Sub
    If RowIndex < 1 Or RowIndex > UBound(mProjectData, 1) Then Exit Sub

    FormMode True
    mRowIndex = RowIndex
    mCurrentProjectID = SafeConvertToString(mProjectData(RowIndex, 1))
    txtProjectID.Value = mCurrentProjectID
    txtProjectCode.Value = SafeConvertToString(mProjectData(RowIndex, 2))
    txtProjectName.Value = SafeConvertToString(mProjectData(RowIndex, 3))
    txtClientName.Value = SafeConvertToString(mProjectData(RowIndex, 4))
    txtLocation.Value = SafeConvertToString(mProjectData(RowIndex, 5))
    txtStartDate.Value = SafeConvertToString(mProjectData(RowIndex, 6))
    txtEndDate.Value = SafeConvertToString(mProjectData(RowIndex, 7))
    txtContractValue.Value = SafeConvertToString(mProjectData(RowIndex, 8))
    txtBudgetValue.Value = SafeConvertToString(mProjectData(RowIndex, 9))
    cboStatus.Value = SafeConvertToString(mProjectData(RowIndex, 10))
    cboProgress.Value = SafeConvertToString(mProjectData(RowIndex, 11))
    txtNotes.Value = SafeConvertToString(mProjectData(RowIndex, 12))
End Sub

Private Sub cmdAdd_Click()
    FormMode True
    ClearFields
    txtProjectCode.Value = GenerateProjectCode()
    txtProjectCode.SetFocus
End Sub

Private Sub cmdEdit_Click()
    If lstProjects.ListIndex < 0 Then
        MsgBox "Please select a project from the list.", vbExclamation, APP_NAME
        Exit Sub
    End If
    PopulateForm lstProjects.ListIndex + 1
End Sub

Private Sub cmdDelete_Click()
    Dim lIdx As Long
    Dim sID As String

    If lstProjects.ListIndex < 0 Then
        MsgBox "Please select a project from the list.", vbExclamation, APP_NAME
        Exit Sub
    End If

    lIdx = lstProjects.ListIndex + 1
    If lIdx < 1 Or lIdx > UBound(mProjectData, 1) Then Exit Sub
    sID = SafeConvertToString(mProjectData(lIdx, 1))
    If Len(sID) = 0 Then Exit Sub

    If MsgBox(MSG_DELETE_CONFIRM, vbYesNo + vbQuestion, APP_NAME) = vbNo Then Exit Sub

    If DeleteProject(sID) Then
        MsgBox "Project deleted successfully.", vbInformation, APP_NAME
        ClearFields
        FormMode False
        LoadProjectList
    Else
        MsgBox "Failed to delete project.", vbExclamation, APP_NAME
    End If
End Sub

Private Sub cmdSave_Click()
    Dim sProjectID As String
    Dim dStart As Date, dEnd As Date
    Dim dContract As Double, dBudget As Double
    Dim dProgress As Double

    If Not ValidateForm Then Exit Sub

    If Len(mCurrentProjectID) > 0 Then
        sProjectID = mCurrentProjectID
    Else
        sProjectID = GenerateGUID()
    End If

    If IsDate(txtStartDate.Value) Then dStart = CDate(txtStartDate.Value)
    If IsDate(txtEndDate.Value) Then dEnd = CDate(txtEndDate.Value)
    If IsNumeric(txtContractValue.Value) Then dContract = CDbl(txtContractValue.Value)
    If IsNumeric(txtBudgetValue.Value) Then dBudget = CDbl(txtBudgetValue.Value)
    If IsNumeric(cboProgress.Value) Then dProgress = CDbl(cboProgress.Value)

    If Len(mCurrentProjectID) > 0 Then
        If UpdateProject(sProjectID, txtProjectCode.Value, txtProjectName.Value, _
                         txtClientName.Value, txtLocation.Value, dStart, dEnd, _
                         dContract, dBudget, cboStatus.Value, dProgress, txtNotes.Value) Then
            MsgBox MSG_SAVE_SUCCESS, vbInformation, APP_NAME
        Else
            MsgBox MSG_SAVE_FAIL, vbExclamation, APP_NAME
            Exit Sub
        End If
    Else
        If CreateProject(sProjectID, txtProjectCode.Value, txtProjectName.Value, _
                         txtClientName.Value, txtLocation.Value, dStart, dEnd, _
                         dContract, dBudget, cboStatus.Value, dProgress, txtNotes.Value) Then
            MsgBox MSG_SAVE_SUCCESS, vbInformation, APP_NAME
        Else
            MsgBox MSG_SAVE_FAIL, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If

    ClearFields
    FormMode False
    LoadProjectList
End Sub

Private Sub cmdCancel_Click()
    FormMode False
    ClearFields
End Sub

Private Sub cmdSearch_Click()
    Dim sFilter As String
    sFilter = InputBox("Enter Project Code or Name to search:", "Search Projects")
    If Len(sFilter) > 0 Then
        LoadProjectList sFilter
    Else
        LoadProjectList
    End If
End Sub

Private Sub cmdRefresh_Click()
    ClearFields
    FormMode False
    LoadProjectList
End Sub

Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Len(Trim$(txtProjectCode.Value)) = 0 Then
        MsgBox "Project Code is required.", vbExclamation, APP_NAME
        txtProjectCode.SetFocus
        Exit Function
    End If

    If Len(Trim$(txtProjectName.Value)) = 0 Then
        MsgBox "Project Name is required.", vbExclamation, APP_NAME
        txtProjectName.SetFocus
        Exit Function
    End If

    If Len(Trim$(txtContractValue.Value)) > 0 Then
        If Not IsNumeric(txtContractValue.Value) Then
            MsgBox "Contract Value must be numeric.", vbExclamation, APP_NAME
            txtContractValue.SetFocus
            Exit Function
        End If
    End If

    If Len(Trim$(txtBudgetValue.Value)) > 0 Then
        If Not IsNumeric(txtBudgetValue.Value) Then
            MsgBox "Budget Value must be numeric.", vbExclamation, APP_NAME
            txtBudgetValue.SetFocus
            Exit Function
        End If
    End If

    If Len(Trim$(txtStartDate.Value)) > 0 Then
        If Not IsDate(txtStartDate.Value) Then
            MsgBox "Start Date is not a valid date.", vbExclamation, APP_NAME
            txtStartDate.SetFocus
            Exit Function
        End If
    End If

    If Len(Trim$(txtEndDate.Value)) > 0 Then
        If Not IsDate(txtEndDate.Value) Then
            MsgBox "End Date is not a valid date.", vbExclamation, APP_NAME
            txtEndDate.SetFocus
            Exit Function
        End If
    End If

    ValidateForm = True
End Function

Private Sub lstProjects_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstProjects.ListIndex >= 0 Then
        PopulateForm lstProjects.ListIndex + 1
    End If
End Sub

