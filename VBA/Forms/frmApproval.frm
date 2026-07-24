Attribute VB_Name = "frmApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSelectedRow As Long
Private mSelectedRecordType As String
Private mSelectedRecordID As String

Private Sub UserForm_Initialize()
    Me.Caption = "DUDS-PFMS - Approval Workflow"
    Me.StartUpPosition = 0
    Me.Left = (Application.Width - Me.Width) / 2
    Me.Top = (Application.Height - Me.Height) / 2
    mSelectedRow = 0
    LoadApprovalList
End Sub

Private Sub LoadApprovalList()
    Dim arrPending As Variant, i As Long, vRec As Variant
    lvwApprovals.ListItems.Clear
    arrPending = GetPendingApprovals(gCurrentRole)
    If Not IsArray(arrPending) Then Exit Sub
    If UBound(arrPending) < LBound(arrPending) Then Exit Sub
    For i = LBound(arrPending) To UBound(arrPending)
        vRec = GetRecord(SHT_APPROVAL, arrPending(i))
        If Not IsEmpty(vRec(0)) Then
            Dim li As ListItem
            Set li = lvwApprovals.ListItems.Add()
            li.Tag = arrPending(i)
            li.SubItems(1) = SafeConvertToString(vRec(COL_APR_RECORDTYPE + 1))
            li.SubItems(2) = SafeConvertToString(vRec(COL_APR_RECORDID + 1))
            li.SubItems(3) = SafeConvertToString(vRec(COL_APR_LEVEL + 1))
            li.SubItems(4) = SafeConvertToString(vRec(COL_APR_STATUS + 1))
        End If
    Next i
    If lvwApprovals.ListItems.Count > 0 Then
        lvwApprovals.SelectedItem = lvwApprovals.ListItems(1)
    End If
End Sub

Private Sub lvwApprovals_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mSelectedRow = CLng(Item.Tag)
    Dim vRec As Variant: vRec = GetRecord(SHT_APPROVAL, mSelectedRow)
    If Not IsEmpty(vRec(0)) Then
        mSelectedRecordType = SafeConvertToString(vRec(COL_APR_RECORDTYPE + 1))
        mSelectedRecordID = SafeConvertToString(vRec(COL_APR_RECORDID + 1))
        lblRecordType.Caption = "Type: " & mSelectedRecordType
        lblRecordID.Caption = "ID: " & mSelectedRecordID
        lblLevel.Caption = "Level: " & SafeConvertToString(vRec(COL_APR_LEVEL + 1))
        LoadRecordDetail mSelectedRecordType, mSelectedRecordID
    End If
End Sub

Private Sub LoadRecordDetail(ByVal RecordType As String, ByVal RecordID As String)
    Dim lRow As Long, vRec As Variant
    lstDetail.Clear
    Select Case RecordType
        Case "Budget":
            lRow = FindRecord(SHT_BUDGETHEADER, COL_BH_ID, RecordID)
            If lRow > 0 Then
                vRec = GetRecord(SHT_BUDGETHEADER, lRow)
                lstDetail.AddItem "Budget No: " & SafeConvertToString(vRec(COL_BH_BUDGETNO + 1))
                lstDetail.AddItem "Description: " & SafeConvertToString(vRec(COL_BH_DESCRIPTION + 1))
                lstDetail.AddItem "Total Amount: " & FormatCurrencyIDR(SafeConvertToDouble(vRec(COL_BH_TOTALAMOUNT + 1)))
                lstDetail.AddItem "Fiscal Year: " & SafeConvertToString(vRec(COL_BH_FISCALYEAR + 1))
                lstDetail.AddItem "Status: " & SafeConvertToString(vRec(COL_BH_STATUS + 1))
            End If
        Case "Project":
            lRow = FindRecord(SHT_PROJECTS, COL_PROJ_ID, RecordID)
            If lRow > 0 Then
                vRec = GetRecord(SHT_PROJECTS, lRow)
                lstDetail.AddItem "Code: " & SafeConvertToString(vRec(COL_PROJ_CODE + 1))
                lstDetail.AddItem "Name: " & SafeConvertToString(vRec(COL_PROJ_NAME + 1))
                lstDetail.AddItem "Budget: " & FormatCurrencyIDR(SafeConvertToDouble(vRec(COL_PROJ_BUDGET + 1)))
                lstDetail.AddItem "Status: " & SafeConvertToString(vRec(COL_PROJ_STATUS + 1))
            End If
        Case Else:
            lstDetail.AddItem "Record details not available for type: " & RecordType
    End Select
End Sub

Private Sub cmdApprove_Click()
    If mSelectedRow = 0 Then MsgBox "Please select an approval record.", vbExclamation: Exit Sub
    Dim sComments As String: sComments = txtComments.Text
    If ApproveRecord(mSelectedRow, sComments) Then
        If IsFullyApproved(mSelectedRecordType, mSelectedRecordID) Then
            UpdateRecordStatus mSelectedRecordType, mSelectedRecordID, APR_APPROVED
            LockApprovedRecords mSelectedRecordType, mSelectedRecordID
        End If
        MsgBox "Record approved successfully.", vbInformation, APP_NAME
        txtComments.Text = ""
        LoadApprovalList
        lvwApprovals.SetFocus
    Else
        MsgBox "Failed to approve record.", vbExclamation, APP_NAME
    End If
End Sub

Private Sub cmdReject_Click()
    If mSelectedRow = 0 Then MsgBox "Please select an approval record.", vbExclamation: Exit Sub
    Dim sComments As String: sComments = txtComments.Text
    If Len(Trim$(sComments)) = 0 Then
        MsgBox "Please provide comments when rejecting.", vbExclamation: Exit Sub
    End If
    If RejectRecord(mSelectedRow, sComments) Then
        UpdateRecordStatus mSelectedRecordType, mSelectedRecordID, APR_REJECTED
        MsgBox "Record rejected.", vbInformation, APP_NAME
        txtComments.Text = ""
        LoadApprovalList
        lvwApprovals.SetFocus
    Else
        MsgBox "Failed to reject record.", vbExclamation, APP_NAME
    End If
End Sub

Private Sub cmdRefresh_Click()
    LoadApprovalList
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

