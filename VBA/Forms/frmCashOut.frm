Attribute VB_Name = "frmCashOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRowIndex As Long
Private mCashOutData As Variant

Private Sub UserForm_Initialize()
    Me.Caption = "DUDS-PFMS - Cash Out"
    Me.StartUpPosition = 0
    Me.Left = (Application.Width - Me.Width) / 2
    Me.Top = (Application.Height - Me.Height) / 2
    ClearFields
    LoadProjectCombo
    LoadCashOutList
End Sub

Private Sub ClearFields()
    txtCashOutID.Value = ""
    txtVoucherNo.Value = GenerateVoucherNo()
    txtDate.Value = Format$(Now, "dd/mm/yyyy")
    cboProject.Value = ""
    txtAccount.Value = "Expense"
    txtVendor.Value = ""
    txtDescription.Value = ""
    txtAmount.Value = ""
    mRowIndex = 0
End Sub

Private Sub LoadProjectCombo()
    Dim arrData As Variant
    Dim i As Long

    On Error Resume Next
    cboProject.Clear
    arrData = LoadProjects()
    If IsArray(arrData) Then
        For i = LBound(arrData, 1) To UBound(arrData, 1)
            cboProject.AddItem SafeConvertToString(arrData(i, 2))
        Next i
    End If
    On Error GoTo 0
End Sub

Private Sub LoadCashOutList(Optional ByVal Filter As String = "")
    Dim i As Long
    Dim sDisplay As String

    lstCashOut.Clear
    mCashOutData = LoadCashOut()

    If IsArray(mCashOutData) Then
        If UBound(mCashOutData, 1) = 0 And IsEmpty(mCashOutData(1, 1)) Then Exit Sub
    Else
        Exit Sub
    End If

    For i = LBound(mCashOutData, 1) To UBound(mCashOutData, 1)
        sDisplay = SafeConvertToString(mCashOutData(i, 2)) & " | " & _
                   Format$(mCashOutData(i, 3), "dd/mm/yyyy") & " | " & _
                   SafeConvertToString(mCashOutData(i, 4)) & " | " & _
                   SafeConvertToString(mCashOutData(i, 6)) & " | " & _
                   FormatCurrencyIDR(SafeConvertToDouble(mCashOutData(i, 8)))
        If Len(Filter) > 0 Then
            If InStr(1, SafeConvertToString(mCashOutData(i, 2)), Filter, vbTextCompare) > 0 Then
                lstCashOut.AddItem sDisplay
            End If
        Else
            lstCashOut.AddItem sDisplay
        End If
    Next i
End Sub

Private Sub PopulateForm(ByVal RowIndex As Long)
    If Not IsArray(mCashOutData) Then Exit Sub
    If RowIndex < 1 Or RowIndex > UBound(mCashOutData, 1) Then Exit Sub

    mRowIndex = RowIndex
    txtCashOutID.Value = SafeConvertToString(mCashOutData(RowIndex, 1))
    txtVoucherNo.Value = SafeConvertToString(mCashOutData(RowIndex, 2))
    txtDate.Value = Format$(mCashOutData(RowIndex, 3), "dd/mm/yyyy")
    cboProject.Value = SafeConvertToString(mCashOutData(RowIndex, 4))
    txtAccount.Value = SafeConvertToString(mCashOutData(RowIndex, 5))
    txtVendor.Value = SafeConvertToString(mCashOutData(RowIndex, 6))
    txtDescription.Value = SafeConvertToString(mCashOutData(RowIndex, 7))
    txtAmount.Value = SafeConvertToString(mCashOutData(RowIndex, 8))
End Sub

Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Len(Trim$(txtVoucherNo.Value)) = 0 Then
        MsgBox "Voucher No is required.", vbExclamation, APP_NAME
        txtVoucherNo.SetFocus
        Exit Function
    End If

    If Not IsDate(txtDate.Value) Then
        MsgBox "Invalid date format.", vbExclamation, APP_NAME
        txtDate.SetFocus
        Exit Function
    End If

    If Len(Trim$(cboProject.Value)) = 0 Then
        MsgBox "Please select a project.", vbExclamation, APP_NAME
        cboProject.SetFocus
        Exit Function
    End If

    If Len(Trim$(txtDescription.Value)) = 0 Then
        MsgBox "Description is required.", vbExclamation, APP_NAME
        txtDescription.SetFocus
        Exit Function
    End If

    If Len(Trim$(txtAmount.Value)) = 0 Then
        MsgBox "Amount is required.", vbExclamation, APP_NAME
        txtAmount.SetFocus
        Exit Function
    End If

    If Not IsNumeric(txtAmount.Value) Then
        MsgBox "Amount must be numeric.", vbExclamation, APP_NAME
        txtAmount.SetFocus
        Exit Function
    End If

    If CDbl(txtAmount.Value) <= 0 Then
        MsgBox "Amount must be greater than zero.", vbExclamation, APP_NAME
        txtAmount.SetFocus
        Exit Function
    End If

    ValidateForm = True
End Function

Private Sub cmdSave_Click()
    Dim sID As String
    Dim dDate As Date
    Dim dAmount As Double

    If Not ValidateForm Then Exit Sub

    If Len(Trim$(txtCashOutID.Value)) > 0 Then
        sID = Trim$(txtCashOutID.Value)
    Else
        sID = GenerateGUID()
    End If

    dDate = CDate(txtDate.Value)
    dAmount = CDbl(txtAmount.Value)

    If CreateCashOut(sID, Trim$(txtVoucherNo.Value), dDate, _
                     Trim$(cboProject.Value), Trim$(txtAccount.Value), _
                     Trim$(txtVendor.Value), Trim$(txtDescription.Value), dAmount) Then
        MsgBox MSG_SAVE_SUCCESS, vbInformation, APP_NAME
        ClearFields
        LoadCashOutList
    Else
        MsgBox MSG_SAVE_FAIL, vbExclamation, APP_NAME
    End If
End Sub

Private Sub cmdCancel_Click()
    ClearFields
End Sub

Private Sub cmdRefresh_Click()
    ClearFields
    LoadCashOutList
End Sub

Private Sub lstCashOut_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstCashOut.ListIndex >= 0 Then
        PopulateForm lstCashOut.ListIndex + 1
    End If
End Sub

