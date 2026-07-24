Attribute VB_Name = "frmCashIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRowIndex As Long
Private mCashInData As Variant

Private Sub UserForm_Initialize()
    Me.Caption = "DUDS-PFMS - Cash In"
    Me.StartUpPosition = 0
    Me.Left = (Application.Width - Me.Width) / 2
    Me.Top = (Application.Height - Me.Height) / 2
    ClearFields
    LoadProjectCombo
    LoadCashInList
End Sub

Private Sub ClearFields()
    txtCashInID.Value = ""
    txtReceiptNo.Value = GenerateReceiptNo()
    txtDate.Value = Format$(Now, "dd/mm/yyyy")
    cboProject.Value = ""
    txtAccount.Value = "Cash"
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

Private Sub LoadCashInList(Optional ByVal Filter As String = "")
    Dim i As Long
    Dim sDisplay As String

    lstCashIn.Clear
    mCashInData = LoadCashIn()

    If IsArray(mCashInData) Then
        If UBound(mCashInData, 1) = 0 And IsEmpty(mCashInData(1, 1)) Then Exit Sub
    Else
        Exit Sub
    End If

    For i = LBound(mCashInData, 1) To UBound(mCashInData, 1)
        sDisplay = SafeConvertToString(mCashInData(i, 2)) & " | " & _
                   Format$(mCashInData(i, 3), "dd/mm/yyyy") & " | " & _
                   SafeConvertToString(mCashInData(i, 4)) & " | " & _
                   FormatCurrencyIDR(SafeConvertToDouble(mCashInData(i, 7)))
        If Len(Filter) > 0 Then
            If InStr(1, SafeConvertToString(mCashInData(i, 2)), Filter, vbTextCompare) > 0 Then
                lstCashIn.AddItem sDisplay
            End If
        Else
            lstCashIn.AddItem sDisplay
        End If
    Next i
End Sub

Private Sub PopulateForm(ByVal RowIndex As Long)
    If Not IsArray(mCashInData) Then Exit Sub
    If RowIndex < 1 Or RowIndex > UBound(mCashInData, 1) Then Exit Sub

    mRowIndex = RowIndex
    txtCashInID.Value = SafeConvertToString(mCashInData(RowIndex, 1))
    txtReceiptNo.Value = SafeConvertToString(mCashInData(RowIndex, 2))
    txtDate.Value = Format$(mCashInData(RowIndex, 3), "dd/mm/yyyy")
    cboProject.Value = SafeConvertToString(mCashInData(RowIndex, 4))
    txtAccount.Value = SafeConvertToString(mCashInData(RowIndex, 5))
    txtDescription.Value = SafeConvertToString(mCashInData(RowIndex, 6))
    txtAmount.Value = SafeConvertToString(mCashInData(RowIndex, 7))
End Sub

Private Function ValidateForm() As Boolean
    ValidateForm = False

    If Len(Trim$(txtReceiptNo.Value)) = 0 Then
        MsgBox "Receipt No is required.", vbExclamation, APP_NAME
        txtReceiptNo.SetFocus
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

    If Len(Trim$(txtCashInID.Value)) > 0 Then
        sID = Trim$(txtCashInID.Value)
    Else
        sID = GenerateGUID()
    End If

    dDate = CDate(txtDate.Value)
    dAmount = CDbl(txtAmount.Value)

    If CreateCashIn(sID, Trim$(txtReceiptNo.Value), dDate, _
                    Trim$(cboProject.Value), Trim$(txtAccount.Value), _
                    Trim$(txtDescription.Value), dAmount) Then
        MsgBox MSG_SAVE_SUCCESS, vbInformation, APP_NAME
        ClearFields
        LoadCashInList
    Else
        MsgBox MSG_SAVE_FAIL, vbExclamation, APP_NAME
    End If
End Sub

Private Sub cmdCancel_Click()
    ClearFields
End Sub

Private Sub cmdRefresh_Click()
    ClearFields
    LoadCashInList
End Sub

Private Sub lstCashIn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstCashIn.ListIndex >= 0 Then
        PopulateForm lstCashIn.ListIndex + 1
    End If
End Sub

