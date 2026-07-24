Attribute VB_Name = "CashRepository"
Option Explicit

Private Const TBL_CASHIN As String = "tblCashIn"
Private Const TBL_CASHOUT As String = "tblCashOut"
Private Const TBL_JOURNAL As String = "tblJournal"

'=============================================================================
' Cash In
'=============================================================================

Public Function CreateCashIn(ByVal CashInID As String, _
                             ByVal ReceiptNo As String, _
                             ByVal TxnDate As Date, _
                             ByVal Project As String, _
                             ByVal Account As String, _
                             ByVal Description As String, _
                             ByVal Amount As Double) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lRow As Long

    On Error GoTo CreateCashIn_Err

    Set ws = ThisWorkbook.Worksheets(SHT_CASHIN)
    Set lo = ws.ListObjects(TBL_CASHIN)

    lRow = lo.ListRows.AddAlwaysNewRow.Position

    With lo.DataBodyRange
        .Cells(lRow, COL_CASHIN_ID + 1).Value = CashInID
        .Cells(lRow, COL_CASHIN_RECEIPT_NO + 1).Value = ReceiptNo
        .Cells(lRow, COL_CASHIN_DATE + 1).Value = TxnDate
        .Cells(lRow, COL_CASHIN_PROJECT + 1).Value = Project
        .Cells(lRow, COL_CASHIN_ACCOUNT + 1).Value = Account
        .Cells(lRow, COL_CASHIN_DESC + 1).Value = Description
        .Cells(lRow, COL_CASHIN_AMOUNT + 1).Value = Amount
        .Cells(lRow, COL_CASHIN_CREATED_BY + 1).Value = gCurrentUser
        .Cells(lRow, COL_CASHIN_CREATED_ON + 1).Value = Now()
    End With

    PostCashInJournal CashInID, TxnDate, Project, Account, Description, Amount

    CreateCashIn = True
    Exit Function

CreateCashIn_Err:
    LogError "CashRepository.CreateCashIn", Err.Description
    CreateCashIn = False
End Function

Public Function GetCashIn(ByVal RowNum As Long) As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo GetCashIn_Err

    Set ws = ThisWorkbook.Worksheets(SHT_CASHIN)
    Set lo = ws.ListObjects(TBL_CASHIN)

    If lo.DataBodyRange Is Nothing Then
        GetCashIn = Array()
        Exit Function
    End If

    GetCashIn = Application.Index(lo.DataBodyRange, RowNum, 0)
    Exit Function

GetCashIn_Err:
    LogError "CashRepository.GetCashIn", Err.Description
    GetCashIn = Array()
End Function

Public Function LoadCashIn() As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo LoadCashIn_Err

    Set ws = ThisWorkbook.Worksheets(SHT_CASHIN)
    Set lo = ws.ListObjects(TBL_CASHIN)

    If lo.DataBodyRange Is Nothing Then
        LoadCashIn = Array()
        Exit Function
    End If

    LoadCashIn = lo.DataBodyRange.Value
    Exit Function

LoadCashIn_Err:
    LogError "CashRepository.LoadCashIn", Err.Description
    LoadCashIn = Array()
End Function

Public Function GenerateReceiptNo() As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lCount As Long

    On Error GoTo GenerateReceiptNo_Err

    Set ws = ThisWorkbook.Worksheets(SHT_CASHIN)
    Set lo = ws.ListObjects(TBL_CASHIN)

    lCount = 1
    If Not lo.DataBodyRange Is Nothing Then
        lCount = lo.DataBodyRange.Rows.Count + 1
    End If

    GenerateReceiptNo = "RCV-" & Format$(lCount, "0000")
    Exit Function

GenerateReceiptNo_Err:
    LogError "CashRepository.GenerateReceiptNo", Err.Description
    GenerateReceiptNo = "RCV-0001"
End Function

'=============================================================================
' Cash Out
'=============================================================================

Public Function CreateCashOut(ByVal CashOutID As String, _
                              ByVal VoucherNo As String, _
                              ByVal TxnDate As Date, _
                              ByVal Project As String, _
                              ByVal Account As String, _
                              ByVal Vendor As String, _
                              ByVal Description As String, _
                              ByVal Amount As Double) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lRow As Long

    On Error GoTo CreateCashOut_Err

    Set ws = ThisWorkbook.Worksheets(SHT_CASHOUT)
    Set lo = ws.ListObjects(TBL_CASHOUT)

    lRow = lo.ListRows.AddAlwaysNewRow.Position

    With lo.DataBodyRange
        .Cells(lRow, COL_CASHOUT_ID + 1).Value = CashOutID
        .Cells(lRow, COL_CASHOUT_VOUCHER_NO + 1).Value = VoucherNo
        .Cells(lRow, COL_CASHOUT_DATE + 1).Value = TxnDate
        .Cells(lRow, COL_CASHOUT_PROJECT + 1).Value = Project
        .Cells(lRow, COL_CASHOUT_ACCOUNT + 1).Value = Account
        .Cells(lRow, COL_CASHOUT_VENDOR + 1).Value = Vendor
        .Cells(lRow, COL_CASHOUT_DESC + 1).Value = Description
        .Cells(lRow, COL_CASHOUT_AMOUNT + 1).Value = Amount
        .Cells(lRow, COL_CASHOUT_CREATED_BY + 1).Value = gCurrentUser
        .Cells(lRow, COL_CASHOUT_CREATED_ON + 1).Value = Now()
    End With

    PostCashOutJournal CashOutID, TxnDate, Project, Account, Vendor, Description, Amount

    CreateCashOut = True
    Exit Function

CreateCashOut_Err:
    LogError "CashRepository.CreateCashOut", Err.Description
    CreateCashOut = False
End Function

Public Function GetCashOut(ByVal RowNum As Long) As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo GetCashOut_Err

    Set ws = ThisWorkbook.Worksheets(SHT_CASHOUT)
    Set lo = ws.ListObjects(TBL_CASHOUT)

    If lo.DataBodyRange Is Nothing Then
        GetCashOut = Array()
        Exit Function
    End If

    GetCashOut = Application.Index(lo.DataBodyRange, RowNum, 0)
    Exit Function

GetCashOut_Err:
    LogError "CashRepository.GetCashOut", Err.Description
    GetCashOut = Array()
End Function

Public Function LoadCashOut() As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo LoadCashOut_Err

    Set ws = ThisWorkbook.Worksheets(SHT_CASHOUT)
    Set lo = ws.ListObjects(TBL_CASHOUT)

    If lo.DataBodyRange Is Nothing Then
        LoadCashOut = Array()
        Exit Function
    End If

    LoadCashOut = lo.DataBodyRange.Value
    Exit Function

LoadCashOut_Err:
    LogError "CashRepository.LoadCashOut", Err.Description
    LoadCashOut = Array()
End Function

Public Function GenerateVoucherNo() As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lCount As Long

    On Error GoTo GenerateVoucherNo_Err

    Set ws = ThisWorkbook.Worksheets(SHT_CASHOUT)
    Set lo = ws.ListObjects(TBL_CASHOUT)

    lCount = 1
    If Not lo.DataBodyRange Is Nothing Then
        lCount = lo.DataBodyRange.Rows.Count + 1
    End If

    GenerateVoucherNo = "VCH-" & Format$(lCount, "0000")
    Exit Function

GenerateVoucherNo_Err:
    LogError "CashRepository.GenerateVoucherNo", Err.Description
    GenerateVoucherNo = "VCH-0001"
End Function

'=============================================================================
' Journal Posting
'=============================================================================

Private Sub PostCashInJournal(ByVal CashInID As String, _
                              ByVal TxnDate As Date, _
                              ByVal Project As String, _
                              ByVal Account As String, _
                              ByVal Description As String, _
                              ByVal Amount As Double)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lRow As Long
    Dim sJournalID1 As String, sJournalID2 As String
    Dim sRef As String
    Dim sDesc As String

    On Error GoTo PostCashInJournal_Err

    Set ws = ThisWorkbook.Worksheets(SHT_JOURNAL)
    Set lo = ws.ListObjects(TBL_JOURNAL)

    sJournalID1 = GenerateGUID()
    sJournalID2 = GenerateGUID()
    sRef = CashInID
    sDesc = "Cash In: " & Description

    ' Debit entry - Cash/Bank account
    lRow = lo.ListRows.AddAlwaysNewRow.Position
    With lo.DataBodyRange
        .Cells(lRow, COL_JOURNAL_ID + 1).Value = sJournalID1
        .Cells(lRow, COL_JOURNAL_DATE + 1).Value = TxnDate
        .Cells(lRow, COL_JOURNAL_SOURCE + 1).Value = "Cash In"
        .Cells(lRow, COL_JOURNAL_REF_NO + 1).Value = sRef
        .Cells(lRow, COL_JOURNAL_PROJECT + 1).Value = Project
        .Cells(lRow, COL_JOURNAL_ACCOUNT + 1).Value = "Cash"
        .Cells(lRow, COL_JOURNAL_DESC + 1).Value = sDesc
        .Cells(lRow, COL_JOURNAL_DEBIT + 1).Value = Amount
        .Cells(lRow, COL_JOURNAL_CREDIT + 1).Value = 0
        .Cells(lRow, COL_JOURNAL_CREATED_BY + 1).Value = gCurrentUser
        .Cells(lRow, COL_JOURNAL_CREATED_ON + 1).Value = Now()
    End With

    ' Credit entry - Income/Account
    lRow = lo.ListRows.AddAlwaysNewRow.Position
    With lo.DataBodyRange
        .Cells(lRow, COL_JOURNAL_ID + 1).Value = sJournalID2
        .Cells(lRow, COL_JOURNAL_DATE + 1).Value = TxnDate
        .Cells(lRow, COL_JOURNAL_SOURCE + 1).Value = "Cash In"
        .Cells(lRow, COL_JOURNAL_REF_NO + 1).Value = sRef
        .Cells(lRow, COL_JOURNAL_PROJECT + 1).Value = Project
        .Cells(lRow, COL_JOURNAL_ACCOUNT + 1).Value = Account
        .Cells(lRow, COL_JOURNAL_DESC + 1).Value = sDesc
        .Cells(lRow, COL_JOURNAL_DEBIT + 1).Value = 0
        .Cells(lRow, COL_JOURNAL_CREDIT + 1).Value = Amount
        .Cells(lRow, COL_JOURNAL_CREATED_BY + 1).Value = gCurrentUser
        .Cells(lRow, COL_JOURNAL_CREATED_ON + 1).Value = Now()
    End With

    Exit Sub

PostCashInJournal_Err:
    LogError "CashRepository.PostCashInJournal", Err.Description
End Sub

Private Sub PostCashOutJournal(ByVal CashOutID As String, _
                               ByVal TxnDate As Date, _
                               ByVal Project As String, _
                               ByVal Account As String, _
                               ByVal Vendor As String, _
                               ByVal Description As String, _
                               ByVal Amount As Double)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lRow As Long
    Dim sJournalID1 As String, sJournalID2 As String
    Dim sRef As String
    Dim sDesc As String

    On Error GoTo PostCashOutJournal_Err

    Set ws = ThisWorkbook.Worksheets(SHT_JOURNAL)
    Set lo = ws.ListObjects(TBL_JOURNAL)

    sJournalID1 = GenerateGUID()
    sJournalID2 = GenerateGUID()
    sRef = CashOutID
    sDesc = "Cash Out: " & Description

    ' Debit entry - Expense/Account
    lRow = lo.ListRows.AddAlwaysNewRow.Position
    With lo.DataBodyRange
        .Cells(lRow, COL_JOURNAL_ID + 1).Value = sJournalID1
        .Cells(lRow, COL_JOURNAL_DATE + 1).Value = TxnDate
        .Cells(lRow, COL_JOURNAL_SOURCE + 1).Value = "Cash Out"
        .Cells(lRow, COL_JOURNAL_REF_NO + 1).Value = sRef
        .Cells(lRow, COL_JOURNAL_PROJECT + 1).Value = Project
        .Cells(lRow, COL_JOURNAL_ACCOUNT + 1).Value = Account
        .Cells(lRow, COL_JOURNAL_DESC + 1).Value = sDesc
        .Cells(lRow, COL_JOURNAL_DEBIT + 1).Value = Amount
        .Cells(lRow, COL_JOURNAL_CREDIT + 1).Value = 0
        .Cells(lRow, COL_JOURNAL_CREATED_BY + 1).Value = gCurrentUser
        .Cells(lRow, COL_JOURNAL_CREATED_ON + 1).Value = Now()
    End With

    ' Credit entry - Cash/Bank
    lRow = lo.ListRows.AddAlwaysNewRow.Position
    With lo.DataBodyRange
        .Cells(lRow, COL_JOURNAL_ID + 1).Value = sJournalID2
        .Cells(lRow, COL_JOURNAL_DATE + 1).Value = TxnDate
        .Cells(lRow, COL_JOURNAL_SOURCE + 1).Value = "Cash Out"
        .Cells(lRow, COL_JOURNAL_REF_NO + 1).Value = sRef
        .Cells(lRow, COL_JOURNAL_PROJECT + 1).Value = Project
        .Cells(lRow, COL_JOURNAL_ACCOUNT + 1).Value = "Cash"
        .Cells(lRow, COL_JOURNAL_DESC + 1).Value = sDesc
        .Cells(lRow, COL_JOURNAL_DEBIT + 1).Value = 0
        .Cells(lRow, COL_JOURNAL_CREDIT + 1).Value = Amount
        .Cells(lRow, COL_JOURNAL_CREATED_BY + 1).Value = gCurrentUser
        .Cells(lRow, COL_JOURNAL_CREATED_ON + 1).Value = Now()
    End With

    Exit Sub

PostCashOutJournal_Err:
    LogError "CashRepository.PostCashOutJournal", Err.Description
End Sub

'=============================================================================
' Load Journal
'=============================================================================

Public Function LoadJournal() As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo LoadJournal_Err

    Set ws = ThisWorkbook.Worksheets(SHT_JOURNAL)
    Set lo = ws.ListObjects(TBL_JOURNAL)

    If lo.DataBodyRange Is Nothing Then
        LoadJournal = Array()
        Exit Function
    End If

    LoadJournal = lo.DataBodyRange.Value
    Exit Function

LoadJournal_Err:
    LogError "CashRepository.LoadJournal", Err.Description
    LoadJournal = Array()
End Function

