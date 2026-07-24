Attribute VB_Name = "JournalService"
Option Explicit

Private Const TBL_HEADER As String = "tblJournalHeader"
Private Const TBL_LINE As String = "tblJournalLine"
Private Const TBL_LEDGER As String = "tblLedger"

Public Function CreateJournal(ByVal JournalDate As Date, ByVal Source As String, ByVal Reference As String, ByVal Description As String, ByRef Lines() As Variant) As String
    Dim sHeaderID As String, sJournalNo As String, ws As Worksheet, lo As ListObject, lRow As Long, i As Long
    Dim wsLine As Worksheet, loLine As ListObject, dTotalDebit As Double, dTotalCredit As Double
    dTotalDebit = 0: dTotalCredit = 0
    On Error GoTo CreateJournal_Err
    If Not ValidateJournalLines(Lines, dTotalDebit, dTotalCredit) Then CreateJournal = "": Exit Function
    sHeaderID = GenerateGUID(): sJournalNo = GenerateJournalNo()
    Set ws = ThisWorkbook.Worksheets(SHT_JOURNALHEADER): Set lo = ws.ListObjects(TBL_HEADER)
    lRow = lo.ListRows.AddAlwaysNewRow.Position
    With lo.DataBodyRange
        .Cells(lRow, COL_JH_ID + 1).Value = sHeaderID
        .Cells(lRow, COL_JH_JOURNALNO + 1).Value = sJournalNo
        .Cells(lRow, COL_JH_DATE + 1).Value = JournalDate
        .Cells(lRow, COL_JH_SOURCE + 1).Value = Source
        .Cells(lRow, COL_JH_REF + 1).Value = Reference
        .Cells(lRow, COL_JH_DESC + 1).Value = Description
        .Cells(lRow, COL_JH_POSTED + 1).Value = JH_DRAFT
        .Cells(lRow, COL_JH_LOCKED + 1).Value = "No"
        .Cells(lRow, COL_JH_CREATED_BY + 1).Value = gCurrentUser
        .Cells(lRow, COL_JH_CREATED_ON + 1).Value = Now()
        .Cells(lRow, COL_JH_POSTED + 1).Value = JH_DRAFT
        .Cells(lRow, COL_JH_LOCKED + 1).Value = "No"
    End With
    Set wsLine = ThisWorkbook.Worksheets(SHT_JOURNALLINE): Set loLine = wsLine.ListObjects(TBL_LINE)
    For i = LBound(Lines, 1) To UBound(Lines, 1)
        lRow = loLine.ListRows.AddAlwaysNewRow.Position
        With loLine.DataBodyRange
            .Cells(lRow, COL_JL_ID + 1).Value = GenerateGUID()
            .Cells(lRow, COL_JL_HEADER_ID + 1).Value = sHeaderID
            .Cells(lRow, COL_JL_ACCOUNT_CODE + 1).Value = Lines(i, 0)
            .Cells(lRow, COL_JL_DESC + 1).Value = Lines(i, 1)
            .Cells(lRow, COL_JL_DEBIT + 1).Value = Lines(i, 2)
            .Cells(lRow, COL_JL_CREDIT + 1).Value = Lines(i, 3)
            .Cells(lRow, COL_JL_PROJECT + 1).Value = Lines(i, 4)
            .Cells(lRow, COL_JL_CREATED_BY + 1).Value = gCurrentUser
            .Cells(lRow, COL_JL_CREATED_ON + 1).Value = Now()
        End With
    Next i
    CreateJournal = sHeaderID
    LogInfo "JournalService.CreateJournal", "Journal " & sJournalNo & " created", "Journal"
    Exit Function
CreateJournal_Err: LogError "JournalService.CreateJournal", Err.Description: CreateJournal = ""
End Function

Private Function ValidateJournalLines(ByRef Lines() As Variant, ByRef TotalDebit As Double, ByRef TotalCredit As Double) As Boolean
    Dim i As Long
    If Not IsArray(Lines) Then ValidateJournalLines = False: Exit Function
    If UBound(Lines, 1) < 2 Then ValidateJournalLines = False: Exit Function
    TotalDebit = 0: TotalCredit = 0
    For i = LBound(Lines, 1) To UBound(Lines, 1)
        TotalDebit = TotalDebit + SafeConvertToDouble(Lines(i, 2))
        TotalCredit = TotalCredit + SafeConvertToDouble(Lines(i, 3))
    Next i
    If Round(TotalDebit, FIN_DECIMAL_PLACES) <> Round(TotalCredit, FIN_DECIMAL_PLACES) Then ValidateJournalLines = False: Exit Function
    If TotalDebit <= 0 Then ValidateJournalLines = False: Exit Function
    ValidateJournalLines = True
End Function

Public Function PostJournal(ByVal HeaderID As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, i As Long, lRow As Long, lLineRow As Long
    Dim wsLine As Worksheet, loLine As ListObject, wsL As Worksheet, loL As ListObject
    Dim sAccountCode As String, sAccountName As String, sDesc As String, sRef As String
    Dim dDebit As Double, dCredit As Double, sProject As String, dBal As Double
    Dim dDate As Date, dOpenBal As Double, iFY As Integer, bIsDebit As Boolean
    On Error GoTo PostJournal_Err
    Set ws = ThisWorkbook.Worksheets(SHT_JOURNALHEADER): Set lo = ws.ListObjects(TBL_HEADER)
    If lo.DataBodyRange Is Nothing Then PostJournal = False: Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_JH_ID + 1).Value = HeaderID Then
            If lo.DataBodyRange.Cells(i, COL_JH_POSTED + 1).Value = JH_POSTED Then PostJournal = False: Exit Function
            If lo.DataBodyRange.Cells(i, COL_JH_LOCKED + 1).Value = "Yes" Then PostJournal = False: Exit Function
            dDate = CDate(lo.DataBodyRange.Cells(i, COL_JH_DATE + 1).Value)
            sRef = SafeConvertToString(lo.DataBodyRange.Cells(i, COL_JH_JOURNALNO + 1).Value)
            Exit For
        End If
    Next i
    Set wsLine = ThisWorkbook.Worksheets(SHT_JOURNALLINE): Set loLine = wsLine.ListObjects(TBL_LINE)
    Set wsL = ThisWorkbook.Worksheets(SHT_LEDGER): Set loL = wsL.ListObjects(TBL_LEDGER)
    iFY = FiscalYear(dDate)
    For i = 1 To loLine.DataBodyRange.Rows.Count
        If loLine.DataBodyRange.Cells(i, COL_JL_HEADER_ID + 1).Value = HeaderID Then
            sAccountCode = SafeConvertToString(loLine.DataBodyRange.Cells(i, COL_JL_ACCOUNT_CODE + 1).Value)
            sDesc = SafeConvertToString(loLine.DataBodyRange.Cells(i, COL_JL_DESC + 1).Value)
            dDebit = SafeConvertToDouble(loLine.DataBodyRange.Cells(i, COL_JL_DEBIT + 1).Value)
            dCredit = SafeConvertToDouble(loLine.DataBodyRange.Cells(i, COL_JL_CREDIT + 1).Value)
            sProject = SafeConvertToString(loLine.DataBodyRange.Cells(i, COL_JL_PROJECT + 1).Value)
            sAccountName = GetAccountName(sAccountCode)
            dOpenBal = GetAccountBalance(sAccountCode)
            dBal = dOpenBal + dDebit - dCredit
            lRow = loL.ListRows.AddAlwaysNewRow.Position
            With loL.DataBodyRange
                .Cells(lRow, COL_LEDGER_ID + 1).Value = GenerateGUID()
                .Cells(lRow, COL_LEDGER_DATE + 1).Value = dDate
                .Cells(lRow, COL_LEDGER_ACCOUNT_CODE + 1).Value = sAccountCode
                .Cells(lRow, COL_LEDGER_ACCOUNT_NAME + 1).Value = sAccountName
                .Cells(lRow, COL_LEDGER_DESC + 1).Value = sDesc
                .Cells(lRow, COL_LEDGER_REF + 1).Value = sRef
                .Cells(lRow, COL_LEDGER_DEBIT + 1).Value = dDebit
                .Cells(lRow, COL_LEDGER_CREDIT + 1).Value = dCredit
                .Cells(lRow, COL_LEDGER_BALANCE + 1).Value = Round(dBal, FIN_DECIMAL_PLACES)
                .Cells(lRow, COL_LEDGER_FISCAL_YEAR + 1).Value = iFY
                .Cells(lRow, COL_LEDGER_CREATED_BY + 1).Value = gCurrentUser
                .Cells(lRow, COL_LEDGER_CREATED_ON + 1).Value = Now()
            End With
            UpdateAccountBalance sAccountCode, dDebit - dCredit, (dDebit > dCredit)
        End If
    Next i
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_JH_ID + 1).Value = HeaderID Then
            lo.DataBodyRange.Cells(i, COL_JH_POSTED + 1).Value = JH_POSTED
            Exit For
        End If
    Next i
    PostJournal = True
    LogInfo "JournalService.PostJournal", "Journal " & sRef & " posted", "Journal"
    Exit Function
PostJournal_Err: LogError "JournalService.PostJournal", Err.Description: PostJournal = False
End Function

Public Function ReverseJournal(ByVal HeaderID As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, i As Long, lLineRow As Long
    Dim wsLine As Worksheet, loLine As ListObject, sNewHeaderID As String
    Dim dDate As Date, sSource As String, sRef As String, sDesc As String
    Dim arrLines() As Variant, n As Long, j As Long
    On Error GoTo ReverseJournal_Err
    Set ws = ThisWorkbook.Worksheets(SHT_JOURNALHEADER): Set lo = ws.ListObjects(TBL_HEADER)
    If lo.DataBodyRange Is Nothing Then ReverseJournal = False: Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_JH_ID + 1).Value = HeaderID Then
            If lo.DataBodyRange.Cells(i, COL_JH_POSTED + 1).Value <> JH_POSTED Then ReverseJournal = False: Exit Function
            If lo.DataBodyRange.Cells(i, COL_JH_LOCKED + 1).Value = "Yes" Then ReverseJournal = False: Exit Function
            dDate = CDate(lo.DataBodyRange.Cells(i, COL_JH_DATE + 1).Value)
            sSource = "Reverse: " & SafeConvertToString(lo.DataBodyRange.Cells(i, COL_JH_SOURCE + 1).Value)
            sRef = SafeConvertToString(lo.DataBodyRange.Cells(i, COL_JH_REF + 1).Value)
            sDesc = "Reverse of " & SafeConvertToString(lo.DataBodyRange.Cells(i, COL_JH_DESC + 1).Value)
            lo.DataBodyRange.Cells(i, COL_JH_LOCKED + 1).Value = "Yes"
            Exit For
        End If
    Next i
    Set wsLine = ThisWorkbook.Worksheets(SHT_JOURNALLINE): Set loLine = wsLine.ListObjects(TBL_LINE)
    n = 0
    If Not loLine.DataBodyRange Is Nothing Then
        For i = 1 To loLine.DataBodyRange.Rows.Count
            If loLine.DataBodyRange.Cells(i, COL_JL_HEADER_ID + 1).Value = HeaderID Then n = n + 1
        Next i
    End If
    ReDim arrLines(1 To n, 0 To 4)
    j = 0
    For i = 1 To loLine.DataBodyRange.Rows.Count
        If loLine.DataBodyRange.Cells(i, COL_JL_HEADER_ID + 1).Value = HeaderID Then
            j = j + 1
            arrLines(j, 0) = loLine.DataBodyRange.Cells(i, COL_JL_ACCOUNT_CODE + 1).Value
            arrLines(j, 1) = "Reverse: " & SafeConvertToString(loLine.DataBodyRange.Cells(i, COL_JL_DESC + 1).Value)
            arrLines(j, 2) = loLine.DataBodyRange.Cells(i, COL_JL_CREDIT + 1).Value
            arrLines(j, 3) = loLine.DataBodyRange.Cells(i, COL_JL_DEBIT + 1).Value
            arrLines(j, 4) = loLine.DataBodyRange.Cells(i, COL_JL_PROJECT + 1).Value
        End If
    Next i
    sNewHeaderID = CreateJournal(dDate, sSource, sRef, sDesc, arrLines)
    If Len(sNewHeaderID) > 0 Then
        PostJournal sNewHeaderID
        ReverseJournal = True
    Else
        ReverseJournal = False
    End If
    Exit Function
ReverseJournal_Err: LogError "JournalService.ReverseJournal", Err.Description: ReverseJournal = False
End Function

Public Function LockJournal(ByVal HeaderID As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, i As Long
    On Error GoTo LockJournal_Err
    Set ws = ThisWorkbook.Worksheets(SHT_JOURNALHEADER): Set lo = ws.ListObjects(TBL_HEADER)
    If lo.DataBodyRange Is Nothing Then LockJournal = False: Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If lo.DataBodyRange.Cells(i, COL_JH_ID + 1).Value = HeaderID Then
            lo.DataBodyRange.Cells(i, COL_JH_LOCKED + 1).Value = "Yes"
            LockJournal = True: Exit Function
        End If
    Next i
    LockJournal = False
LockJournal_Err: LogError "JournalService.LockJournal", Err.Description: LockJournal = False
End Function

Public Function GenerateJournalNo() As String
    Dim ws As Worksheet, lo As ListObject, i As Long, lMax As Long
    Dim sPrefix As String
    On Error GoTo GenerateJournalNo_Err
    sPrefix = "JRN-" & Format$(Now, "yyMM") & "-"
    Set ws = ThisWorkbook.Worksheets(SHT_JOURNALHEADER): Set lo = ws.ListObjects(TBL_HEADER)
    lMax = 0
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.DataBodyRange.Rows.Count
            Dim sNo As String: sNo = SafeConvertToString(lo.DataBodyRange.Cells(i, COL_JH_JOURNALNO + 1).Value)
            If Left$(sNo, Len(sPrefix)) = sPrefix Then
                Dim lNum As Long: lNum = SafeConvertToLong(Mid$(sNo, Len(sPrefix) + 1))
                If lNum >
