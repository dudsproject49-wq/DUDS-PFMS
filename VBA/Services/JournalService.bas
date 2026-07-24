Attribute VB_Name = "JournalService"
Option Explicit

Public Function CreateJournal(ByVal JournalDate As Date, ByVal Source As String, ByVal Reference As String, ByVal Description As String, ByRef Lines() As Variant) As String
    Dim sHeaderID As String, sJournalNo As String, ws As Worksheet, lRow As Long, i As Long
    Dim dTotalDebit As Double, dTotalCredit As Double
    On Error GoTo CreateJournal_Err
    If Not ValidateJournalLines(Lines, dTotalDebit, dTotalCredit) Then CreateJournal = "": Exit Function
    sHeaderID = GenerateGUID(): sJournalNo = GenerateJournalNo()
    Set ws = ThisWorkbook.Worksheets(SHT_JOURNALHEADER)
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lRow, COL_JH_ID + 1).Value = sHeaderID
    ws.Cells(lRow, COL_JH_JOURNALNO + 1).Value = sJournalNo
    ws.Cells(lRow, COL_JH_DATE + 1).Value = JournalDate
    ws.Cells(lRow, COL_JH_SOURCE + 1).Value = Source
    ws.Cells(lRow, COL_JH_REF + 1).Value = Reference
    ws.Cells(lRow, COL_JH_DESC + 1).Value = Description
    ws.Cells(lRow, COL_JH_POSTED + 1).Value = JH_DRAFT
    ws.Cells(lRow, COL_JH_LOCKED + 1).Value = "No"
    ws.Cells(lRow, COL_JH_CREATED_BY + 1).Value = gCurrentUser
    ws.Cells(lRow, COL_JH_CREATED_ON + 1).Value = Now()
    Set ws = ThisWorkbook.Worksheets(SHT_JOURNALLINE)
    For i = LBound(Lines, 1) To UBound(Lines, 1)
        lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(lRow, COL_JL_ID + 1).Value = GenerateGUID()
        ws.Cells(lRow, COL_JL_HEADER_ID + 1).Value = sHeaderID
        ws.Cells(lRow, COL_JL_ACCOUNT_CODE + 1).Value = Lines(i, 0)
        ws.Cells(lRow, COL_JL_DESC + 1).Value = Lines(i, 1)
        ws.Cells(lRow, COL_JL_DEBIT + 1).Value = Lines(i, 2)
        ws.Cells(lRow, COL_JL_CREDIT + 1).Value = Lines(i, 3)
        ws.Cells(lRow, COL_JL_PROJECT + 1).Value = Lines(i, 4)
        ws.Cells(lRow, COL_JL_CREATED_BY + 1).Value = gCurrentUser
        ws.Cells(lRow, COL_JL_CREATED_ON + 1).Value = Now()
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
    Dim wsJH As Worksheet, wsJL As Worksheet, wsLedger As Worksheet
    Dim lLastJH As Long, lLastJL As Long, lLastL As Long, i As Long, j As Long
    Dim dDate As Date, sRef As String, sAcctCode As String, sDesc As String
    Dim dDebit As Double, dCredit As Double, sProject As String
    Dim sAcctName As String, dOpenBal As Double, dBal As Double, iFY As Integer
    Dim lFoundRow As Long
    On Error GoTo PostJournal_Err
    Set wsJH = ThisWorkbook.Worksheets(SHT_JOURNALHEADER)
    Set wsJL = ThisWorkbook.Worksheets(SHT_JOURNALLINE)
    Set wsLedger = ThisWorkbook.Worksheets(SHT_LEDGER)
    lLastJH = wsJH.Cells(wsJH.Rows.Count, 1).End(xlUp).Row
    lFoundRow = 0
    For i = 2 To lLastJH
        If wsJH.Cells(i, COL_JH_ID + 1).Value = HeaderID Then
            If wsJH.Cells(i, COL_JH_POSTED + 1).Value = JH_POSTED Then PostJournal = False: Exit Function
            If wsJH.Cells(i, COL_JH_LOCKED + 1).Value = "Yes" Then PostJournal = False: Exit Function
            dDate = CDate(wsJH.Cells(i, COL_JH_DATE + 1).Value)
            sRef = SafeConvertToString(wsJH.Cells(i, COL_JH_JOURNALNO + 1).Value)
            lFoundRow = i
            Exit For
        End If
    Next i
    If lFoundRow = 0 Then PostJournal = False: Exit Function
    iFY = FiscalYear(dDate)
    lLastJL = wsJL.Cells(wsJL.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lLastJL
        If wsJL.Cells(i, COL_JL_HEADER_ID + 1).Value = HeaderID Then
            sAcctCode = SafeConvertToString(wsJL.Cells(i, COL_JL_ACCOUNT_CODE + 1).Value)
            sDesc = SafeConvertToString(wsJL.Cells(i, COL_JL_DESC + 1).Value)
            dDebit = SafeConvertToDouble(wsJL.Cells(i, COL_JL_DEBIT + 1).Value)
            dCredit = SafeConvertToDouble(wsJL.Cells(i, COL_JL_CREDIT + 1).Value)
            sProject = SafeConvertToString(wsJL.Cells(i, COL_JL_PROJECT + 1).Value)
            sAcctName = GetAcctName(sAcctCode)
            dOpenBal = GetAcctBalance(sAcctCode)
            dBal = Round(dOpenBal + dDebit - dCredit, FIN_DECIMAL_PLACES)
            lLastL = wsLedger.Cells(wsLedger.Rows.Count, 1).End(xlUp).Row + 1
            wsLedger.Cells(lLastL, COL_LEDGER_ID + 1).Value = GenerateGUID()
            wsLedger.Cells(lLastL, COL_LEDGER_DATE + 1).Value = dDate
            wsLedger.Cells(lLastL, COL_LEDGER_ACCOUNT_CODE + 1).Value = sAcctCode
            wsLedger.Cells(lLastL, COL_LEDGER_ACCOUNT_NAME + 1).Value = sAcctName
            wsLedger.Cells(lLastL, COL_LEDGER_DESC + 1).Value = sDesc
            wsLedger.Cells(lLastL, COL_LEDGER_REF + 1).Value = sRef
            wsLedger.Cells(lLastL, COL_LEDGER_DEBIT + 1).Value = dDebit
            wsLedger.Cells(lLastL, COL_LEDGER_CREDIT + 1).Value = dCredit
            wsLedger.Cells(lLastL, COL_LEDGER_BALANCE + 1).Value = dBal
            wsLedger.Cells(lLastL, COL_LEDGER_FISCAL_YEAR + 1).Value = iFY
            wsLedger.Cells(lLastL, COL_LEDGER_CREATED_BY + 1).Value = gCurrentUser
            wsLedger.Cells(lLastL, COL_LEDGER_CREATED_ON + 1).Value = Now()
            UpdateAcctBalance sAcctCode, dDebit, dCredit
        End If
    Next i
    wsJH.Cells(lFoundRow, COL_JH_POSTED + 1).Value = JH_POSTED
    PostJournal = True
    LogInfo "JournalService.PostJournal", "Journal " & sRef & " posted", "Journal"
    Exit Function
PostJournal_Err: LogError "JournalService.PostJournal", Err.Description: PostJournal = False
End Function

Public Function GenerateJournalNo() As String
    Dim ws As Worksheet, lLast As Long, i As Long, lMax As Long
    Dim sPrefix As String, sNo As String, lNum As Long
    On Error GoTo GenerateJournalNo_Err
    sPrefix = "JRN-" & Format$(Now, "yyMM") & "-"
    Set ws = ThisWorkbook.Worksheets(SHT_JOURNALHEADER)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lMax = 0
    For i = 2 To lLast
        sNo = SafeConvertToString(ws.Cells(i, COL_JH_JOURNALNO + 1).Value)
        If Left$(sNo, Len(sPrefix)) = sPrefix Then
            lNum = SafeConvertToLong(Mid$(sNo, Len(sPrefix) + 1))
            If lNum > lMax Then lMax = lNum
        End If
    Next i
    GenerateJournalNo = sPrefix & Format$(lMax + 1, "0000")
    Exit Function
GenerateJournalNo_Err: LogError "JournalService.GenerateJournalNo", Err.Description: GenerateJournalNo = ""
End Function

Public Function ReverseJournal(ByVal HeaderID As String) As Boolean
    Dim wsJH As Worksheet, wsJL As Worksheet, lLastJH As Long, lLastJL As Long
    Dim i As Long, j As Long, n As Long, lFoundRow As Long
    Dim dDate As Date, sSource As String, sRef As String, sDesc As String
    Dim arrLines() As Variant, sNewHeaderID As String
    On Error GoTo ReverseJournal_Err
    Set wsJH = ThisWorkbook.Worksheets(SHT_JOURNALHEADER)
    Set wsJL = ThisWorkbook.Worksheets(SHT_JOURNALLINE)
    lLastJH = wsJH.Cells(wsJH.Rows.Count, 1).End(xlUp).Row
    lFoundRow = 0
    For i = 2 To lLastJH
        If wsJH.Cells(i, COL_JH_ID + 1).Value = HeaderID Then
            If wsJH.Cells(i, COL_JH_POSTED + 1).Value <> JH_POSTED Then ReverseJournal = False: Exit Function
            If wsJH.Cells(i, COL_JH_LOCKED + 1).Value = "Yes" Then ReverseJournal = False: Exit Function
            dDate = CDate(wsJH.Cells(i, COL_JH_DATE + 1).Value)
            sSource = "Reverse: " & SafeConvertToString(wsJH.Cells(i, COL_JH_SOURCE + 1).Value)
            sRef = SafeConvertToString(wsJH.Cells(i, COL_JH_REF + 1).Value)
            sDesc = "Reverse of " & SafeConvertToString(wsJH.Cells(i, COL_JH_DESC + 1).Value)
            wsJH.Cells(i, COL_JH_LOCKED + 1).Value = "Yes"
            lFoundRow = i
            Exit For
        End If
    Next i
    If lFoundRow = 0 Then ReverseJournal = False: Exit Function
    lLastJL = wsJL.Cells(wsJL.Rows.Count, 1).End(xlUp).Row
    n = 0
    For i = 2 To lLastJL
        If wsJL.Cells(i, COL_JL_HEADER_ID + 1).Value = HeaderID Then n = n + 1
    Next i
    If n = 0 Then ReverseJournal = False: Exit Function
    ReDim arrLines(1 To n, 0 To 4)
    j = 0
    For i = 2 To lLastJL
        If wsJL.Cells(i, COL_JL_HEADER_ID + 1).Value = HeaderID Then
            j = j + 1
            arrLines(j, 0) = wsJL.Cells(i, COL_JL_ACCOUNT_CODE + 1).Value
            arrLines(j, 1) = "Reverse: " & SafeConvertToString(wsJL.Cells(i, COL_JL_DESC + 1).Value)
            arrLines(j, 2) = wsJL.Cells(i, COL_JL_CREDIT + 1).Value
            arrLines(j, 3) = wsJL.Cells(i, COL_JL_DEBIT + 1).Value
            arrLines(j, 4) = wsJL.Cells(i, COL_JL_PROJECT + 1).Value
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

Private Function GetAcctName(ByVal AcctCode As String) As String
    Dim ws As Worksheet, lLast As Long, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lLast
        If ws.Cells(i, COL_ACCT_CODE + 1).Value = AcctCode Then GetAcctName = SafeConvertToString(ws.Cells(i, COL_ACCT_NAME + 1).Value): Exit Function
    Next i
    GetAcctName = ""
    On Error GoTo 0
End Function

Private Function GetAcctBalance(ByVal AcctCode As String) As Double
    Dim ws As Worksheet, lLast As Long, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lLast
        If ws.Cells(i, COL_ACCT_CODE + 1).Value = AcctCode Then GetAcctBalance = SafeConvertToDouble(ws.Cells(i, COL_ACCT_CURRENT_BAL + 1).Value): Exit Function
    Next i
    GetAcctBalance = 0
    On Error GoTo 0
End Function

Private Sub UpdateAcctBalance(ByVal AcctCode As String, ByVal dDebit As Double, ByVal dCredit As Double)
    Dim ws As Worksheet, lLast As Long, i As Long, dCurr As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_ACCOUNT)
    lLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lLast
        If ws.Cells(i, COL_ACCT_CODE + 1).Value = AcctCode Then
            dCurr = SafeConvertToDouble(ws.Cells(i, COL_ACCT_CURRENT_BAL + 1).Value)
            ws.Cells(i, COL_ACCT_CURRENT_BAL + 1).Value = Round(dCurr + dDebit - dCredit, FIN_DECIMAL_PLACES)
            Exit For
        End If
    Next i
    On Error GoTo 0
End Sub
