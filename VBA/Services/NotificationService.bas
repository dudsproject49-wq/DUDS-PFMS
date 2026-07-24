Attribute VB_Name = "NotificationService"
Option Explicit

Private Type tNotification
    sType As String
    sMessage As String
    sSeverity As String
End Type

Private mNotifications() As tNotification
Private mNotifCount As Long

Public Function GetNotificationCount() As Long
    GetNotificationCount = mNotifCount
End Function

Public Function GetNotificationType(ByVal Index As Long) As String
    If Index >= 0 And Index < mNotifCount Then GetNotificationType = mNotifications(Index).sType
End Function

Public Function GetNotificationMessage(ByVal Index As Long) As String
    If Index >= 0 And Index < mNotifCount Then GetNotificationMessage = mNotifications(Index).sMessage
End Function

Public Function GetNotificationSeverity(ByVal Index As Long) As String
    If Index >= 0 And Index < mNotifCount Then GetNotificationSeverity = mNotifications(Index).sSeverity
End Function

Public Sub RefreshNotifications()
    mNotifCount = 0
    ReDim mNotifications(0 To 99)
    
    CheckBudgetExceeds
    CheckBudgetExceeded
    CheckCashBalance
    CheckOverdueInvoices
    CheckProjectDeadline
    CheckPendingApprovals
End Sub

Private Sub AddNotif(ByVal sType As String, ByVal sMessage As String, ByVal sSeverity As String)
    If mNotifCount >= UBound(mNotifications) Then
        ReDim Preserve mNotifications(0 To mNotifCount + 50)
    End If
    mNotifications(mNotifCount).sType = sType
    mNotifications(mNotifCount).sMessage = sMessage
    mNotifications(mNotifCount).sSeverity = sSeverity
    mNotifCount = mNotifCount + 1
End Sub

Private Sub CheckBudgetExceeds()
    Dim ws As Worksheet, lo As ListObject, i As Long
    Dim dActual As Double, dPlanned As Double, dPct As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_BUDGET): Set lo = ws.ListObjects("tblBudget")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.DataBodyRange.Rows.Count
        dPlanned = SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_BUDGET_PLANNED + 1).Value)
        dActual = SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_BUDGET_ACTUAL + 1).Value)
        If dPlanned > 0 Then
            dPct = dActual / dPlanned
            If dPct >= NOTIF_BUDGET_PCT And dPct < 1 Then
                AddNotif "Budget Warning", _
                    "Budget for " & SafeConvertToString(lo.DataBodyRange.Cells(i, COL_BUDGET_CAT + 1).Value) & _
                    " is at " & Format$(dPct * 100, "0.0") & "%", "Warning"
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub CheckBudgetExceeded()
    Dim ws As Worksheet, lo As ListObject, i As Long
    Dim dActual As Double, dPlanned As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_BUDGET): Set lo = ws.ListObjects("tblBudget")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.DataBodyRange.Rows.Count
        dPlanned = SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_BUDGET_PLANNED + 1).Value)
        dActual = SafeConvertToDouble(lo.DataBodyRange.Cells(i, COL_BUDGET_ACTUAL + 1).Value)
        If dPlanned > 0 And dActual > dPlanned Then
            AddNotif "Budget Exceeded", _
                SafeConvertToString(lo.DataBodyRange.Cells(i, COL_BUDGET_CAT + 1).Value) & _
                " has exceeded budget by " & FormatCurrencyIDR(dActual - dPlanned), "Critical"
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub CheckCashBalance()
    Dim dCash As Double: dCash = GetCashPosition()
    If dCash < NOTIF_CASH_LOW Then
        AddNotif "Low Cash Balance", _
            "Cash balance is " & FormatCurrencyIDR(dCash) & _
            ". Below threshold of " & FormatCurrencyIDR(NOTIF_CASH_LOW), "Critical"
    End If
End Sub

Private Sub CheckOverdueInvoices()
    Dim ws As Worksheet, lo As ListObject, i As Long
    Dim dDueDate As Date, dAmount As Double, sVendor As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_CASHOUT): Set lo = ws.ListObjects("tblCashOut")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.DataBodyRange.Rows.Count
        dDueDate = CDate(lo.DataBodyRange.Cells(i, 3).Value)
        dAmount = SafeConvertToDouble(lo.DataBodyRange.Cells(i, 8).Value)
        sVendor = SafeConvertToString(lo.DataBodyRange.Cells(i, 6).Value)
        If dDueDate < Date And dAmount > 0 Then
            AddNotif "Invoice Overdue", _
                "Invoice from " & sVendor & " for " & FormatCurrencyIDR(dAmount) & _
                " was due on " & Format$(dDueDate, "dd/mm/yyyy"), "Warning"
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub CheckProjectDeadline()
    Dim ws As Worksheet, lo As ListObject, i As Long
    Dim dEndDate As Date, sName As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS): Set lo = ws.ListObjects("tblProject")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.DataBodyRange.Rows.Count
        dEndDate = CDate(lo.DataBodyRange.Cells(i, COL_PROJ_END_DATE + 1).Value)
        sName = SafeConvertToString(lo.DataBodyRange.Cells(i, COL_PROJ_NAME + 1).Value)
        If dEndDate >= Date And DateDiff("d", Date, dEndDate) <= NOTIF_DEADLINE_DAYS Then
            AddNotif "Deadline Approaching", _
                "Project '" & sName & "' deadline is " & Format$(dEndDate, "dd/mm/yyyy") & _
                " (" & DateDiff("d", Date, dEndDate) & " days remaining)", "Info"
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub CheckPendingApprovals()
    Dim arrPending As Variant, i As Long
    arrPending = GetPendingApprovals(gCurrentRole)
    If IsArray(arrPending) Then
        If UBound(arrPending) >= LBound(arrPending) Then
            AddNotif "Approval Pending", _
                "You have " & UBound(arrPending) - LBound(arrPending) + 1 & _
                " pending approval(s) waiting for action.", "Info"
        End If
    End If
End Sub

Public Sub ShowNotificationSummary()
    Dim sMsg As String, i As Long
    RefreshNotifications
    If mNotifCount = 0 Then
        MsgBox "No notifications at this time.", vbInformation, APP_NAME
        Exit Sub
    End If
    sMsg = "NOTIFICATIONS (" & mNotifCount & "):" & vbCrLf & vbCrLf
    For i = 0 To mNotifCount - 1
        sMsg = sMsg & "[" & mNotifications(i).sSeverity & "] " & mNotifications(i).sType & vbCrLf
        sMsg = sMsg & mNotifications(i).sMessage & vbCrLf & vbCrLf
    Next i
    MsgBox sMsg, vbInformation, APP_NAME
End Sub

Public Sub WriteNotificationsToSheet(ByVal TargetSheet As String)
    Dim ws As Worksheet, i As Long, lRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TargetSheet)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = TargetSheet
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "Notifications - " & Format$(Now, "dd/mm/yyyy hh:mm")
    ws.Range("A1:C1").Font.Bold = True
    ws.Cells(3, 1).Value = "Severity": ws.Cells(3, 2).Value = "Type"
    ws.Cells(3, 3).Value = "Message"
    ws.Range("A3:C3").Font.Bold = True: ws.Range("A3:C3").Interior.Color = CLR_HEADER
    RefreshNotifications
    lRow = 4
    For i = 0 To mNotifCount - 1
        ws.Cells(lRow, 1).Value = mNotifications(i).sSeverity
        ws.Cells(lRow, 2).Value = mNotifications(i).sType
        ws.Cells(lRow, 3).Value = mNotifications(i).sMessage
        lRow = lRow + 1
    Next i
    ws.Columns("A:C").AutoFit
End Sub

