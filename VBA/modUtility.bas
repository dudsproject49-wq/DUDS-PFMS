Attribute VB_Name = "modUtility"
Option Explicit

'=============================================================================
' Module      : modUtility
' Project     : DUDS-PFMS (Project Financial Management System)
' Description : Centralized error handling, logging, and general-purpose
'               utility functions used across the application.
'=============================================================================

'------------------------------------------------------------------------------
' Error Handling
'------------------------------------------------------------------------------

Public Sub HandleError(ByVal ProcName As String, ByVal ErrNum As Long, _
                       ByVal ErrDesc As String, Optional ByVal ErrLine As Long = 0)
    '-----------------------------------------------------------------------
    ' Central error handler. Logs the error to _Log sheet and displays a
    ' user-friendly message. Called from error-handling blocks throughout
    ' the application.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.HandleError"
    
    Dim wsLog As Worksheet
    Dim lNextRow As Long
    
    On Error Resume Next
    
    ' Try to log to the _Log worksheet
    Set wsLog = GetLogSheet()
    If Not wsLog Is Nothing Then
        lNextRow = NextEmptyRow(wsLog, 1)
        
        wsLog.Cells(lNextRow, 1).Value = Now()               ' Timestamp
        wsLog.Cells(lNextRow, 2).Value = ProcName             ' Source procedure
        wsLog.Cells(lNextRow, 3).Value = ErrNum               ' Error number
        wsLog.Cells(lNextRow, 4).Value = ErrDesc              ' Error description
        wsLog.Cells(lNextRow, 5).Value = ErrLine              ' Line number (if available)
        wsLog.Cells(lNextRow, 6).Value = Environ("USERNAME")  ' Windows user
        
        ' Auto-fit columns on first use
        If lNextRow = 2 Then
            wsLog.Columns("A:F").AutoFit
        End If
    End If
    
    ' Also write to a text log file as fallback
    WriteErrorLog ProcName, ErrNum, ErrDesc
    
    ' Show user message
    Select Case ErrNum
        Case 1004:  ' Application-defined / object-defined
            MsgBox "An unexpected error occurred in " & ProcName & "." & vbCrLf & _
                   "Please contact your system administrator.", _
                   vbCritical, APP_NAME
        Case 91:    ' Object variable or With block variable not set
            MsgBox "System component not initialized. Please restart the application.", _
                   vbExclamation, APP_NAME
        Case 13:    ' Type mismatch
            MsgBox "Data type mismatch in " & ProcName & "." & vbCrLf & _
                   "Please verify your input.", vbExclamation, APP_NAME
        Case Else:
            MsgBox "Error " & ErrNum & " in " & ProcName & ":" & vbCrLf & _
                   ErrDesc & vbCrLf & vbCrLf & _
                   "Line: " & IIf(ErrLine > 0, ErrLine, "N/A"), _
                   vbCritical, APP_NAME
    End Select
    
    On Error GoTo 0
End Sub

Public Function LogError(ByVal Source As String, ByVal Description As String, _
                         Optional ByVal Category As String = "General") As Long
    '-----------------------------------------------------------------------
    ' Lightweight error logging function. Returns the row number where the
    ' log entry was written, or 0 if logging failed.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.LogError"
    
    Dim wsLog As Worksheet
    Dim lRow As Long
    
    On Error GoTo LogError_Err
    
    Set wsLog = GetLogSheet()
    If wsLog Is Nothing Then
        LogError = 0
        Exit Function
    End If
    
    lRow = NextEmptyRow(wsLog, 1)
    
    With wsLog
        .Cells(lRow, 1).Value = Now()
        .Cells(lRow, 2).Value = Source
        .Cells(lRow, 3).Value = "ERROR"
        .Cells(lRow, 4).Value = Description
        .Cells(lRow, 5).Value = Category
        .Cells(lRow, 6).Value = Environ("USERNAME")
    End With
    
    LogError = lRow
    
LogError_Exit:
    Exit Function
    
LogError_Err:
    LogError = 0
    Resume LogError_Exit
End Function

Public Sub LogInfo(ByVal Source As String, ByVal Description As String, _
                   Optional ByVal Category As String = "Info")
    '-----------------------------------------------------------------------
    ' Log informational messages to the _Log sheet.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.LogInfo"
    
    Dim wsLog As Worksheet
    Dim lRow As Long
    
    On Error GoTo LogInfo_Err
    
    Set wsLog = GetLogSheet()
    If wsLog Is Nothing Then Exit Sub
    
    lRow = NextEmptyRow(wsLog, 1)
    
    With wsLog
        .Cells(lRow, 1).Value = Now()
        .Cells(lRow, 2).Value = Source
        .Cells(lRow, 3).Value = "INFO"
        .Cells(lRow, 4).Value = Description
        .Cells(lRow, 5).Value = Category
        .Cells(lRow, 6).Value = Environ("USERNAME")
    End With
    
LogInfo_Exit:
    Exit Sub
    
LogInfo_Err:
    Resume LogInfo_Exit
End Sub

Private Sub WriteErrorLog(ByVal ProcName As String, ByVal ErrNum As Long, _
                          ByVal ErrDesc As String)
    '-----------------------------------------------------------------------
    ' Writes error details to a text file as a fallback logging mechanism.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.WriteErrorLog"
    
    Dim sLogPath As String
    Dim iFile As Integer
    
    On Error Resume Next
    
    sLogPath = ThisWorkbook.Path & "\" & PATH_LOG_FILE
    iFile = FreeFile
    
    Open sLogPath For Append As #iFile
    Print #iFile, Now() & "|" & ProcName & "|" & ErrNum & "|" & ErrDesc & "|" & Environ("USERNAME")
    Close #iFile
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Worksheet Helpers
'------------------------------------------------------------------------------

Private Function GetLogSheet() As Worksheet
    '-----------------------------------------------------------------------
    ' Returns a reference to the _Log worksheet. Returns Nothing if the
    ' sheet does not exist (e.g., database not initialized yet).
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.GetLogSheet"
    
    On Error Resume Next
    Set GetLogSheet = ThisWorkbook.Worksheets(SHT_LOG)
    On Error GoTo 0
End Function

Private Function NextEmptyRow(ByRef ws As Worksheet, _
                              ByVal ColNum As Long) As Long
    '-----------------------------------------------------------------------
    ' Finds the next empty row in a given column of a worksheet.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.NextEmptyRow"
    
    On Error Resume Next
    NextEmptyRow = ws.Cells(ws.Rows.Count, ColNum).End(xlUp).Row + 1
    If NextEmptyRow < 2 Then NextEmptyRow = 2
    On Error GoTo 0
End Function

Public Function SheetExists(ByVal SheetName As String) As Boolean
    '-----------------------------------------------------------------------
    ' Checks whether a worksheet with the given name exists in the workbook.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.SheetExists"
    
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SheetName)
    SheetExists = (Err.Number = 0)
    Set ws = Nothing
    On Error GoTo 0
End Function

Public Function GetNamedRange(ByVal RangeName As String) As Range
    '-----------------------------------------------------------------------
    ' Safely retrieves a named range from the workbook scope.
    ' Returns Nothing if the name doesn't exist.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.GetNamedRange"
    
    On Error Resume Next
    Set GetNamedRange = ThisWorkbook.Names(RangeName).RefersToRange
    On Error GoTo 0
End Function

Public Sub SetNamedRange(ByVal RangeName As String, ByRef TargetRange As Range)
    '-----------------------------------------------------------------------
    ' Creates or updates a workbook-scoped named range pointing to a range.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.SetNamedRange"
    
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=RangeName, RefersTo:=TargetRange
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' String Utilities
'------------------------------------------------------------------------------

Public Function IsEmptyOrNull(ByVal Value As Variant) As Boolean
    '-----------------------------------------------------------------------
    ' Returns True if Value is Empty, Null, or a zero-length string.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.IsEmptyOrNull"
    
    IsEmptyOrNull = True
    
    If IsEmpty(Value) Then Exit Function
    If IsNull(Value) Then Exit Function
    If VarType(Value) = vbString Then
        If Len(Trim$(Value)) = 0 Then Exit Function
    End If
    
    IsEmptyOrNull = False
End Function

Public Function Coalesce(ByVal ParamArray Values() As Variant) As Variant
    '-----------------------------------------------------------------------
    ' Returns the first non-null, non-empty value from the parameter array.
    ' Similar to SQL COALESCE.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.Coalesce"
    
    Dim v As Variant
    
    For Each v In Values
        If Not IsEmptyOrNull(v) Then
            Coalesce = v
            Exit Function
        End If
    Next v
    
    Coalesce = vbNullString
End Function

Public Function TrimAll(ByVal Text As String) As String
    '-----------------------------------------------------------------------
    ' Trims leading/trailing spaces and collapses multiple spaces to one.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.TrimAll"
    
    Dim re As Object
    
    TrimAll = Trim$(Text)
    
    ' Collapse multiple spaces using replace approach (no RegEx dependency)
    Do While InStr(TrimAll, "  ") > 0
        TrimAll = Replace(TrimAll, "  ", " ")
    Loop
End Function

Public Function LeftOf(ByVal Text As String, ByVal Delimiter As String) As String
    '-----------------------------------------------------------------------
    ' Returns the substring to the left of the first occurrence of Delimiter.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.LeftOf"
    
    Dim lPos As Long
    
    lPos = InStr(Text, Delimiter)
    If lPos > 0 Then
        LeftOf = Left$(Text, lPos - 1)
    Else
        LeftOf = Text
    End If
End Function

Public Function RightOf(ByVal Text As String, ByVal Delimiter As String) As String
    '-----------------------------------------------------------------------
    ' Returns the substring to the right of the first occurrence of Delimiter.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.RightOf"
    
    Dim lPos As Long
    
    lPos = InStr(Text, Delimiter)
    If lPos > 0 Then
        RightOf = Mid$(Text, lPos + Len(Delimiter))
    Else
        RightOf = Text
    End If
End Function

Public Function GenerateGUID() As String
    '-----------------------------------------------------------------------
    ' Generates a pseudo-GUID string using VBA's internal type library.
    ' Format: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.GenerateGUID"
    
    Dim objTypeLib As Object
    Dim sGUID As String
    
    On Error Resume Next
    Set objTypeLib = CreateObject("Scriptlet.TypeLib")
    If Err.Number = 0 Then
        sGUID = objTypeLib.GUID
        GenerateGUID = Left$(sGUID, Len(sGUID) - 2)  ' Remove trailing CrLf
    Else
        ' Fallback using Timer and Rnd
        Randomize
        GenerateGUID = "{" & Format(Now, "yyyymmdd") & "-" & _
                       Format(Timer * 1000, "000000") & "-" & _
                       Format(Int(Rnd * 10000), "0000") & "-" & _
                       Format(Int(Rnd * 10000), "0000") & "-" & _
                       Format(Int(Rnd * 1000000), "000000") & "}"
    End If
    
    Set objTypeLib = Nothing
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Date / Time Utilities
'------------------------------------------------------------------------------

Public Function FiscalYear(ByVal InputDate As Date) As Integer
    '-----------------------------------------------------------------------
    ' Returns the fiscal year for a given date.
    ' Fiscal year starts January 1 (calendar year). Adjust as needed.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.FiscalYear"
    
    FiscalYear = Year(InputDate)
    ' If fiscal year starts July 1, uncomment below:
    ' If Month(InputDate) >= 7 Then
    '     FiscalYear = Year(InputDate) + 1
    ' End If
End Function

Public Function FormatCurrencyIDR(ByVal Amount As Double) As String
    '-----------------------------------------------------------------------
    ' Formats a numeric amount as Indonesian Rupiah currency string.
    ' Example: FormatCurrencyIDR(1500000) -> "Rp1.500.000,00"
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.FormatCurrencyIDR"
    
    Dim sFormatted As String
    
    ' Format with thousand separator and 2 decimals
    sFormatted = Format$(Amount, "#,##0." & String$(FIN_DECIMAL_PLACES, "0"))
    
    ' Replace comma with period and vice versa for IDR format
    sFormatted = Replace(sFormatted, ".", "|")   ' Temp placeholder
    sFormatted = Replace(sFormatted, ",", ".")   ' Decimal separator -> period
    sFormatted = Replace(sFormatted, "|", ",")   ' Thousand separator -> comma
    
    FormatCurrencyIDR = FIN_CURRENCY_SYMBOL & sFormatted
End Function

Public Function CalculateVAT(ByVal NetAmount As Double) As Double
    '-----------------------------------------------------------------------
    ' Calculates the VAT amount for a given net amount.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.CalculateVAT"
    
    CalculateVAT = Round(NetAmount * FIN_VAT_RATE, FIN_DECIMAL_PLACES)
End Function

Public Function CalculateTotalWithVAT(ByVal NetAmount As Double) As Double
    '-----------------------------------------------------------------------
    ' Returns the gross total (net + VAT).
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.CalculateTotalWithVAT"
    
    CalculateTotalWithVAT = Round(NetAmount + CalculateVAT(NetAmount), FIN_DECIMAL_PLACES)
End Function

'------------------------------------------------------------------------------
' Validation
'------------------------------------------------------------------------------

Public Function IsValidEmail(ByVal Email As String) As Boolean
    '-----------------------------------------------------------------------
    ' Simple email format validation.
    ' Checks for: text@text.text (at least one dot after @)
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.IsValidEmail"
    
    Dim lAtPos As Long
    Dim lDotPos As Long
    
    Email = Trim$(Email)
    If Len(Email) = 0 Then
        IsValidEmail = False
        Exit Function
    End If
    
    lAtPos = InStr(1, Email, "@")
    lDotPos = InStrRev(Email, ".")
    
    IsValidEmail = (lAtPos > 1) And (lDotPos > lAtPos + 1) And _
                   (lDotPos < Len(Email))
End Function

Public Function IsNumericRange(ByVal Value As Variant, _
                               ByVal MinVal As Double, _
                               ByVal MaxVal As Double) As Boolean
    '-----------------------------------------------------------------------
    ' Checks if Value is numeric and lies within [MinVal, MaxVal].
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.IsNumericRange"
    
    If Not IsNumeric(Value) Then
        IsNumericRange = False
        Exit Function
    End If
    
    IsNumericRange = (CDbl(Value) >= MinVal) And (CDbl(Value) <= MaxVal)
End Function

Public Function IsValidDate(ByVal DateValue As Variant) As Boolean
    '-----------------------------------------------------------------------
    ' Returns True if the variant is a valid date.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.IsValidDate"
    
    IsValidDate = IsDate(DateValue)
End Function

'------------------------------------------------------------------------------
' Data Helpers
'------------------------------------------------------------------------------

Public Function SafeConvertToString(ByVal Value As Variant) As String
    '-----------------------------------------------------------------------
    ' Converts a variant to string, returning "" for Null/Empty.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.SafeConvertToString"
    
    If IsEmptyOrNull(Value) Then
        SafeConvertToString = vbNullString
    Else
        SafeConvertToString = CStr(Value)
    End If
End Function

Public Function SafeConvertToDouble(ByVal Value As Variant) As Double
    '-----------------------------------------------------------------------
    ' Converts a variant to Double, returning 0 for Null/Empty/non-numeric.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.SafeConvertToDouble"
    
    If IsEmptyOrNull(Value) Or Not IsNumeric(Value) Then
        SafeConvertToDouble = 0#
    Else
        SafeConvertToDouble = CDbl(Value)
    End If
End Function

Public Function SafeConvertToLong(ByVal Value As Variant) As Long
    '-----------------------------------------------------------------------
    ' Converts a variant to Long, returning 0 for Null/Empty/non-numeric.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.SafeConvertToLong"
    
    If IsEmptyOrNull(Value) Or Not IsNumeric(Value) Then
        SafeConvertToLong = 0
    Else
        SafeConvertToLong = CLng(Value)
    End If
End Function

Public Function SafeConvertToDate(ByVal Value As Variant) As Date
    '-----------------------------------------------------------------------
    ' Converts a variant to Date, returning #1/1/1900# for invalid values.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.SafeConvertToDate"
    
    If IsEmptyOrNull(Value) Or Not IsDate(Value) Then
        SafeConvertToDate = DateSerial(1900, 1, 1)
    Else
        SafeConvertToDate = CDate(Value)
    End If
End Function

'------------------------------------------------------------------------------
' Array / Collection Helpers
'------------------------------------------------------------------------------

Public Function ArrayContains(ByRef Arr() As Variant, ByVal Value As Variant) As Boolean
    '-----------------------------------------------------------------------
    ' Returns True if Value is found in the 1-based array.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.ArrayContains"
    
    Dim i As Long
    
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) = Value Then
            ArrayContains = True
            Exit Function
        End If
    Next i
    
    ArrayContains = False
End Function

Public Function InStrArray(ByVal SearchString As String, _
                           ByRef Values() As String, _
                           Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Long
    '-----------------------------------------------------------------------
    ' Returns the index (0-based) of SearchString in Values(), or -1 if not found.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.InStrArray"
    
    Dim i As Long
    
    InStrArray = -1
    
    For i = LBound(Values) To UBound(Values)
        If StrComp(Values(i), SearchString, Compare) = 0 Then
            InStrArray = i
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' System Helpers
'------------------------------------------------------------------------------

Public Function IsWorkbookInitialized() As Boolean
    '-----------------------------------------------------------------------
    ' Quick check to see if the database system sheets exist.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.IsWorkbookInitialized"
    
    IsWorkbookInitialized = SheetExists(SHT_CONFIG) And _
                            SheetExists(SHT_USERS) And _
                            SheetExists(SHT_PROJECTS) And _
                            SheetExists(SHT_TRANSACTIONS) And _
                            SheetExists(SHT_BUDGET) And _
                            SheetExists(SHT_LOG) And _
                            SheetExists(SHT_AUDIT)
End Function

Public Function GetComputerName() As String
    '-----------------------------------------------------------------------
    ' Returns the local computer name.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.GetComputerName"
    
    GetComputerName = Environ("COMPUTERNAME")
End Function

Public Function GetUserName() As String
    '-----------------------------------------------------------------------
    ' Returns the current Windows user name.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.GetUserName"
    
    GetUserName = Environ("USERNAME")
End Function

Public Sub EnsureBackupFolder()
    '-----------------------------------------------------------------------
    ' Creates the backup folder if it doesn't exist.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.EnsureBackupFolder"
    
    Dim sBackupPath As String
    
    sBackupPath = ThisWorkbook.Path & "\" & PATH_BACKUP_FOLDER
    
    On Error Resume Next
    If Dir(sBackupPath, vbDirectory) = "" Then
        MkDir sBackupPath
    End If
    On Error GoTo 0
End Sub

Public Sub ApplicationWait(ByVal Seconds As Double)
    '-----------------------------------------------------------------------
    ' Pauses execution for a specified number of seconds without using
    ' DoEvents in a tight loop.
    '-----------------------------------------------------------------------
    Const FnName As String = "modUtility.ApplicationWait"
    
    Dim dEnd As Double
    
    dEnd = Timer + Seconds
    
    Do While Timer < dEnd
        DoEvents
    Loop
End Sub

