Attribute VB_Name = "modSecurity"
Option Explicit

'=============================================================================
' Module      : modSecurity
' Project     : DUDS-PFMS (Project Financial Management System)
' Description : User authentication, password hashing, session management,
'               and role-based access control (RBAC).
'=============================================================================

'------------------------------------------------------------------------------
' User Authentication
'------------------------------------------------------------------------------

Public Function AuthenticateUser(ByVal UserName As String, _
                                 ByVal Password As String) As Boolean
    '-----------------------------------------------------------------------
    ' Validates a username/password combination against the _Users table.
    ' Returns True if authentication succeeds. On success, calls
    ' SetCurrentUser to establish the session.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.AuthenticateUser"
    
    Dim lRow As Long
    Dim sStoredHash As String
    Dim sInputHash As String
    Dim sActive As String
    
    On Error GoTo AuthenticateUser_Err
    
    ' Trim inputs
    UserName = Trim$(UserName)
    Password = Trim$(Password)
    
    ' Validate input
    If Len(UserName) = 0 Or Len(Password) = 0 Then
        AuthenticateUser = False
        Exit Function
    End Function
    
    ' Find user by username
    lRow = FindRecord(SHT_USERS, COL_USER_NAME, UserName)
    
    If lRow = 0 Then
        ' User not found
        LogFailedAttempt UserName, "User not found"
        AuthenticateUser = False
        Exit Function
    End If
    
    ' Check if account is active
    Dim arrRecord As Variant
    arrRecord = GetRecord(SHT_USERS, lRow)
    
    If UBound(arrRecord) >= COL_USER_ACTIVE Then
        sActive = SafeConvertToString(arrRecord(COL_USER_ACTIVE))
        If StrComp(sActive, "Yes", vbTextCompare) <> 0 Then
            LogFailedAttempt UserName, "Account inactive"
            MsgBox "Your account has been deactivated. Please contact your administrator.", _
                   vbExclamation, APP_NAME
            AuthenticateUser = False
            Exit Function
        End If
    End If
    
    ' Hash the input password
    sInputHash = HashPassword(Password)
    
    ' Get stored hash
    If UBound(arrRecord) >= COL_USER_PASSWORD Then
        sStoredHash = SafeConvertToString(arrRecord(COL_USER_PASSWORD))
    Else
        LogFailedAttempt UserName, "Corrupt record"
        AuthenticateUser = False
        Exit Function
    End If
    
    ' Compare hashes
    If StrComp(sInputHash, sStoredHash, vbBinaryCompare) = 0 Then
        ' Authentication successful
        Dim sUserID As String
        Dim sRole As String
        
        sUserID = SafeConvertToString(arrRecord(COL_USER_ID))
        sRole = SafeConvertToString(arrRecord(COL_USER_ROLE))
        
        ' Update last login timestamp
        UpdateLastLogin lRow
        
        ' Establish session
        SetCurrentUser UserName, sRole, sUserID
        
        ' Reset login attempts
        gLoginAttempts = 0
        
        ' Log success
        LogInfo FnName, "User '" & UserName & "' authenticated successfully as " & sRole, "Security"
        
        AuthenticateUser = True
    Else
        ' Password mismatch
        gLoginAttempts = gLoginAttempts + 1
        LogFailedAttempt UserName, "Invalid password (attempt " & gLoginAttempts & ")"
        
        ' Check if account should be locked
        If gLoginAttempts >= SEC_MAX_LOGIN_ATTEMPTS Then
            LockAccount UserName
            MsgBox "Account locked due to multiple failed login attempts." & vbCrLf & _
                   "Please try again in " & SEC_LOCKOUT_MINUTES & " minutes.", _
                   vbCritical, APP_NAME
        Else
            MsgBox MSG_LOGIN_FAIL, vbExclamation, APP_NAME
        End If
        
        AuthenticateUser = False
    End If
    
    Exit Function

AuthenticateUser_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    AuthenticateUser = False
End Function

Public Function LogoutUser() As Boolean
    '-----------------------------------------------------------------------
    ' Logs out the current user and clears the session.
    ' Returns True.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.LogoutUser"
    
    On Error GoTo LogoutUser_Err
    
    If Len(gCurrentUser) > 0 Then
        LogInfo FnName, "User '" & gCurrentUser & "' logged out.", "Security"
    End If
    
    ClearCurrentUser
    LogoutUser = True
    Exit Function

LogoutUser_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    LogoutUser = False
End Function

Public Function ChangePassword(ByVal UserName As String, _
                               ByVal OldPassword As String, _
                               ByVal NewPassword As String) As Boolean
    '-----------------------------------------------------------------------
    ' Changes a user's password after verifying the old password.
    ' Validates new password strength before updating.
    ' Returns True on success.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.ChangePassword"
    
    Dim lRow As Long
    Dim sOldHash As String
    Dim sNewHash As String
    Dim arrValues(0 To 7) As Variant
    Dim arrOldRecord As Variant
    Dim i As Long
    
    On Error GoTo ChangePassword_Err
    
    ' Validate new password strength
    If Not ValidatePasswordStrength(NewPassword) Then
        MsgBox "Password must be at least " & SEC_MIN_PASSWORD_LEN & " characters long.", _
               vbExclamation, APP_NAME
        ChangePassword = False
        Exit Function
    End If
    
    ' Find the user
    lRow = FindRecord(SHT_USERS, COL_USER_NAME, UserName)
    If lRow = 0 Then
        MsgBox MSG_RECORD_NOT_FOUND, vbExclamation, APP_NAME
        ChangePassword = False
        Exit Function
    End If
    
    ' Get old record
    arrOldRecord = GetRecord(SHT_USERS, lRow)
    
    ' Verify old password
    sOldHash = HashPassword(OldPassword)
    If StrComp(SafeConvertToString(arrOldRecord(COL_USER_PASSWORD)), sOldHash, vbBinaryCompare) <> 0 Then
        MsgBox "Current password is incorrect.", vbExclamation, APP_NAME
        ChangePassword = False
        Exit Function
    End If
    
    ' Build the updated record (preserve all fields, update password)
    For i = LBound(arrOldRecord) To UBound(arrOldRecord)
        arrValues(i) = arrOldRecord(i)
    Next i
    arrValues(COL_USER_PASSWORD) = HashPassword(NewPassword)
    
    ' Update the record
    If UpdateRecord(SHT_USERS, lRow, arrValues) Then
        LogInfo FnName, "Password changed for user '" & UserName & "'", "Security"
        MsgBox "Password changed successfully.", vbInformation, APP_NAME
        ChangePassword = True
    Else
        ChangePassword = False
    End If
    
    Exit Function

ChangePassword_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    ChangePassword = False
End Function

'------------------------------------------------------------------------------
' Password Hashing
'------------------------------------------------------------------------------

Public Function HashPassword(ByVal PlainText As String) As String
    '-----------------------------------------------------------------------
    ' Generates a salted hash of the password using a simple but effective
    ' algorithm. Uses the SEC_SALT constant combined with the password,
    ' then applies multiple rounds of character code transformation.
    '
    ' NOTE: For production use with sensitive data, consider replacing this
    ' with a more robust hashing mechanism via Windows Cryptography API.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.HashPassword"
    
    Dim sSalted As String
    Dim i As Long
    Dim lHash As Long
    Dim sResult As String
    
    On Error GoTo HashPassword_Err
    
    ' Combine password with salt
    sSalted = SEC_SALT & PlainText & SEC_SALT
    
    ' Initialize hash value
    lHash = 5381
    
    ' Apply djb2 hash algorithm variant
    For i = 1 To Len(sSalted)
        lHash = ((lHash * 33) Xor AscW(Mid$(sSalted, i, 1))) And &H7FFFFFFF
    Next i
    
    ' Apply additional rounds for obfuscation
    For i = 1 To 10
        lHash = ((lHash * 33) Xor (lHash \ 997)) And &H7FFFFFFF
    Next i
    
    ' Convert to hex string for storage
    sResult = Right$("00000000" & Hex$(lHash), 8)
    
    ' Add a version prefix for future hash algorithm migration
    HashPassword = "DUDS1:" & sResult
    
    Exit Function

HashPassword_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    HashPassword = "DUDS1:00000000"
End Function

Public Function ValidatePasswordStrength(ByVal Password As String) As Boolean
    '-----------------------------------------------------------------------
    ' Validates that the password meets minimum strength requirements.
    ' Current rules:
    '   - Minimum length: SEC_MIN_PASSWORD_LEN
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.ValidatePasswordStrength"
    
    ValidatePasswordStrength = (Len(Password) >= SEC_MIN_PASSWORD_LEN)
End Function

'------------------------------------------------------------------------------
' Account Management
'------------------------------------------------------------------------------

Public Function CreateUser(ByVal UserName As String, _
                           ByVal Password As String, _
                           ByVal Role As String, _
                           Optional ByVal Email As String = "") As Boolean
    '-----------------------------------------------------------------------
    ' Creates a new user account. Requires Admin or Manager role.
    ' Returns True on success.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.CreateUser"
    
    Dim arrValues(0 To 7) As Variant
    Dim lResult As Long
    
    On Error GoTo CreateUser_Err
    
    ' Check permissions (only Admin and Manager can create users)
    RequireRole ROLE_MANAGER
    
    ' Validate inputs
    UserName = Trim$(UserName)
    If Len(UserName) = 0 Then
        MsgBox "Username cannot be empty.", vbExclamation, APP_NAME
        CreateUser = False
        Exit Function
    End If
    
    If Not ValidatePasswordStrength(Password) Then
        MsgBox "Password must be at least " & SEC_MIN_PASSWORD_LEN & " characters long.", _
               vbExclamation, APP_NAME
        CreateUser = False
        Exit Function
    End If
    
    ' Check if username already exists
    If FindRecord(SHT_USERS, COL_USER_NAME, UserName) > 0 Then
        MsgBox "Username '" & UserName & "' already exists.", vbExclamation, APP_NAME
        CreateUser = False
        Exit Function
    End If
    
    ' Validate role
    Select Case Role
        Case ROLE_ADMIN, ROLE_MANAGER, ROLE_FINANCE, ROLE_USER, ROLE_VIEWER:
            ' Valid role
        Case Else:
            MsgBox "Invalid role specified.", vbExclamation, APP_NAME
            CreateUser = False
            Exit Function
    End Select
    
    ' Validate email if provided
    If Len(Email) > 0 Then
        If Not IsValidEmail(Email) Then
            MsgBox "Invalid email format.", vbExclamation, APP_NAME
            CreateUser = False
            Exit Function
        End If
    End If
    
    ' Build record
    arrValues(0) = GenerateGUID()                          ' UserID
    arrValues(1) = UserName                                ' UserName
    arrValues(2) = HashPassword(Password)                  ' PasswordHash
    arrValues(3) = Role                                    ' Role
    arrValues(4) = Email                                   ' Email
    arrValues(5) = "Yes"                                   ' Active
    arrValues(6) = Format$(Now(), "yyyy-mm-dd hh:mm:ss")   ' CreatedOn
    arrValues(7) = ""                                      ' LastLogin
    
    ' Insert record
    lResult = InsertRecord(SHT_USERS, arrValues)
    
    If lResult > 0 Then
        LogInfo FnName, "User '" & UserName & "' created with role " & Role, "Security"
        MsgBox "User '" & UserName & "' created successfully.", vbInformation, APP_NAME
        CreateUser = True
    Else
        MsgBox MSG_SAVE_FAIL, vbCritical, APP_NAME
        CreateUser = False
    End If
    
    Exit Function

CreateUser_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    CreateUser = False
End Function

Public Function UpdateUser(ByVal UserID As String, _
                           ByVal NewRole As String, _
                           Optional ByVal NewEmail As String = "", _
                           Optional ByVal Active As String = "Yes") As Boolean
    '-----------------------------------------------------------------------
    ' Updates an existing user's role, email, and active status.
    ' Requires Admin role. Returns True on success.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.UpdateUser"
    
    Dim lRow As Long
    Dim arrRecord As Variant
    Dim arrValues(0 To 7) As Variant
    Dim i As Long
    
    On Error GoTo UpdateUser_Err
    
    ' Only Admin can update users
    RequireRole ROLE_ADMIN
    
    ' Find the user
    lRow = FindRecord(SHT_USERS, COL_USER_ID, UserID)
    If lRow = 0 Then
        MsgBox MSG_RECORD_NOT_FOUND, vbExclamation, APP_NAME
        UpdateUser = False
        Exit Function
    End If
    
    ' Get current record
    arrRecord = GetRecord(SHT_USERS, lRow)
    
    ' Validate role
    Select Case NewRole
        Case ROLE_ADMIN, ROLE_MANAGER, ROLE_FINANCE, ROLE_USER, ROLE_VIEWER:
            ' Valid role
        Case Else:
            MsgBox "Invalid role specified.", vbExclamation, APP_NAME
            UpdateUser = False
            Exit Function
    End Select
    
    ' Validate email if provided
    If Len(NewEmail) > 0 Then
        If Not IsValidEmail(NewEmail) Then
            MsgBox "Invalid email format.", vbExclamation, APP_NAME
            UpdateUser = False
            Exit Function
        End If
    End If
    
    ' Build updated record
    For i = LBound(arrRecord) To UBound(arrRecord)
        arrValues(i) = arrRecord(i)
    Next i
    arrValues(COL_USER_ROLE) = NewRole
    arrValues(COL_USER_EMAIL) = NewEmail
    arrValues(COL_USER_ACTIVE) = Active
    
    ' Update
    If UpdateRecord(SHT_USERS, lRow, arrValues) Then
        LogInfo FnName, "User '" & SafeConvertToString(arrRecord(COL_USER_NAME)) & "' updated", "Security"
        MsgBox "User updated successfully.", vbInformation, APP_NAME
        UpdateUser = True
    Else
        MsgBox MSG_SAVE_FAIL, vbCritical, APP_NAME
        UpdateUser = False
    End If
    
    Exit Function

UpdateUser_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    UpdateUser = False
End Function

Public Function DeleteUser(ByVal UserID As String) As Boolean
    '-----------------------------------------------------------------------
    ' Deletes a user account. Requires Admin role.
    ' Returns True on success.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.DeleteUser"
    
    Dim lRow As Long
    Dim arrRecord As Variant
    Dim sUserName As String
    
    On Error GoTo DeleteUser_Err
    
    ' Only Admin can delete
    RequireRole ROLE_ADMIN
    
    ' Find the user
    lRow = FindRecord(SHT_USERS, COL_USER_ID, UserID)
    If lRow = 0 Then
        MsgBox MSG_RECORD_NOT_FOUND, vbExclamation, APP_NAME
        DeleteUser = False
        Exit Function
    End If
    
    ' Get user name for logging
    arrRecord = GetRecord(SHT_USERS, lRow)
    sUserName = SafeConvertToString(arrRecord(COL_USER_NAME))
    
    ' Confirm deletion
    If MsgBox("Are you sure you want to delete user '" & sUserName & "'?", _
              vbYesNo + vbQuestion, APP_NAME) <> vbYes Then
        DeleteUser = False
        Exit Function
    End If
    
    ' Cannot delete yourself
    If StrComp(sUserName, gCurrentUser, vbTextCompare) = 0 Then
        MsgBox "You cannot delete your own account.", vbExclamation, APP_NAME
        DeleteUser = False
        Exit Function
    End If
    
    ' Delete record
    If DeleteRecord(SHT_USERS, lRow) Then
        LogInfo FnName, "User '" & sUserName & "' deleted", "Security"
        MsgBox "User deleted successfully.", vbInformation, APP_NAME
        DeleteUser = True
    Else
        MsgBox "Failed to delete user.", vbCritical, APP_NAME
        DeleteUser = False
    End If
    
    Exit Function

DeleteUser_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    DeleteUser = False
End Function

'------------------------------------------------------------------------------
' Internal Helpers
'------------------------------------------------------------------------------

Private Sub LogFailedAttempt(ByVal UserName As String, _
                             ByVal Reason As String)
    '-----------------------------------------------------------------------
    ' Logs a failed authentication attempt to the _Log sheet.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.LogFailedAttempt"
    
    LogInfo "modSecurity.AuthenticateUser", _
            "Failed login for '" & UserName & "': " & Reason, _
            "Security"
End Sub

Private Sub LockAccount(ByVal UserName As String)
    '-----------------------------------------------------------------------
    ' Locks a user account by setting Active to "No" after too many
    ' failed login attempts.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.LockAccount"
    
    Dim lRow As Long
    Dim arrRecord As Variant
    Dim arrValues(0 To 7) As Variant
    Dim i As Long
    
    On Error Resume Next
    
    lRow = FindRecord(SHT_USERS, COL_USER_NAME, UserName)
    If lRow = 0 Then Exit Sub
    
    arrRecord = GetRecord(SHT_USERS, lRow)
    If UBound(arrRecord) < COL_USER_ACTIVE Then Exit Sub
    
    ' Build updated record
    For i = LBound(arrRecord) To UBound(arrRecord)
        arrValues(i) = arrRecord(i)
    Next i
    arrValues(COL_USER_ACTIVE) = "No"
    
    UpdateRecord SHT_USERS, lRow, arrValues
    LogInfo "modSecurity.LockAccount", _
            "Account '" & UserName & "' locked due to excessive failed logins", _
            "Security"
    
    On Error GoTo 0
End Sub

Private Sub UpdateLastLogin(ByVal RowNum As Long)
    '-----------------------------------------------------------------------
    ' Updates the LastLogin timestamp for a user record.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.UpdateLastLogin"
    
    Dim wsUsers As Worksheet
    
    On Error Resume Next
    
    Set wsUsers = ThisWorkbook.Worksheets(SHT_USERS)
    If wsUsers Is Nothing Then Exit Sub
    
    wsUsers.Cells(RowNum, COL_USER_LASTLOGIN + 1).Value = _
        Format$(Now(), "yyyy-mm-dd hh:mm:ss")
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Security Diagnostics
'------------------------------------------------------------------------------

Public Function GetCurrentUserInfo() As String
    '-----------------------------------------------------------------------
    ' Returns a formatted string with current user session information.
    ' Useful for display in dashboards or About forms.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.GetCurrentUserInfo"
    
    Dim sResult As String
    
    If IsUserLoggedIn() Then
        sResult = "User: " & gCurrentUser & vbCrLf & _
                  "Role:  " & gCurrentRole & vbCrLf & _
                  "Session started: " & Format$(gSessionStart, "yyyy-mm-dd hh:mm:ss") & vbCrLf & _
                  "Session expires: " & Format$(DateAdd("n", SEC_SESSION_TIMEOUT, gSessionStart), "hh:mm:ss")
    Else
        sResult = "Not logged in."
    End If
    
    GetCurrentUserInfo = sResult
End Function

Public Function IsAccountLocked(ByVal UserName As String) As Boolean
    '-----------------------------------------------------------------------
    ' Checks whether a user account is currently locked (inactive).
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.IsAccountLocked"
    
    Dim lRow As Long
    Dim arrRecord As Variant
    
    On Error GoTo IsAccountLocked_Err
    
    lRow = FindRecord(SHT_USERS, COL_USER_NAME, UserName)
    If lRow = 0 Then
        IsAccountLocked = False
        Exit Function
    End If
    
    arrRecord = GetRecord(SHT_USERS, lRow)
    
    If UBound(arrRecord) >= COL_USER_ACTIVE Then
        IsAccountLocked = (StrComp(SafeConvertToString(arrRecord(COL_USER_ACTIVE)), "Yes", vbTextCompare) <> 0)
    Else
        IsAccountLocked = False
    End If
    
    Exit Function

IsAccountLocked_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    IsAccountLocked = False
End Function

Public Function ListUsers() As String()
    '-----------------------------------------------------------------------
    ' Returns an array of all usernames (for UI dropdowns, etc.).
    ' Requires Manager role or higher.
    '-----------------------------------------------------------------------
    Const FnName As String = "modSecurity.ListUsers"
    
    Dim arrAllRecords As Variant
    Dim arrNames() As String
    Dim i As Long
    Dim lCount As Long
    
    On Error GoTo ListUsers_Err
    
    ' Check permission
    If Not HasRole(ROLE_MANAGER) Then
        ListUsers = Array()
        Exit Function
    End If
    
    arrAllRecords = GetAllRecords(SHT_USERS)
    
    If UBound(arrAllRecords) < LBound(arrAllRecords) Then
        ListUsers = Array()
        Exit Function
    End If
    
    ' Extract usernames
    ReDim arrNames(LBound(arrAllRecords) To UBound(arrAllRecords))
    lCount = 0
    
    For i = LBound(arrAllRecords) To UBound(arrAllRecords)
        If IsArray(arrAllRecords(i)) Then
            arrNames(lCount) = SafeConvertToString(arrAllRecords(i)(COL_USER_NAME))
            lCount = lCount + 1
        End If
    Next i
    
    If lCount > 0 Then
        ReDim Preserve arrNames(0 To lCount - 1)
    End If
    
    ListUsers = arrNames
    Exit Function

ListUsers_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    ListUsers = Array()
End Function

