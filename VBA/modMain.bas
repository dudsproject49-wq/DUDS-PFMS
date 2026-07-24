Attribute VB_Name = "modMain"
Option Explicit

'=============================================================================
' Module      : modMain
' Project     : DUDS-PFMS (Project Financial Management System)
' Description : Application entry point. Declares global variables and
'               handles Workbook startup/shutdown events.
'=============================================================================

'------------------------------------------------------------------------------
' Global Variables
'------------------------------------------------------------------------------

Public gAppInitialized  As Boolean        ' Whether the app has been fully initialized
Public gCurrentUser     As String         ' Username of the currently logged-in user
Public gCurrentRole     As String         ' Role of the currently logged-in user
Public gCurrentUserID   As String         ' UserID (GUID) of the currently logged-in user
Public gDBPath          As String         ' Full path to the workbook database file
Public gSessionStart    As Date           ' Timestamp when the session started
Public gLoginAttempts   As Long           ' Failed login attempt counter
Public gAppState        As String         ' Current application state string
Public gIsShuttingDown  As Boolean        ' Flag to prevent re-entrant shutdown

'------------------------------------------------------------------------------
' Constants (local to this module)
'------------------------------------------------------------------------------
Private Const MOD_NAME As String = "modMain"

'------------------------------------------------------------------------------
' Workbook Events (to be placed in ThisWorkbook code-behind)
'------------------------------------------------------------------------------
' NOTE: The actual Workbook_Open and Workbook_BeforeClose subs below are
'       meant to be called from the ThisWorkbook object's event handlers.
'       Copy the event procedure bodies below into ThisWorkbook:
'
'       Private Sub Workbook_Open()
'           StartupApp
'       End Sub
'
'       Private Sub Workbook_BeforeClose(Cancel As Boolean)
'           ShutdownApp
'       End Sub
'------------------------------------------------------------------------------

Public Sub StartupApp()
    '-----------------------------------------------------------------------
    ' Main startup routine. Called from Workbook_Open.
    ' Initializes the database, sets up the environment, and displays the
    ' login form or welcome message.
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".StartupApp"
    
    On Error GoTo StartupApp_Err
    
    ' Prevent double initialization
    If gAppInitialized Then Exit Sub
    
    ' Initialize global state
    gAppInitialized = False
    gIsShuttingDown = False
    gLoginAttempts = 0
    gCurrentUser = ""
    gCurrentRole = ""
    gCurrentUserID = ""
    gAppState = "Starting"
    gDBPath = ThisWorkbook.FullName
    gSessionStart = Now()
    
    ' Configure application environment
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .DisplayStatusBar = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = False
        .Interactive = True
        .Cursor = xlWait
    End With
    
    ' Initialize the database (create system sheets etc.)
    If Not InitializeDatabase() Then
        GoTo StartupApp_Fail
    End If
    
    ' Run startup checks
    If Not RunStartupChecks() Then
        GoTo StartupApp_Fail
    End If
    
    ' Mark as initialized
    gAppInitialized = True
    gAppState = "Ready"
    
    ' Restore application settings
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Cursor = xlDefault
        .StatusBar = MSG_WELCOME
    End With
    
    ' Log startup
    LogInfo FnName, MSG_WELCOME, "Startup"
    
    ' Display welcome message
    StatusBarMessage MSG_WELCOME, False
    
    GoTo StartupApp_Exit

StartupApp_Fail:
    gAppState = "Failed"
    MsgBox "Application failed to start. Please contact your administrator.", _
           vbCritical, APP_NAME
    GoTo StartupApp_Exit

StartupApp_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    GoTo StartupApp_Fail

StartupApp_Exit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
End Sub

Public Sub ShutdownApp()
    '-----------------------------------------------------------------------
    ' Main shutdown routine. Called from Workbook_BeforeClose.
    ' Performs cleanup, logs the shutdown, and releases resources.
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".ShutdownApp"
    
    On Error GoTo ShutdownApp_Err
    
    ' Prevent re-entrant shutdown
    If gIsShuttingDown Then Exit Sub
    gIsShuttingDown = True
    gAppState = "ShuttingDown"
    
    ' Log the shutdown
    LogInfo FnName, MSG_SHUTDOWN, "Shutdown"
    
    ' Save the workbook
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
    
    ' Reset global variables
    gAppInitialized = False
    gCurrentUser = ""
    gCurrentRole = ""
    gCurrentUserID = ""
    gLoginAttempts = 0
    gAppState = "Terminated"
    gIsShuttingDown = False
    
    GoTo ShutdownApp_Exit

ShutdownApp_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    gIsShuttingDown = False

ShutdownApp_Exit:
    Application.StatusBar = False
End Sub

'------------------------------------------------------------------------------
' Startup Checks
'------------------------------------------------------------------------------

Private Function RunStartupChecks() As Boolean
    '-----------------------------------------------------------------------
    ' Runs a series of validation checks to ensure the environment is ready.
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".RunStartupChecks"
    
    On Error GoTo RunStartupChecks_Err
    
    ' Check that the workbook is not in design mode
    If Application.VBE.ActiveVBProject.Protection = vbext_pp_locked Then
        MsgBox "VBA project is locked. Please unlock to run this application.", _
               vbExclamation, APP_NAME
        RunStartupChecks = False
        Exit Function
    End If
    
    ' Check Excel version (minimum 2013)
    If Val(Application.Version) < 15 Then
        MsgBox "This application requires Excel 2013 or later.", _
               vbExclamation, APP_NAME
        RunStartupChecks = False
        Exit Function
    End If
    
    ' Ensure backup folder exists
    EnsureBackupFolder
    
    ' Check if we have a default admin user; if not, create one
    If Not DefaultAdminExists() Then
        If Not CreateDefaultAdmin() Then
            MsgBox "Failed to create default admin account.", _
                   vbCritical, APP_NAME
            RunStartupChecks = False
            Exit Function
        End If
    End If
    
    RunStartupChecks = True
    Exit Function

RunStartupChecks_Err:
    HandleError FnName, Err.Number, Err.Description, Erl()
    RunStartupChecks = False
End Function

'------------------------------------------------------------------------------
' Session Management
'------------------------------------------------------------------------------

Public Sub SetCurrentUser(ByVal UserName As String, _
                          ByVal UserRole As String, _
                          ByVal UserID As String)
    '-----------------------------------------------------------------------
    ' Updates the global session variables after a successful login.
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".SetCurrentUser"
    
    gCurrentUser = UserName
    gCurrentRole = UserRole
    gCurrentUserID = UserID
    gSessionStart = Now()
    gLoginAttempts = 0
    gAppState = "LoggedIn"
    
    ' Update the config sheet with current user info
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Worksheets(SHT_CONFIG)
    If Not wsConfig Is Nothing Then
        ' Update the last login info
        Dim lRow As Long
        lRow = FindRecord(SHT_CONFIG, 0, "Last_Login_User")
        If lRow > 0 Then
            wsConfig.Cells(lRow, 2).Value = UserName
        End If
    End If
    On Error GoTo 0
    
    LogInfo FnName, "User '" & UserName & "' logged in as " & UserRole, "Security"
End Sub

Public Sub ClearCurrentUser()
    '-----------------------------------------------------------------------
    ' Clears the session variables on logout.
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".ClearCurrentUser"
    
    LogInfo MOD_NAME & ".ClearCurrentUser", _
            "User '" & gCurrentUser & "' logged out.", "Security"
    
    gCurrentUser = ""
    gCurrentRole = ""
    gCurrentUserID = ""
    gSessionStart = DateSerial(1900, 1, 1)
    gAppState = "Ready"
End Sub

Public Function IsSessionValid() As Boolean
    '-----------------------------------------------------------------------
    ' Checks whether the current user session is still valid (not expired).
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".IsSessionValid"
    
    If Len(gCurrentUser) = 0 Then
        IsSessionValid = False
        Exit Function
    End If
    
    ' Check session timeout
    If DateDiff("n", gSessionStart, Now()) > SEC_SESSION_TIMEOUT Then
        IsSessionValid = False
        MsgBox MSG_SESSION_EXP, vbExclamation, APP_NAME
        Exit Function
    End If
    
    IsSessionValid = True
End Function

'------------------------------------------------------------------------------
' Default Admin User
'------------------------------------------------------------------------------

Private Function DefaultAdminExists() As Boolean
    '-----------------------------------------------------------------------
    ' Checks whether a default admin user already exists in _Users.
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".DefaultAdminExists"
    
    Dim lRow As Long
    
    lRow = FindRecord(SHT_USERS, COL_USER_NAME, "admin")
    DefaultAdminExists = (lRow > 0)
End Function

Private Function CreateDefaultAdmin() As Boolean
    '-----------------------------------------------------------------------
    ' Creates the default admin account with initial credentials.
    ' Username: admin, Password: admin123 (should be changed on first login)
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".CreateDefaultAdmin"
    
    Dim arrValues(0 To 7) As Variant
    
    arrValues(0) = GenerateGUID()                    ' UserID
    arrValues(1) = "admin"                            ' UserName
    arrValues(2) = HashPassword("admin123")           ' PasswordHash
    arrValues(3) = ROLE_ADMIN                         ' Role
    arrValues(4) = "admin@duds-pfms.local"            ' Email
    arrValues(5) = "Yes"                              ' Active
    arrValues(6) = Format$(Now(), "yyyy-mm-dd hh:mm:ss")  ' CreatedOn
    arrValues(7) = ""                                 ' LastLogin
    
    Dim lResult As Long
    lResult = InsertRecord(SHT_USERS, arrValues)
    
    CreateDefaultAdmin = (lResult > 0)
End Function

'------------------------------------------------------------------------------
' UI Helpers
'------------------------------------------------------------------------------

Public Sub StatusBarMessage(ByVal Message As String, _
                            Optional ByVal IsError As Boolean = False)
    '-----------------------------------------------------------------------
    ' Displays a message in the Excel status bar.
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".StatusBarMessage"
    
    On Error Resume Next
    If IsError Then
        Application.StatusBar = "!!! " & Message & " !!!"
    Else
        Application.StatusBar = Message
    End If
    On Error GoTo 0
End Sub

Public Sub ShowSplashMessage()
    '-----------------------------------------------------------------------
    ' Displays a simple welcome dialog at startup.
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".ShowSplashMessage"
    
    MsgBox MSG_WELCOME & vbCrLf & vbCrLf & _
           "Version: " & APP_VERSION & vbCrLf & _
           "Company: " & APP_COMPANY, _
           vbInformation, APP_FULL_NAME
End Sub

'------------------------------------------------------------------------------
' Global State Helpers
'------------------------------------------------------------------------------

Public Function IsUserLoggedIn() As Boolean
    '-----------------------------------------------------------------------
    ' Quick check whether a user is currently logged in.
    '-----------------------------------------------------------------------
    IsUserLoggedIn = (Len(gCurrentUser) > 0) And gAppInitialized
End Function

Public Function HasRole(ByVal RequiredRole As String) As Boolean
    '-----------------------------------------------------------------------
    ' Checks if the currently logged-in user has the specified role.
    ' Roles hierarchy: Admin > Manager > Finance > User > Viewer
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".HasRole"
    
    Dim arrRoleHierarchy As Variant
    Dim i As Long
    Dim lCurrentLevel As Long
    Dim lRequiredLevel As Long
    
    ' Define role hierarchy (higher index = higher privilege)
    arrRoleHierarchy = Array(ROLE_VIEWER, ROLE_USER, ROLE_FINANCE, _
                             ROLE_MANAGER, ROLE_ADMIN)
    
    ' Find levels
    lCurrentLevel = -1
    lRequiredLevel = -1
    
    For i = LBound(arrRoleHierarchy) To UBound(arrRoleHierarchy)
        If arrRoleHierarchy(i) = gCurrentRole Then lCurrentLevel = i
        If arrRoleHierarchy(i) = RequiredRole Then lRequiredLevel = i
    Next i
    
    ' If either role is not found, deny access
    If lCurrentLevel = -1 Or lRequiredLevel = -1 Then
        HasRole = False
        Exit Function
    End If
    
    ' User must have at least the required privilege level
    HasRole = (lCurrentLevel >= lRequiredLevel)
End Function

Public Sub RequireRole(ByVal RequiredRole As String)
    '-----------------------------------------------------------------------
    ' Checks role and shows access denied message if not authorized.
    ' Raises an error to stop execution if insufficient privileges.
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".RequireRole"
    
    If Not HasRole(RequiredRole) Then
        MsgBox MSG_ACCESS_DENIED, vbCritical, APP_NAME
        Err.Raise vbObjectError + 1001, FnName, MSG_ACCESS_DENIED
    End If
End Sub

'------------------------------------------------------------------------------
' Dashboard Auto-Refresh (called via Application.OnTime from frmDashboard)
'------------------------------------------------------------------------------

Public Sub AutoRefreshDashboard()
    '-----------------------------------------------------------------------
    ' Callback for Application.OnTime scheduled refresh from frmDashboard.
    ' This must be a Public Sub in a standard module for OnTime to work.
    '-----------------------------------------------------------------------
    Const FnName As String = MOD_NAME & ".AutoRefreshDashboard"
    
    On Error Resume Next
    
    ' Check if dashboard form is loaded and visible
    Dim frm As Object
    For Each frm In VBA.UserForms
        If TypeName(frm) = "frmDashboard" Then
            If frm.Visible Then
                frm.RefreshDashboard
            End If
            Exit For
        End If
    Next frm
    
    On Error GoTo 0
End Sub

