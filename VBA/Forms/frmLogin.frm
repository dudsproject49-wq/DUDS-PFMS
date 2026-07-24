VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin
   Caption         =   "DUDS-PFMS - Login"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   BeginProperty Font
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   4095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword
      Appearance      =   0  'Flat
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox txtUsername
      Appearance      =   0  'Flat
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CheckBox chkRemember
      BackColor       =   &H00F0F0F0&
      Caption         =   "Remember Me"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin
      BackColor       =   &H001F4D7A&
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblStatus
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ""
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   1140
      Width           =   3975
   End
   Begin VB.Label lblPassword
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00333333&
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1950
      Width           =   1455
   End
   Begin VB.Label lblUsername
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00333333&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label lblTitle
      BackStyle       =   0  'Transparent
      Caption         =   "Project Financial Management System"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   10.5
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label lblAppName
      BackStyle       =   0  'Transparent
      Caption         =   "DUDS-PFMS"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Width           =   3135
   End
   Begin VB.Shape shpTopBar
      BackColor       =   &H001F4D7A&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   4680
   End
   Begin VB.Shape shpBottomLine
      BorderColor     =   &H00CCCCCC&
      Height          =   15
      Left            =   0
      Top             =   3480
      Width           =   4680
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Me.Icon = ThisWorkbook.Application.Icon
    On Error GoTo 0
    gLoginAttempts = 0
    ClearCurrentUser
    LoadSavedCredentials
End Sub

Private Sub Form_Activate()
    If Len(txtUsername.Text) = 0 Then
        txtUsername.SetFocus
    Else
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdExit_Click
End Sub

Private Sub txtUsername_GotFocus()
    txtUsername.SelStart = 0
    txtUsername.SelLength = Len(txtUsername.Text)
    lblStatus.Caption = ""
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
    lblStatus.Caption = ""
End Sub

Private Sub cmdLogin_Click()
    Dim sUsername As String
    Dim sPassword As String
    
    cmdLogin.Enabled = False
    cmdLogin.Caption = "Signing in..."
    DoEvents
    
    sUsername = Trim$(txtUsername.Text)
    sPassword = txtPassword.Text
    
    If Len(sUsername) = 0 Then
        lblStatus.Caption = "Please enter your username."
        cmdLogin.Enabled = True
        cmdLogin.Caption = "Login"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    If Len(sPassword) = 0 Then
        lblStatus.Caption = "Please enter your password."
        cmdLogin.Enabled = True
        cmdLogin.Caption = "Login"
        txtPassword.SetFocus
        Exit Sub
    End If
    
    If AuthenticateUser(sUsername, sPassword) Then
        SaveCredentials
        Me.Hide
        frmDashboard.Show vbModal
        Unload Me
    Else
        lblStatus.Caption = MSG_LOGIN_FAIL
        txtPassword.Text = ""
        txtPassword.SetFocus
    End If
    
    cmdLogin.Enabled = True
    cmdLogin.Caption = "Login"
End Sub

Private Sub cmdExit_Click()
    Dim lReply As VbMsgBoxResult
    If Len(txtUsername.Text) > 0 Or Len(txtPassword.Text) > 0 Then
        lReply = MsgBox("Are you sure you want to exit?", _
                        vbYesNo + vbQuestion, APP_NAME)
        If lReply <> vbYes Then Exit Sub
    End If
    ThisWorkbook.Close SaveChanges:=False
End Sub

Private Sub SaveCredentials()
    SaveSetting APP_NAME, "Login", "RememberMe", CStr(chkRemember.Value)
    If chkRemember.Value = vbChecked Then
        SaveSetting APP_NAME, "Login", "LastUsername", txtUsername.Text
    Else
        SaveSetting APP_NAME, "Login", "LastUsername", ""
    End If
End Sub

Private Sub LoadSavedCredentials()
    Dim sRemember As String
    Dim sSavedUser As String
    sRemember = GetSetting(APP_NAME, "Login", "RememberMe", "0")
    If sRemember = "1" Then
        chkRemember.Value = vbChecked
        sSavedUser = GetSetting(APP_NAME, "Login", "LastUsername", "")
        txtUsername.Text = sSavedUser
        If Len(sSavedUser) > 0 Then txtPassword.SetFocus
    End If
End Sub

