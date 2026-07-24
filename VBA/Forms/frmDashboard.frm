VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard
   Caption         =   "DUDS-PFMS - Dashboard"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   BeginProperty Font
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   8295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   10200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraNav
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   1455
      Begin VB.CommandButton cmdLogout
         BackColor       =   &H00E0E0E0&
         Caption         =   "Logout"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdReports
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reports"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton cmdJournal
         BackColor       =   &H00E0E0E0&
         Caption         =   "Journal"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdBudget
         BackColor       =   &H00E0E0E0&
         Caption         =   "Budget"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdCashOut
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cash Out"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdCashIn
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cash In"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdProjects
         BackColor       =   &H00E0E0E0&
         Caption         =   "Projects"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblNavTitle
         BackStyle       =   0  'Transparent
         Caption         =   "Navigation"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00333333&
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame fraKPIProjects
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00CCCCCC&
      ForeColor       =   &H00CCCCCC&
      Height          =   1575
      Left            =   8400
      TabIndex        =   18
      Top             =   1320
      Width           =   1455
      Begin VB.Label lblActiveProjectsValue
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001F4D7A&
         Height          =   495
         Left            =   60
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblActiveProjectsLabel
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Active Projects"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   480
         Width           =   1335
      End
      Begin VB.Shape shpProjectsTop
         BackColor       =   &H001F4D7A&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   0
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame fraKPIProfit
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00CCCCCC&
      ForeColor       =   &H00CCCCCC&
      Height          =   1575
      Left            =   5760
      TabIndex        =   14
      Top             =   1320
      Width           =   2055
      Begin VB.Label lblProfitValue
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rp0"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblProfitLabel
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Profit"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
      Begin VB.Shape shpProfitTop
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   0
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame fraKPICashOut
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00CCCCCC&
      ForeColor       =   &H00CCCCCC&
      Height          =   1575
      Left            =   3120
      TabIndex        =   10
      Top             =   1320
      Width           =   2055
      Begin VB.Label lblCashOutValue
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rp0"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblCashOutLabel
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Out"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.Shape shpCashOutTop
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   0
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame fraKPICashIn
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00CCCCCC&
      ForeColor       =   &H00CCCCCC&
      Height          =   1575
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
      Begin VB.Label lblCashInValue
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rp0"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblCashInLabel
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash In"
         BeginProperty Font
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.Shape shpCashInTop
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   0
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Shape shpTopBar
      BackColor       =   &H001F4D7A&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   10200
   End
   Begin VB.Label lblRole
      BackStyle       =   0  'Transparent
      Caption         =   "Role: Admin"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lblUser
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome, User"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lblAppTitle
      BackStyle       =   0  'Transparent
      Caption         =   "DUDS-PFMS | Dashboard"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label lblVersion
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.0"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AAAAAA&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblWelcome
      BackStyle       =   0  'Transparent
      Caption         =   "Project Financial Management System"
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CCCCCC&
      Height          =   300
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label lblStatus
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
      ForeColor       =   &H00666666&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   6000
      Width           =   9495
   End
   Begin VB.Shape shpVersionDot
      BackColor       =   &H00AAAAAA&
      BorderStyle     =   0  'Transparent
      Height          =   15
      Left            =   1080
      Top             =   135
      Width           =   15
   End
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    If Not IsSessionValid() Then
        Me.Hide
        frmLogin.Show vbModal
        Unload Me
        Exit Sub
    End If
    LoadDashboard
End Sub

Private Sub Form_Activate()
    RefreshDashboard
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then RefreshDashboard
End Sub

Public Sub LoadDashboard()
    lblUser.Caption = "Welcome, " & gCurrentUser
    lblRole.Caption = "Role: " & gCurrentRole
    lblVersion.Caption = "v" & APP_VERSION
    lblStatus.Caption = "Dashboard loaded at " & Format$(Now(), "hh:mm:ss")
    RefreshDashboard
End Sub

Public Sub RefreshDashboard()
    Dim wsTxn As Worksheet
    Dim lLastRow As Long
    Dim i As Long
    Dim dCashIn As Double
    Dim dCashOut As Double
    Dim lActive As Long
    Dim wsProj As Worksheet

    On Error Resume Next

    Set wsTxn = ThisWorkbook.Worksheets(SHT_TRANSACTIONS)
    If Not wsTxn Is Nothing Then
        lLastRow = wsTxn.Cells(wsTxn.Rows.Count, COL_TXN_TYPE + 1).End(xlUp).Row
        dCashIn = 0: dCashOut = 0
        For i = 2 To lLastRow
            Select Case SafeConvertToString(wsTxn.Cells(i, COL_TXN_TYPE + 1).Value)
                Case TXN_INCOME
                    dCashIn = dCashIn + SafeConvertToDouble(wsTxn.Cells(i, COL_TXN_TOTAL + 1).Value)
                Case TXN_EXPENSE
                    dCashOut = dCashOut + SafeConvertToDouble(wsTxn.Cells(i, COL_TXN_TOTAL + 1).Value)
            End Select
        Next i
    End If

    lblCashInValue.Caption = FormatCurrencyIDR(dCashIn)
    lblCashOutValue.Caption = FormatCurrencyIDR(dCashOut)
    lblProfitValue.Caption = FormatCurrencyIDR(dCashIn - dCashOut)
    If (dCashIn - dCashOut) >= 0 Then
        lblProfitValue.ForeColor = &H00008000
    Else
        lblProfitValue.ForeColor = &H000000C0
    End If

    Set wsProj = ThisWorkbook.Worksheets(SHT_PROJECTS)
    If Not wsProj Is Nothing Then
        lLastRow = wsProj.Cells(wsProj.Rows.Count, COL_PROJ_STATUS + 1).End(xlUp).Row
        lActive = 0
        For i = 2 To lLastRow
            If StrComp(SafeConvertToString(wsProj.Cells(i, COL_PROJ_STATUS + 1).Value), _
                       STATUS_ACTIVE, vbTextCompare) = 0 Then
                lActive = lActive + 1
            End If
        Next i
    End If
    lblActiveProjectsValue.Caption = CStr(lActive)

    lblStatus.Caption = "Last refreshed at " & Format$(Now(), "hh:mm:ss")
    On Error GoTo 0
End Sub

Private Sub cmdProjects_Click()
    MsgBox "Projects module - Coming soon.", vbInformation, APP_NAME
End Sub

Private Sub cmdCashIn_Click()
    MsgBox "Cash In module - Coming soon.", vbInformation, APP_NAME
End Sub

Private Sub cmdCashOut_Click()
    MsgBox "Cash Out module - Coming soon.", vbInformation, APP_NAME
End Sub

Private Sub cmdBudget_Click()
    MsgBox "Budget module - Coming soon.", vbInformation, APP_NAME
End Sub

Private Sub cmdJournal_Click()
    MsgBox "Journal module - Coming soon.", vbInformation, APP_NAME
End Sub

Private Sub cmdReports_Click()
    MsgBox "Reports module - Coming soon.", vbInformation, APP_NAME
End Sub

Private Sub cmdLogout_Click()
    Dim lReply As VbMsgBoxResult
    lReply = MsgBox("Are you sure you want to logout?", vbYesNo + vbQuestion, APP_NAME)
    If lReply <> vbYes Then Exit Sub
    LogoutUser
    Me.Hide
    frmLogin.Show vbModal
    Unload Me
End Sub

Public Sub Logout()
    cmdLogout_Click
End Sub

