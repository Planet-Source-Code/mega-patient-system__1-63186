VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin prjPatientSystem.ucGradContainer fraBackGround 
      Height          =   2385
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   4207
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Authorization"
      CaptionAlignment=   2
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderIcon      =   "frmLogin.frx":2CFA
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   0
         Top             =   540
         Width           =   1725
      End
      Begin VB.CheckBox chkMask 
         Caption         =   "Unmask Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4050
         TabIndex        =   4
         Top             =   1860
         WhatsThisHelpID =   3055
         Width           =   1140
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2850
         TabIndex        =   3
         Top             =   1860
         WhatsThisHelpID =   3056
         Width           =   1140
      End
      Begin VB.TextBox txtUserID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   870
         Width           =   3165
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   3165
      End
      Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1740
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   53
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "System Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   930
         Width           =   1905
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   1230
         Width           =   1905
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'-- Local Variables
Private m_lCtr      As Long
Private m_lCanceled As Long
Dim m_lMod          As Long     '-- Load/Free Library.
'

Private Sub chkMask_Click()
    If chkMask.Value = 1 Then
        txtPassword.PasswordChar = vbNullString
    Else
        txtPassword.PasswordChar = "*"
    End If
End Sub

Private Sub cmdCancel_Click()
    m_lCanceled = 1
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
    Dim cn  As ADODB.Connection
    Dim rs  As ADODB.Recordset
    
On Error GoTo errHandler
    
    '-- System Date.
    If Not isValidData(txtDate, "System Date", 2) Then Exit Sub
    
    '-- User ID.
    If Not isValidData(txtUserID, "User ID") Then Exit Sub
    
    '-- Password.
'    If Trim$(txtPassword) = vbNullString Then
'        MsgBox "Kindly enter password.", vbExclamation
'        txtPassword.SetFocus
'        Exit Sub
'    End If
    
    Set cn = New ADODB.Connection
    cn.Open g_sConnectionString
    
    Set rs = New ADODB.Recordset
    g_sSQL = "SELECT * FROM tblUser WHERE UserID = '" & txtUserID & "' AND Password = '" & txtPassword & "'"
    rs.Open g_sSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
    
        '-- Set global user info.
        With g_tUser
            .ID = rs.Fields("ID")
            .Level = rs.Fields("Level")
            .UserID = rs.Fields("UserID")
            .Name = rs.Fields("Name")
            .Password = rs.Fields("Password")
        End With
        
        '-- Set global date.
        g_sSysDate = txtDate
        
        With frmMain
            .cmdUser.Caption = "Log-Off as " & g_tUser.UserID
            .cmdUser.ToolTipText = "Log-Off as " & g_tUser.UserID
            .cmdDate.Caption = txtDate
            .Show
        End With
        
        '-- Update ini file.
        Call SaveINI(INI_NAME, "Login", "UserID", txtUserID)
        
        Call Unload(Me)
        
    Else
        
        '-- Increment login counter.
        m_lCtr = m_lCtr + 1
        MsgBox "Check user name and password." & vbCr & vbCr & "Attempts: " & m_lCtr, vbExclamation
        If m_lCtr >= 3 Then '-- Reached maximum attempts.
            Call Unload(Me)
            m_lCanceled = 1 '-- Cancel Login.
        Else
            txtUserID.SetFocus
        End If
        
    End If
    
errHandler:
    Set rs = Nothing
    Set cn = Nothing
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmLogin", "cmdOk_Click", True)
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errHandler

    '-- (*) KBID 309366 (http://support.microsoft.com/default.aspx?scid=kb;en-us;309366)
    m_lMod = LoadLibrary("shell32.dll")
    Call InitCommonControls
    
    Call GetGradientColor(Me.hWnd)
    
    '-- Set Login Counter and Cancel Flag.
    m_lCtr = 0
    m_lCanceled = 0

    '-- DB Connection.
    Set g_cn = New ADODB.Connection
    g_cn.Open g_sConnectionString
        
    Call CenterForm(Me)
    
    Call initFrame(fraBackground)
    
    '-- Default.
    txtDate = Format(Date, FMT_DT)
    If g_tSetting.LastLog = 1 Then
        txtUserID = GetINI(INI_NAME, "Login", "UserID", vbNullString)
    End If
    
'-- (*) From vbAccelerator
'--     http://www.vbaccelerator.com/home/VB/Code/Libraries/XP_Visual_Styles/Preventing_Crashes_at_Shutdown/article.asp
errHandler:
    If Err.Number = 364 Then '-- Object Unloaded.
        Err.Clear
        Call Unload(Me)
    ElseIf Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmLogin", "Form_Load", True)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FreeLibrary(m_lMod)
    
    '-- Reset LogOff flag.
    If g_lLogOff = 1 Then g_lLogOff = 0
    
    If m_lCanceled = 1 Then End
End Sub

Private Sub txtDate_LostFocus()
On Error GoTo errHandler

    txtDate = Format(txtDate, FMT_DT)

errHandler:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub
